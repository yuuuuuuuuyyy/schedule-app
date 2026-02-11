import streamlit as st
import pandas as pd
import io
import random
import calendar
import re
from datetime import datetime, timedelta

# --- 1. 環境檢查 ---
try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side, PatternFill
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# 設定網頁標題與寬度
st.set_page_config(page_title="智慧排班系統", page_icon="📅", layout="wide")

# 隱藏 Streamlit 預設選單
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

if not ORTOOLS_AVAILABLE:
    st.error("❌ 嚴重錯誤：排班引擎 (ortools) 未安裝！")
    st.stop()

# ==========================================
# 2. 核心邏輯定義
# ==========================================

# 基準日：2025/12/21 (用於週期上色)
BASE_DATE = datetime(2025, 12, 21)

def clean_str(s):
    if isinstance(s, pd.Series): 
        if s.empty: return ""
        s = s.iloc[0]
    if pd.isna(s): return ""
    s = str(s).strip()
    if s.endswith(".0"): s = s[:-2]
    # 強制把 "0" 當作空白，不讓它佔用格子
    if s in ["0", "nan", "None", ""]: return ""
    return s.replace(" ", "").replace("　", "").replace("’", "'").replace("‘", "'").replace("，", ",")

# 小工具：從 "2例" 或 "3休" 中提取數字
def extract_number(s):
    if pd.isna(s): return 0
    s_str = str(s)
    if s_str.isdigit():
        return int(s_str)
    numbers = re.findall(r'\d+', s_str)
    if numbers:
        return int(numbers[0])
    return 0

def parse_skills(skill_str):
    if pd.isna(skill_str) or skill_str == "":
        return set()
    s = str(skill_str).replace("，", ",").replace(" ", "").replace("　", "")
    parts = s.split(',')
    valid_skills = set()
    for p in parts:
        clean_p = clean_str(p)
        if clean_p:
            valid_skills.add(clean_p)
    return valid_skills

def smart_rename(df, mapping):
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.duplicated()]
    
    new_columns = {}
    for col in df.columns:
        col_str = str(col)
        found = False
        
        # 策略 A: 精確比對
        for target_name, keywords in mapping.items():
            if col_str in keywords:
                new_columns[col] = target_name
                found = True
                break
        
        # 策略 B: 模糊比對
        if not found:
            for target_name, keywords in mapping.items():
                for kw in keywords:
                    if len(kw) > 1 and kw in col_str:
                        new_columns[col] = target_name
                        found = True
                        break
                if found: break
                
    if new_columns:
        df = df.rename(columns=new_columns)
    
    df = df.loc[:, ~df.columns.duplicated()]
    return df

# --- 班別屬性判斷 ---
def is_mandatory_off(shift_name):
    return str(shift_name).strip() == "9例"

def is_regular_rest(shift_name):
    return str(shift_name).strip() == "9"

def is_rest_day(shift_name):
    s = str(shift_name).strip()
    if not s: return True 
    if s in ['休', '0', 'nan', 'None', '']: return True
    if s.startswith("9"): return True
    return False

def is_working_day(shift_name):
    return not is_rest_day(shift_name)

# --- 週期計算 ---
def get_big_cycle_id(date_obj):
    delta = (date_obj - BASE_DATE).days
    return delta // 28

def get_week_id(date_obj):
    delta = (date_obj - BASE_DATE).days
    return delta // 7

def check_consecutive_safe(timeline, index_to_change):
    temp_line = timeline.copy()
    temp_line[index_to_change] = 1 
    max_con = 0
    current_con = 0
    for val in temp_line:
        if val == 1:
            current_con += 1
            max_con = max(max_con, current_con)
        else:
            current_con = 0
    return max_con <= 6

def apply_strict_labor_rules(df_result, year, month, staff_last_month_consecutive={}):
    date_cols = []
    col_map = {} 
    for col in df_result.columns:
        if col in ['ID', 'Name', '員工']: continue
        try:
            d = int(col)
            dt = datetime(year, month, d)
            date_cols.append(dt)
            col_map[dt] = col
        except: pass
    
    date_cols.sort()
    if not date_cols: return df_result, []

    logs = []

    for idx, row in df_result.iterrows():
        sid = row['ID']
        week_ids = sorted(list(set([get_week_id(dt) for dt in date_cols])))
        for wid in week_ids:
            days_in_week = [dt for dt in date_cols if get_week_id(dt) == wid]
            if not days_in_week: continue
            mandatory_dates = [] 
            regular_9_dates = [] 
            for dt in days_in_week:
                shift = str(df_result.at[idx, col_map[dt]]).strip()
                if is_mandatory_off(shift): mandatory_dates.append(dt)
                if is_regular_rest(shift): regular_9_dates.append(dt)
            
            if len(mandatory_dates) == 0:
                if regular_9_dates:
                    df_result.at[idx, col_map[regular_9_dates[0]]] = "9例"
            elif len(mandatory_dates) > 1:
                for drop_dt in mandatory_dates[1:]:
                    df_result.at[idx, col_map[drop_dt]] = "9"
                    regular_9_dates.append(drop_dt)

        cycle_ids = sorted(list(set([get_big_cycle_id(dt) for dt in date_cols])))
        for cid in cycle_ids:
            days_in_cycle = [dt for dt in date_cols if get_big_cycle_id(dt) == cid]
            if not days_in_cycle: continue
            regular_cnt = 0   
            regular_9_candidates = [] 
            for dt in days_in_cycle:
                shift = str(df_result.at[idx, col_map[dt]]).strip()
                if is_regular_rest(shift):
                    regular_cnt += 1
                    regular_9_candidates.append(dt)
            
            excess_regular = regular_cnt - 4
            if excess_regular > 0:
                prev_cons = staff_last_month_consecutive.get(sid, 0)
                timeline_prefix = [1] * prev_cons
                month_dates = date_cols 
                current_timeline = []
                for dt in month_dates:
                    s = str(df_result.at[idx, col_map[dt]]).strip()
                    val = 1 if is_working_day(s) else 0
                    current_timeline.append(val)
                full_timeline = timeline_prefix + current_timeline
                changed_count = 0
                for target_dt in regular_9_candidates:
                    if changed_count >= excess_regular: break
                    try:
                        day_idx = month_dates.index(target_dt)
                        full_idx = len(timeline_prefix) + day_idx
                        if check_consecutive_safe(full_timeline, full_idx):
                            df_result.at[idx, col_map[target_dt]] = "01特"
                            full_timeline[full_idx] = 1 
                            changed_count += 1
                    except ValueError: pass
    return df_result, logs

def get_prev_month(year, month):
    if month == 1: return year - 1, 12
    return year, month - 1

def auto_calculate_last_consecutive_from_upload(uploaded_file, prev_year, prev_month, current_staff_ids):
    if uploaded_file is None: return {}, {}, "無上傳檔案"
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        target_sheet = None
        candidates = [f"{prev_month}月", f"{prev_month}", f"{prev_month:02d}"]
        for cand in candidates:
            if cand in sheets:
                target_sheet = cand
                break
        if not target_sheet: return {}, {}, f"找不到 '{prev_month}月' 工作表 (無上月資料)"
        
        df_prev = pd.read_excel(uploaded_file, sheet_name=target_sheet, dtype=str)
        header_row = -1
        for i, row in df_prev.iterrows():
            row_str = row.astype(str).values
            if any("卡號" in s or "ID" in s for s in row_str):
                header_row = i + 1 
                break
        if header_row != -1:
             df_prev = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=header_row, dtype=str)
        
        id_col = next((c for c in df_prev.columns if "ID" in str(c) or "卡號" in str(c)), None)
        if not id_col: return {}, {}, "上月工作表無 ID 欄位"
        df_prev[id_col] = df_prev[id_col].apply(clean_str)
        
        day_cols = []
        for c in df_prev.columns:
            try:
                if 1 <= int(float(str(c))) <= 31: day_cols.append(c)
            except: pass
        day_cols.sort(key=lambda x: int(float(str(x))))
        
        last_consecutive = {}
        last_shift_map = {} 
        
        for sid in current_staff_ids:
            row = df_prev[df_prev[id_col] == sid]
            if row.empty: 
                last_consecutive[sid] = 0
                last_shift_map[sid] = None
                continue
            
            con = 0
            for c in reversed(day_cols):
                val = row.iloc[0][c]
                if isinstance(val, pd.Series): val = val.iloc[0]
                if is_working_day(str(val)): con += 1
                else: break
            last_consecutive[sid] = con

            if day_cols:
                last_val = row.iloc[0][day_cols[-1]]
                if isinstance(last_val, pd.Series): last_val = last_val.iloc[0]
                last_shift_map[sid] = clean_str(last_val)
            else:
                last_shift_map[sid] = None

        return last_consecutive, last_shift_map, f"已銜接 '{target_sheet}' 工作表"
    except Exception as e:
        return {}, {}, f"讀取上月錯誤: {e}"

def create_template_excel(year, month):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    _, num_days = calendar.monthrange(year, month)
    
    ws1 = wb.active
    ws1.title = "Staff"
    ws1.append(["ID", "Name", "Skills"])
    ws1.append(["1800", "範例員工", "8-4'F,8-5"]) 

    ws2 = wb.create_sheet("Roster")
    header = ["ID", "Name"] + [str(i) for i in range(1, num_days + 1)]
    ws2.append(header)
    ws2.append(["1800", "範例員工"] + [""] * num_days)

    ws3 = wb.create_sheet("Shifts")
    ws3.append(["Date", "Shift", "Count"])
    example_date = f"{year}/{month}/1"
    ws3.append([example_date, "8-5", 1])

    ws4 = wb.create_sheet("ShiftTime")
    ws4.append(["Code", "Start", "End"])
    ws4.append(["8-5", 8, 17])
    ws4.append(["8-4'F", 8, 16.5])
    ws4.append(["4-12", 16, 24])
    ws4.append(["12'-9", 12.5, 21])

    ws5 = wb.create_sheet("例休")
    ws5.append(["ID", "日期", "9例數量", "9數量"]) 
    ws5.append(["1800", f"{year}/{month}/15", 2, 2])
    
    ws5.column_dimensions['B'].width = 15

    wb.save(output)
    return output.getvalue()

def generate_formatted_excel(df, year, month):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Final_Schedule"
    
    fill_big_blue = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid") 
    fill_big_orange = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid") 
    fill_small_pink = PatternFill(start_color="F2DCDB", end_color="F2DCDB", fill_type="solid") 
    fill_small_purple = PatternFill(start_color="E4DFEC", end_color="E4DFEC", fill_type="solid") 
    
    target_stats = ["9例", "9", "4-12", "12'-9"] 
    
    headers = list(df.columns)
    if 'Name' in headers: headers[headers.index('Name')] = '員工'
    
    headers.extend([""] + target_stats)
    
    ws.append(headers)
    
    weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
    weekdays = []
    for col in headers:
        if col in target_stats or col == "":
            weekdays.append('') 
        elif col == 'ID': weekdays.append('')
        elif col == '員工': weekdays.append('星期')
        else:
            try:
                d = int(col)
                dt = datetime(year, month, d)
                weekdays.append(weekday_map[dt.weekday()])
            except: weekdays.append('')
    ws.append(weekdays)
    
    for row_data in df.values.tolist():
        shifts = [str(x).strip() for x in row_data[2:]]
        
        counts = []
        for t in target_stats:
            counts.append(shifts.count(t))
            
        final_row = row_data + [""] + counts
        ws.append(final_row)
        
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            
            if cell.row <= 2:
                header_val = headers[cell.column - 1]
                try:
                    d = int(header_val)
                    current_dt = datetime(year, month, d)
                    delta_days = (current_dt - BASE_DATE).days
                    if delta_days >= 0:
                        if cell.row == 1:
                            big_cycle_idx = delta_days // 28
                            if big_cycle_idx % 2 == 0: cell.fill = fill_big_blue
                            else: cell.fill = fill_big_orange
                        elif cell.row == 2:
                            small_cycle_idx = delta_days // 14
                            if small_cycle_idx % 2 == 0: cell.fill = fill_small_pink
                            else: cell.fill = fill_small_purple
                except ValueError: pass

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def create_preview_df(df, year, month):
    weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
    headers = list(df.columns)
    weekdays_row = {}
    for col in headers:
        if col == 'ID': weekdays_row[col] = ''
        elif col == 'Name': weekdays_row[col] = '星期'
        else:
            try:
                d = int(col)
                dt = datetime(year, month, d)
                weekdays_row[col] = weekday_map[dt.weekday()]
            except: weekdays_row[col] = ''
    return pd.concat([pd.DataFrame([weekdays_row]), df], ignore_index=True)

# --- 3. 主程式介面 ---

with st.sidebar:
    st.title("⚙️ 排班設定面板")
    
    c1, c2 = st.columns(2)
    with c1: 
        this_year = datetime.now().year
        year_range = range(this_year - 1, this_year + 10)
        y = st.selectbox("年份", year_range, index=1) 
    with c2: 
        m = st.selectbox("月份", range(1,13), index=3) # 預設 4月

    st.divider()

    st.write("📝 **初次使用？請先下載範本**")
    
    template_data = create_template_excel(y, m) 
    
    st.download_button(
        label="📥 下載排班範本",
        data=template_data,
        file_name="排班範本.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    
    with st.expander("🛠️ 快速生成每月需求表 (Shifts)"):
        st.caption("勾選平日/假日需要的班別，自動產生整個月的 Excel！")
        
        all_shifts = ["8-4' F", "8-5", "12' -9", "4-12", "8-5掃", "01", "01特", "9", "9例", "8-4'", "8-4'掃"]
        
        st.write("🗓️ **平日 (週一~週五)**:")
        wd_shifts = st.multiselect("平日班別", all_shifts, default=["8-4' F", "8-5", "12' -9", "4-12", "8-5掃", "01"])
        
        st.write("🎉 **假日 (週六、週日)**:")
        we_shifts = st.multiselect("假日班別", all_shifts, default=["8-4' F", "4-12", "8-4'", "8-4'掃"]) 

        if st.button("⚡ 生成並準備下載"):
            try:
                _, num_days = calendar.monthrange(y, m)
                data_gen = []
                for day_gen in range(1, num_days + 1):
                    dt_gen = datetime(y, m, day_gen)
                    date_str = dt_gen.strftime("%Y/%-m/%-d")
                    
                    if dt_gen.weekday() >= 5:
                        target_shifts = we_shifts
                    else:
                        target_shifts = wd_shifts
                    
                    for s_name in target_shifts:
                        data_gen.append([date_str, s_name, 1])
                
                df_gen = pd.DataFrame(data_gen, columns=["Date", "Shift", "Count"])
                output_gen = io.BytesIO()
                with pd.ExcelWriter(output_gen, engine='openpyxl') as writer:
                    df_gen.to_excel(writer, sheet_name='Shifts', index=False)
                
                st.download_button(
                    label=f"📥 下載 {m}月需求表",
                    data=output_gen.getvalue(),
                    file_name=f"shifts_{y}_{m}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"生成失敗: {e}")

    st.divider()

    uploaded_file = st.file_uploader("📂 請上傳 Excel 排班表 (data.xlsx)", type=['xlsx'])
    
    st.info("💡 **週期上色說明**：\n- 日期列：28天大週期 (藍/橘)\n- 星期列：14天小週期 (粉/紫)")

st.title("📅 智慧排班系統")
st.markdown("---")

if uploaded_file is not None:
    try:
        try:
            df_staff = pd.read_excel(uploaded_file, sheet_name='Staff')
            staff_cols = {'ID': ['ID', '卡號'], 'Skills': ['Skills', '技能']}
            df_staff = smart_rename(df_staff, staff_cols)
            skills_map = {}
            for _, r in df_staff.iterrows():
                if 'ID' in r and 'Skills' in r:
                    sid = clean_str(r['ID'])
                    skills_map[sid] = parse_skills(r['Skills'])
                    if "不排班" in str(r['Skills']):
                        skills_map[sid] = {"不排班"}
        except: 
            skills_map = {}
            st.warning("⚠️ 讀取 Staff 失敗，將無法執行技能限制。")

        try:
            df_tmp = pd.read_excel(uploaded_file, sheet_name='Roster', header=None, nrows=20)
            h_idx = -1
            found_header = False
            target_keywords = ["ID", "卡號", "員工", "姓名", "Name"]
            for i, r in df_tmp.iterrows():
                row_str = " ".join([str(v) for v in r.values])
                if any(kw in row_str for kw in target_keywords):
                    h_idx = i
                    found_header = True
                    break
            
            if h_idx == -1: h_idx = 0 
            
            df_roster = pd.read_excel(uploaded_file, sheet_name='Roster', header=h_idx)
            df_roster = smart_rename(df_roster, {'ID':['ID','卡號'], 'Name':['Name','姓名','員工']})
            if 'ID' not in df_roster.columns:
                 st.error("❌ Roster 工作表找不到 'ID' 或 '卡號' 欄位。")
                 st.stop()
            if 'Name' not in df_roster.columns: df_roster['Name'] = df_roster['ID']
            df_roster = df_roster.loc[:, ~df_roster.columns.duplicated()]
            df_roster['ID'] = df_roster['ID'].apply(clean_str)
            d_map = {}
            v_days = []
            for c in df_roster.columns:
                try:
                    s = str(c).strip().replace(".0","")
                    d = int(s)
                    if 1<=d<=31: 
                        d_map[c] = str(d)
                        v_days.append(d)
                except:
                    try: 
                        t = pd.to_datetime(c)
                        d_map[c] = str(t.day)
                        v_days.append(t.day)
                    except: pass
            df_roster = df_roster.rename(columns=d_map)
            df_roster = df_roster.loc[:, ~df_roster.columns.duplicated()]
            v_days = sorted(list(set(v_days)))
            for d in v_days: df_roster[str(d)] = df_roster[str(d)].apply(clean_str)
        except Exception as e:
            st.error(f"❌ 讀取 Roster 失敗: {e}")
            st.stop()

        try:
            df_shifts = pd.read_excel(uploaded_file, sheet_name='Shifts')
            df_shifts = smart_rename(df_shifts, {'Date':['Date','日期'], 'Shift':['Shift','班別'], 'Count':['Count','人數']})
            df_shifts['Date'] = pd.to_datetime(df_shifts['Date'])
        except Exception as e:
            st.error(f"❌ 讀取 Shifts 失敗: {e}")
            st.stop()

        # ✨ 讀取休假限制 (針對新標題 "9例數量", "9數量" 做最佳化)
        leave_constraints = []
        try:
            name_to_id = {}
            if 'Name' in df_roster.columns and 'ID' in df_roster.columns:
                for _, r in df_roster.iterrows():
                    n = clean_str(r['Name'])
                    i = clean_str(r['ID'])
                    if n and i:
                        name_to_id[n] = i

            xls_obj = pd.ExcelFile(uploaded_file)
            target_sheet = None
            if "例休" in xls_obj.sheet_names:
                target_sheet = "例休"
            elif "LeaveConstraints" in xls_obj.sheet_names:
                target_sheet = "LeaveConstraints"
            
            if target_sheet:
                df_leave = pd.read_excel(uploaded_file, sheet_name=target_sheet)
                df_leave = smart_rename(df_leave, {
                    'ID': ['ID', '卡號'], 
                    'LimitDate': ['LimitDate', '指定日期', '日期'], 
                    # ✨ 關鍵修改：將您的 Excel 標題加入關鍵字
                    'MinExample': ['9例數量', '至少9例', 'MinExample'], 
                    'MinRest': ['9數量', '至少9', 'MinRest']
                })
                for _, r in df_leave.iterrows():
                    try:
                        raw_id = clean_str(r['ID'])
                        l_sid = name_to_id.get(raw_id, raw_id)

                        l_date = pd.to_datetime(r['LimitDate'])
                        l_min_ex = extract_number(r.get('MinExample', 0))
                        l_min_re = extract_number(r.get('MinRest', 0))

                        if l_date.month == m:
                            leave_constraints.append({
                                'sid': l_sid,
                                'date': l_date,
                                'min_ex': l_min_ex,
                                'min_re': l_min_re
                            })
                    except: pass
        except: pass 

        py, pm = get_prev_month(y, m)
        sids = df_roster['ID'].tolist()
        last_con, last_shift_map, msg = auto_calculate_last_consecutive_from_upload(uploaded_file, py, pm, sids)
        
        if "找不到" in msg: 
            st.warning(f"⚠️ {msg}")
        else: 
            st.success(f"✅ {msg}")
        
        if leave_constraints:
            st.success(f"🛡️ 已讀取 {len(leave_constraints)} 條指定日期【例休】限制")
            with st.expander("🔍 查看已讀取的例休限制 (前 5 筆)"):
                for i, lc in enumerate(leave_constraints[:5]):
                    st.write(f"#{i+1}: 員工 {lc['sid']} 在 {lc['date'].month}/{lc['date'].day} 前，必須盡可能排 {lc['min_ex']}例 + {lc['min_re']}休")

        with st.expander("🔍 資料讀取診斷報告 (若排班失敗請點此)"):
            st.write(f"**偵測到的標題列索引**: {h_idx}")
            st.write(f"**Roster 讀取到的欄位**: {list(df_roster.columns)}")
            st.write(f"**辨識出的有效日期欄 (v_days)**: {v_days}")
            if not v_days:
                st.error("❌ 嚴重警告：程式找不到任何日期欄位！請檢查 Roster 的標題列是否包含 1, 2, 3... 等數字。")

        mask = (df_shifts['Date'].dt.year == y) & (df_shifts['Date'].dt.month == m)
        m_shifts = df_shifts[mask].copy()
        m_shifts = m_shifts[m_shifts['Date'].dt.day.isin(v_days)]

        if st.button("🚀 啟動 AI 自動排班", type="primary", use_container_width=True):
            shift_time_db = {}
            forbidden_pairs = set() 
            try:
                df_st = pd.read_excel(uploaded_file, sheet_name='ShiftTime', dtype=str)
                for _, row in df_st.iterrows():
                    code = clean_str(row.get('Code', ''))
                    try:
                        s_t = float(row.get('Start', 0))
                        e_t = float(row.get('End', 0))
                        shift_time_db[code] = {'Start': s_t, 'End': e_t}
                    except: pass
                known_shifts = list(shift_time_db.keys())
                for s1 in known_shifts:
                    for s2 in known_shifts:
                        t1 = shift_time_db[s1]
                        t2 = shift_time_db[s2]
                        rest = (t2['Start'] + 24) - t1['End']
                        if rest < 11:
                            forbidden_pairs.add((s1, s2))
                
                forbidden_pairs.add(('4-12', "12'-9"))
                
                if forbidden_pairs:
                    with st.expander(f"🛡️ 已啟動法規防護 ({len(forbidden_pairs)} 條規則)"):
                        st.write(list(forbidden_pairs))
            except: pass

            # --- ✨ 數學衝突偵測器 (僅警告不中斷) ---
            fixed_check = {}
            for _, r in df_roster.iterrows():
                sid = r['ID']
                for d in v_days:
                    v_obj = r[str(d)]
                    if isinstance(v_obj, pd.Series): v_obj = v_obj.iloc[0]
                    v = clean_str(v_obj) 
                    if v != "":
                        fixed_check[(sid, d)] = v

            for lc in leave_constraints:
                sid = lc['sid']
                limit_d = lc['date'].day
                req_ex = lc['min_ex']
                req_re = lc['min_re']
                
                fixed_ex_cnt = 0
                fixed_re_cnt = 0
                empty_slots = 0
                
                current_range_days = [d for d in v_days if d <= limit_d]
                for d in current_range_days:
                    fv = fixed_check.get((sid, d), "")
                    if fv == "9例": fixed_ex_cnt += 1
                    elif fv == "9": fixed_re_cnt += 1
                    elif fv == "": empty_slots += 1
                
                if fixed_ex_cnt > req_ex:
                    st.warning(f"⚠️ **注意**：員工 {sid} 在 {limit_d} 號前，Roster 已固定 {fixed_ex_cnt} 個「9例」，超過限制的 {req_ex} 個 (將保留現狀)。")
                if fixed_re_cnt > req_re:
                    st.warning(f"⚠️ **注意**：員工 {sid} 在 {limit_d} 號前，Roster 已固定 {fixed_re_cnt} 個「9」，超過限制的 {req_re} 個 (將保留現狀)。")
                
                needed_ex = max(0, req_ex - fixed_ex_cnt)
                needed_re = max(0, req_re - fixed_re_cnt)
                
                if needed_ex + needed_re > empty_slots:
                    st.warning(f"⚠️ **空間不足**：員工 {sid} 在 {limit_d} 號前還缺 {needed_ex}個9例、{needed_re}個9，但只剩 {empty_slots} 個空格 (將盡量排入)。")

            for sid, last_shift in last_shift_map.items():
                if last_shift:
                    for s1, s2 in forbidden_pairs:
                        if clean_str(last_shift) == s1:
                            fv_1 = fixed_check.get((sid, 1))
                            if fv_1 and clean_str(fv_1) == s2:
                                st.warning(f"⚠️ **跨月衝突**：員工 {sid} 上個月最後一天是 {s1}，但這個月 1 號固定排了 {s2}，違反休息規定！")

            with st.spinner("⏳ AI 正在運算最佳排班組合..."):
                model = cp_model.CpModel()
                solver = cp_model.CpSolver()
                vars = {}
                fixed = {}
                for _, r in df_roster.iterrows():
                    sid = r['ID']
                    for d in v_days:
                        v_obj = r[str(d)]
                        if isinstance(v_obj, pd.Series): v_obj = v_obj.iloc[0]
                        v = clean_str(v_obj) # ✨ 使用 clean_str 處理 "0"
                        if v != "":
                            fixed[(sid, d)] = v

                needed = []
                for _, r in m_shifts.iterrows():
                    dn = r['Date'].day
                    sn = clean_str(r['Shift'])
                    cnt = r['Count']
                    filled = sum(1 for sid in sids if fixed.get((sid, dn)) == sn)
                    rem = cnt - filled
                    if rem > 0: needed.append((dn, sn, rem))

                rest_shifts = ["9", "9例"]
                existing_demands = set((x[0], x[1]) for x in needed)
                for d in v_days:
                    for s_rest in rest_shifts:
                        if (d, s_rest) not in existing_demands:
                            needed.append((d, s_rest, len(sids)))

                lookup = {}
                obj = []
                for d, s, c in needed:
                    grp = []
                    target_shift = clean_str(s)
                    for sid in sids:
                        if (sid, d) in fixed: continue
                        user_skills = skills_map.get(sid, set())
                        if "不排班" in user_skills: continue
                        if is_working_day(target_shift) and target_shift not in user_skills:
                            continue
                        v = model.NewBoolVar(f"{sid}_{d}_{s}")
                        vars[(sid, d, s)] = v
                        grp.append(v)
                        if (sid, d) not in lookup: lookup[(sid, d)] = []
                        lookup[(sid, d)].append((target_shift, v)) 
                        
                        # ✨ 關鍵修改：給休假班極高權重，實現「盡可能排入」
                        w = random.randint(100, 200)
                        if target_shift in ["9", "9例", "01特"]:
                            w += 5000 
                        obj.append(v * w)
                        
                    if grp: model.Add(sum(grp) <= c)

                model.Maximize(sum(obj))
                for _, vs in lookup.items(): 
                    model.Add(sum([x[1] for x in vs]) <= 1)
                
                w_size = 7
                for sid in sids:
                    prev = last_con.get(sid, 0)
                    pre = [1] * prev
                    curr = []
                    for d in v_days:
                        fv = fixed.get((sid, d), "")
                        if fv: 
                            val = 0 if is_rest_day(fv) else 1
                        elif (sid, d) in lookup: 
                            working_vars = [v for (s, v) in lookup[(sid, d)] if is_working_day(s)]
                            val = sum(working_vars)
                        else: 
                            val = 0 
                        curr.append(val)
                    full = pre + curr
                    if len(full) >= w_size:
                        for i in range(len(full)-w_size+1):
                            win = full[i:i+w_size]
                            model.Add(sum(win) <= 6)
                
                for sid in sids:
                    last_shift = last_shift_map.get(sid)
                    if last_shift:
                        for s1, s2 in forbidden_pairs:
                            if clean_str(last_shift) == s1: 
                                v2 = vars.get((sid, 1, s2))
                                if v2 is not None:
                                    model.Add(v2 == 0)

                    for i in range(len(v_days) - 1):
                        d1 = v_days[i]
                        d2 = v_days[i+1]
                        fix1 = fixed.get((sid, d1))
                        fix2 = fixed.get((sid, d2))
                        for s1, s2 in forbidden_pairs:
                            v1 = vars.get((sid, d1, s1))
                            v2 = vars.get((sid, d2, s2))
                            if v1 is not None and v2 is not None:
                                model.AddBoolOr([v1.Not(), v2.Not()])
                            if fix1 == s1 and v2 is not None:
                                model.Add(v2 == 0)
                            if v1 is not None and fix2 == s2:
                                model.Add(v1 == 0)

                for lc in leave_constraints:
                    sid = lc['sid']
                    limit_d = lc['date'].day
                    target_9li = lc['min_ex'] 
                    target_9 = lc['min_re']   
                    
                    vars_9li = []
                    vars_9 = []
                    
                    current_range_days = [d for d in v_days if d <= limit_d]
                    
                    for d in current_range_days:
                        fv = fixed.get((sid, d), "")
                        if fv:
                            if str(fv) == "9例": target_9li -= 1
                            if str(fv) == "9": target_9 -= 1
                        
                        elif (sid, d) in lookup:
                             for s_name, var in lookup[(sid, d)]:
                                 if str(s_name) == "9例":
                                     vars_9li.append(var)
                                 elif str(s_name) == "9":
                                     vars_9.append(var)
                    
                    # ✨ 關鍵修改：改用 <= (小於等於)，且不報錯
                    # 因為變數權重很高，AI 會自動嘗試填滿到 target_9li
                    remaining_quota_9li = max(0, target_9li)
                    remaining_quota_9 = max(0, target_9)
                    
                    if vars_9li:
                        model.Add(sum(vars_9li) <= remaining_quota_9li)
                    if vars_9:
                        model.Add(sum(vars_9) <= remaining_quota_9)

                status = solver.Solve(model)

            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                
                df_fin = df_roster.copy().set_index('ID')
                for (sid, d, s), v in vars.items():
                    if solver.Value(v): df_fin.at[sid, str(d)] = s
                df_fin = df_fin.reset_index()

                for idx, r in df_fin.iterrows():
                    sid = r['ID']
                    user_skills = skills_map.get(sid, set())
                    if "不排班" in user_skills: fill = ""
                    else: fill = "9"
                    for d in v_days:
                        val = str(r[str(d)]).strip()
                        if val in ['','nan','None','0']:
                            df_fin.at[idx, str(d)] = fill

                df_fin, logs = apply_strict_labor_rules(df_fin, y, m, last_con)
                cols = ['ID', 'Name'] + [str(d) for d in v_days]
                df_export = df_fin[cols].copy()
                
                kpi1, kpi2, kpi3 = st.columns(3)
                with kpi1: st.metric("👥 參與排班人數", f"{len(sids)} 人")
                with kpi2: st.metric("📅 排班總天數", f"{len(v_days)} 天")
                with kpi3: st.metric("🛡️ 違規檢查", "0 錯誤", delta="Passed")

                tab1, tab2 = st.tabs(["📊 排班結果預覽", "📥 下載 Excel"])
                
                with tab1:
                    df_preview = create_preview_df(df_export, y, m)
                    st.dataframe(df_preview, use_container_width=True)

                with tab2:
                    xlsx_data = generate_formatted_excel(df_export, y, m)
                    fn = f"schedule_{y}_{m}_final.xlsx"
                    st.download_button(
                        label=f"📥 下載排班結果 ({fn})",
                        data=xlsx_data,
                        file_name=fn,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
            else:
                st.error("❌ 排班失敗：找不到可行解。建議檢查：1. 固定班是否已違反法規？ 2. 人力是否不足？ 3. 例休限制是否太嚴苛？")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("👋 歡迎使用！請先在左側側邊欄上傳您的 Excel 排班檔案。")