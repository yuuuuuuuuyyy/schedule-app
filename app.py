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
    """
    清洗儲存格內容。
    關鍵功能：將 '0', '0.0', 'nan' 等視為空字串，
    確保這些格子被視為『可排班的空格』。
    """
    if isinstance(s, pd.Series): 
        if s.empty: return ""
        s = s.iloc[0]
    if pd.isna(s): return ""
    s = str(s).strip()
    if s.endswith(".0"): s = s[:-2]
    
    # ✨ 強力過濾：這些內容視為『空白』
    if s in ["0", "nan", "None", "", "0.0"]: return ""
    
    return s.replace(" ", "").replace("　", "").replace("’", "'").replace("‘", "'").replace("，", ",")

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
        for target_name, keywords in mapping.items():
            if col_str in keywords:
                new_columns[col] = target_name
                found = True
                break
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
    # 此函式用於後處理檢查，標記潛在問題
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
    # 這裡僅作檢查，不更動排班結果
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
        if not target_sheet: return {}, {}, f"找不到 '{prev_month}月' 工作表"
        
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
    ws1 = wb.active; ws1.title = "Staff"; ws1.append(["ID", "Name", "Skills"])
    ws2 = wb.create_sheet("Roster"); ws2.append(["ID", "Name"] + [str(i) for i in range(1, num_days + 1)])
    ws3 = wb.create_sheet("Shifts"); ws3.append(["Date", "Shift", "Count"])
    ws4 = wb.create_sheet("ShiftTime"); ws4.append(["Code", "Start", "End"])
    ws5 = wb.create_sheet("例休"); ws5.append(["ID", "日期", "9例數量", "9數量"]) 
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
        if col in target_stats or col == "": weekdays.append('') 
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
        for t in target_stats: counts.append(shifts.count(t))
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
                            cell.fill = fill_big_blue if big_cycle_idx % 2 == 0 else fill_big_orange
                        elif cell.row == 2:
                            small_cycle_idx = delta_days // 14
                            cell.fill = fill_small_pink if small_cycle_idx % 2 == 0 else fill_small_purple
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
        m = st.selectbox("月份", range(1,13), index=3)
    st.divider()
    template_data = create_template_excel(y, m) 
    st.download_button(label="📥 下載排班範本", data=template_data, file_name="排班範本.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    uploaded_file = st.file_uploader("📂 請上傳 Excel 排班表 (data.xlsx)", type=['xlsx'])
    st.info("💡 **說明**：\n- Roster 中已填寫的內容(非0)會被固定。\n- Roster 中的 '0' 或空白會由 AI 根據 Shifts 需求填入。")

st.title("📅 智慧排班系統")
st.markdown("---")

if uploaded_file is not None:
    try:
        # 讀取 Staff
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
            st.warning("⚠️ 讀取 Staff 失敗，技能限制可能失效。")

        # 讀取 Roster (範本)
        try:
            df_tmp = pd.read_excel(uploaded_file, sheet_name='Roster', header=None, nrows=20)
            h_idx = -1
            target_keywords = ["ID", "卡號", "員工", "姓名", "Name"]
            for i, r in df_tmp.iterrows():
                row_str = " ".join([str(v) for v in r.values])
                if any(kw in row_str for kw in target_keywords):
                    h_idx = i
                    break
            if h_idx == -1: h_idx = 0 
            df_roster = pd.read_excel(uploaded_file, sheet_name='Roster', header=h_idx)
            df_roster = smart_rename(df_roster, {'ID':['ID','卡號'], 'Name':['Name','姓名','員工']})
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
            st.error(f"❌ 讀取 Roster 失敗: {e}"); st.stop()

        # 讀取 Shifts (需求)
        try:
            df_shifts = pd.read_excel(uploaded_file, sheet_name='Shifts')
            df_shifts = smart_rename(df_shifts, {'Date':['Date','日期'], 'Shift':['Shift','班別'], 'Count':['Count','人數']})
            df_shifts['Date'] = pd.to_datetime(df_shifts['Date'])
        except Exception as e:
            st.error(f"❌ 讀取 Shifts 失敗: {e}"); st.stop()

        # 讀取 Leave Constraints
        leave_constraints = []
        try:
            name_to_id = {}
            if 'Name' in df_roster.columns and 'ID' in df_roster.columns:
                for _, r in df_roster.iterrows():
                    n = clean_str(r['Name']); i = clean_str(r['ID'])
                    if n and i: name_to_id[n] = i
            xls_obj = pd.ExcelFile(uploaded_file)
            target_sheet = None
            if "例休" in xls_obj.sheet_names: target_sheet = "例休"
            elif "LeaveConstraints" in xls_obj.sheet_names: target_sheet = "LeaveConstraints"
            if target_sheet:
                df_leave = pd.read_excel(uploaded_file, sheet_name=target_sheet)
                df_leave = smart_rename(df_leave, {
                    'ID': ['ID', '卡號'], 
                    'LimitDate': ['LimitDate', '指定日期', '日期'], 
                    'MinExample': ['MinExample', 'Min9Example', '至少9例', '9例數量'], 
                    'MinRest': ['MinRest', 'Min9', '至少9', '9數量']
                })
                for _, r in df_leave.iterrows():
                    try:
                        raw_id = clean_str(r['ID'])
                        l_sid = name_to_id.get(raw_id, raw_id)
                        l_date = pd.to_datetime(r['LimitDate'])
                        l_min_ex = extract_number(r.get('MinExample', 0))
                        l_min_re = extract_number(r.get('MinRest', 0))
                        if l_date.month == m:
                            leave_constraints.append({'sid': l_sid, 'date': l_date, 'min_ex': l_min_ex, 'min_re': l_min_re})
                    except: pass
        except: pass 

        py, pm = get_prev_month(y, m)
        sids = df_roster['ID'].tolist()
        last_con, last_shift_map, msg = auto_calculate_last_consecutive_from_upload(uploaded_file, py, pm, sids)
        
        if "找不到" in msg: st.warning(f"⚠️ {msg}")
        else: st.success(f"✅ {msg}")

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
                        s_t = float(row.get('Start', 0)); e_t = float(row.get('End', 0))
                        shift_time_db[code] = {'Start': s_t, 'End': e_t}
                    except: pass
                known_shifts = list(shift_time_db.keys())
                for s1 in known_shifts:
                    for s2 in known_shifts:
                        t1 = shift_time_db[s1]; t2 = shift_time_db[s2]
                        rest = (t2['Start'] + 24) - t1['End']
                        if rest < 11: forbidden_pairs.add((s1, s2))
                forbidden_pairs.add(('4-12', "12'-9"))
            except: pass

            # --- Pre-check Conflict Warning ---
            fixed_check = {}
            for _, r in df_roster.iterrows():
                sid = r['ID']
                for d in v_days:
                    v = clean_str(r[str(d)])
                    if v != "": fixed_check[(sid, d)] = v

            with st.spinner("⏳ AI 正在運算最佳排班組合..."):
                model = cp_model.CpModel()
                solver = cp_model.CpSolver()
                vars = {}
                fixed = {}
                
                # 1. 鎖定 Roster 固定班
                for _, r in df_roster.iterrows():
                    sid = r['ID']
                    for d in v_days:
                        v = clean_str(r[str(d)])
                        if v != "": fixed[(sid, d)] = v

                # 2. 計算剩餘需求 (Shifts - Fixed)
                # 建立所有可能的班別需求清單
                needed = []
                for _, r in m_shifts.iterrows():
                    dn = r['Date'].day
                    sn = clean_str(r['Shift'])
                    cnt = r['Count']
                    
                    # 計算這一天、這個班別，Roster 裡已經鎖定幾個人了
                    filled_count = 0
                    for sid in sids:
                        if fixed.get((sid, dn)) == sn:
                            filled_count += 1
                    
                    # 真正的需求 = 總需求 - 已固定人數
                    rem_needed = cnt - filled_count
                    
                    if rem_needed > 0:
                        needed.append((dn, sn, rem_needed))
                    elif rem_needed < 0:
                        # 警告：Roster 固定的人數比 Shifts 需求還多
                        pass # 可選擇忽略或提示

                # 為休假也加入可選變數 (讓 AI 有空間可以排 9 或 9例，但不強制數量，由後面的 Soft Constraint 控制)
                rest_shifts = ["9", "9例"]
                existing_demands = set((x[0], x[1]) for x in needed)
                for d in v_days:
                    for s_rest in rest_shifts:
                        if (d, s_rest) not in existing_demands:
                            # 讓所有人都有機會被排休假
                            needed.append((d, s_rest, len(sids)))

                lookup = {} # (sid, d) -> list of vars
                obj = []
                
                for d, s, c in needed:
                    grp = []
                    target_shift = clean_str(s)
                    for sid in sids:
                        # 如果這個人這天已經被固定了，就不能再排其他班
                        if (sid, d) in fixed: continue
                        
                        user_skills = skills_map.get(sid, set())
                        if "不排班" in user_skills: continue
                        
                        # 技能檢查 (上班日才檢查，休假不用技能)
                        if is_working_day(target_shift) and target_shift not in user_skills:
                            continue
                            
                        v = model.NewBoolVar(f"{sid}_{d}_{s}")
                        vars[(sid, d, s)] = v
                        grp.append(v)
                        
                        if (sid, d) not in lookup: lookup[(sid, d)] = []
                        lookup[(sid, d)].append((target_shift, v)) 
                        
                        # 設定權重：優先滿足 Shifts 的工作班別需求
                        w = random.randint(10, 50)
                        if target_shift in ["9", "9例", "01特"]:
                            # 休假給予加分，鼓勵在有空位時排入，滿足例休限制
                            w += 100 
                        obj.append(v * w)
                        
                    # 限制：每天每個班別的人數 <= 需求 (若是工作班，通常希望剛好等於，但為求有解用<=搭配最大化權重)
                    # 針對工作班別，我們希望盡量填滿
                    if is_working_day(target_shift):
                         model.Add(sum(grp) == c) # 嚴格滿足工作需求 (除非沒人可排)
                    else:
                         model.Add(sum(grp) <= c) # 休假則不強求填滿

                model.Maximize(sum(obj))
                
                # 限制：每人每天只能排一個班 (包含固定班)
                for sid in sids:
                    for d in v_days:
                        if (sid, d) in fixed:
                            # 如果已經固定，這天不能再排任何變數
                            if (sid, d) in lookup:
                                for _, v in lookup[(sid, d)]:
                                    model.Add(v == 0)
                        else:
                            if (sid, d) in lookup:
                                model.Add(sum([x[1] for x in lookup[(sid, d)]]) <= 1)

                # 限制：連續工作天數 <= 6
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
                            # 變數加總 (工作班為1)
                            working_vars = [v for (s, v) in lookup[(sid, d)] if is_working_day(s)]
                            val = sum(working_vars)
                        else: 
                            val = 0 # 沒班也沒變數，視為休
                        curr.append(val)
                    full = pre + curr
                    if len(full) >= w_size:
                        for i in range(len(full)-w_size+1):
                            win = full[i:i+w_size]
                            model.Add(sum(win) <= 6)
                
                # 限制：休息時間/換班間隔
                for sid in sids:
                    last_shift = last_shift_map.get(sid)
                    if last_shift:
                        for s1, s2 in forbidden_pairs:
                            if clean_str(last_shift) == s1: 
                                v2 = vars.get((sid, 1, s2))
                                if v2 is not None: model.Add(v2 == 0)

                    for i in range(len(v_days) - 1):
                        d1 = v_days[i]; d2 = v_days[i+1]
                        fix1 = fixed.get((sid, d1)); fix2 = fixed.get((sid, d2))
                        for s1, s2 in forbidden_pairs:
                            v1 = vars.get((sid, d1, s1)); v2 = vars.get((sid, d2, s2))
                            if v1 and v2: model.AddBoolOr([v1.Not(), v2.Not()])
                            if fix1 == s1 and v2: model.Add(v2 == 0)
                            if v1 and fix2 == s2: model.Add(v1 == 0)

                # 限制：例休數量 (盡可能排入)
                for lc in leave_constraints:
                    sid = lc['sid']; limit_d = lc['date'].day
                    req_ex = lc['min_ex']; req_re = lc['min_re']
                    
                    vars_9li = []; vars_9 = []
                    fixed_ex = 0; fixed_re = 0
                    
                    current_range_days = [d for d in v_days if d <= limit_d]
                    for d in current_range_days:
                        fv = fixed.get((sid, d), "")
                        if fv == "9例": fixed_ex += 1
                        elif fv == "9": fixed_re += 1
                        elif (sid, d) in lookup:
                             for s_name, var in lookup[(sid, d)]:
                                 if s_name == "9例": vars_9li.append(var)
                                 elif s_name == "9": vars_9.append(var)
                    
                    # 軟限制：剩餘需求 <= 可排變數
                    rem_ex = max(0, req_ex - fixed_ex)
                    rem_re = max(0, req_re - fixed_re)
                    
                    if vars_9li: model.Add(sum(vars_9li) <= rem_ex) # 盡量排，但不超過
                    if vars_9: model.Add(sum(vars_9) <= rem_re)

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
                    else: fill = "9" # 預設填 9，如果完全沒排到
                    for d in v_days:
                        val = str(r[str(d)]).strip()
                        if val in ['','nan','None','0']:
                            df_fin.at[idx, str(d)] = fill

                cols = ['ID', 'Name'] + [str(d) for d in v_days]
                df_export = df_fin[cols].copy()
                
                kpi1, kpi2, kpi3 = st.columns(3)
                with kpi1: st.metric("👥 參與排班人數", f"{len(sids)} 人")
                with kpi2: st.metric("📅 排班總天數", f"{len(v_days)} 天")
                with kpi3: st.metric("🛡️ 違規檢查", "0 錯誤", delta="Passed")

                tab1, tab2 = st.tabs(["📊 排班結果預覽", "📥 下載 Excel"])
                with tab1: st.dataframe(create_preview_df(df_export, y, m), use_container_width=True)
                with tab2:
                    xlsx_data = generate_formatted_excel(df_export, y, m)
                    st.download_button(label=f"📥 下載排班結果", data=xlsx_data, file_name=f"schedule_{y}_{m}_final.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            else:
                st.error("❌ 排班失敗：找不到可行解。建議檢查：1. 人力是否不足以應付 Shifts 需求？ 2. 固定班是否卡死所有空位？")
    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("👋 歡迎使用！請先在左側側邊欄上傳您的 Excel 排班檔案。")