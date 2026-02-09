import streamlit as st
import pandas as pd
import io
import random
from datetime import datetime, timedelta

# --- 1. 環境檢查 ---
try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

st.set_page_config(page_title="變形工時排班系統 (技能嚴格版)", layout="wide")
st.title("📅 智慧排班系統 (技能限制 + 嚴格法規修正)")

if not ORTOOLS_AVAILABLE:
    st.error("❌ 嚴重錯誤：排班引擎 (ortools) 未安裝！")
    st.stop()

# ==========================================
# 2. 核心邏輯定義
# ==========================================

BASE_DATE = datetime(2025, 12, 21)

def clean_str(s):
    if isinstance(s, pd.Series): 
        if s.empty: return ""
        s = s.iloc[0]
    if pd.isna(s): return ""
    s = str(s).strip()
    if s.endswith(".0"): s = s[:-2]
    if s in ["0", "nan", "None", ""]: return ""
    # 統一全形符號與空格
    return s.replace(" ", "").replace("　", "").replace("’", "'").replace("‘", "'").replace("，", ",")

def parse_skills(skill_str):
    """將技能字串解析為 set，並進行標準化處理"""
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
    new_columns = {}
    for col in df.columns:
        col_str = str(col)
        for target_name, keywords in mapping.items():
            for kw in keywords:
                if kw in col_str:
                    new_columns[col] = target_name
                    break
    if new_columns:
        df = df.rename(columns=new_columns)
    return df

# --- 班別屬性判斷 ---

def is_mandatory_off(shift_name):
    return str(shift_name).strip() == "9例"

def is_regular_rest(shift_name):
    return str(shift_name).strip() == "9"

def is_rest_day(shift_name):
    s = str(shift_name).strip()
    if not s: return True 
    # 這裡確保 '01', '01特' 不會被誤判為休假 (因為 startswith('0') 不是 '0')
    if s in ['休', '0', 'nan', 'None']: return True
    return s.startswith("9")

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
        
        # Rule 2: 每週一例
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

        # Rule 1/5/6: 總量控制
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
    if uploaded_file is None: return {}, "無上傳檔案"
    
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        
        target_sheet = None
        candidates = [f"{prev_month}月", f"{prev_month}", f"{prev_month:02d}"]
        
        for cand in candidates:
            if cand in sheets:
                target_sheet = cand
                break
        
        if not target_sheet:
            return {}, f"找不到 '{prev_month}月' 工作表 (無上月資料)"

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
        if not id_col: return {}, "上月工作表無 ID 欄位"
        
        df_prev[id_col] = df_prev[id_col].apply(clean_str)
        
        day_cols = []
        for c in df_prev.columns:
            try:
                if 1 <= int(float(str(c))) <= 31: day_cols.append(c)
            except: pass
        day_cols.sort(key=lambda x: int(float(str(x))))
        
        res = {}
        for sid in current_staff_ids:
            row = df_prev[df_prev[id_col] == sid]
            if row.empty: res[sid] = 0; continue
            con = 0
            for c in reversed(day_cols):
                if is_working_day(str(row.iloc[0][c])): con += 1
                else: break
            res[sid] = con
        return res, f"已銜接 '{target_sheet}' 工作表"
        
    except Exception as e:
        return {}, f"讀取上月錯誤: {e}"

def generate_formatted_excel(df, year, month):
    output = io.BytesIO()
    # 修正：改用 openpyxl 引擎，避免 xlsxwriter 報錯
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 先轉成 DataFrame
        # 這裡不直接用 writer，我們用 openpyxl 原生操作來排版
        pass 

    wb = openpyxl.Workbook()
    ws = wb.active
    
    weekday_map = {0: '一', 1: '二', 2: '三', 3: '四', 4: '五', 5: '六', 6: '日'}
    
    headers = list(df.columns)
    if 'Name' in headers: headers[headers.index('Name')] = '員工'
    
    weekdays = []
    for col in headers:
        if col == 'ID': weekdays.append('')
        elif col == '員工': weekdays.append('星期')
        else:
            try:
                d = int(col)
                dt = datetime(year, month, d)
                weekdays.append(weekday_map[dt.weekday()])
            except: weekdays.append('')
    
    ws.append(headers)
    ws.append(weekdays)
    
    for r in df.values.tolist():
        ws.append(r)
        
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
            
    # 重新存入 BytesIO
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

uploaded_file = st.file_uploader("📂 請上傳 Excel 排班表 (data.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. 讀取 Staff (含技能解析)
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

        # 2. 讀取 Roster
        try:
            df_tmp = pd.read_excel(uploaded_file, sheet_name='Roster', header=None, nrows=15)
            h_idx = -1
            for i, r in df_tmp.iterrows():
                if any("卡號" in str(v) for v in r.values): h_idx = i; break
            
            if h_idx == -1: 
                h_idx = 0

            df_roster = pd.read_excel(uploaded_file, sheet_name='Roster', header=h_idx)
            df_roster = smart_rename(df_roster, {'ID':['ID','卡號'], 'Name':['Name','姓名','員工']})
            
            if 'ID' not in df_roster.columns:
                 st.error("❌ Roster 工作表找不到 'ID' 或 '卡號' 欄位，請檢查 Excel 標題。")
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
            v_days = sorted(list(set(v_days)))
            for d in v_days: df_roster[str(d)] = df_roster[str(d)].apply(clean_str)

        except Exception as e:
            st.error(f"❌ 讀取 Roster 失敗: {e}")
            st.stop()

        # 3. 讀取 Shifts
        try:
            df_shifts = pd.read_excel(uploaded_file, sheet_name='Shifts')
            df_shifts = smart_rename(df_shifts, {'Date':['Date','日期'], 'Shift':['Shift','班別'], 'Count':['Count','人數']})
            df_shifts['Date'] = pd.to_datetime(df_shifts['Date'])
        except Exception as e:
            st.error(f"❌ 讀取 Shifts 失敗: {e}")
            st.stop()

        # UI
        years = sorted(df_shifts['Date'].dt.year.unique())
        c1, c2 = st.columns(2)
        with c1: y = st.selectbox("年份", years if years else [2026])
        with c2: m = st.selectbox("月份", range(1,13), index=3)

        py, pm = get_prev_month(y, m)
        sids = df_roster['ID'].tolist()
        last_con, msg = auto_calculate_last_consecutive_from_upload(uploaded_file, py, pm, sids)
        if "找不到" in msg: st.warning(msg)
        else: st.success(msg)

        mask = (df_shifts['Date'].dt.year == y) & (df_shifts['Date'].dt.month == m)
        m_shifts = df_shifts[mask].copy()
        m_shifts = m_shifts[m_shifts['Date'].dt.day.isin(v_days)]

        st.divider()
        st.info("💡 **規則說明：** 嚴格限制班別需在員工技能清單內。9/9例/01特 依規則自動填補。")

        if st.button("🚀 啟動變形工時排班 (含技能限制)", type="primary"):
            # ==========================================
            # 🔥 步驟 0：讀取 ShiftTime 並計算禁止組合
            # ==========================================
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
                
                if forbidden_pairs:
                    st.warning(f"🛡️ 已啟動法規防護：自動偵測並禁止 {len(forbidden_pairs)} 組休息不足的班別組合 (如 晚班接早班)。")
            except Exception as e:
                st.info("ℹ️ 未偵測到 ShiftTime 分頁，略過休息時間檢查。")

            st.write("1️⃣ AI 排班運算中...")
            
            model = cp_model.CpModel()
            solver = cp_model.CpSolver()
            vars = {}
            
            fixed = {}
            for _, r in df_roster.iterrows():
                sid = r['ID']
                for d in v_days:
                    v = r[str(d)]
                    if v != "": fixed[(sid, d)] = v

            needed = []
            for _, r in m_shifts.iterrows():
                dn = r['Date'].day
                sn = clean_str(r['Shift'])
                cnt = r['Count']
                filled = sum(1 for sid in sids if fixed.get((sid, dn)) == sn)
                rem = cnt - filled
                if rem > 0: needed.append((dn, sn, rem))

            lookup = {}
            obj = []
            for d, s, c in needed:
                grp = []
                target_shift = clean_str(s)
                
                for sid in sids:
                    # 1. 檢查固定班
                    if (sid, d) in fixed: continue
                    
                    # 2. 檢查是否不排班
                    user_skills = skills_map.get(sid, set())
                    if "不排班" in user_skills: continue
                    
                    # 3. 嚴格技能檢查
                    if is_working_day(target_shift) and target_shift not in user_skills:
                        continue

                    v = model.NewBoolVar(f"{sid}_{d}_{s}")
                    vars[(sid, d, s)] = v
                    grp.append(v)
                    if (sid, d) not in lookup: lookup[(sid, d)] = []
                    lookup[(sid, d)].append(v)
                    obj.append(v * random.randint(100, 200)) 
                
                if grp: model.Add(sum(grp) <= c)

            model.Maximize(sum(obj))

            for _, vs in lookup.items(): model.Add(sum(vs) <= 1)
            
            # ========================================================
            # 🔥 關鍵修正：嚴格限制連續上班 <= 6天 (移除了跳過邏輯)
            # ========================================================
            w_size = 7
            for sid in sids:
                prev = last_con.get(sid, 0)
                pre = [1] * prev
                curr = []
                for d in v_days:
                    fv = fixed.get((sid, d), "")
                    if fv: 
                        # 01, 01特 視為上班 (1)
                        val = 0 if is_rest_day(fv) else 1
                    elif (sid, d) in lookup: 
                        val = sum(lookup[(sid, d)])
                    else: 
                        val = 0 # 該日無排班需求
                    curr.append(val)
                
                full = pre + curr
                if len(full) >= w_size:
                    for i in range(len(full)-w_size+1):
                        win = full[i:i+w_size]
                        # ⚠️ 之前這裡有一行 'if all(...) continue' 是 bug 的元兇，已經移除！
                        # 現在強制所有 7 天的區間，上班日加總都不能超過 6 天
                        model.Add(sum(win) <= 6)
            
            # ==========================================
            # 🔥 步驟 1：將休息時間限制加入 Solver
            # ==========================================
            for sid in sids:
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

            status = solver.Solve(model)

            if status in [cp_model.OPTIMAL, cp_model.FEASIBLE]:
                st.write("2️⃣ 填補空白 & 法規嚴格修正中...")
                
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
                
                st.success("🎉 排班完成！(已修正連續上班Bug + 解決下載錯誤)")
                
                df_preview = create_preview_df(df_export, y, m)
                st.dataframe(df_preview)
                
                # 使用 openpyxl 產生下載檔案
                xlsx_data = generate_formatted_excel(df_export, y, m)
                
                fn = f"schedule_{y}_{m}_final.xlsx"
                st.download_button(f"📥 下載 Excel ({fn})", xlsx_data, fn, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("❌ 排班失敗：找不到可行解。")
                st.info("建議檢查：1. 固定班是否已連續上班超過 6 天？ 2. 需求班別是否超過員工技能供給？")

    except Exception as e:
        st.error(f"Error: {e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("👋 請先在上方上傳您的 Excel 排班檔案。")