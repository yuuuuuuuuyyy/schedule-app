import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# 設定網頁標題與寬度
st.set_page_config(page_title="智慧排班系統", layout="wide")

st.title("📅 智慧排班系統 (自動休息間隔檢查版)")
st.markdown("---")

# 上傳檔案區域
uploaded_file = st.file_uploader("📂 請上傳 Excel 排班表 (需包含 'ShiftTime' 分頁)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # ==========================================
        # 1. 讀取資料
        # ==========================================
        # 讀取主要排班表 (預設讀取第一個分頁)
        df = pd.read_excel(uploaded_file, sheet_name=0, header=1)  # 假設標題在第2行(Index 1)
        
        # 清理資料：移除全空的欄位與列
        df = df.dropna(how='all').dropna(axis=1, how='all')
        
        # 抓取員工名單 (假設 ID 欄位存在，或者直接取前兩欄當作資訊)
        # 這裡假設第2欄是員工姓名，第3欄開始是日期
        # 如果你的格式不同，請根據實際 Excel 調整
        employee_names = df.iloc[:, 1].astype(str).tolist() # 員工姓名
        date_columns = df.columns[2:] # 日期欄位 (從第3欄開始)
        
        num_employees = len(employee_names)
        num_days = len(date_columns)
        
        st.write(f"✅ 偵測到 **{num_employees}** 位員工，需排班天數 **{num_days}** 天。")

        # 收集所有出現過的班別代號 (包含預排的和空格)
        unique_shifts = set()
        for col in date_columns:
            unique_shifts.update(df[col].dropna().astype(str).unique())
            
        # 移除可能讀到的 'nan' 字串
        if 'nan' in unique_shifts:
            unique_shifts.remove('nan')
            
        # 建立班別對應表 (Map shift name to integer ID)
        # 0 保留給 "空班/未排班" (如果不希望有空班，邏輯需調整)
        shift_list = sorted(list(unique_shifts))
        shift_map = {shift: i for i, shift in enumerate(shift_list)}
        
        # 顯示偵測到的班別
        st.info(f"📋 偵測到的班別代號：{', '.join(shift_list)}")

        # ==========================================
        # 2. 建立 OR-Tools 模型
        # ==========================================
        model = cp_model.CpModel()
        shifts = {} # 變數：shifts[(員工, 天, 班別)]

        # 建立變數
        for e in range(num_employees):
            for d in range(num_days):
                for s in range(len(shift_list)):
                    shifts[(e, d, s)] = model.NewBoolVar(f'shift_e{e}_d{d}_s{s}')

        # 限制 1：每天每人只能排 1 個班 (Exactly one shift per day)
        for e in range(num_employees):
            for d in range(num_days):
                model.Add(sum(shifts[(e, d, s)] for s in range(len(shift_list))) == 1)

        # 限制 2：遵守 Excel 既有的預排班表 (Hard constraints)
        # 如果 Excel 格子裡已經有填字，就必須固定，不能改
        for e in range(num_employees):
            for d, col in enumerate(date_columns):
                val = str(df.iloc[e, d + 2]) # +2 是因為前兩欄是 ID/姓名
                if val != 'nan' and val in shift_map:
                    target_shift_idx = shift_map[val]
                    model.Add(shifts[(e, d, target_shift_idx)] == 1)

        # ==========================================
        # 🔥 限制 3：讀取 ShiftTime 並自動加入休息時間限制
        # ==========================================
        try:
            # 讀取 ShiftTime 分頁
            df_shift_time = pd.read_excel(uploaded_file, sheet_name='ShiftTime')
            
            # 建立時間查詢表
            # 格式: {'4-12': {'Start': 16, 'End': 24}, ...}
            shift_time_db = {}
            for idx, row in df_shift_time.iterrows():
                # 強制轉成字串並去除前後空白，避免 '12-9 ' 對應不到 '12-9'
                code = str(row['Code']).strip()
                try:
                    s_start = float(row['Start'])
                    s_end = float(row['End'])
                    shift_time_db[code] = {'Start': s_start, 'End': s_end}
                except:
                    continue # 略過格式錯誤的行

            # 找出所有「休息不足 11 小時」的組合
            forbidden_pairs = []
            
            # 檢查所有可能的班別配對 (Shift A -> Shift B)
            for s1_name in shift_list:
                for s2_name in shift_list:
                    # 只檢查有在時間表裡的班別
                    if s1_name in shift_time_db and s2_name in shift_time_db:
                        end_time_d1 = shift_time_db[s1_name]['End']
                        start_time_d2 = shift_time_db[s2_name]['Start']
                        
                        # 計算休息時間：(隔天開始 + 24) - 前天結束
                        rest_hours = (start_time_d2 + 24) - end_time_d1
                        
                        if rest_hours < 11:
                            forbidden_pairs.append((s1_name, s2_name))

            st.write(f"🛡️ **法規防護網啟動**：已自動封鎖 {len(forbidden_pairs)} 組休息不足的班別組合。")
            with st.expander("查看被禁止的接班組合 (點擊展開)"):
                for p in forbidden_pairs:
                    st.caption(f"❌ {p[0]} (結束 {shift_time_db[p[0]]['End']}) ➜ 接 ➜ {p[1]} (開始 {shift_time_db[p[1]]['Start']}) [休息 { (shift_time_db[p[1]]['Start']+24) - shift_time_db[p[0]]['End'] } 小時]")

            # 將限制加入模型
            for e in range(num_employees):
                for d in range(num_days - 1): # 檢查每一天跟它的「隔天」
                    for s1_name, s2_name in forbidden_pairs:
                        # 取得這兩個班別在模型中的數字 ID
                        if s1_name in shift_map and s2_name in shift_map:
                            idx1 = shift_map[s1_name]
                            idx2 = shift_map[s2_name]
                            
                            # 邏輯：(今天不是 s1) OR (明天不是 s2)
                            model.AddBoolOr([
                                shifts[(e, d, idx1)].Not(),
                                shifts[(e, d + 1, idx2)].Not()
                            ])

        except ValueError:
            st.warning("⚠️ 警告：找不到 'ShiftTime' 分頁。程式將只執行基本排班，無法檢查休息時間。")
        except Exception as ex:
            st.error(f"讀取班別時間發生錯誤: {ex}")

        # ==========================================
        # 3. 求解與輸出
        # ==========================================
        solver = cp_model.CpSolver()
        # 設定求解時間上限 (避免卡死)
        solver.parameters.max_time_in_seconds = 30.0
        
        if st.button("🚀 開始排班運算", type="primary"):
            with st.spinner("正在運算最佳排班組合... (這可能需要幾秒鐘)"):
                status = solver.Solve(model)
            
            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                st.success("🎉 排班完成！符合所有規則。")
                
                # 建立結果 DataFrame
                result_data = []
                for e in range(num_employees):
                    row = [df.iloc[e, 0], df.iloc[e, 1]] # ID, Name
                    for d in range(num_days):
                        # 找出這天被選中的班別
                        for s in range(len(shift_list)):
                            if solver.Value(shifts[(e, d, s)]) == 1:
                                row.append(shift_list[s])
                                break
                    result_data.append(row)
                
                # 加上欄位名稱
                result_df = pd.DataFrame(result_data, columns=df.columns)
                
                # 顯示結果
                st.dataframe(result_df)
                
                # 下載按鈕
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    result_df.to_excel(writer, index=False, sheet_name='Final_Schedule')
                    
                st.download_button(
                    label="📥 下載排班結果 Excel",
                    data=buffer.getvalue(),
                    file_name="排班結果.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("❌ 找不到可行解！可能是限制太嚴格，或 Excel 中的預排班別已經違反了休息規則。")
                st.info("建議檢查：是否有員工被手動排了 '晚班接早班'，導致程式無解。")

    except Exception as e:
        st.error(f"發生錯誤，請檢查 Excel 格式是否正確：{e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("👋 請先在上方上傳您的 Excel 排班檔案。")