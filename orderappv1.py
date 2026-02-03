import streamlit as st
import pandas as pd
import msoffcrypto
import io
import chardet
from datetime import datetime, timedelta

st.set_page_config(page_title="Order App 2.0 by Okross Frank", page_icon="📊", layout="wide")

def try_decrypt(file_stream, password):
    decrypted_buffer = io.BytesIO()
    try:
        file_stream.seek(0)
        office_file = msoffcrypto.OfficeFile(file_stream)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted_buffer)
        decrypted_buffer.seek(0)
        return decrypted_buffer
    except:
        file_stream.seek(0)
        return file_stream

def read_excel_comprehensive(file, use_pass=False, password=""):
    ext = file.name.split('.')[-1].lower()
    try:
        content = file
        if use_pass and ext == 'xlsx':
            content = try_decrypt(file, password)
        engine = 'openpyxl' if ext == 'xlsx' else 'xlrd'
        df = pd.read_excel(content, engine=engine)
        df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]
        return df
    except Exception as e:
        st.error(f"檔案 {file.name} 讀取失敗: {e}")
        return None

# --- 介面設定 ---
with st.sidebar:
    st.title("🛡️ 參數設定")
    shop_url = st.text_input("1. 店鋪網址", value="https://www.etmall.com.tw/ms/172448")
    platform_name = st.text_input("2. 電商平台英文名稱", value="ETMall")
    st.divider()
    use_pass = st.checkbox("3. 檔案有密碼", value=True)
    excel_pass = st.text_input("輸入密碼", value="123456", type="password")
    st.divider()
    f_old = st.checkbox("4. 排除 >350天舊單", value=True)
    st.info("💡 系統會自動排除：無快遞單號、包含『勿拍』字樣、重複、或狀態異常之訂單。")

st.header("📦 Order App v2.0 訂單格式合併轉換")

uploaded_files = st.file_uploader("請上傳 B 檔 (對帳單) 與 C 檔 (訂單清單)", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

if uploaded_files and shop_url:
    if st.button("🚀 執行分析並產生報告", type="primary"):
        b_list, c_list = [], []
        
        for f in uploaded_files:
            df = read_excel_comprehensive(f, use_pass, excel_pass)
            if df is None: continue
            cols = df.columns.tolist()
            if "渠道单号" in cols or "渠道單號" in cols:
                target_col = "渠道单号" if "渠道单号" in cols else "渠道單號"
                df["join_key"] = df[target_col].astype(str).apply(lambda x: x.split('-')[-1] if '-' in x else x)
                b_list.append(df)
            elif "客户订单号" in cols or "客戶訂單號" in cols:
                target_col = "客户订单号" if "客户订单号" in cols else "客戶訂單號"
                df["join_key"] = df[target_col].astype(str).str.strip()
                c_data = df.rename(columns={"快递单号": "快递单号", "快遞單號": "快递单号", "快递公司": "快递公司", "快遞公司": "快递公司"})
                c_list.append(c_data[["join_key", "快递单号", "快递公司"]])

        if b_list and c_list:
            df_b_all = pd.concat(b_list, ignore_index=True)
            df_c_all = pd.concat(c_list, ignore_index=True).drop_duplicates(subset=["join_key"])
            
            # 合併原始數據
            raw_merged = pd.merge(df_b_all, df_c_all, on="join_key", how="left")
            
            # --- 開始過濾與統計 ---
            total_initial_count = len(raw_merged)
            
            # 轉換金額為數值
            raw_merged['amount'] = pd.to_numeric(raw_merged.get('支付总金额', 0), errors='coerce').fillna(0)
            
            # 1. 判定排除條件
            # a. 無物流單號
            mask_no_tracking = raw_merged['快递单号'].isna() | (raw_merged['快递单号'].astype(str).str.strip() == "")
            # b. 包含勿拍
            mask_dont_buy = raw_merged['前台传入商品名称'].astype(str).str.contains("勿拍", na=False)
            # c. 重複訂單 (保留第一筆)
            mask_duplicate = raw_merged.duplicated(subset=["join_key"], keep='first')
            
            # 合併所有排除條件
            is_excluded = mask_no_tracking | mask_dont_buy | mask_duplicate
            
            # 分拆成功與排除的 DataFrame
            success_df = raw_merged[~is_excluded].copy()
            excluded_df = raw_merged[is_excluded].copy()
            
            # --- 計算統計值 ---
            success_count = len(success_df)
            success_sum = success_df['amount'].sum()
            
            excluded_count = len(excluded_df)
            excluded_sum = excluded_df['amount'].sum()
            
            # --- 顯示統計摘要 ---
            st.subheader("📊 處理結果摘要")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("成功筆數", f"{success_count} 筆")
            m2.metric("成功總金額 (USD)", f"${success_sum:,.2f}")
            m3.metric("排除筆數", f"{excluded_count} 筆", delta=f"-{excluded_count}", delta_color="inverse")
            m4.metric("排除總金額 (USD)", f"${excluded_sum:,.2f}")

            # --- 建立 D 檔 ---
            d_df = pd.DataFrame()
            d_df['订单编号'] = success_df['join_key']
            date_col = next((c for c in ["渠道订单创建时间", "渠道訂單創建時間"] if c in success_df.columns), None)
            d_df['订单日期'] = pd.to_datetime(success_df[date_col], errors='coerce').dt.strftime('%Y-%m-%d') if date_col else ""
            d_df['订单币种'] = success_df.get('支付币种', 'USD')
            d_df['订单金额'] = success_df['amount']
            d_df['商品名称'] = success_df.get('前台传入商品名称', '')
            d_df['商品数量'] = pd.to_numeric(success_df.get('商品数量', 1), errors='coerce').fillna(1)
            d_df['商品单价'] = (d_df['订单金额'] / d_df['商品数量'].replace(0, 1)).round(2)
            d_df['店铺网址'] = shop_url
            d_df['快递单号'] = success_df['快递单号']
            d_df['物流企业名称'] = success_df['快递公司']
            d_df['电商平台英文名称'] = platform_name
            
            # 預覽與下載
            st.divider()
            st.subheader("📝 預覽成功訂單 (前 5 筆)")
            st.dataframe(d_df.head())

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                header = d_df.columns.tolist()
                v_line = ["version", "20201013"] + [""] * (len(header) - 2)
                pd.DataFrame([v_line]).to_excel(writer, index=False, header=False, startrow=0)
                pd.DataFrame([header]).to_excel(writer, index=False, header=False, startrow=1)
                d_df.to_excel(writer, index=False, header=False, startrow=2)
            
            st.download_button(
                label=f"📥 下載 {platform_name} 上傳檔 (D)",
                data=buf.getvalue(),
                file_name=f"{platform_name}_D_{datetime.now().strftime('%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            # 顯示排除原因（可選）
            with st.expander("查看被排除的訂單原因"):
                st.write("以下訂單因：無物流單號、包含『勿拍』或重複而被剔除。")
                st.dataframe(excluded_df[['join_key', '前台传入商品名称', 'amount', '快递单号']])
        else:

            st.error("❌ 找不到對應的 B 檔與 C 檔欄位，請檢查上傳內容。")

