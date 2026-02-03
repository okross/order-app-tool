import streamlit as st
import pandas as pd
import msoffcrypto
import io
from datetime import datetime, timedelta

st.set_page_config(page_title="Order App 3.0 - 訂單轉換", page_icon="💰", layout="wide")

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

# --- 側邊欄 ---
with st.sidebar:
    st.title("🛡️ 參數設定")
    shop_url = st.text_input("1. 店鋪網址", value="https://www.etmall.com.tw/")
    platform_name = st.text_input("2. 電商平台名稱", value="ETMall")
    exchange_rate = st.number_input("3. 匯率 (1 USD = ? NTD)", value=32.0, step=0.1)
    st.divider()
    use_pass = st.checkbox("4. 檔案有密碼", value=True)
    excel_pass = st.text_input("輸入密碼", value="123456", type="password")
    f_return = st.checkbox("5. 排除銷退訂單", value=True)

st.header("📦 Order App v3.0 - 財務統計強化版")

uploaded_files = st.file_uploader("上傳訂單 Excel", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

if uploaded_files and shop_url:
    if st.button("🚀 開始分析", type="primary"):
        final_rows = []
        
        for f in uploaded_files:
            df = read_excel_comprehensive(f, use_pass, excel_pass)
            if df is None: continue
            
            # Etmall 邏輯
            if "出貨指示日" in df.columns:
                for _, row in df.iterrows():
                    tracking = str(row.get('配送單號', '')).strip()
                    if tracking in ["", "nan"]: continue
                    
                    if f_return:
                        ret = str(row.get('銷退狀態', '')).strip()
                        if ret not in ["", "nan"]: continue
                    
                    qty = pd.to_numeric(row.get('數量', 1), errors='coerce') or 1
                    price = pd.to_numeric(row.get('售價', 0), errors='coerce') or 0
                    
                    final_rows.append({
                        '订单编号': row.get('訂單編號'),
                        '订单日期': pd.to_datetime(row.get('出貨指示日')).strftime('%Y-%m-%d') if pd.notna(row.get('出貨指示日')) else "",
                        '订单币种': 'TWD',
                        '订单金额': qty * price,
                        '商品名称': row.get('商品名稱'),
                        '商品数量': qty,
                        '商品单价': price,
                        '店铺网址': shop_url,
                        '快递单号': tracking,
                        '物流企业名称': row.get('貨運公司'),
                        '电商平台英文名称': platform_name
                    })

        if final_rows:
            result_df = pd.DataFrame(final_rows).drop_duplicates(subset=["订单编号"])
            success_count = len(result_df)
            success_sum_ntd = result_df['订单金額' if '订单金額' in result_df.columns else '订单金额'].sum()
            success_sum_usd = success_sum_ntd / exchange_rate

            # --- 顯示修改後的數據統計 ---
            st.subheader("📊 數據統計摘要")
            c1, c2 = st.columns(2)
            c1.metric("成功處理筆數", f"{success_count} 筆")
            c2.metric("總成交金額 (NTD)", f"{success_sum_ntd:,.0f} NTD")
            
            st.write(f"💵 **約合美金：** `${success_sum_usd:,.2f} USD` (以匯率 {exchange_rate} 計算)")

            # 下載按鈕
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                header = result_df.columns.tolist()
                v_line = ["version", "20201013"] + [""] * (len(header) - 2)
                pd.DataFrame([v_line]).to_excel(writer, index=False, header=False, startrow=0)
                pd.DataFrame([header]).to_excel(writer, index=False, header=False, startrow=1)
                result_df.to_excel(writer, index=False, header=False, startrow=2)
            
            st.divider()
            st.download_button(f"📥 下載 {platform_name} 格式檔", buf.getvalue(), f"{platform_name}_v3.xlsx", type="primary")
            st.dataframe(result_df.head())
