import streamlit as st
import pandas as pd
import msoffcrypto
import io
import chardet
from datetime import datetime, timedelta

st.set_page_config(page_title="Order App 3.0 - Etmall 全新支援版", page_icon="🛍️", layout="wide")

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
        # 強力清理欄位：拿掉空格與換行
        df.columns = [str(c).strip().replace('\n', '').replace(' ', '') for c in df.columns]
        return df
    except Exception as e:
        st.error(f"檔案 {file.name} 讀取失敗: {e}")
        return None

# --- 側邊欄介面 ---
with st.sidebar:
    st.title("🛡️ 參數設定")
    shop_url = st.text_input("1. 店鋪網址", value="https://www.etmall.com.tw/")
    platform_name = st.text_input("2. 電商平台英文名稱", value="ETMall")
    st.divider()
    use_pass = st.checkbox("3. 檔案有密碼 (僅限對帳單 xlsx)", value=True)
    excel_pass = st.text_input("輸入密碼", value="123456", type="password")
    st.divider()
    f_return = st.checkbox("4. 排除『銷退/取消』訂單", value=True)
    f_old = st.checkbox("5. 排除 >350天舊單", value=True)

st.header("📦 Order App v3.0 - 訂單自動化轉換")
st.info("支援『Etmall 直接轉換』或『分銷商雙檔合併』。系統將根據欄位自動判斷。")

uploaded_files = st.file_uploader("請上傳訂單 Excel 檔案", type=['xlsx', 'xls', 'csv'], accept_multiple_files=True)

if uploaded_files and shop_url:
    if st.button("🚀 開始分析資料", type="primary"):
        etmall_list = []
        b_list, c_list = [], []
        
        for f in uploaded_files:
            df = read_excel_comprehensive(f, use_pass, excel_pass)
            if df is None: continue
            cols = df.columns.tolist()
            
            # --- 判斷邏輯 A: Etmall 新格式 (單檔) ---
            if "出貨指示日" in cols and "訂單編號" in cols:
                st.info(f"✅ 偵測到 Etmall 格式: {f.name}")
                etmall_list.append(df)
            
            # --- 判斷邏輯 B: 舊版分銷商對帳 (需要合併) ---
            elif "渠道单号" in cols or "渠道單號" in cols:
                target_col = "渠道单号" if "渠道单号" in cols else "渠道單號"
                df["join_key"] = df[target_col].astype(str).apply(lambda x: x.split('-')[-1] if '-' in x else x)
                b_list.append(df)
                st.info(f"✅ 偵測到分銷商對帳單: {f.name}")
            
            elif "客户订单号" in cols or "客戶訂單號" in cols:
                target_col = "客户订单号" if "客户订单号" in cols else "客戶訂單號"
                df["join_key"] = df[target_col].astype(str).str.strip()
                c_data = df.rename(columns={"快递单号": "快递单号", "快遞單號": "快递單號", "快递公司": "快递公司", "快遞公司": "快递公司"})
                c_list.append(c_data)
                st.info(f"✅ 偵測到訂單清單: {f.name}")

        final_rows = []

        # 處理 Etmall 格式
        if etmall_list:
            for df in etmall_list:
                for _, row in df.iterrows():
                    # 排除邏輯
                    tracking = str(row.get('配送單號', '')).strip()
                    if tracking == "" or tracking == "nan": continue # 無物流單號排除
                    
                    if f_return:
                        return_status = str(row.get('銷退狀態', '')).strip()
                        if return_status != "" and return_status != "nan": continue # 有銷退資訊排除
                    
                    if "勿拍" in str(row.get('商品名稱', '')): continue
                    
                    # 計算
                    qty = pd.to_numeric(row.get('數量', 1), errors='coerce') or 1
                    unit_price = pd.to_numeric(row.get('售價', 0), errors='coerce') or 0
                    total_amt = qty * unit_price
                    
                    final_rows.append({
                        '订单编号': row.get('訂單編號'),
                        '订单日期': pd.to_datetime(row.get('出貨指示日')).strftime('%Y-%m-%d') if pd.notna(row.get('出貨指示日')) else "",
                        '订单币种': 'TWD',
                        '订单金额': total_amt,
                        '商品名称': row.get('商品名稱'),
                        '商品数量': qty,
                        '商品单价': unit_price,
                        '店铺网址': shop_url,
                        '快递单号': tracking,
                        '物流企业名称': row.get('貨運公司'),
                        '电商平台英文名称': platform_name
                    })

        # 處理舊版雙檔合併格式
        if b_list and c_list:
            df_b = pd.concat(b_list, ignore_index=True)
            df_c = pd.concat(c_list, ignore_index=True).drop_duplicates(subset=["join_key"])
            merged = pd.merge(df_b, df_c, on="join_key", how="left")
            for _, row in merged.iterrows():
                tracking = str(row.get('快递单号', '')).strip()
                if tracking == "" or tracking == "nan": continue
                
                qty = pd.to_numeric(row.get('商品数量', 1), errors='coerce') or 1
                total_amt = pd.to_numeric(row.get('支付总金额', 0), errors='coerce') or 0
                
                final_rows.append({
                    '订单编号': row.get('join_key'),
                    '订单日期': pd.to_datetime(row.get('渠道订单创建时间')).strftime('%Y-%m-%d') if pd.notna(row.get('渠道订单创建时间')) else "",
                    '订单币种': 'USD',
                    '订单金额': total_amt,
                    '商品名称': row.get('前台传入商品名称'),
                    '商品数量': qty,
                    '商品单价': round(total_amt / qty, 2) if qty != 0 else 0,
                    '店铺网址': shop_url,
                    '快递单号': tracking,
                    '物流企业名称': row.get('快递公司'),
                    '电商平台英文名称': platform_name
                })

        # --- 輸出結果 ---
        if final_rows:
            result_df = pd.DataFrame(final_rows).drop_duplicates(subset=["订单编号"])
            
            # 統計
            success_count = len(result_df)
            success_sum = result_df['订单金额'].sum()
            
            st.subheader("📊 數據統計摘要")
            c1, c2 = st.columns(2)
            c1.metric("成功處理筆數", f"{success_count} 筆")
            c2.metric("總成交金額", f"${success_sum:,.2f}")

            # 產生下載檔
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                header = result_df.columns.tolist()
                v_line = ["version", "20201013"] + [""] * (len(header) - 2)
                pd.DataFrame([v_line]).to_excel(writer, index=False, header=False, startrow=0)
                pd.DataFrame([header]).to_excel(writer, index=False, header=False, startrow=1)
                result_df.to_excel(writer, index=False, header=False, startrow=2)
            
            st.divider()
            st.download_button(
                label=f"📥 下載 {platform_name} 格式檔",
                data=buf.getvalue(),
                file_name=f"{platform_name}_D_{datetime.now().strftime('%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.dataframe(result_df.head())
        else:
            st.error("❌ 未能產出有效資料，請檢查檔案內容或過濾條件。")
