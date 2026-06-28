import streamlit as st
import pandas as pd
import tempfile
import os
import sys
import asyncio

if sys.platform == 'win32':
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

from validators import clean_stock_symbols, is_financial_stock, check_duplicates
from cafef_ppe_client import CafeFPPEClient
from fireant_client import FireAntClient
from fireant_mapper import FireantMapper, build_unified_data
import asyncio
from excel_inspector import ExcelInspector
from excel_writer import ExcelWriter
from html_parser import build_raw_data_from_htmls

st.set_page_config(page_title="CafeF Financial Data Tool", layout="wide")

st.title("Tự động lấy dữ liệu tài chính từ CafeF")
st.markdown("Công cụ hỗ trợ tải dữ liệu tài chính (Bảng cân đối kế toán, KQKD, LCTT) và điền tự động vào file Excel mẫu.")

# Khởi tạo session state
if "crawled" not in st.session_state:
    st.session_state.crawled = False
if "mapped_data_list" not in st.session_state:
    st.session_state.mapped_data_list = []
if "logs" not in st.session_state:
    st.session_state.logs = []
if "tmp_path" not in st.session_state:
    st.session_state.tmp_path = ""

# --- Phần 1: Upload File ---
st.header("1. Upload file Excel mẫu")
uploaded_file = st.file_uploader("Chọn file Excel mẫu (.xlsx)", type=["xlsx"])

# --- Phần 2: Nhập mã cổ phiếu ---
st.header("2. Nhập mã cổ phiếu")
tickers_input = st.text_area("Danh sách mã (cách nhau bởi dấu phẩy, khoảng trắng hoặc xuống dòng):", value="", placeholder="Ví dụ: VCG, HPG, FPT, MWG")

# --- Phần 3: Chọn thời gian ---
st.header("3. Chọn thời gian")
col1, col2 = st.columns(2)
with col1:
    start_year = st.number_input("Năm bắt đầu", min_value=2000, max_value=2050, value=2021)
with col2:
    end_year = st.number_input("Năm kết thúc", min_value=2000, max_value=2050, value=2025)

# --- Phần 4: Nguồn Dữ Liệu & Tùy chọn ---
st.header("4. Nguồn Dữ Liệu & Tùy chọn")
data_source = st.radio("Nguồn dữ liệu:", [
    "Tự động lấy qua FireAnt API (Fallback PPE CafeF)", 
    "Tải lên file HTML thủ công (BCTC, CDKT, LCTT)"
])

uploaded_htmls = []
if "Thủ công" in data_source:
    uploaded_htmls = st.file_uploader("Upload các file HTML (Có thể chọn nhiều file cùng lúc)", type=["html"], accept_multiple_files=True)

st.subheader("Tùy chọn khác")
col3, col4 = st.columns(2)
with col3:
    is_headless = st.checkbox("Chạy ngầm (Headless browser)", value=True, help="Tắt để mở trình duyệt xem tool đang làm gì hoặc để đăng nhập thủ công.")
    user_data_dir = st.text_input("Thư mục lưu phiên đăng nhập (Tuỳ chọn):", value="", placeholder="Ví dụ: ./fireant_session")
with col4:
    unit_info = st.info("Hệ thống tự động chuẩn hóa đơn vị về **Triệu đồng** cho tất cả nguồn.")
ignore_duplicates = st.checkbox("Bỏ qua mã đã có trong file (Không ghi thêm)", value=True)

btn_crawl = st.button("Crawl dữ liệu", type="primary")

# --- Nút Crawl Dữ Liệu ---
if btn_crawl:
    if not uploaded_file:
        st.error("Vui lòng upload file Excel mẫu!")
        st.stop()
        
    if start_year > end_year:
        st.error("Năm bắt đầu không thể lớn hơn năm kết thúc!")
        st.stop()
        
    tickers = clean_stock_symbols(tickers_input)
    if not tickers:
        st.error("Vui lòng nhập ít nhất 1 mã cổ phiếu hợp lệ!")
        st.stop()
        
    st.info(f"Đã nhận diện {len(tickers)} mã cổ phiếu: {', '.join(tickers)}")
    
    # Save uploaded file to temp
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.getvalue())
        st.session_state.tmp_path = tmp.name
        
    # Inspect Excel
    with st.spinner("Đang kiểm tra file Excel..."):
        try:
            inspector = ExcelInspector(st.session_state.tmp_path)
            existing_symbols, max_id = inspector.get_existing_symbols_and_max_id()
        except Exception as e:
            st.error(f"Lỗi khi đọc file Excel: {e}")
            st.stop()
            
    st.write(f"- ID lớn nhất hiện tại trong file: **{max_id}**")
    
    new_tickers, duplicate_tickers = check_duplicates(tickers, existing_symbols)
    if duplicate_tickers:
        st.warning(f"Các mã đã có trong file: {', '.join(duplicate_tickers)}")
        if ignore_duplicates:
            tickers = new_tickers
            st.info(f"Sẽ chỉ xử lý {len(tickers)} mã mới.")
            
    if not tickers:
        st.warning("Không có mã nào cần xử lý mới. Dừng!")
        st.stop()
        
    st.header("Trạng thái Crawl")
    progress_bar = st.progress(0)
    status_text = st.empty()
    


    unified_data_list = []
    logs = []
    
    if "Tự động" in data_source:
        if user_data_dir and not os.path.exists(user_data_dir):
            os.makedirs(user_data_dir, exist_ok=True)
        # Chỉ lấy FireAnt token một lần
        status_text.text("Đang khởi tạo FireAnt API...")
        client = FireAntClient(headless=is_headless, user_data_dir=user_data_dir if user_data_dir else None)
        
        # Pre-fetch token once so threads don't race to open browsers
        try:
            loop = asyncio.get_event_loop()
        except RuntimeError:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
        loop.run_until_complete(client._get_token())
        
        ppe_client = CafeFPPEClient()
    
    import concurrent.futures
    
    def process_ticker(ticker):
        source_name = "N/A"
        raw_data = {"status": "MISSING_OR_NO_PERMISSION"}
        ppe_res = {"status": "MISSING", "data": {}}
        
        if "Tự động" in data_source:
            raw_data = client.scrape_symbol_financials(ticker, start_year, end_year)
            source_name = "FireAnt API"
            
            if raw_data.get("status") == "OK":
                ppe_res = ppe_client.get_ppe(ticker, start_year, end_year)
        else:
            ticker_htmls = []
            for f in uploaded_htmls:
                if len(tickers) == 1 or ticker.lower() in f.name.lower():
                    ticker_htmls.append(f.getvalue().decode('utf-8'))
                    
            if ticker_htmls:
                from html_parser import build_raw_data_from_htmls
                raw_data = build_raw_data_from_htmls(ticker_htmls, ticker, "FireAnt")
                source_name = "Thủ công"
                
        mapper = FireantMapper(raw_data)
        unified_dict = build_unified_data(mapper, ticker, start_year, end_year, unit_multiplier=1.0 / 1000000.0)
        
        if "Tự động" in data_source:
            if ppe_res["status"] == "OK":
                for y, val in ppe_res["data"].items():
                    if y in unified_dict["years"]:
                        unified_dict["years"][y]["ppe_gross"] = {"value": val, "source": "CafeF"}
            else:
                for y in unified_dict["years"]:
                    unified_dict["years"][y]["ppe_gross"] = {"value": None, "source": "MISSING_FROM_CAFEF"}
                    
        return ticker, raw_data, unified_dict

    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        futures = {executor.submit(process_ticker, t): t for t in tickers}
        completed_count = 0
        
        for future in concurrent.futures.as_completed(futures):
            ticker = futures[future]
            completed_count += 1
            status_text.text(f"Đang xử lý {ticker} ({completed_count}/{len(tickers)})...")
            
            try:
                _, raw_data, unified_dict = future.result()
                
                if raw_data.get("status") == "MISSING_OR_NO_PERMISSION":
                    st.warning(f"Mã {ticker} không lấy được dữ liệu tài chính.")
                    
                unified_data_list.append(unified_dict)
                
                for year in range(start_year, end_year + 1):
                    yd = unified_dict["years"].get(year, {})
                    for k, f in [("Tổng tài sản đầu kỳ", "total_assets_prev"), ("Tổng tài sản năm t", "total_assets"), ("Nợ phải trả năm t", "liabilities"), ("Các khoản phải thu ngắn hạn", "short_term_receivables"), ("Hàng tồn kho", "inventory"), ("Doanh thu thuần", "revenue"), ("Lợi nhuận sau thuế", "net_profit_after_tax"), ("Dòng tiền từ hoạt động kinh doanh", "operating_cash_flow")]:
                        val_key = "total_assets" if f == "total_assets_prev" else f
                        src_dict = unified_dict["years"].get(year - 1, {}) if "_prev" in f else yd
                        info = src_dict.get(val_key, {})
                        logs.append({"ticker": ticker, "year": year, "sheet": "DATA", "column": k, "source": info.get("source", ""), "value": info.get("value"), "status": "OK" if info.get("value") is not None else ("MISSING_OR_NO_PERMISSION" if raw_data.get("status") != "OK" else "MISSING"), "note": "", "endpoint": "API /full-financial-reports"})
                    
                    for k, f in [("Lợi nhuận sau thuế", "net_profit_after_tax"), ("Dòng tiền HĐKD", "operating_cash_flow"), ("Khoản phải thu ngắn hạn năm t", "short_term_receivables"), ("PPE (Nguyên giá TSCĐ)", "ppe_gross"), ("Doanh thu năm t", "revenue"), ("Doanh thu năm t-1", "revenue_prev"), ("ROA", "roa"), ("Tổng tài sản đầu kỳ", "total_assets_prev")]:
                        val_key = f.replace("_prev", "") if "_prev" in f else f
                        src_dict = unified_dict["years"].get(year - 1, {}) if "_prev" in f else yd
                        info = src_dict.get(val_key, {})
                        logs.append({"ticker": ticker, "year": year, "sheet": "DATA_DA", "column": k, "source": info.get("source", ""), "value": info.get("value"), "status": "OK" if info.get("value") is not None else ("MISSING_OR_NO_PERMISSION" if raw_data.get("status") != "OK" else "MISSING"), "note": "", "endpoint": "API /full-financial-reports"})
                        
            except Exception as e:
                st.error(f"Lỗi xử lý mã {ticker}: {e}")
                
            progress_bar.progress(completed_count / len(tickers))
        
    status_text.text("Crawl hoàn tất!")
    st.session_state.unified_data_list = unified_data_list
    st.session_state.logs = logs
    st.session_state.crawled = True
    
if st.session_state.crawled:
    st.success("Đã có dữ liệu! Tiến hành xuất Excel.")
    
    st.header("5. Các chỉ tiêu bị thiếu (MISSING)")
    missing_logs = [log for log in st.session_state.logs if "MISSING" in log["status"]]
    if missing_logs:
        st.dataframe(pd.DataFrame(missing_logs))
    else:
        st.info("Tuyệt vời! Không có chỉ tiêu nào bị thiếu.")
        
    st.header("6. Xuất Excel")
    if st.button("Xuất Excel", type="primary"):
        with st.spinner("Đang ghi dữ liệu vào Excel..."):
            try:
                writer = ExcelWriter(st.session_state.tmp_path)
                writer.append_data(st.session_state.unified_data_list, st.session_state.logs, start_year, end_year)
                
                out_path = st.session_state.tmp_path.replace(".xlsx", "_FILLED.xlsx")
                writer.save(out_path)
                
                st.success("Đã ghi dữ liệu thành công!")
                
                stats = writer.stats
                st.write("### Báo cáo ghi dữ liệu:")
                for sheet_key, sheet_name in [("DATA", "DATA "), ("DATA_DA", "DATA DA")]:
                    s = stats[sheet_key]
                    if s["written"] > 0:
                        st.write(f"- **Sheet `{sheet_name}`**: Ghi {s['written']} dòng (Từ dòng {s['start_row']} đến {s['end_row']})")
                    else:
                        st.write(f"- **Sheet `{sheet_name}`**: Không có dòng nào được ghi.")
                        
                with open(out_path, "rb") as f:
                    st.download_button(
                        label="📥 7. Tải file Excel kết quả",
                        data=f.read(),
                        file_name="COBRP_2026_FILLED_FROM_FIREANT.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"Lỗi khi ghi file: {e}")
