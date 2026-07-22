import subprocess
from datetime import datetime, timedelta
from fasthtml.common import *
import pandas as pd
import os
from datetime import datetime

# 1. Khởi tạo ứng dụng FastHTML với Pico CSS
app, rt = fast_app(pico=True)

EXCEL_FILE = "danh_sach_don_vi.xlsx"

# 2. Hàm lấy thời gian cập nhật file Excel gần nhất
import openpyxl
from datetime import timedelta

def get_file_update_time():
    try:
        if os.path.exists(EXCEL_FILE):
            wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True)
            props = wb.properties
            mod_time = props.modified or props.created
            wb.close()  # Đóng file ngay sau khi lấy thông tin
            
            if mod_time:
                if mod_time.tzinfo is None:
                    vn_time = mod_time + timedelta(hours=7)
                else:
                    vn_time = mod_time.astimezone()
                return vn_time.strftime("%d/%m/%Y %H:%M")
    except Exception as e:
        print(f"Lỗi đọc metadata Excel: {e}")
        
    return "Chưa xác định"

# 3. Hàm làm sạch số điện thoại và mã số thuế
def clean_phone(val):
    val = str(val).strip()
    if val.endswith('.0'):
        val = val[:-2]
    if val and val != "nan" and not val.startswith('0') and len(val) in [9, 10]:
        val = '0' + val
    return "" if val == "nan" else val

# 4. Hàm đọc dữ liệu từ Excel
def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE, dtype=str)
        df = df.fillna("")
        
        phone_cols = ['DIEN_THOAI_NLH', 'DIEN_THOAI_CQT', 'MS_THUE']
        for col in phone_cols:
            if col in df.columns:
                df[col] = df[col].apply(clean_phone)
                
        return df
    except Exception as e:
        print(f"Lỗi đọc file Excel: {e}")
        return pd.DataFrame(columns=[
            "MA_DONVI", "TEN_DONVI", "DC_LIEN_HE", 
            "MS_THUE", "DIEN_THOAI_NLH", "NGAY_TANG", 
            "CHUYEN_QUAN_THU", "DIEN_THOAI_CQT"
        ])

df = load_data()
LAST_UPDATE = get_file_update_time()

# 5. Hàm render bảng kết quả
def render_results(results):
    if results.empty:
        return Article(P("⚠️ Không tìm thấy kết quả phù hợp.", style="color: orange;"))
    
    rows = []
    for _, row in results.iterrows():
        rows.append(
            Tr(
                Td(B(row.get('MA_DONVI', ''))),
                Td(row.get('TEN_DONVI', '')),
                Td(row.get('DC_LIEN_HE', '')),
                Td(row.get('MS_THUE', '')),
                Td(row.get('DIEN_THOAI_NLH', '')),
                Td(row.get('NGAY_TANG', '')),
                Td(row.get('CHUYEN_QUAN_THU', '')),
                Td(row.get('DIEN_THOAI_CQT', ''))
            )
        )
    
    return Div(
        P(f"✅ Tìm thấy {len(results)} kết quả (hiển thị tối đa 20):", style="color: green; font-weight: bold;"),
        Div(
            Table(
                Thead(
                    Tr(
                        Th("Mã đơn vị"), 
                        Th("Tên đơn vị"), 
                        Th("Địa chỉ liên hệ"), 
                        Th("Mã số thuế"),
                        Th("SĐT NLH"),
                        Th("Ngày tăng"),
                        Th("Chuyên quản thu"),
                        Th("SĐT CQT")
                    )
                ),
                Tbody(*rows)
            ),
            style="overflow-x: auto;"
        )
    )

# 6. Giao diện trang chủ (GET /)
@rt("/")
def get():
    return Titled(
        "🔍 Tra cứu thông tin đơn vị",
        Main(
            Article(
                # Thẻ Div sắp xếp Tiêu đề bên trái và Ngày update căn lề bên phải
                Div(
                    H3("Nhập thông tin cần tìm", style="margin-bottom: 0;"),
                    Small(
                        f"📅 Dữ liệu cập nhật: {LAST_UPDATE}", 
                        style="color: #666; font-style: italic; white-space: nowrap;"
                    ),
                    style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; margin-bottom: 15px;"
                ),
                P("Tìm theo Mã đơn vị, Tên, Địa chỉ, MST, SĐT, Chuyên quản hoặc SĐT CQT:"),
                Input(
                    type="search",
                    name="query",
                    placeholder="Nhập từ khóa tìm kiếm...",
                    hx_post="/search",
                    hx_trigger="keyup changed delay:250ms",
                    hx_target="#result-area"
                )
            ),
            Div(id="result-area"),
            cls="container"
        )
    )

# 7. Xử lý tìm kiếm (POST /search)
@rt("/search")
def post(query: str):
    query_str = query.strip().lower()
    
    if not query_str:
        return Div()
    
    mask = (
        df['MA_DONVI'].str.lower().str.contains(query_str) |
        df['TEN_DONVI'].str.lower().str.contains(query_str) |
        df['DC_LIEN_HE'].str.lower().str.contains(query_str) |
        df['MS_THUE'].str.lower().str.contains(query_str) |
        df['DIEN_THOAI_NLH'].str.lower().str.contains(query_str) |
        df['CHUYEN_QUAN_THU'].str.lower().str.contains(query_str)
    )
    
    if 'DIEN_THOAI_CQT' in df.columns:
        mask = mask | df['DIEN_THOAI_CQT'].str.lower().str.contains(query_str)

    filtered_df = df[mask].head(20)
    return render_results(filtered_df)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app:app", host="0.0.0.0", port=8000, reload=True)
