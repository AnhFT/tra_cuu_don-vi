from fasthtml.common import *
import pandas as pd

# 1. Khởi tạo ứng dụng FastHTML
app, rt = fast_app(hdrs=(picolink,))

# 2. Đọc và chuẩn hóa dữ liệu Excel
EXCEL_FILE = "danh_sach_don_vi.xlsx"

def load_data():
    try:
        df = pd.read_excel(EXCEL_FILE)
        return df.fillna("").astype(str)
    except Exception as e:
        print(f"Lỗi đọc file Excel: {e}")
        return pd.DataFrame(columns=[
            "MA_DONVI", "TEN_DONVI", "DC_LIEN_HE", 
            "MS_THUE", "DIEN_THOAI_NLH", "NGAY_TANG", 
            "CHUYEN_QUAN_THU", "DIEN_THOAI_CQT"
        ])

df = load_data()

# 3. Render bảng hiển thị kết quả
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

# 4. Giao diện trang chủ
@rt("/")
def get():
    return Titled(
        "🔍 Tra cứu thông tin đơn vị",
        Main(
            Article(
                H3("Nhập thông tin cần tìm"),
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

# 5. Xử lý logic tìm kiếm
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

serve()
