import os
import glob
from tkinter import Tk, filedialog, messagebox
import pandas as pd


def gop_file_excel():
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    # 1. Chọn thư mục nguồn
    thu_muc_nguon = filedialog.askdirectory(
        title="Chọn thư mục chứa các file Excel cần gộp"
    )
    if not thu_muc_nguon:
        messagebox.showwarning("Thông báo", "Bạn chưa chọn thư mục nào!")
        return

    # 2. Tìm danh sách file
    danh_sach_file = (
        glob.glob(os.path.join(thu_muc_nguon, "*.xlsx"))
        + glob.glob(os.path.join(thu_muc_nguon, "*.xls"))
        + glob.glob(os.path.join(thu_muc_nguon, "*.csv"))
    )

    if not danh_sach_file:
        messagebox.showerror(
            "Lỗi", "Không tìm thấy file Excel hoặc CSV nào trong thư mục này!"
        )
        return

    # 3. Chọn nơi lưu file kết quả
    file_output = filedialog.asksaveasfilename(
        title="Chọn nơi lưu file kết quả",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile="File_Gop_Tong_Hop.xlsx",
    )
    if not file_output:
        messagebox.showwarning("Thông báo", "Bạn đã hủy chọn nơi lưu file!")
        return

    # Danh sách để lưu lại vết các file đã gộp thành công nhằm hiển thị thông báo chi tiết
    cac_file_da_gop = []

    # 4. Tiến hành gộp dữ liệu
    try:
        with pd.ExcelWriter(file_output, engine="openpyxl") as writer:
            for duong_dan_file in danh_sach_file:
                ten_file = os.path.basename(duong_dan_file)
                ten_sheet = os.path.splitext(ten_file)[0]
                ten_sheet = "".join(
                    c for c in ten_sheet if c not in r"\/*?:[]"
                )[:31]

                # Đọc dữ liệu dạng chữ (dtype=str) để giữ nguyên số 0 ở đầu
                if duong_dan_file.endswith(".csv"):
                    df = pd.read_csv(
                        duong_dan_file, encoding="utf-8-sig", dtype=str
                    )
                else:
                    df = pd.read_excel(duong_dan_file, dtype=str)

                df.to_excel(writer, sheet_name=ten_sheet, index=False)

                # --- Đã khôi phục dòng thông báo chi tiết trên Terminal ---
                thong_bao = f"Đã gộp file: {ten_file} -> Sheet: {ten_sheet}"
                print(thong_bao)
                cac_file_da_gop.append(thong_bao)

        # Tạo nội dung thông báo chi tiết cho hộp thoại popup cuối cùng
        chi_tiet_ket_qua = "\n".join(cac_file_da_gop)
        messagebox.showinfo(
            "Thành công",
            f"Đã gộp thành công vào file:\n{file_output}\n\nChi tiết các sheet đã tạo:\n{chi_tiet_ket_qua}",
        )

    except Exception as e:
        messagebox.showerror("Lỗi hệ thống", f"Đã xảy ra lỗi khi xử lý:\n{str(e)}")


if __name__ == "__main__":
    gop_file_excel()