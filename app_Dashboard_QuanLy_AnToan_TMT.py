import streamlit as st
import pandas as pd
import os
import openpyxl

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & BẢO MẬT ĐĂNG NHẬP
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Quản Lý An Toàn TMT - Version 1.0 20260925",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Cấu hình tài khoản đăng nhập
USER_CREDENTIALS = {
    "tmt": "123456",     # Username: tmt | Pass: 123456
    "admin": "123456"
}

# Quản lý Session State đăng nhập
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

# MÀN HÌNH ĐĂNG NHẬP
if not st.session_state.logged_in:
    st.title("🛡️ HỆ THỐNG QUẢN LÝ AN TOÀN TMT")
    st.caption("📌 Version 1.0 20260925")
    st.markdown("---")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("🔐 Đăng Nhập Hệ Thống")
        with st.form("login_form"):
            user_input = st.text_input("Tên đăng nhập:", value="tmt")
            pass_input = st.text_input("Mật khẩu:", type="password")
            submit_login = st.form_submit_button("🔑 Đăng Nhập", type="primary", use_container_width=True)
            
            if submit_login:
                if user_input in USER_CREDENTIALS and USER_CREDENTIALS[user_input] == pass_input:
                    st.session_state.logged_in = True
                    st.session_state.username = user_input
                    st.success("Đăng nhập thành công!")
                    st.rerun()
                else:
                    st.error("❌ Tên đăng nhập hoặc mật khẩu không chính xác!")
    st.stop()

# ---------------------------------------------------------
# 2. KHỞI TẠO ĐƯỜNG DẪN & DỮ LIỆU
# ---------------------------------------------------------
EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(EXCEL_DIR):
    EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

EXCEL_PATH_WEB = os.path.join(EXCEL_DIR, "DanhMuc_CongCu_WEB_TMT.xlsx")
EXCEL_PATH_QUAN_LY = os.path.join(EXCEL_DIR, "QuanLy_AnToan_TMT.xlsx")
EXCEL_PATH_BC_DINH_KY = os.path.join(EXCEL_DIR, "DanhSach_BaoCao_DinhKy_TMT.xlsx")
EXCEL_PATH_GSHEET = os.path.join(EXCEL_DIR, "DanhSach_Gsheet_TMT.xlsx")

CATEGORIES = [
    "DTTU_01 AT", "DTTU_01 AT 01 Bao cao", "DTTU_01 AT 01 Bao cao 2026",
    "DTTU_02 PCTT", "DTTU_03 PCCC", "DTTU_04 HL", "DTTU_05 ATDTXD",
    "DTTU_ATGT", "DTTU_CNTT", "DTTU_DCAT", "DTTU_DCNN va Cac loai xe",
    "DTTU_DGRR", "DTTU_HNTH cac loai", "DTTU_KIEM TRA",
    "DTTU_KIEM TRA-Thuc hien Kien Nghi", "DTTU_UCKC",
    "DTTU_UCKC dien tap cac loai", "DTTU_khac 01 PHOI HOP CAC TO",
    "DTTU_khac 02 XEM DE BIET CTY", "DTTU_khac 03 ATD dia phuong",
    "DTTU_khac 03 XEM DE BIET dia phuong", "Quy dinh 0000 Discussion",
    "Quy dinh GOV", "Quy dinh PCTN", "Quy dinh PCTN file tham khao cac Doi",
    "Quy dinh SPC va EVN", "Quy dinh trao doi EVN-SPC-PCTN"
]

def reindex_df(df):
    if not df.empty:
        df = df.reset_index(drop=True)
        df["STT"] = df.index + 1
    return df

def auto_fit_columns(workbook):
    for sheetname in workbook.sheetnames:
        worksheet = workbook[sheetname]
        for col in worksheet.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    val_str = str(cell.value)
                    if val_str.startswith('=HYPERLINK'):
                        max_len = max(max_len, 20)
                    else:
                        max_len = max(max_len, len(val_str))
            adjusted_width = max(max_len + 4, 15)
            worksheet.column_dimensions[col_letter].width = min(adjusted_width, 60)

# HEADER CỦA HỆ THỐNG
st.title("🛡️ Hệ Thống Quản Lý An Toàn & Công Tác Chuyên Môn TMT")
st.markdown("##### 🏷️ **Version 1.0 20260925**")
st.markdown("---")

# ---------------------------------------------------------
# 3. SIDEBAR (THANH ĐIỀU HƯỚNG & ĐĂNG XUẤT)
# ---------------------------------------------------------
st.sidebar.markdown(f"👤 **Xin chào:** `{st.session_state.username}`")

col_btn1, col_btn2 = st.sidebar.columns(2)
with col_btn1:
    if st.button("🚪 Đăng xuất", use_container_width=True):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.rerun()

with col_btn2:
    if st.button("🧹 Xóa Cache", use_container_width=True):
        st.cache_data.clear()
        st.toast("Đã xóa cache thành công!", icon="🎉")

st.sidebar.markdown("---")
st.sidebar.header("📂 PHÂN MỤC CHÍNH")
main_menu = st.sidebar.radio(
    "Chọn phân mục làm việc:",
    [
        "1 🌐 DS WEBsites_CV", 
        "2 📋 DM QL Files", 
        "3 📊 DS BCdinhky_CV",
        "4 🟢 DS Gsheet_CV"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption(f"📌 Thư mục lưu trữ Excel:\n`{EXCEL_DIR}`")

# ---------------------------------------------------------
# 4. KHO DỮ LIỆU SESSION STATE
# ---------------------------------------------------------
if "web_tools_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_WEB):
        try:
            st.session_state.web_tools_df = reindex_df(pd.read_excel(EXCEL_PATH_WEB))
        except Exception:
            pass
    if "web_tools_df" not in st.session_state:
        st.session_state.web_tools_df = pd.DataFrame([
            {"STT": 1, "Mô tả WEB": "Hệ thống D-Office", "Link truy cập": "https://doffice.evn.com.vn", "Ghi chú": "Quản lý văn bản điều hành"},
            {"STT": 2, "Mô tả WEB": "Cổng thông tin Điện lực", "Link truy cập": "https://evnspc.vn", "Ghi chú": "Tra cứu quy định & chỉ đạo"}
        ])

if "data_store" not in st.session_state:
    st.session_state.data_store = {}
    if os.path.exists(EXCEL_PATH_QUAN_LY):
        try:
            excel_file = pd.ExcelFile(EXCEL_PATH_QUAN_LY)
            for idx, cat in enumerate(CATEGORIES):
                sheet_name = f"MKT_{idx+1}"
                if sheet_name in excel_file.sheet_names:
                    df_read = pd.read_excel(excel_file, sheet_name=sheet_name)
                    st.session_state.data_store[cat] = reindex_df(df_read)
                else:
                    st.session_state.data_store[cat] = pd.DataFrame(columns=["STT", "Thư mục / Hồ sơ", "Link xem", "Ghi chú"])
        except Exception:
            pass

    if not st.session_state.data_store:
        for cat in CATEGORIES:
            st.session_state.data_store[cat] = pd.DataFrame([
                {"STT": 1, "Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link xem": "https://drive.google.com", "Ghi chú": "Cập nhật định kỳ"}
            ])

if "bc_dinhky_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_BC_DINH_KY):
        try:
            st.session_state.bc_dinhky_df = reindex_df(pd.read_excel(EXCEL_PATH_BC_DINH_KY))
        except Exception:
            pass
    if "bc_dinhky_df" not in st.session_state:
        st.session_state.bc_dinhky_df = pd.DataFrame([
            {"STT": 1, "Tên Báo Cáo / Công Việc": "Báo cáo công tác An toàn định kỳ Quý", "Tần suất": "Hàng Quý", "Đơn vị nhận": "Công ty Điện lực", "Link biểu mẫu": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 20 cuối quý"},
            {"STT": 2, "Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link biểu mẫu": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 25 hàng tháng"}
        ])

if "gsheet_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_GSHEET):
        try:
            st.session_state.gsheet_df = reindex_df(pd.read_excel(EXCEL_PATH_GSHEET))
        except Exception:
            pass
    if "gsheet_df" not in st.session_state:
        st.session_state.gsheet_df = pd.DataFrame([
            {"STT": 1, "Mô tả Google Sheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link Google Sheet": "https://docs.google.com/spreadsheets", "Ghi chú": "Dùng chung phòng An Toàn"},
            {"STT": 2, "Mô tả Google Sheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link Google Sheet": "https://docs.google.com/spreadsheets", "Ghi chú": "Cập nhật trực tuyến"}
        ])

# ---------------------------------------------------------
# 5. GIAO DIỆN CÁC MỤC LÀM VIỆC
# ---------------------------------------------------------
# MỤC 1: DS WEBsites_CV
if main_menu == "1 🌐 DS WEBsites_CV":
    st.subheader("🌐 Bảng Danh Sách WEBsites_CV")
    st.markdown("*(Sửa dữ liệu trực tiếp trong bảng -> Bấm **💾 Lưu cập nhật** bên dưới)*")

    edited_web_df = st.data_editor(
        st.session_state.web_tools_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Mô tả WEB", "Link truy cập", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Mô tả WEB": st.column_config.TextColumn("Mô tả WEB", width="large"),
            "Link truy cập": st.column_config.LinkColumn("Link truy cập", display_text="🔗 Truy cập Web", width="medium"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium")
        },
        key="editor_web"
    )

    if st.button("💾 Lưu cập nhật DS WEBsites_CV", type="primary"):
        st.session_state.web_tools_df = reindex_df(edited_web_df)
        st.success("Đã lưu cập nhật danh sách WEBsites thành công!")
        st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel WEBsites_CV")
    
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        if st.button("📥 Xuất toàn bộ Excel WEBsites_CV", type="primary", use_container_width=True):
            try:
                with pd.ExcelWriter(EXCEL_PATH_WEB, engine='openpyxl') as writer:
                    st.session_state.web_tools_df.to_excel(writer, sheet_name="WEBSITES", index=False)
                    auto_fit_columns(writer.book)
                st.success(f"Đã xuất thành công tại: `{EXCEL_PATH_WEB}`")
            except Exception as e:
                st.error(f"Lỗi xuất file: {e}")

    with col_w2:
        up_w = st.file_uploader("Nhập file Excel WEBsites để cập nhật:", type=["xlsx", "xls"], key="up_w")
        if up_w:
            try:
                df_u = pd.read_excel(up_w)
                st.session_state.web_tools_df = reindex_df(df_u)
                st.success("Đã đồng bộ dữ liệu WEBsites thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi nhập file: {e}")

# MỤC 2: DM QL Files
elif main_menu == "2 📋 DM QL Files":
    st.sidebar.subheader("📋 Danh Mục Quản Lý")
    selected_cat = st.sidebar.radio("Chọn mảng công việc:", CATEGORIES)

    if selected_cat:
        st.subheader(f"📂 Quản Lý Hồ Sơ: {selected_cat}")
        st.markdown("*(Sửa dữ liệu trực tiếp trong bảng -> Bấm **💾 Lưu cập nhật** bên dưới)*")

        current_df = st.session_state.data_store[selected_cat]

        edited_df = st.data_editor(
            current_df,
            num_rows="dynamic",
            use_container_width=True,
            column_order=["STT", "Thư mục / Hồ sơ", "Link xem", "Ghi chú"],
            column_config={
                "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
                "Thư mục / Hồ sơ": st.column_config.TextColumn("Thư mục / Hồ sơ", width="large"),
                "Link xem": st.column_config.LinkColumn("Link xem", display_text="🔗 Mở xem", width="medium"),
                "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium")
            },
            key=f"editor_{selected_cat}"
        )

        if st.button("💾 Lưu cập nhật Mảng Công Việc", type="primary"):
            st.session_state.data_store[selected_cat] = reindex_df(edited_df)
            st.success(f"Đã lưu cập nhật cho **{selected_cat}**!")
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel Quản Lý Hồ Sơ")

    col_q1, col_q2 = st.columns(2)
    with col_q1:
        if st.button("📥 Xuất toàn bộ Excel DM QL Files", type="primary", use_container_width=True):
            try:
                with pd.ExcelWriter(EXCEL_PATH_QUAN_LY, engine='openpyxl') as writer:
                    summary_list = []
                    for idx, cat in enumerate(CATEGORIES):
                        sheet_code = f"MKT_{idx+1}"
                        summary_list.append({
                            "STT": idx + 1,
                            "Mảng công việc": cat,
                            "Số lượng hồ sơ": len(st.session_state.data_store[cat]),
                            "Liên kết Sheet": f'=HYPERLINK("#\'{sheet_code}\'!A1", "👉 Đến Sheet {sheet_code}")'
                        })
                    pd.DataFrame(summary_list).to_excel(writer, sheet_name="SHEET_TONG", index=False)

                    for idx, cat in enumerate(CATEGORIES):
                        sheet_code = f"MKT_{idx+1}"
                        df_s = reindex_df(st.session_state.data_store[cat])
                        df_s.to_excel(writer, sheet_name=sheet_code, index=False, startrow=2)

                    wb = writer.book
                    for idx, cat in enumerate(CATEGORIES):
                        sheet_code = f"MKT_{idx+1}"
                        ws = wb[sheet_code]
                        c = ws['A1']
                        c.value = "🏠 Bấm vào đây để về SHEET_TONG"
                        c.hyperlink = "#'SHEET_TONG'!A1"
                        c.style = "Hyperlink"

                    auto_fit_columns(wb)

                st.balloons()
                st.success(f"Đã xuất thành công tại: `{EXCEL_PATH_QUAN_LY}`")
            except Exception as e:
                st.error(f"Lỗi xuất file Excel: {e}")

    with col_q2:
        up_q = st.file_uploader("Nhập file Excel Quản Lý để cập nhật:", type=["xlsx", "xls"], key="up_q")
        if up_q:
            try:
                excel_u = pd.ExcelFile(up_q)
                for idx, cat in enumerate(CATEGORIES):
                    sheet_code = f"MKT_{idx+1}"
                    if sheet_code in excel_u.sheet_names:
                        df_r = pd.read_excel(excel_u, sheet_name=sheet_code, skiprows=lambda x: x == 0)
                        st.session_state.data_store[cat] = reindex_df(df_r.dropna(how="all"))
                st.success("Đã đồng bộ dữ liệu Quản Lý Hồ Sơ thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi nhập file Excel: {e}")

# MỤC 3: DS BCdinhky_CV
elif main_menu == "3 📊 DS BCdinhky_CV":
    st.subheader("📊 Bảng Danh Sách Báo Cáo Định Kỳ & Công Việc")
    st.markdown("*(Sửa dữ liệu trực tiếp trong bảng -> Bấm **💾 Lưu cập nhật** bên dưới)*")

    edited_bc_df = st.data_editor(
        st.session_state.bc_dinhky_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Tên Báo Cáo / Công Việc", "Tần suất", "Đơn vị nhận", "Link biểu mẫu", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Tên Báo Cáo / Công Việc": st.column_config.TextColumn("Tên Báo Cáo / Công Việc", width="large"),
            "Tần suất": st.column_config.SelectboxColumn("Tần suất", options=["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"], width="medium"),
            "Đơn vị nhận": st.column_config.TextColumn("Đơn vị nhận", width="medium"),
            "Link biểu mẫu": st.column_config.LinkColumn("Link biểu mẫu", display_text="🔗 Tải / Xem Biểu Mẫu", width="medium"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium")
        },
        key="editor_bc_dinhky"
    )

    if st.button("💾 Lưu cập nhật DS Báo Cáo Định Kỳ", type="primary"):
        st.session_state.bc_dinhky_df = reindex_df(edited_bc_df)
        st.success("Đã lưu cập nhật danh sách Báo Cáo Định Kỳ thành công!")
        st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel Báo Cáo Định Kỳ")

    col_bc1, col_bc2 = st.columns(2)
    with col_bc1:
        if st.button("📥 Xuất toàn bộ Excel Báo Cáo Định Kỳ", type="primary", use_container_width=True):
            try:
                with pd.ExcelWriter(EXCEL_PATH_BC_DINH_KY, engine='openpyxl') as writer:
                    st.session_state.bc_dinhky_df.to_excel(writer, sheet_name="BAO_CAO_DINH_KY", index=False)
                    auto_fit_columns(writer.book)
                st.success(f"Đã xuất file thành công tại:\n`{EXCEL_PATH_BC_DINH_KY}`")
            except Exception as e:
                st.error(f"Lỗi xuất file: {e}")

    with col_bc2:
        up_bc = st.file_uploader("Nhập file Excel Báo Cáo Định Kỳ để cập nhật:", type=["xlsx", "xls"], key="up_bc")
        if up_bc:
            try:
                df_bc_up = pd.read_excel(up_bc)
                st.session_state.bc_dinhky_df = reindex_df(df_bc_up)
                st.success("Đã cập nhật danh sách Báo Cáo Định Kỳ mới thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file Excel: {e}")

# MỤC 4: DS Gsheet_CV
elif main_menu == "4 🟢 DS Gsheet_CV":
    st.subheader("🟢 Bảng Danh Sách Google Sheets_CV")
    st.markdown("*(Sửa dữ liệu trực tiếp trong bảng -> Bấm **💾 Lưu cập nhật** bên dưới)*")

    edited_gsheet_df = st.data_editor(
        st.session_state.gsheet_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Mô tả Google Sheet", "Link Google Sheet", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Mô tả Google Sheet": st.column_config.TextColumn("Mô tả Google Sheet", width="large"),
            "Link Google Sheet": st.column_config.LinkColumn("Link Google Sheet", display_text="🔗 Mở Google Sheet", width="medium"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="medium")
        },
        key="editor_gsheet"
    )

    if st.button("💾 Lưu cập nhật DS Google Sheets", type="primary"):
        st.session_state.gsheet_df = reindex_df(edited_gsheet_df)
        st.success("Đã lưu cập nhật danh sách Google Sheets thành công!")
        st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel Google Sheets_CV")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        if st.button("📥 Xuất toàn bộ Excel Google Sheets_CV", type="primary", use_container_width=True):
            try:
                with pd.ExcelWriter(EXCEL_PATH_GSHEET, engine='openpyxl') as writer:
                    st.session_state.gsheet_df.to_excel(writer, sheet_name="GSHEETS", index=False)
                    auto_fit_columns(writer.book)
                st.success(f"Đã xuất file thành công tại:\n`{EXCEL_PATH_GSHEET}`")
            except Exception as e:
                st.error(f"Lỗi xuất file: {e}")

    with col_g2:
        up_g = st.file_uploader("Nhập file Excel Google Sheets để cập nhật:", type=["xlsx", "xls"], key="up_g")
        if up_g:
            try:
                df_g_up = pd.read_excel(up_g)
                st.session_state.gsheet_df = reindex_df(df_g_up)
                st.success("Đã cập nhật danh sách Google Sheets mới thành công!")
                st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file Excel: {e}")
