import streamlit as st
import pandas as pd
import os
import openpyxl

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & CUSTOM CSS (TỐI ƯU GIAO DIỆN & TỐI ĐA HÓA KHÔNG GIAN)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Quản Lý An Toàn TMT - Version 1.0 20260925",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS tinh chỉnh Sidebar gọn gàng & đổi nút Radio thành Nút Bấm Đẹp
st.markdown("""
    <style>
    /* 1. Ép vùng nội dung chính chiếm tối đa độ rộng màn hình (99%) */
    .main .block-container {
        max-width: 99% !important;
        padding-left: 0.5rem !important;
        padding-right: 0.5rem !important;
        padding-top: 0.5rem !important;
        padding-bottom: 0.5rem !important;
    }

    /* 2. Thu gọn tối đa độ rộng của Sidebar */
    [data-testid="stSidebar"] {
        min-width: 200px !important;
        max-width: 210px !important;
        background-color: #f8f9fa;
        border-right: 1px solid #e9ecef;
    }

    /* Giảm lề bên trong Sidebar */
    [data-testid="stSidebar"] > div:first-child {
        padding: 0.8rem 0.5rem !important;
    }

    /* 3. Header Card tiêu đề gọn nhẹ */
    .header-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 8px 15px;
        border-radius: 6px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        margin-bottom: 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .header-card h1 {
        color: #ffffff !important;
        font-size: 16px !important;
        font-weight: 700 !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    .version-badge {
        background-color: rgba(255,255,255,0.2);
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
    }

    /* 4. Tùy chỉnh Nút bấm Sidebar */
    .sidebar-menu-btn button {
        width: 100% !important;
        text-align: left !important;
        justify-content: flex-start !important;
        padding: 6px 10px !important;
        font-size: 13px !important;
        margin-bottom: 4px !important;
        border-radius: 6px !important;
    }

    /* Hiệu ứng đường dẫn thư mục gọn gàng */
    .path-box {
        background-color: #eef2f7;
        border: 1px dashed #cbd5e1;
        padding: 6px;
        border-radius: 5px;
        font-size: 10px;
        word-break: break-all;
        color: #475569;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 2. BẢO MẬT ĐĂNG NHẬP
# ---------------------------------------------------------
USER_CREDENTIALS = {
    "ttm": "123456",
    "admin": "123456"
}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1.2, 1.6, 1.2])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
            <div style="background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center;">
                <h3 style="color: #1e3c72; margin-bottom: 5px;">🛡️ HỆ THỐNG QUẢN LÝ AN TOÀN</h3>
                <span style="background: #e7f1ff; color: #0d6efd; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 11px;">Version 1.0 20260925</span>
                <hr style="margin: 12px 0;">
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("##### 🔐 Đăng Nhập Tài Khoản")
            user_input = st.text_input("Tên đăng nhập:", value="tmt")
            pass_input = st.text_input("Mật khẩu:", type="password")
            submit_login = st.form_submit_button("🔑 ĐĂNG NHẬP", type="primary", use_container_width=True)
            
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
# 3. KHỞI TẠO ĐƯỜNG DẪN & DỮ LIỆU
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

DEFAULT_WEBSITES = [
    {'STT': 1, 'Mô tả WEB': 'D-Office', 'Link truy cập': 'https://doffice.evn.com.vn', 'Ghi chú': 'Công văn / văn bản EVN'},
    {'STT': 2, 'Mô tả WEB': 'Công cụ web trực tuyến', 'Link truy cập': 'https://www.congcuweb.net/', 'Ghi chú': 'Hiệu chỉnh tên công văn / văn bản'},
    {'STT': 3, 'Mô tả WEB': 'QLAT SPC', 'Link truy cập': 'https://giamsatantoan.evnspc.vn/Home/Index', 'Ghi chú': 'Quản lý giám sát an toàn SPC'},
    {'STT': 4, 'Mô tả WEB': 'Lịch tuần', 'Link truy cập': 'https://lichtuan.evnspc.vn', 'Ghi chú': 'Công ty Điện lực Tây Ninh'},
    {'STT': 5, 'Mô tả WEB': 'Hệ thống PMIS', 'Link truy cập': 'https://pmis.evn.com.vn', 'Ghi chú': 'Quản lý vận hành thiết bị & lưới điện'},
    {'STT': 6, 'Mô tả WEB': 'Tritm.la Dashboard 2026 DTTU ', 'Link truy cập': 'https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit?gid=964445540#gid=964445540', 'Ghi chú': 'Google sheet CV'},
    {'STT': 7, 'Mô tả WEB': 'Hệ thống Giám sát Thiên tai Việt Nam', 'Link truy cập': 'https://vndms.gov.vn/', 'Ghi chú': 'Cảnh báo và phòng chống thiên tai'}
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

# ---------------------------------------------------------
# 4. KHO DỮ LIỆU SESSION STATE
# ---------------------------------------------------------
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "3 📊 DS BCdinhky_CV"

if "web_tools_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_WEB):
        try:
            st.session_state.web_tools_df = reindex_df(pd.read_excel(EXCEL_PATH_WEB))
        except Exception:
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))
    else:
        st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))

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

# HEADER CỦA HỆ THỐNG
st.markdown("""
    <div class="header-card">
        <h1>🛡️ Hệ Thống Quản Lý An Toàn & Công Tác Chuyên Môn TMT</h1>
        <span class="version-badge">Version 1.0</span>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. SIDEBAR MỚI (TỐI ƯU CÁC NÚT CHỌN & TIẾT KIỆM DIỆN TÍCH)
# ---------------------------------------------------------
st.sidebar.markdown(f"👤 **User:** `{st.session_state.username}`")

col_btn1, col_btn2 = st.sidebar.columns(2)
with col_btn1:
    if st.button("🚪 Thoát", use_container_width=True, type="secondary"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.rerun()

with col_btn2:
    if st.button("🧹 Cache", use_container_width=True, type="secondary"):
        st.cache_data.clear()
        st.toast("Đã xóa cache!", icon="🎉")

st.sidebar.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)
st.sidebar.markdown("**📁 MỤC LÀM VIỆC**")

# Danh sách menu các mục
menu_options = [
    ("1 🌐 DS WEBsites_CV", "1 🌐 DS WEBsites_CV"),
    ("2 📋 DM QL Files", "2 📋 DM QL Files"),
    ("3 📊 DS BCdinhky_CV", "3 📊 DS BCdinhky_CV"),
    ("4 🟢 DS Gsheet_CV", "4 🟢 DS Gsheet_CV")
]

# Tạo danh sách Nút Bấm thay thế Radio Button
for label, key_val in menu_options:
    is_active = (st.session_state.active_tab == key_val)
    btn_type = "primary" if is_active else "secondary"
    
    st.sidebar.markdown('<div class="sidebar-menu-btn">', unsafe_allow_html=True)
    if st.sidebar.button(label, key=f"menu_{key_val}", type=btn_type, use_container_width=True):
        st.session_state.active_tab = key_val
        st.rerun()
    st.sidebar.markdown('</div>', unsafe_allow_html=True)

main_menu = st.session_state.active_tab

st.sidebar.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
st.sidebar.markdown("**📌 Lưu trữ Excel:**")
st.sidebar.markdown(f'<div class="path-box">{EXCEL_DIR}</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 6. GIAO DIỆN HIỂN THỊ CHÍNH (ĐÃ CĂN CHỈNH ĐỘ RỘNG BẢNG)
# ---------------------------------------------------------
# MỤC 1: DS WEBsites_CV
if main_menu == "1 🌐 DS WEBsites_CV":
    st.subheader("🌐 Bảng Danh Sách WEBsites_CV")

    edited_web_df = st.data_editor(
        st.session_state.web_tools_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Mô tả WEB", "Link truy cập", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Mô tả WEB": st.column_config.TextColumn("Mô tả WEB", width="medium"),
            "Link truy cập": st.column_config.LinkColumn("Link truy cập", display_text="🔗 Truy cập Web", width="small"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="large")
        },
        key="editor_web"
    )

    col_save, col_reset = st.columns([2, 1])
    with col_save:
        if st.button("💾 Lưu Cập Nhật DS WEBsites_CV", type="primary", use_container_width=True):
            st.session_state.web_tools_df = reindex_df(edited_web_df)
            st.success("Đã lưu cập nhật danh sách WEBsites thành công!")
            st.rerun()
    with col_reset:
        if st.button("🔄 Khôi Phục Mặc Định", type="secondary", use_container_width=True):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))
            st.success("Đã khôi phục danh sách Web mặc định!")
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Dữ Liệu Excel")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        if st.button("📥 Xuất Toàn Bộ Excel WEBsites_CV", type="primary", use_container_width=True):
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
    selected_cat = st.sidebar.selectbox("📂 Chọn mảng công việc:", CATEGORIES)

    if selected_cat:
        st.subheader(f"📂 Quản Lý Hồ Sơ: {selected_cat}")

        current_df = st.session_state.data_store[selected_cat]

        edited_df = st.data_editor(
            current_df,
            num_rows="dynamic",
            use_container_width=True,
            column_order=["STT", "Thư mục / Hồ sơ", "Link xem", "Ghi chú"],
            column_config={
                "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
                "Thư mục / Hồ sơ": st.column_config.TextColumn("Thư mục / Hồ sơ", width="medium"),
                "Link xem": st.column_config.LinkColumn("Link xem", display_text="🔗 Mở xem", width="small"),
                "Ghi chú": st.column_config.TextColumn("Ghi chú", width="large")
            },
            key=f"editor_{selected_cat}"
        )

        if st.button("💾 Lưu Cập Nhật Mảng Công Việc", type="primary"):
            st.session_state.data_store[selected_cat] = reindex_df(edited_df)
            st.success(f"Đã lưu cập nhật cho **{selected_cat}**!")
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Dữ Liệu Excel")

    col_q1, col_q2 = st.columns(2)
    with col_q1:
        if st.button("📥 Xuất Toàn Bộ Excel DM QL Files", type="primary", use_container_width=True):
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

    edited_bc_df = st.data_editor(
        st.session_state.bc_dinhky_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Tên Báo Cáo / Công Việc", "Tần suất", "Đơn vị nhận", "Link biểu mẫu", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Tên Báo Cáo / Công Việc": st.column_config.TextColumn("Tên Báo Cáo / Công Việc", width="large"),
            "Tần suất": st.column_config.SelectboxColumn("Tần suất", options=["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"], width="small"),
            "Đơn vị nhận": st.column_config.TextColumn("Đơn vị nhận", width="medium"),
            "Link biểu mẫu": st.column_config.LinkColumn("Link biểu mẫu", display_text="🔗 Tải / Xem Biểu Mẫu", width="small"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="large")
        },
        key="editor_bc_dinhky"
    )

    if st.button("💾 Lưu Cập Nhật DS Báo Cáo Định Kỳ", type="primary"):
        st.session_state.bc_dinhky_df = reindex_df(edited_bc_df)
        st.success("Đã lưu cập nhật danh sách Báo Cáo Định Kỳ thành công!")
        st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Dữ Liệu Excel")

    col_bc1, col_bc2 = st.columns(2)
    with col_bc1:
        if st.button("📥 Xuất Toàn Bộ Excel Báo Cáo Định Kỳ", type="primary", use_container_width=True):
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

    edited_gsheet_df = st.data_editor(
        st.session_state.gsheet_df,
        num_rows="dynamic",
        use_container_width=True,
        column_order=["STT", "Mô tả Google Sheet", "Link Google Sheet", "Ghi chú"],
        column_config={
            "STT": st.column_config.NumberColumn("STT", format="%d", width="small"),
            "Mô tả Google Sheet": st.column_config.TextColumn("Mô tả Google Sheet", width="medium"),
            "Link Google Sheet": st.column_config.LinkColumn("Link Google Sheet", display_text="🔗 Mở Google Sheet", width="small"),
            "Ghi chú": st.column_config.TextColumn("Ghi chú", width="large")
        },
        key="editor_gsheet"
    )

    if st.button("💾 Lưu Cập Nhật DS Google Sheets", type="primary"):
        st.session_state.gsheet_df = reindex_df(edited_gsheet_df)
        st.success("Đã lưu cập nhật danh sách Google Sheets thành công!")
        st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Dữ Liệu Excel")

    col_g1, col_g2 = st.columns(2)
    with col_g1:
        if st.button("📥 Xuất Toàn Bộ Excel Google Sheets_CV", type="primary", use_container_width=True):
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
