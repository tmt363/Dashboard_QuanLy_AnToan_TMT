import streamlit as st
import pandas as pd
import os
import openpyxl
from streamlit_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & CUSTOM CSS (TỐI ƯU TOÀN MÀN HÌNH)
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Quản Lý An Toàn TMT",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Ép giao diện chính chiếm 99% độ rộng màn hình */
    .main .block-container {
        max-width: 99% !important;
        padding: 0.5rem 0.8rem !important;
    }

    /* Thu gọn Sidebar tối đa */
    [data-testid="stSidebar"] {
        min-width: 200px !important;
        max-width: 210px !important;
        background-color: #f8f9fa;
        border-right: 1px solid #e9ecef;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding: 0.8rem 0.5rem !important;
    }

    /* Header Card sang trọng */
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
    }
    .version-badge {
        background-color: rgba(255,255,255,0.2);
        padding: 2px 8px;
        border-radius: 12px;
        font-size: 11px;
        font-weight: 600;
    }

    /* CSS cho Nút Bấm trong AG-Grid */
    .ag-link-btn {
        background-color: #0d6efd;
        color: white !important;
        padding: 3px 10px;
        border-radius: 4px;
        text-decoration: none !important;
        font-size: 12px;
        font-weight: 600;
        display: inline-block;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .ag-link-btn:hover {
        background-color: #0b5ed7;
        color: white !important;
    }

    /* Path Box trong Sidebar */
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
USER_CREDENTIALS = {"tmt": "123456", "admin": "123456"}

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
                <span style="background: #e7f1ff; color: #0d6efd; padding: 3px 10px; border-radius: 12px; font-weight: 600; font-size: 11px;">Version 1.0</span>
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
                    st.error("❌ Mật khẩu hoặc Tên đăng nhập không đúng!")
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
    {'STT': 7, 'Mô tả WEB': 'Hệ thống Giám sát Thiên tai Việt Nam', 'Link truy cập': 'https://vndms.gov.vn/', 'Ghi chú': 'Cảnh báo và phòng chống thiên tai'},
    {'STT': 8, 'Mô tả WEB': 'Hệ thống HRMS', 'Link truy cập': 'https://hrms.evn.com.vn', 'Ghi chú': 'Quản lý lao động tiền lương'},
    {'STT': 9, 'Mô tả WEB': 'Hệ thống E-Learning', 'Link truy cập': 'https://elearning.evn.com.vn', 'Ghi chú': 'Huấn luyện an toàn & thi trực tuyến'},
    {'STT': 10, 'Mô tả WEB': 'Cổng Dịch vụ công Quốc gia', 'Link truy cập': 'https://dichvucong.gov.vn', 'Ghi chú': 'Thực hiện thủ tục hành chính PCCC/ĐTXD'},
    {'STT': 11, 'Mô tả WEB': 'Cổng Thông tin Bộ Công Thương', 'Link truy cập': 'https://moit.gov.vn', 'Ghi chú': 'Theo dõi văn bản quy phạm kỹ thuật'},
    {'STT': 12, 'Mô tả WEB': 'Cổng Báo cáo Phòng chống thiên tai', 'Link truy cập': 'https://pctt.evn.com.vn', 'Ghi chú': 'Cập nhật tình hình PCTT & TKCN'},
    {'STT': 13, 'Mô tả WEB': 'Hệ thống Quản lý Đầu tư Xây dựng (IMIS)', 'Link truy cập': 'https://imis.evn.com.vn', 'Ghi chú': 'Theo dõi an toàn dự án ĐTXD'},
    {'STT': 14, 'Mô tả WEB': 'Hệ thống Thông tin Báo cáo EVN', 'Link truy cập': 'https://baocao.evn.com.vn', 'Ghi chú': 'Tổng hợp chỉ tiêu an toàn - kỹ thuật'},
    {'STT': 15, 'Mô tả WEB': 'Lưu trữ Hồ sơ / Biểu mẫu TMT', 'Link truy cập': 'https://drive.google.com', 'Ghi chú': 'Kho lưu trữ dữ liệu dùng chung TMT'},
    {'STT': 16, 'Mô tả WEB': 'Thư viện Quy chuẩn - Quy định An toàn', 'Link truy cập': 'https://drive.google.com', 'Ghi chú': 'Tra cứu tài liệu an toàn PCCC & ĐT'}
]

def reindex_df(df):
    if not df.empty:
        df = df.reset_index(drop=True)
        df["STT"] = df.index + 1
    return df

# Khởi tạo dữ liệu
if "active_tab" not in st.session_state:
    st.session_state.active_tab = "1 🌐 DS WEBsites_CV"

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

# HEADER HỆ THỐNG
st.markdown("""
    <div class="header-card">
        <h1>🛡️ Hệ Thống Quản Lý An Toàn & Công Tác Chuyên Môn TMT</h1>
        <span class="version-badge">Version 1.0</span>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 4. SIDEBAR ĐAN TRANG MỚI GỌN GÀNG
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

menu_options = [
    ("1 🌐 DS WEBsites_CV", "1 🌐 DS WEBsites_CV"),
    ("2 📋 DM QL Files", "2 📋 DM QL Files"),
    ("3 📊 DS BCdinhky_CV", "3 📊 DS BCdinhky_CV"),
    ("4 🟢 DS Gsheet_CV", "4 🟢 DS Gsheet_CV")
]

for label, key_val in menu_options:
    is_active = (st.session_state.active_tab == key_val)
    btn_type = "primary" if is_active else "secondary"
    if st.sidebar.button(label, key=f"menu_{key_val}", type=btn_type, use_container_width=True):
        st.session_state.active_tab = key_val
        st.rerun()

main_menu = st.session_state.active_tab

st.sidebar.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
st.sidebar.markdown("**📌 Thư mục Excel:**")
st.sidebar.markdown(f'<div class="path-box">{EXCEL_DIR}</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 5. HÀM HIỂN THỊ BẢNG AG-GRID TỐI ƯU KÍCH THƯỚC CỘT CHUẨN
# ---------------------------------------------------------
def render_custom_aggrid(df, text_col, link_col, btn_text="🔗 Truy cập Web"):
    gb = GridOptionsBuilder.from_dataframe(df)
    
    # Cấu hình Cột STT: Thu gọn tuyệt đối (60px)
    gb.configure_column("STT", headerName="STT", width=60, pinned="left", type=["numericColumn"], cellStyle={'textAlign': 'center'})
    
    # Cấu hình Cột Nội Dung/Mô Tả: Vừa vặn (260px)
    gb.configure_column(text_col, headerName=text_col, width=260, editable=True)
    
    # Cấu hình Cột Link: Thu gọn vừa nút bấm (140px)
    link_renderer = JsCode(f"""
        function(params) {{
            if (!params.value) return '';
            return `<a class="ag-link-btn" href="${{params.value}}" target="_blank">{btn_text}</a>`;
        }}
    """)
    gb.configure_column(link_col, headerName=link_col, width=140, cellRenderer=link_renderer, editable=True)
    
    # Cấu hình Cột Ghi Chú: Mở rộng chiếm hết toàn bộ màn hình còn lại (Flex = 1)
    gb.configure_column("Ghi chú", headerName="Ghi chú", flex=1, minWidth=300, editable=True)
    
    # Bật tính năng chỉnh sửa
    gb.configure_default_column(resizable=True, filter=True)
    grid_options = gb.build()
    
    grid_response = AgGrid(
        df,
        gridOptions=grid_options,
        allow_unsafe_jscode=True,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
        theme="alpine",
        height=450,
        fit_columns_on_grid_load=False
    )
    return grid_response['data']

# ---------------------------------------------------------
# 6. HIỂN THỊ CÁC MỤC VỚI AG-GRID
# ---------------------------------------------------------
if main_menu == "1 🌐 DS WEBsites_CV":
    st.subheader("🌐 Bảng Danh Sách WEBsites_CV")
    st.caption("💡 *Mẹo: Anh có thể nhấp đôi trực tiếp vào ô để sửa dữ liệu.*")

    updated_df = render_custom_aggrid(
        st.session_state.web_tools_df, 
        text_col="Mô tả WEB", 
        link_col="Link truy cập",
        btn_text="🔗 Truy cập Web"
    )

    col_save, col_reset = st.columns([2, 1])
    with col_save:
        if st.button("💾 Lưu Cập Nhật DS WEBsites_CV", type="primary", use_container_width=True):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(updated_df))
            st.success("Đã lưu cập nhật thành công!")
            st.rerun()
    with col_reset:
        if st.button("🔄 Khôi Phục Mặc Định", type="secondary", use_container_width=True):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))
            st.success("Đã khôi phục danh sách mặc định!")
            st.rerun()

elif main_menu == "2 📋 DM QL Files":
    selected_cat = st.sidebar.selectbox("📂 Chọn mảng công việc:", CATEGORIES)
    if selected_cat:
        st.subheader(f"📂 Quản Lý Hồ Sơ: {selected_cat}")
        current_df = st.session_state.data_store[selected_cat]

        updated_df = render_custom_aggrid(
            current_df, 
            text_col="Thư mục / Hồ sơ", 
            link_col="Link xem",
            btn_text="🔗 Mở xem"
        )

        if st.button("💾 Lưu Cập Nhật Mảng Công Việc", type="primary"):
            st.session_state.data_store[selected_cat] = reindex_df(pd.DataFrame(updated_df))
            st.success(f"Đã lưu cập nhật cho **{selected_cat}**!")
            st.rerun()

elif main_menu == "3 📊 DS BCdinhky_CV":
    st.subheader("📊 Bảng Danh Sách Báo Cáo Định Kỳ & Công Việc")

    updated_df = render_custom_aggrid(
        st.session_state.bc_dinhky_df, 
        text_col="Tên Báo Cáo / Công Việc", 
        link_col="Link biểu mẫu",
        btn_text="🔗 Tải Biểu Mẫu"
    )

    if st.button("💾 Lưu Cập Nhật DS Báo Cáo Định Kỳ", type="primary"):
        st.session_state.bc_dinhky_df = reindex_df(pd.DataFrame(updated_df))
        st.success("Đã lưu cập nhật Báo Cáo Định Kỳ thành công!")
        st.rerun()

elif main_menu == "4 🟢 DS Gsheet_CV":
    st.subheader("🟢 Bảng Danh Sách Google Sheets_CV")

    updated_df = render_custom_aggrid(
        st.session_state.gsheet_df, 
        text_col="Mô tả Google Sheet", 
        link_col="Link Google Sheet",
        btn_text="🔗 Mở GSheet"
    )

    if st.button("💾 Lưu Cập Nhật DS Google Sheets", type="primary"):
        st.session_state.gsheet_df = reindex_df(pd.DataFrame(updated_df))
        st.success("Đã lưu cập nhật Google Sheets thành công!")
        st.rerun()
