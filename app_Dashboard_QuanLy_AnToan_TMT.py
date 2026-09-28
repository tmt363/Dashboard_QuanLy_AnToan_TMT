import streamlit as st
import pandas as pd
import os
import io

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & CUSTOM CSS TỐI ƯU GIAO DIỆN CHUẨN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Quản lý an toàn TTM version 1.0 20260928",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Ép giao diện tràn viền tối đa */
    .main .block-container {
        max-width: 99% !important;
        padding: 0.5rem 0.8rem !important;
    }

    /* Mở rộng Sidebar */
    [data-testid="stSidebar"] {
        min-width: 250px !important;
        max-width: 270px !important;
        background-color: #f8f9fa;
        border-right: 1px solid #e9ecef;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding: 0.8rem 0.6rem !important;
    }

    /* Đảm bảo nút bấm Sidebar căn lề đẹp */
    [data-testid="stSidebar"] .stButton > button {
        text-align: left !important;
        justify-content: flex-start !important;
        padding-left: 12px !important;
    }

    /* Header Card tiêu đề */
    .header-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 10px 18px;
        border-radius: 6px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
        margin-bottom: 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .header-card h1 {
        color: #ffffff !important;
        font-size: 18px !important;
        font-weight: 700 !important;
        margin: 0 !important;
    }
    .version-badge {
        background-color: rgba(255,255,255,0.22);
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: 600;
        white-space: nowrap;
    }

    /* Nút bấm liên kết trực tiếp trong bảng */
    .btn-link-action {
        display: inline-block;
        background-color: #0d6efd;
        color: #ffffff !important;
        padding: 3px 8px;
        border-radius: 4px;
        text-decoration: none !important;
        font-size: 12px;
        font-weight: 600;
        text-align: center;
        white-space: nowrap;
        margin-right: 4px;
        margin-bottom: 2px;
    }
    .btn-link-action:hover {
        background-color: #0b5ed7;
    }

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
                <h3 style="color: #1e3c72; margin-bottom: 5px;">🛡️ QUẢN LÝ AN TOÀN TTM</h3>
                <span style="background: #e7f1ff; color: #0d6efd; padding: 4px 12px; border-radius: 12px; font-weight: 600; font-size: 12px;">Version 1.0 20260928</span>
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
                    st.error("❌ Mật khẩu hoặc Tên đăng nhập không chính xác!")
    st.stop()

# ---------------------------------------------------------
# 3. KHỞI TẠO ĐƯỜNG DẪN & DỮ LIỆU MẶC ĐỊNH
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
    {'STT': 1, 'Mô tả WEB': 'D-Office', 'Link 1': 'https://doffice.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Công văn / văn bản EVN'},
    {'STT': 2, 'Mô tả WEB': 'Công cụ web trực tuyến', 'Link 1': 'https://www.congcuweb.net/', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Hiệu chỉnh tên công văn / văn bản'},
    {'STT': 3, 'Mô tả WEB': 'QLAT SPC', 'Link 1': 'https://giamsatantoan.evnspc.vn/Home/Index', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Quản lý giám sát an toàn SPC'},
    {'STT': 4, 'Mô tả WEB': 'Lịch tuần', 'Link 1': 'https://lichtuan.evnspc.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Công ty Điện lực Tây Ninh'},
    {'STT': 5, 'Mô tả WEB': 'Hệ thống PMIS', 'Link 1': 'https://pmis.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Quản lý vận hành thiết bị & lưới điện'},
    {'STT': 6, 'Mô tả WEB': 'Tritm.la Dashboard 2026 DTTU ', 'Link 1': 'https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit?gid=964445540#gid=964445540', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Google sheet CV'},
    {'STT': 7, 'Mô tả WEB': 'Hệ thống Giám sát Thiên tai Việt Nam', 'Link 1': 'https://vndms.gov.vn/', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Cảnh báo và phòng chống thiên tai'},
    {'STT': 8, 'Mô tả WEB': 'Hệ thống HRMS', 'Link 1': 'https://hrms.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Quản lý lao động tiền lương'},
    {'STT': 9, 'Mô tả WEB': 'Hệ thống E-Learning', 'Link 1': 'https://elearning.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Huấn luyện an toàn & thi trực tuyến'},
    {'STT': 10, 'Mô tả WEB': 'Cổng Dịch vụ công Quốc gia', 'Link 1': 'https://dichvucong.gov.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Thực hiện thủ tục hành chính PCCC/ĐTXD'},
    {'STT': 11, 'Mô tả WEB': 'Cổng Thông tin Bộ Công Thương', 'Link 1': 'https://moit.gov.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Theo dõi văn bản quy phạm kỹ thuật'},
    {'STT': 12, 'Mô tả WEB': 'Cổng Báo cáo Phòng chống thiên tai', 'Link 1': 'https://pctt.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Cập nhật tình hình PCTT & TKCN'},
    {'STT': 13, 'Mô tả WEB': 'Hệ thống Quản lý Đầu tư Xây dựng (IMIS)', 'Link 1': 'https://imis.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Theo dõi an toàn dự án ĐTXD'},
    {'STT': 14, 'Mô tả WEB': 'Hệ thống Thông tin Báo cáo EVN', 'Link 1': 'https://baocao.evn.com.vn', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Tổng hợp chỉ tiêu an toàn - kỹ thuật'},
    {'STT': 15, 'Mô tả WEB': 'Lưu trữ Hồ sơ / Biểu mẫu TMT', 'Link 1': 'https://drive.google.com', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Kho lưu trữ dữ liệu dùng chung TMT'},
    {'STT': 16, 'Mô tả WEB': 'Thư viện Quy chuẩn - Quy định An toàn', 'Link 1': 'https://drive.google.com', 'Link 2': '', 'Link 3': '', 'Ghi chú': 'Tra cứu tài liệu an toàn PCCC & ĐT'}
]

def reindex_df(df):
    if not df.empty:
        df = df.reset_index(drop=True)
        df["STT"] = df.index + 1
        for col in ['Link 1', 'Link 2', 'Link 3']:
            if col not in df.columns:
                df[col] = ""
            df[col] = df[col].fillna("").astype(str)
    return df

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='DATA')
    return output.getvalue()

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
                    st.session_state.data_store[cat] = pd.DataFrame(columns=["STT", "Thư mục / Hồ sơ", "Link 1", "Link 2", "Link 3", "Ghi chú"])
        except Exception:
            pass

    if not st.session_state.data_store:
        for cat in CATEGORIES:
            st.session_state.data_store[cat] = reindex_df(pd.DataFrame([
                {"STT": 1, "Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật định kỳ"}
            ]))

if "bc_dinhky_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_BC_DINH_KY):
        try:
            st.session_state.bc_dinhky_df = reindex_df(pd.read_excel(EXCEL_PATH_BC_DINH_KY))
        except Exception:
            pass
    if "bc_dinhky_df" not in st.session_state:
        st.session_state.bc_dinhky_df = reindex_df(pd.DataFrame([
            {"STT": 1, "Tên Báo Cáo / Công Việc": "Báo cáo công tác An toàn định kỳ Quý", "Tần suất": "Hàng Quý", "Đơn vị nhận": "Công ty Điện lực", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Nộp trước ngày 20 cuối quý"},
            {"STT": 2, "Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Nộp trước ngày 25 hàng tháng"}
        ]))

if "gsheet_df" not in st.session_state:
    if os.path.exists(EXCEL_PATH_GSHEET):
        try:
            st.session_state.gsheet_df = reindex_df(pd.read_excel(EXCEL_PATH_GSHEET))
        except Exception:
            pass
    if "gsheet_df" not in st.session_state:
        st.session_state.gsheet_df = reindex_df(pd.DataFrame([
            {"STT": 1, "Mô tả Google Sheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Dùng chung phòng An Toàn"},
            {"STT": 2, "Mô tả Google Sheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật trực tuyến"}
        ]))

# HEADER HỆ THỐNG
st.markdown("""
    <div class="header-card">
        <h1>🛡️ Quản lý an toàn TTM</h1>
        <span class="version-badge">Version 1.0 20260928</span>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 4. SIDEBAR CHUYÊN NGHIỆP
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
    ("2 📋 DM QL Files_CV", "2 📋 DM QL Files_CV"),
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
# 5. DIALOGS CHỈNH SỬA / THÊM MỚI (HỖ TRỢ TỐI ĐA 3 LINKS)
# ---------------------------------------------------------
@st.dialog("✏️ Chỉnh sửa mục")
def edit_item_dialog(df_key, row_idx, title_col, cat_key=None):
    if cat_key:
        df = st.session_state.data_store[cat_key]
    else:
        df = getattr(st.session_state, df_key)

    item = df.iloc[row_idx]
    
    val_title = st.text_input(f"{title_col}:", value=item.get(title_col, ""))
    val_l1 = st.text_input("Link 1:", value=item.get("Link 1", ""))
    val_l2 = st.text_input("Link 2 (nếu có):", value=item.get("Link 2", ""))
    val_l3 = st.text_input("Link 3 (nếu có):", value=item.get("Link 3", ""))
    val_note = st.text_input("Ghi chú:", value=item.get("Ghi chú", ""))

    if st.button("💾 Cập nhật", type="primary", use_container_width=True):
        df.at[row_idx, title_col] = val_title
        df.at[row_idx, "Link 1"] = val_l1
        df.at[row_idx, "Link 2"] = val_l2
        df.at[row_idx, "Link 3"] = val_l3
        df.at[row_idx, "Ghi chú"] = val_note
        
        if cat_key:
            st.session_state.data_store[cat_key] = reindex_df(df)
        else:
            setattr(st.session_state, df_key, reindex_df(df))
        st.success("Đã cập nhật thành công!")
        st.rerun()

@st.dialog("➕ Thêm mới")
def add_item_dialog(df_key, title_col, cat_key=None):
    val_title = st.text_input(f"{title_col}:")
    val_l1 = st.text_input("Link 1:")
    val_l2 = st.text_input("Link 2 (nếu có):")
    val_l3 = st.text_input("Link 3 (nếu có):")
    val_note = st.text_input("Ghi chú:")

    if st.button("➕ Thêm mới", type="primary", use_container_width=True):
        if cat_key:
            df = st.session_state.data_store[cat_key]
        else:
            df = getattr(st.session_state, df_key)

        new_row = {
            "STT": len(df) + 1,
            title_col: val_title,
            "Link 1": val_l1,
            "Link 2": val_l2,
            "Link 3": val_l3,
            "Ghi chú": val_note
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        
        if cat_key:
            st.session_state.data_store[cat_key] = reindex_df(df)
        else:
            setattr(st.session_state, df_key, reindex_df(df))
        st.success("Đã thêm mới thành công!")
        st.rerun()

# ---------------------------------------------------------
# 6. HÀM HIỂN THỊ BẢNG VỚI NÚT THAO TÁC (CÂY VIẾT ✏️)
# ---------------------------------------------------------
def render_interactive_table(df, df_key, title_col, cat_key=None):
    # Nút thêm mới phía trên
    col_add, col_empty = st.columns([1.5, 4])
    with col_add:
        if st.button(f"➕ Thêm mới", key=f"add_btn_{df_key}_{cat_key}", type="primary"):
            add_item_dialog(df_key, title_col, cat_key)

    # Tiêu đề bảng
    col_stt, col_t, col_links, col_note, col_action = st.columns([0.6, 3, 2.5, 3, 1.2])
    with col_stt: st.markdown("**STT**")
    with col_t: st.markdown(f"**{title_col}**")
    with col_links: st.markdown("**Links truy cập**")
    with col_note: st.markdown("**Ghi chú**")
    with col_action: st.markdown("**Thao tác**")
    st.markdown("<hr style='margin: 4px 0 10px 0;'>", unsafe_allow_html=True)

    # Lặp qua từng dòng dữ liệu
    for idx, row in df.iterrows():
        c_stt, c_t, c_links, c_note, c_act = st.columns([0.6, 3, 2.5, 3, 1.2])
        
        with c_stt:
            st.write(f"**{row.get('STT', idx+1)}**")
            
        with c_t:
            st.write(row.get(title_col, ""))
            
        with c_links:
            # Tạo các nút Link 1, Link 2, Link 3 nếu có
            l1, l2, l3 = str(row.get("Link 1", "")), str(row.get("Link 2", "")), str(row.get("Link 3", ""))
            btns_html = ""
            if l1 and l1 != "nan":
                btns_html += f'<a class="btn-link-action" href="{l1}" target="_blank">Link 1</a>'
            if l2 and l2 != "nan":
                btns_html += f'<a class="btn-link-action" href="{l2}" target="_blank">Link 2</a>'
            if l3 and l3 != "nan":
                btns_html += f'<a class="btn-link-action" href="{l3}" target="_blank">Link 3</a>'
            st.markdown(btns_html if btns_html else "-", unsafe_allow_html=True)
            
        with c_note:
            st.write(row.get("Ghi chú", ""))
            
        with c_act:
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                # NÚT CÂY VIẾT ✏️ SỬA
                if st.button("✏️", key=f"edit_{df_key}_{cat_key}_{idx}", help="Chỉnh sửa dòng này"):
                    edit_item_dialog(df_key, idx, title_col, cat_key)
            with btn_col2:
                # NÚT THÙNG RÁC 🗑️ XÓA
                if st.button("🗑️", key=f"del_{df_key}_{cat_key}_{idx}", help="Xóa dòng này"):
                    if cat_key:
                        st.session_state.data_store[cat_key] = reindex_df(df.drop(idx))
                    else:
                        setattr(st.session_state, df_key, reindex_df(df.drop(idx)))
                    st.rerun()

        st.markdown("<hr style='margin: 4px 0;'>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 7. XUẤT NHẬP EXCEL TOOL
# ---------------------------------------------------------
def render_io_excel_tools(df, current_key, file_prefix):
    st.markdown("---")
    col_up, col_down = st.columns([1.2, 1])
    
    with col_up:
        st.markdown("##### 📥 Tải lên / Thay thế dữ liệu từ file Excel (.xlsx)")
        uploaded_file = st.file_uploader(f"Chọn file Excel [{file_prefix}]", type=["xlsx", "xls"], key=f"uploader_{current_key}")
        if uploaded_file is not None:
            try:
                new_df = pd.read_excel(uploaded_file)
                new_df = reindex_df(new_df)
                if st.button("🔥 Xác nhận đè dữ liệu mới", type="primary", key=f"btn_confirm_{current_key}"):
                    if current_key == "web":
                        st.session_state.web_tools_df = new_df
                    elif current_key == "bc":
                        st.session_state.bc_dinhky_df = new_df
                    elif current_key == "gsheet":
                        st.session_state.gsheet_df = new_df
                    else:
                        st.session_state.data_store[current_key] = new_df
                    st.balloons()
                    st.success("Tải dữ liệu từ Excel thành công!")
                    st.rerun()
            except Exception as e:
                st.error(f"Lỗi đọc file Excel: {e}")

    with col_down:
        st.markdown("##### 📤 Xuất dữ liệu ra file Excel")
        excel_bytes = to_excel_bytes(df)
        st.download_button(
            label="💾 Tải file Excel về máy",
            data=excel_bytes,
            file_name=f"{file_prefix}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ---------------------------------------------------------
# 8. MỤC HIỂN THỊ CHÍNH
# ---------------------------------------------------------
if main_menu == "1 🌐 DS WEBsites_CV":
    st.subheader("🌐 Bảng Danh Sách WEBsites_CV")
    
    render_interactive_table(
        st.session_state.web_tools_df, 
        df_key="web_tools_df", 
        title_col="Mô tả WEB"
    )

    col_rst1, col_rst2 = st.columns([3, 1])
    with col_rst2:
        if st.button("🔄 Khôi phục 16 Web mặc định", type="secondary"):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))
            st.toast("Đã khôi phục 16 Web mặc định!", icon="🎉")
            st.rerun()

    render_io_excel_tools(st.session_state.web_tools_df, "web", "1 DS WEBsites_CV out_20260928 macdinh")

elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.sidebar.selectbox("📂 Chọn mảng công việc:", CATEGORIES)
    if selected_cat:
        st.subheader(f"📂 Quản Lý Hồ Sơ: {selected_cat}")
        current_df = st.session_state.data_store[selected_cat]

        render_interactive_table(
            current_df, 
            df_key="data_store", 
            title_col="Thư mục / Hồ sơ",
            cat_key=selected_cat
        )

        render_io_excel_tools(current_df, selected_cat, f"HoSo_{selected_cat}")

elif main_menu == "3 📊 DS BCdinhky_CV":
    st.subheader("📊 Bảng Danh Sách Báo Cáo Định Kỳ & Công Việc")

    render_interactive_table(
        st.session_state.bc_dinhky_df, 
        df_key="bc_dinhky_df", 
        title_col="Tên Báo Cáo / Công Việc"
    )

    render_io_excel_tools(st.session_state.bc_dinhky_df, "bc", "DanhSach_BaoCao_DinhKy_TMT")

elif main_menu == "4 🟢 DS Gsheet_CV":
    st.subheader("🟢 Bảng Danh Sách Google Sheets_CV")

    render_interactive_table(
        st.session_state.gsheet_df, 
        df_key="gsheet_df", 
        title_col="Mô tả Google Sheet"
    )

    render_io_excel_tools(st.session_state.gsheet_df, "gsheet", "DanhSach_Gsheet_TMT")
