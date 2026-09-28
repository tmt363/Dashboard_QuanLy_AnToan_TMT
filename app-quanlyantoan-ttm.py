import streamlit as st
import pandas as pd
import os
import io

# ---------------------------------------------------------
# 1. CẤU HÌNH TRANG & CUSTOM CSS
# ---------------------------------------------------------
st.set_page_config(
    page_title="Hệ Thống Quản Lý An Toàn TMT",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main .block-container {
        max-width: 99% !important;
        padding-top: 0.5rem !important;
        padding-bottom: 1rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
    }

    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0.8rem !important;
        padding-left: 0.8rem !important;
        padding-right: 0.8rem !important;
    }

    [data-testid="stSidebar"] {
        min-width: 260px !important;
        max-width: 280px !important;
        background-color: #f8f9fa;
        border-right: 1px solid #e9ecef;
    }

    .sticky-header-container {
        position: sticky;
        top: 2.8rem;
        z-index: 999;
        background-color: #ffffff;
        padding-top: 4px;
        padding-bottom: 6px;
        margin-bottom: 8px;
        border-bottom: 2px solid #cbd5e1;
    }

    .header-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        padding: 8px 16px;
        border-radius: 6px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        height: 42px;
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
        border-radius: 10px;
        font-size: 11px;
        font-weight: 600;
    }

    .page-subheading {
        margin-top: 6px !important;
        font-size: 17px !important;
        font-weight: 700 !important;
        color: #1e3c72 !important;
    }

    .path-box-inside {
        background-color: #f1f5f9;
        border: 1px dashed #cbd5e1;
        padding: 4px 6px;
        border-radius: 4px;
        font-size: 10px;
        word-break: break-all;
        color: #475569;
        margin-top: 4px;
    }

    .admin-box {
        background-color: #f8fafc;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
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
if "show_admin_panel" not in st.session_state:
    st.session_state.show_admin_panel = False
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "Light Mode"

# MÀN HÌNH ĐĂNG NHẬP (ĐÃ ĐỔI TÊN ĐÚNG VỚI BÊN TRONG)
if not st.session_state.logged_in:
    col1, col2, col3 = st.columns([1.2, 1.6, 1.2])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""
            <div style="background: white; padding: 25px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center;">
                <h3 style="color: #1e3c72; margin-bottom: 5px;">🛡️ HỆ THỐNG QUẢN LÝ AN TOÀN - TTM</h3>
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
                    st.error("❌ Mật khẩu hoặc Tên đăng nhập không chính xác!")
    st.stop()

# ---------------------------------------------------------
# 3. DỮ LIỆU & ĐƯỜNG DẪN
# ---------------------------------------------------------
EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(EXCEL_DIR):
    EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

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

# ĐỦ 9 WEBSITES MẶC ĐỊNH
DEFAULT_WEBSITES = [
    {'STT': 1, 'Mô tả WEB': 'D-Office', 'Link 1': 'https://doffice.evn.com.vn', 'Ghi chú': 'Công văn / văn bản EVN'},
    {'STT': 2, 'Mô tả WEB': 'Công cụ web trực tuyến', 'Link 1': 'https://www.congcuweb.net/', 'Ghi chú': 'Hiệu chỉnh tên công văn / văn bản'},
    {'STT': 3, 'Mô tả WEB': 'QLAT SPC', 'Link 1': 'https://giamsatantoan.evnspc.vn/Home/Index', 'Ghi chú': 'Quản lý giám sát an toàn SPC'},
    {'STT': 4, 'Mô tả WEB': 'Lịch tuần', 'Link 1': 'https://lichtuan.evnspc.vn', 'Ghi chú': 'Công ty Điện lực Tây Ninh'},
    {'STT': 5, 'Mô tả WEB': 'Hệ thống PMIS', 'Link 1': 'https://pmis.evn.com.vn', 'Ghi chú': 'Quản lý vận hành thiết bị & lưới điện'},
    {'STT': 6, 'Mô tả WEB': 'Tritm.la Dashboard 2026 DTTU', 'Link 1': 'https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit?gid=964445540#gid=964445540', 'Ghi chú': 'Google sheet CV'},
    {'STT': 7, 'Mô tả WEB': 'Hệ thống Giám sát Thiên tai Việt Nam', 'Link 1': 'https://vndms.gov.vn/', 'Ghi chú': 'Cảnh báo và phòng chống thiên tai'},
    {'STT': 8, 'Mô tả WEB': 'Hệ thống HRMS', 'Link 1': 'https://hrms.evn.com.vn', 'Ghi chú': 'Quản lý lao động tiền lương'},
    {'STT': 9, 'Mô tả WEB': 'Hệ thống E-Learning', 'Link 1': 'https://elearning.evn.com.vn', 'Ghi chú': 'Huấn luyện an toàn & thi trực tuyến'}
]

def reindex_df(df):
    if not df.empty:
        df = df.reset_index(drop=True)
        df["STT"] = df.index + 1
    return df

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

if "active_tab" not in st.session_state:
    st.session_state.active_tab = "1 🌐 DS WEBsites_CV"

if "web_tools_df" not in st.session_state:
    st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_WEBSITES))

if "data_store" not in st.session_state:
    st.session_state.data_store = {}
    for cat in CATEGORIES:
        st.session_state.data_store[cat] = pd.DataFrame([
            {"STT": 1, "Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link 1": "https://drive.google.com", "Ghi chú": "Cập nhật định kỳ"}
        ])

if "bc_dinhky_df" not in st.session_state:
    st.session_state.bc_dinhky_df = pd.DataFrame([
        {"STT": 1, "Tên Báo Cáo / Công Việc": "Báo cáo công tác An toàn định kỳ Quý", "Tần suất": "Hàng Quý", "Đơn vị nhận": "Công ty Điện lực", "Link 1": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 20 cuối quý"},
        {"STT": 2, "Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link 1": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 25 hàng tháng"}
    ])

if "gsheet_df" not in st.session_state:
    st.session_state.gsheet_df = pd.DataFrame([
        {"STT": 1, "Mô tả Google Sheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link 1": "https://docs.google.com/spreadsheets", "Ghi chú": "Dùng chung phòng An Toàn"},
        {"STT": 2, "Mô tả Google Sheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link 1": "https://docs.google.com/spreadsheets", "Ghi chú": "Cập nhật trực tuyến"}
    ])

# ---------------------------------------------------------
# 4. SIDEBAR (BỎ HOÀN TOÀN KHỐI USER THÔNG TIN)
# ---------------------------------------------------------
# Nút bật / tắt Quản lý tài khoản (Quản trị dữ liệu)
if st.sidebar.button("⚙️ Quản lý tài khoản", use_container_width=True, type="secondary"):
    st.session_state.show_admin_panel = not st.session_state.show_admin_panel

col_btn1, col_btn2 = st.sidebar.columns(2)
with col_btn1:
    if st.button("🚪 Thoát", use_container_width=True, type="secondary"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.show_admin_panel = False
        st.rerun()

with col_btn2:
    if st.button("🧹 Clear Cache", use_container_width=True, type="secondary"):
        st.cache_data.clear()
        st.toast("Đã xóa cache thành công!", icon="🎉")

st.sidebar.markdown("<hr style='margin: 10px 0 8px 0;'>", unsafe_allow_html=True)

# THƯ MỤC EXCEL
st.sidebar.markdown("<b>📌 Thư mục Excel:</b>", unsafe_allow_html=True)
st.sidebar.markdown(f'<div class="path-box-inside">{EXCEL_DIR}</div>', unsafe_allow_html=True)

st.sidebar.markdown("<hr style='margin: 10px 0 8px 0;'>", unsafe_allow_html=True)

# CHỌN GIAO DIỆN
st.sidebar.markdown("🎨 **CHỌN GIAO DIỆN (THEME)**")
selected_theme = st.sidebar.radio(
    "Theme mode:",
    ["☀️ Light Mode", "🌙 Dark Mode"],
    index=0 if st.session_state.theme_mode == "Light Mode" else 1,
    label_visibility="collapsed"
)
st.session_state.theme_mode = "Light Mode" if "Light" in selected_theme else "Dark Mode"

st.sidebar.markdown("<hr style='margin: 10px 0 8px 0;'>", unsafe_allow_html=True)

# MỤC LÀM VIỆC
st.sidebar.markdown("📁 **MỤC LÀM VIỆC**")
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

# ---------------------------------------------------------
# 5. BANNER CỐ ĐỊNH Ở ĐỈNH MÀN HÌNH
# ---------------------------------------------------------
tab_titles = {
    "1 🌐 DS WEBsites_CV": "Bảng Danh Sách WEBsites_CV",
    "2 📋 DM QL Files_CV": "Danh Mục Quản Lý Files_CV",
    "3 📊 DS BCdinhky_CV": "Bảng Danh Sách Báo Cáo Định Kỳ & Công Việc",
    "4 🟢 DS Gsheet_CV": "Bảng Danh Sách Google Sheets_CV"
}

st.markdown(f"""
    <div class="sticky-header-container">
        <div class="header-card">
            <h1>🛡️ Quản lý an toàn TTM</h1>
            <span class="version-badge">Version 1.0 20260928</span>
        </div>
        <div class="page-subheading">{tab_titles.get(main_menu, "")}</div>
    </div>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 6. DIALOG POP-UP CHỈNH SỬA
# ---------------------------------------------------------
@st.dialog("✏️ Chỉnh sửa mục")
def edit_item_dialog(df_ref, idx, item_type="gsheet", category_name=None):
    row = df_ref.loc[idx]
    
    existing_links = []
    for col in df_ref.columns:
        if "Link" in col and str(row[col]) != "nan" and str(row[col]).strip() != "":
            existing_links.append(str(row[col]))
            
    if not existing_links:
        existing_links = [""]

    if f"link_count_{item_type}_{idx}" not in st.session_state:
        st.session_state[f"link_count_{item_type}_{idx}"] = len(existing_links)

    if item_type == "web":
        title_val = st.text_input("Mô tả WEB:", value=str(row.get("Mô tả WEB", "")))
    elif item_type == "hoso":
        title_val = st.text_input("Thư mục / Hồ sơ:", value=str(row.get("Thư mục / Hồ sơ", "")))
    elif item_type == "bc":
        title_val = st.text_input("Tên Báo Cáo / Công Việc:", value=str(row.get("Tên Báo Cáo / Công Việc", "")))
        tan_suat = st.selectbox("Tần suất:", ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"], index=0)
        don_vi = st.text_input("Đơn vị nhận:", value=str(row.get("Đơn vị nhận", "")))
    else:
        title_val = st.text_input("Mô tả Google Sheet:", value=str(row.get("Mô tả Google Sheet", "")))

    new_links = []
    curr_count = st.session_state[f"link_count_{item_type}_{idx}"]
    for i in range(curr_count):
        init_val = existing_links[i] if i < len(existing_links) else ""
        link_label = f"Link {i+1}:" if i == 0 else f"Link {i+1} (nếu có):"
        link_val = st.text_input(link_label, value=init_val, key=f"inp_link_{item_type}_{idx}_{i}")
        new_links.append(link_val)

    if st.button("➕ Thêm link", type="secondary"):
        st.session_state[f"link_count_{item_type}_{idx}"] += 1
        st.rerun()

    ghichu_val = st.text_input("Ghi chú:", value=str(row.get("Ghi chú", "")))

    if st.button("💾 Cập nhật", type="primary", use_container_width=True):
        if item_type == "web":
            target_df = st.session_state.web_tools_df
            target_df.at[idx, "Mô tả WEB"] = title_val
        elif item_type == "hoso":
            target_df = st.session_state.data_store[category_name]
            target_df.at[idx, "Thư mục / Hồ sơ"] = title_val
        elif item_type == "bc":
            target_df = st.session_state.bc_dinhky_df
            target_df.at[idx, "Tên Báo Cáo / Công Việc"] = title_val
            target_df.at[idx, "Tần suất"] = tan_suat
            target_df.at[idx, "Đơn vị nhận"] = don_vi
        else:
            target_df = st.session_state.gsheet_df
            target_df.at[idx, "Mô tả Google Sheet"] = title_val

        for i, l_val in enumerate(new_links):
            target_df.at[idx, f"Link {i+1}"] = l_val
        target_df.at[idx, "Ghi chú"] = ghichu_val

        st.success("Đã cập nhật!")
        st.rerun()

# ---------------------------------------------------------
# 7. KHU VỰC QUẢN TRỊ DỮ LIỆU (CHỈ HIỆN KHI BẤM "QUẢN LÝ TÀI KHOẢN")
# ---------------------------------------------------------
def render_admin_tools(df, current_key, file_prefix, item_type="gsheet", category_name=None):
    st.markdown("""
        <div class="admin-box">
            <h4 style="color: #1e3c72; margin-top: 0;">⚙️ KHU VỰC QUẢN TRỊ DỮ LIỆU</h4>
    """, unsafe_allow_html=True)
    
    # NÚT THÊM DÒNG MỚI
    if st.button("➕ Thêm mới dòng", type="primary", key=f"btn_add_new_{item_type}_{category_name}"):
        new_idx = len(df)
        if item_type == "web":
            st.session_state.web_tools_df.loc[new_idx] = {"STT": new_idx+1, "Mô tả WEB": "Mô tả mới", "Link 1": "", "Ghi chú": ""}
        elif item_type == "hoso":
            st.session_state.data_store[category_name].loc[new_idx] = {"STT": new_idx+1, "Thư mục / Hồ sơ": "Hồ sơ mới", "Link 1": "", "Ghi chú": ""}
        elif item_type == "bc":
            st.session_state.bc_dinhky_df.loc[new_idx] = {"STT": new_idx+1, "Tên Báo Cáo / Công Việc": "Báo cáo mới", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "", "Link 1": "", "Ghi chú": ""}
        else:
            st.session_state.gsheet_df.loc[new_idx] = {"STT": new_idx+1, "Mô tả Google Sheet": "Sheet mới", "Link 1": "", "Ghi chú": ""}
        st.rerun()

    st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
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

    st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# 8. HÀM HIỂN THỊ BẢNG DỮ LIỆU CHÍNH
# ---------------------------------------------------------
def render_data_table_with_actions(df, title_col, item_type="gsheet", category_name=None):
    if item_type == "bc":
        headers = ["STT", title_col, "Tần suất", "Đơn vị nhận", "Links truy cập", "Ghi chú", "Thao tác"]
        cols_width = [1, 3, 2, 2, 2, 3, 2]
    else:
        headers = ["STT", title_col, "Links truy cập", "Ghi chú", "Thao tác"]
        cols_width = [1, 4, 2, 3, 2]

    cols = st.columns(cols_width)
    for i, h in enumerate(headers):
        cols[i].markdown(f"**{h}**")
    st.markdown("<hr style='margin: 4px 0 10px 0;'>", unsafe_allow_html=True)

    for idx, row in df.iterrows():
        c = st.columns(cols_width)
        c[0].write(f"**{row.get('STT', idx+1)}**")
        c[1].write(str(row.get(title_col, "")))
        
        col_offset = 2
        if item_type == "bc":
            c[2].write(str(row.get("Tần suất", "")))
            c[3].write(str(row.get("Đơn vị nhận", "")))
            col_offset = 4

        link_markdowns = []
        for col_name in df.columns:
            if "Link" in col_name and str(row[col_name]) != "nan" and str(row[col_name]).strip() != "":
                l_url = str(row[col_name])
                link_markdowns.append(f"[{col_name}]({l_url})")
        
        c[col_offset].markdown(" | ".join(link_markdowns) if link_markdowns else "-")
        c[col_offset+1].write(str(row.get("Ghi chú", "")))

        btn_e, btn_d = c[col_offset+2].columns(2)
        if btn_e.button("✏️", key=f"btn_edit_{item_type}_{category_name}_{idx}"):
            edit_item_dialog(df, idx, item_type, category_name)
        if btn_d.button("🗑️", key=f"btn_del_{item_type}_{category_name}_{idx}"):
            if item_type == "web":
                st.session_state.web_tools_df = reindex_df(df.drop(idx))
            elif item_type == "hoso":
                st.session_state.data_store[category_name] = reindex_df(df.drop(idx))
            elif item_type == "bc":
                st.session_state.bc_dinhky_df = reindex_df(df.drop(idx))
            else:
                st.session_state.gsheet_df = reindex_df(df.drop(idx))
            st.rerun()

# ---------------------------------------------------------
# 9. ĐIỀU HƯỚNG MỤC LÀM VIỆC
# ---------------------------------------------------------
if main_menu == "1 🌐 DS WEBsites_CV":
    if st.session_state.show_admin_panel:
        render_admin_tools(st.session_state.web_tools_df, "web", "DanhMuc_CongCu_WEB_TMT", item_type="web")
    render_data_table_with_actions(st.session_state.web_tools_df, title_col="Mô tả WEB", item_type="web")

elif main_menu == "2 📋 DM QL Files_CV":
    st.markdown("##### 📁 Chọn mảng công việc:")
    selected_cat = st.selectbox("Mảng công việc:", CATEGORIES, index=0, label_visibility="collapsed")
    current_df = st.session_state.data_store[selected_cat]
    
    if st.session_state.show_admin_panel:
        render_admin_tools(current_df, selected_cat, f"HoSo_{selected_cat}", item_type="hoso", category_name=selected_cat)
    render_data_table_with_actions(current_df, title_col="Thư mục / Hồ sơ", item_type="hoso", category_name=selected_cat)

elif main_menu == "3 📊 DS BCdinhky_CV":
    if st.session_state.show_admin_panel:
        render_admin_tools(st.session_state.bc_dinhky_df, "bc", "DanhSach_BaoCao_DinhKy_TMT", item_type="bc")
    render_data_table_with_actions(st.session_state.bc_dinhky_df, title_col="Tên Báo Cáo / Công Việc", item_type="bc")

elif main_menu == "4 🟢 DS Gsheet_CV":
    if st.session_state.show_admin_panel:
        render_admin_tools(st.session_state.gsheet_df, "gsheet", "DanhSach_Gsheet_TMT", item_type="gsheet")
    render_data_table_with_actions(st.session_state.gsheet_df, title_col="Mô tả Google Sheet", item_type="gsheet")
