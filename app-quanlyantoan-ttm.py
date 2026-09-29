import streamlit as st
import pandas as pd
import os
import openpyxl
from datetime import datetime
import io 

# ==========================================
# 1. CẤU HÌNH TRANG VÀ SESSION STATE
# ==========================================
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Thư mục mặc định ban đầu
DEFAULT_EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(DEFAULT_EXCEL_DIR):
    DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

# Khởi tạo biến lưu trạng thái Giao diện, Admin và Thư mục
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False
if "excel_dir" not in st.session_state:
    st.session_state.excel_dir = DEFAULT_EXCEL_DIR

# Xử lý CSS Giao diện Mượt mà & Sáng/Tối
css_style = """
    <style>
    .st-emotion-cache-1y4p8pa { padding-top: 2rem; }
    .table-header { font-weight: bold; color: #1E88E5; }
    .row-text { font-size: 14px; }
    </style>
"""
if st.session_state.get("dark_mode"):
    css_style += """
        <style>
        .stApp { background-color: #1E1E1E !important; color: #FFFFFF !important; }
        .stSidebar { background-color: #2D2D2D !important; }
        .table-header { color: #64B5F6 !important; }
        h1, h2, h3, h4, h5, h6, p, span, div, strong { color: #E0E0E0 !important; }
        </style>
    """
st.markdown(css_style, unsafe_allow_html=True)


# ==========================================
# 2. BIẾN VÀ HÀM HỖ TRỢ
# ==========================================
DATE_STR = datetime.now().strftime("%Y%m%d")

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

DEFAULT_14_WEBS = [
    {"Mô tả WEB": "Hệ thống D-Office", "Link 1": "https://doffice.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý văn bản điều hành"},
    {"Mô tả WEB": "Cổng thông tin Điện lực (EVN SPC)", "Link 1": "https://evnspc.vn", "Link 2": "https://www.congcuweb.net/", "Link 3": "", "Ghi chú": "Tra cứu quy định & chỉ đạo"},
    {"Mô tả WEB": "Quản lý An toàn (ECP / ATLD)", "Link 1": "https://giamsatantoan.evnspc.vn/Home/Index", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý giám sát an toàn SPC"},
    {"Mô tả WEB": "Lịch công tác / Lịch tuần", "Link 1": "https://lichtuan.evnspc.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Công ty Điện lực Tây Ninh"},
    {"Mô tả WEB": "Hệ thống PMIS", "Link 1": "https://pmis.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý vận hành thiết bị & lưới điện"},
    {"Mô tả WEB": "Tritm.la Dashboard 2026 DTTU", "Link 1": "https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit?gid=964445540#gid=964445540", "Link 2": "", "Link 3": "", "Ghi chú": "Google sheet CV"},
    {"Mô tả WEB": "Hệ thống Giám sát Thiên tai Việt Nam", "Link 1": "https://vndms.gov.vn/", "Link 2": "", "Link 3": "", "Ghi chú": "Cảnh báo và phòng chống thiên tai"},
    {"Mô tả WEB": "Hệ thống HRMS", "Link 1": "https://hrms.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý lao động tiền lương"},
    {"Mô tả WEB": "Hệ thống E-Learning", "Link 1": "https://elearning.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Huấn luyện an toàn & thi trực tuyến"},
    {"Mô tả WEB": "Cổng Dịch vụ công Quốc gia", "Link 1": "https://dichvucong.gov.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Thực hiện thủ tục hành chính PCCC/ĐTXD"},
    {"Mô tả WEB": "Cổng Thông tin Bộ Công Thương", "Link 1": "https://moit.gov.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Theo dõi văn bản quy phạm kỹ thuật"},
    {"Mô tả WEB": "Cổng Báo cáo Phòng chống thiên tai", "Link 1": "https://pctt.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật tình hình PCTT & TKCN"},
    {"Mô tả WEB": "Hệ thống Quản lý Đầu tư Xây dựng (IMIS)", "Link 1": "https://imis.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Theo dõi an toàn dự án ĐTXD"},
    {"Mô tả WEB": "Hệ thống Thông tin Báo cáo EVN", "Link 1": "https://baocao.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Tổng hợp chỉ tiêu an toàn - kỹ thuật"}
]

def reindex_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.dropna(how='all').reset_index(drop=True)
    if "STT" in df.columns:
        df = df.drop(columns=["STT"])
    df.insert(0, "STT", range(1, len(df) + 1))
    
    for col in ["Link 1", "Link 2", "Link 3"]:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("")
    return df

def render_links_cell(l1, l2, l3):
    links = []
    if l1 and str(l1).strip().startswith("http"):
        links.append(f"[Link 1]({str(l1).strip()})")
    if l2 and str(l2).strip().startswith("http"):
        links.append(f"[Link 2]({str(l2).strip()})")
    if l3 and str(l3).strip().startswith("http"):
        links.append(f"[Link 3]({str(l3).strip()})")
    return " | ".join(links) if links else "-"

def auto_fit_columns(workbook):
    for sheetname in workbook.sheetnames:
        worksheet = workbook[sheetname]
        for col in worksheet.columns:
            max_len = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            adjusted_width = max(max_len + 4, 12)
            worksheet.column_dimensions[col_letter].width = min(adjusted_width, 60)

def generate_excel_download(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        auto_fit_columns(writer.book)
    return output.getvalue()

# Khởi tạo dữ liệu
if "web_tools_df" not in st.session_state:
    st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
if "data_store" not in st.session_state:
    st.session_state.data_store = {}
    for cat in CATEGORIES:
        st.session_state.data_store[cat] = reindex_df(pd.DataFrame([
            {"Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật định kỳ"}
        ]))
if "bc_dinhky_df" not in st.session_state:
    st.session_state.bc_dinhky_df = reindex_df(pd.DataFrame([
        {"Tên Báo Cáo / Công Việc": "Báo cáo công tác An toàn định kỳ Quý", "Tần suất": "Hàng Quý", "Đơn vị nhận": "Công ty Điện lực", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Nộp trước ngày 20 cuối quý"},
        {"Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Nộp trước ngày 25 hàng tháng"}
    ]))
if "gsheet_df" not in st.session_state:
    st.session_state.gsheet_df = reindex_df(pd.DataFrame([
        {"Mô tả Gsheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Dùng chung phòng An Toàn"},
        {"Mô tả Gsheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật trực tuyến"}
    ]))


# ==========================================
# 3. DIALOGS (HỘP THOẠI)
# ==========================================
@st.dialog("🔐 Đăng Nhập Quản Trị Viên")
def admin_login_dialog():
    st.write("Vui lòng nhập mật khẩu để kích hoạt các tính năng thêm/sửa/xóa.")
    pwd = st.text_input("Mật khẩu:", type="password")
    if st.button("Xác nhận", type="primary", use_container_width=True):
        if pwd == "admin123":
            st.session_state.is_admin = True
            st.success("Đăng nhập thành công!")
            st.rerun()
        else:
            st.error("Mật khẩu không chính xác!")

@st.dialog("➕ Thêm mới Website_CV")
def add_web_dialog():
    mota = st.text_input("Mô tả WEB (*):")
    l1 = st.text_input("Link 1 (*):")
    l2 = st.text_input("Link 2 (bổ sung):")
    l3 = st.text_input("Link 3 (bổ sung):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if mota:
            new_row = pd.DataFrame([{"Mô tả WEB": mota, "Link 1": l1, "Link 2": l2, "Link 3": l3, "Ghi chú": ghichu}])
            st.session_state.web_tools_df = reindex_df(pd.concat([st.session_state.web_tools_df, new_row], ignore_index=True))
            st.success("Đã thêm thành công!")
            st.rerun()
        else:
            st.warning("Vui lòng nhập mô tả WEB!")

@st.dialog("✏️ Chỉnh sửa Website_CV")
def edit_web_dialog(idx):
    df = st.session_state.web_tools_df
    row = df.loc[idx]
    mota = st.text_input("Mô tả WEB:", value=row["Mô tả WEB"])
    l1 = st.text_input("Link 1:", value=row["Link 1"])
    l2 = st.text_input("Link 2:", value=row["Link 2"])
    l3 = st.text_input("Link 3:", value=row["Link 3"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.web_tools_df.loc[idx, ["Mô tả WEB", "Link 1", "Link 2", "Link 3", "Ghi chú"]] = [mota, l1, l2, l3, ghichu]
        st.session_state.web_tools_df = reindex_df(st.session_state.web_tools_df)
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Gsheet_CV")
def add_gsheet_dialog():
    mota = st.text_input("Mô tả Gsheet (*):")
    l1 = st.text_input("Link 1 (*):")
    l2 = st.text_input("Link 2 (bổ sung):")
    l3 = st.text_input("Link 3 (bổ sung):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if mota:
            new_row = pd.DataFrame([{"Mô tả Gsheet": mota, "Link 1": l1, "Link 2": l2, "Link 3": l3, "Ghi chú": ghichu}])
            st.session_state.gsheet_df = reindex_df(pd.concat([st.session_state.gsheet_df, new_row], ignore_index=True))
            st.success("Đã thêm Gsheet thành công!")
            st.rerun()
        else:
            st.warning("Vui lòng nhập mô tả Gsheet!")

@st.dialog("✏️ Chỉnh sửa Gsheet_CV")
def edit_gsheet_dialog(idx):
    df = st.session_state.gsheet_df
    row = df.loc[idx]
    mota = st.text_input("Mô tả Gsheet:", value=row["Mô tả Gsheet"])
    l1 = st.text_input("Link 1:", value=row["Link 1"])
    l2 = st.text_input("Link 2:", value=row["Link 2"])
    l3 = st.text_input("Link 3:", value=row["Link 3"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.gsheet_df.loc[idx, ["Mô tả Gsheet", "Link 1", "Link 2", "Link 3", "Ghi chú"]] = [mota, l1, l2, l3, ghichu]
        st.session_state.gsheet_df = reindex_df(st.session_state.gsheet_df)
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Hồ Sơ / Thư Mục")
def add_file_dialog(selected_cat):
    hoso = st.text_input("Tên Thư mục / Hồ sơ (*):")
    l1 = st.text_input("Link 1 (*):")
    l2 = st.text_input("Link 2 (bổ sung):")
    l3 = st.text_input("Link 3 (bổ sung):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if hoso:
            new_row = pd.DataFrame([{"Thư mục / Hồ sơ": hoso, "Link 1": l1, "Link 2": l2, "Link 3": l3, "Ghi chú": ghichu}])
            st.session_state.data_store[selected_cat] = reindex_df(pd.concat([st.session_state.data_store[selected_cat], new_row], ignore_index=True))
            st.success("Đã thêm hồ sơ!")
            st.rerun()
        else:
            st.warning("Vui lòng nhập tên hồ sơ!")

@st.dialog("✏️ Chỉnh sửa Hồ Sơ / Thư Mục")
def edit_file_dialog(selected_cat, idx):
    df = st.session_state.data_store[selected_cat]
    row = df.loc[idx]
    hoso = st.text_input("Tên Thư mục / Hồ sơ:", value=row["Thư mục / Hồ sơ"])
    l1 = st.text_input("Link 1:", value=row["Link 1"])
    l2 = st.text_input("Link 2:", value=row["Link 2"])
    l3 = st.text_input("Link 3:", value=row["Link 3"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.data_store[selected_cat].loc[idx, ["Thư mục / Hồ sơ", "Link 1", "Link 2", "Link 3", "Ghi chú"]] = [hoso, l1, l2, l3, ghichu]
        st.session_state.data_store[selected_cat] = reindex_df(st.session_state.data_store[selected_cat])
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Báo Cáo / Công Việc")
def add_bc_dialog():
    ten_bc = st.text_input("Tên Báo Cáo / Công Việc (*):")
    tan_suat = st.selectbox("Tần suất:", ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"])
    don_vi = st.text_input("Đơn vị nhận:")
    l1 = st.text_input("Link 1 (*):")
    l2 = st.text_input("Link 2 (bổ sung):")
    l3 = st.text_input("Link 3 (bổ sung):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if ten_bc:
            new_row = pd.DataFrame([{"Tên Báo Cáo / Công Việc": ten_bc, "Tần suất": tan_suat, "Đơn vị nhận": don_vi, "Link 1": l1, "Link 2": l2, "Link 3": l3, "Ghi chú": ghichu}])
            st.session_state.bc_dinhky_df = reindex_df(pd.concat([st.session_state.bc_dinhky_df, new_row], ignore_index=True))
            st.success("Đã thêm báo cáo!")
            st.rerun()
        else:
            st.warning("Vui lòng nhập tên báo cáo!")

@st.dialog("✏️ Chỉnh sửa Báo Cáo / Công Việc")
def edit_bc_dialog(idx):
    df = st.session_state.bc_dinhky_df
    row = df.loc[idx]
    ten_bc = st.text_input("Tên Báo Cáo / Công Việc:", value=str(row.get("Tên Báo Cáo / Công Việc", "")))
    tan_suat = st.selectbox("Tần suất:", ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"], index=0)
    don_vi = st.text_input("Đơn vị nhận:", value=str(row.get("Đơn vị nhận", "")))
    l1 = st.text_input("Link 1:", value=str(row.get("Link 1", "")))
    l2 = st.text_input("Link 2:", value=str(row.get("Link 2", "")))
    l3 = st.text_input("Link 3:", value=str(row.get("Link 3", "")))
    ghichu = st.text_input("Ghi chú:", value=str(row.get("Ghi chú", "")))
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.bc_dinhky_df.loc[idx, ["Tên Báo Cáo / Công Việc", "Tần suất", "Đơn vị nhận", "Link 1", "Link 2", "Link 3", "Ghi chú"]] = [ten_bc, tan_suat, don_vi, l1, l2, l3, ghichu]
        st.session_state.bc_dinhky_df = reindex_df(st.session_state.bc_dinhky_df)
        st.success("Đã cập nhật!")
        st.rerun()


# ==========================================
# 4. SIDEBAR (THANH ĐIỀU HƯỚNG BÊN TRÁI)
# ==========================================
with st.sidebar:
    st.markdown("### 👤 Người dùng: `ttm`")
    
    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        if st.button("🚪 Đăng xuất", use_container_width=True):
            st.info("Đã đăng xuất!")
    with col_sb2:
        if st.button("🧹 Xóa Cache", use_container_width=True):
            st.cache_data.clear()
            st.success("Đã xóa cache!")

    col_sb3, col_sb4 = st.columns(2)
    with col_sb3:
        theme_label = "☀️ Sáng" if st.session_state.get("dark_mode") else "🌙 Tối"
        if st.button(theme_label, use_container_width=True):
            st.session_state.dark_mode = not st.session_state.dark_mode
            st.rerun()
            
    with col_sb4:
        if st.session_state.get("is_admin"):
            if st.button("🔓 Thoát Admin", use_container_width=True):
                st.session_state.is_admin = False
                st.rerun()
        else:
            if st.button("🔐 Quản trị", use_container_width=True):
                admin_login_dialog()

    st.divider()
    
    if st.session_state.get("is_admin"):
        st.success("Quyền hiện tại: ADMIN")
        
    st.header("📂 PHÂN MỤC CHÍNH")
    main_menu = st.radio(
        "Điều hướng ứng dụng:",
        [
            "1 🌐 DS WEBsites_CV", 
            "2 📋 DM QL Files_CV", 
            "3 📊 DS BCdinhky_CV",
            "4 🟢 DS Gsheet_CV"
        ],
        label_visibility="collapsed"
    )

    # -----------------------------------------------------
    # CẤU HÌNH THƯ MỤC (Chỉ hiển thị khi là Admin)
    # -----------------------------------------------------
    if st.session_state.get("is_admin"):
        st.divider()
        st.header("⚙️ Cấu Hình Thư Mục")
        st.caption("*(Lưu ý: Mở thư mục chỉ hoạt động khi chạy trên máy tính cá nhân)*")
        
        # Ô nhập liệu cho phép người dùng thay đổi đường dẫn
        new_dir = st.text_input("Đường dẫn lưu file cục bộ:", value=st.session_state.excel_dir)
        
        col_dir1, col_dir2 = st.columns(2)
        with col_dir1:
            if st.button("💾 Xác nhận", type="primary", use_container_width=True):
                st.session_state.excel_dir = new_dir
                st.toast("🎉 Đã cập nhật đường dẫn thư mục!", icon="✅")
        with col_dir2:
            if st.button("📂 Mở thư mục", use_container_width=True):
                if os.path.exists(st.session_state.excel_dir):
                    try:
                        os.startfile(st.session_state.excel_dir)
                    except Exception:
                        st.error("Tính năng này không khả dụng khi chạy trên Web Cloud!")
                else:
                    st.error("Đường dẫn thư mục không tồn tại!")


# ==========================================
# 5. MAIN LAYOUT (GIAO DIỆN CHÍNH)
# ==========================================
st.title("🛡️ Quản Lý An Toàn TTM")
st.caption("📌 Phiên bản hệ thống hiệu chỉnh ngày: 28/09/2026")

m1, m2, m3 = st.columns(3)
m1.metric("🌐 Tổng số Websites", len(st.session_state.web_tools_df))
m2.metric("📊 Tổng số BC Định kỳ", len(st.session_state.bc_dinhky_df))
m3.metric("🟢 Tổng số GSheets", len(st.session_state.gsheet_df))
st.divider()

# ------------------------------------------
# PHẦN 1: DS WEBsites_CV
# ------------------------------------------
if main_menu == "1 🌐 DS WEBsites_CV":
    col_title, col_action1, col_action2 = st.columns([5, 2, 2])
    col_title.subheader("🌐 Danh Sách Các Website Hỗ Trợ CV")
    with col_action1:
        if st.session_state.is_admin:
            if st.button("➕ Thêm mới Website", type="primary", use_container_width=True):
                add_web_dialog()
    with col_action2:
        if st.session_state.is_admin:
            if st.button("🔄 Khôi phục mặc định", use_container_width=True):
                st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
                st.rerun()

    with st.container(border=True):
        h_cols = st.columns([0.7, 3, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3, 3, 3.3])
        h_cols[0].markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h_cols[1].markdown("<span class='table-header'>Mô tả WEB</span>", unsafe_allow_html=True)
        h_cols[2].markdown("<span class='table-header'>Links truy cập</span>", unsafe_allow_html=True)
        h_cols[3].markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        if st.session_state.is_admin:
            h_cols[4].markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_web = st.session_state.web_tools_df
        for idx, row in df_web.iterrows():
            c_cols = st.columns([0.7, 3, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3, 3, 3.3], vertical_alignment="center")
            c_cols[0].markdown(f"**{row['STT']}**")
            c_cols[1].write(row["Mô tả WEB"])
            c_cols[2].markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c_cols[3].write(row["Ghi chú"])
            
            if st.session_state.is_admin:
                btn_edit, btn_del = c_cols[4].columns(2)
                if btn_edit.button("✏️", key=f"ew_{idx}", help="Chỉnh sửa"):
                    edit_web_dialog(idx)
                if btn_del.button("🗑️", key=f"dw_{idx}", help="Xóa"):
                    st.session_state.web_tools_df = reindex_df(df_web.drop(idx))
                    st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #555;'>", unsafe_allow_html=True)

    if st.session_state.is_admin:
        with st.expander("⚙️ Quản lý Nhập / Xuất Excel (WEBsites)", expanded=False):
            col_w1, col_w2 = st.columns(2)
            with col_w1:
                st.markdown("#### Tải dữ liệu xuống máy")
                out_df = reindex_df(st.session_state.web_tools_df)
                excel_data = generate_excel_download(out_df, "WEBSITES")
                st.download_button(
                    label="📥 Tải file Excel",
                    data=excel_data,
                    file_name=f"1_DS_WEBsites_CV_{DATE_STR}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_w2:
                st.markdown("#### Nhập dữ liệu")
                up_w = st.file_uploader("Chọn file Excel:", type=["xlsx", "xls"], key="up_w", label_visibility="collapsed")
                if up_w and st.button("🚀 Cập nhật từ File", type="primary", use_container_width=True):
                    try:
                        df_u = pd.read_excel(up_w)
                        st.session_state.web_tools_df = reindex_df(df_u)
                        st.toast("🎉 Đã cập nhật thành công!", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi: {e}")

# ------------------------------------------
# PHẦN 2: DM QL Files_CV
# ------------------------------------------
elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.selectbox("📌 Chọn Mảng Công Việc:", CATEGORIES)
    
    col_title, col_action = st.columns([7, 3])
    col_title.subheader(f"📂 Hồ Sơ: {selected_cat}")
    with col_action:
        if st.session_state.is_admin:
            if st.button("➕ Thêm Hồ Sơ Mới", type="primary", use_container_width=True):
                add_file_dialog(selected_cat)

    with st.container(border=True):
        h_cols = st.columns([0.7, 3.5, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3.5, 3, 2.8])
        h_cols[0].markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h_cols[1].markdown("<span class='table-header'>Thư mục / Hồ sơ</span>", unsafe_allow_html=True)
        h_cols[2].markdown("<span class='table-header'>Links xem</span>", unsafe_allow_html=True)
        h_cols[3].markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        if st.session_state.is_admin:
            h_cols[4].markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_cat = st.session_state.data_store[selected_cat]
        for idx, row in df_cat.iterrows():
            c_cols = st.columns([0.7, 3.5, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3.5, 3, 2.8], vertical_alignment="center")
            c_cols[0].markdown(f"**{row['STT']}**")
            c_cols[1].write(row["Thư mục / Hồ sơ"])
            c_cols[2].markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c_cols[3].write(row["Ghi chú"])
            
            if st.session_state.is_admin:
                btn_e, btn_d = c_cols[4].columns(2)
                if btn_e.button("✏️", key=f"eq_{idx}"):
                    edit_file_dialog(selected_cat, idx)
                if btn_d.button("🗑️", key=f"dq_{idx}"):
                    st.session_state.data_store[selected_cat] = reindex_df(df_cat.drop(idx))
                    st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #555;'>", unsafe_allow_html=True)

    if st.session_state.is_admin:
        with st.expander(f"⚙️ Quản lý Nhập / Xuất Excel ({selected_cat})", expanded=False):
            col_q1, col_q2 = st.columns(2)
            with col_q1:
                st.markdown("#### Tải dữ liệu xuống máy")
                safe_name = selected_cat.replace(" ", "_")
                out_df = reindex_df(st.session_state.data_store[selected_cat])
                excel_data = generate_excel_download(out_df, "HOSO")
                st.download_button(
                    label="📥 Tải file Excel",
                    data=excel_data,
                    file_name=f"2_DM_QL_Files_CV_{safe_name}_{DATE_STR}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_q2:
                st.markdown("#### Nhập dữ liệu")
                up_q = st.file_uploader("Chọn file Excel:", type=["xlsx", "xls"], key=f"up_q_{selected_cat}", label_visibility="collapsed")
                if up_q and st.button("🚀 Cập nhật từ File", type="primary", use_container_width=True):
                    try:
                        df_qu = pd.read_excel(up_q)
                        st.session_state.data_store[selected_cat] = reindex_df(df_qu)
                        st.toast("🎉 Cập nhật thành công!", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi: {e}")

# ------------------------------------------
# PHẦN 3: DS BCdinhky_CV
# ------------------------------------------
elif main_menu == "3 📊 DS BCdinhky_CV":
    col_title, col_action = st.columns([7, 3])
    col_title.subheader("📊 Quản Lý Báo Cáo Định Kỳ")
    with col_action:
        if st.session_state.is_admin:
            if st.button("➕ Thêm Báo Cáo Mới", type="primary", use_container_width=True):
                add_bc_dialog()

    with st.container(border=True):
        h_cols = st.columns([0.6, 2.5, 1.5, 1.5, 2, 2, 1.2] if st.session_state.is_admin else [0.6, 2.5, 1.5, 1.5, 2, 3])
        h_cols[0].markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h_cols[1].markdown("<span class='table-header'>Tên Báo Cáo / CV</span>", unsafe_allow_html=True)
        h_cols[2].markdown("<span class='table-header'>Tần suất</span>", unsafe_allow_html=True)
        h_cols[3].markdown("<span class='table-header'>Đơn vị nhận</span>", unsafe_allow_html=True)
        h_cols[4].markdown("<span class='table-header'>Links biểu mẫu</span>", unsafe_allow_html=True)
        h_cols[5].markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        if st.session_state.is_admin:
            h_cols[6].markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_bc = st.session_state.bc_dinhky_df
        for idx, row in df_bc.iterrows():
            c_cols = st.columns([0.6, 2.5, 1.5, 1.5, 2, 2, 1.2] if st.session_state.is_admin else [0.6, 2.5, 1.5, 1.5, 2, 3], vertical_alignment="center")
            c_cols[0].markdown(f"**{row['STT']}**")
            c_cols[1].write(str(row.get("Tên Báo Cáo / Công Việc", "")))
            c_cols[2].write(str(row.get("Tần suất", "")))
            c_cols[3].write(str(row.get("Đơn vị nhận", "")))
            c_cols[4].markdown(render_links_cell(row.get("Link 1", ""), row.get("Link 2", ""), row.get("Link 3", "")))
            c_cols[5].write(str(row.get("Ghi chú", "")))
            
            if st.session_state.is_admin:
                btn_e, btn_d = c_cols[6].columns(2)
                if btn_e.button("✏️", key=f"ebc_{idx}"):
                    edit_bc_dialog(idx)
                if btn_d.button("🗑️", key=f"dbc_{idx}"):
                    st.session_state.bc_dinhky_df = reindex_df(df_bc.drop(idx))
                    st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #555;'>", unsafe_allow_html=True)

    if st.session_state.is_admin:
        with st.expander("⚙️ Quản lý Nhập / Xuất Excel (Báo Cáo)", expanded=False):
            col_b1, col_b2 = st.columns(2)
            with col_b1:
                st.markdown("#### Tải dữ liệu xuống máy")
                out_df = reindex_df(st.session_state.bc_dinhky_df)
                excel_data = generate_excel_download(out_df, "BAOCAO")
                st.download_button(
                    label="📥 Tải file Excel",
                    data=excel_data,
                    file_name=f"3_DS_BCdinhky_CV_{DATE_STR}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_b2:
                st.markdown("#### Nhập dữ liệu")
                up_b = st.file_uploader("Chọn file Excel:", type=["xlsx", "xls"], key="up_b", label_visibility="collapsed")
                if up_b and st.button("🚀 Cập nhật từ File", type="primary", use_container_width=True):
                    try:
                        df_raw = pd.read_excel(up_b)
                        st.session_state.bc_dinhky_df = reindex_df(df_raw)
                        st.toast("🎉 Đã cập nhật thành công!", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi: {e}")

# ------------------------------------------
# PHẦN 4: DS Gsheet_CV
# ------------------------------------------
elif main_menu == "4 🟢 DS Gsheet_CV":
    col_title, col_action1, col_action2 = st.columns([5, 2, 2])
    col_title.subheader("🟢 Danh Sách Gsheet Hỗ Trợ CV")
    with col_action1:
        if st.session_state.is_admin:
            if st.button("➕ Thêm mới Gsheet", type="primary", use_container_width=True):
                add_gsheet_dialog()
    with col_action2:
        if st.session_state.is_admin:
            if st.button("🔄 Khôi phục mặc định", use_container_width=True):
                st.session_state.gsheet_df = reindex_df(pd.DataFrame([
                    {"Mô tả Gsheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Dùng chung phòng An Toàn"},
                    {"Mô tả Gsheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật trực tuyến"}
                ]))
                st.rerun()

    with st.container(border=True):
        h_cols = st.columns([0.7, 3, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3, 3, 3.3])
        h_cols[0].markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h_cols[1].markdown("<span class='table-header'>Mô tả Gsheet</span>", unsafe_allow_html=True)
        h_cols[2].markdown("<span class='table-header'>Links truy cập</span>", unsafe_allow_html=True)
        h_cols[3].markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        if st.session_state.is_admin:
            h_cols[4].markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_gsheet = st.session_state.gsheet_df
        for idx, row in df_gsheet.iterrows():
            c_cols = st.columns([0.7, 3, 3, 2.5, 1.5] if st.session_state.is_admin else [0.7, 3, 3, 3.3], vertical_alignment="center")
            c_cols[0].markdown(f"**{row['STT']}**")
            c_cols[1].write(row["Mô tả Gsheet"])
            c_cols[2].markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c_cols[3].write(row["Ghi chú"])
            
            if st.session_state.is_admin:
                btn_edit, btn_del = c_cols[4].columns(2)
                if btn_edit.button("✏️", key=f"eg_{idx}"):
                    edit_gsheet_dialog(idx)
                if btn_del.button("🗑️", key=f"dg_{idx}"):
                    st.session_state.gsheet_df = reindex_df(df_gsheet.drop(idx))
                    st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #555;'>", unsafe_allow_html=True)

    if st.session_state.is_admin:
        with st.expander("⚙️ Quản lý Nhập / Xuất Excel (Gsheet)", expanded=False):
            col_gx1, col_gx2 = st.columns(2)
            with col_gx1:
                st.markdown("#### Tải dữ liệu xuống máy")
                out_df = reindex_df(st.session_state.gsheet_df)
                excel_data = generate_excel_download(out_df, "GSHEETS")
                st.download_button(
                    label="📥 Tải file Excel",
                    data=excel_data,
                    file_name=f"4_DS_Gsheet_CV_{DATE_STR}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            with col_gx2:
                st.markdown("#### Nhập dữ liệu")
                up_g = st.file_uploader("Chọn file Excel:", type=["xlsx", "xls"], key="up_g", label_visibility="collapsed")
                if up_g and st.button("🚀 Cập nhật từ File", type="primary", use_container_width=True):
                    try:
                        df_gu = pd.read_excel(up_g)
                        st.session_state.gsheet_df = reindex_df(df_gu)
                        st.toast("🎉 Đã cập nhật thành công!", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Lỗi: {e}")
