import streamlit as st
import pandas as pd
import os
import openpyxl
from datetime import datetime

# 1. Cấu hình trang Dashboard
st.set_page_config(
    page_title="Hệ Thống Quản Lý An Toàn & Công Tác Chuyên Môn TMT",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Thư mục làm việc
EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(EXCEL_DIR):
    EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

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
    {"Mô tả WEB": "Hệ thống D-Office", "Link truy cập": "https://doffice.evn.com.vn", "Ghi chú": "Quản lý văn bản điều hành"},
    {"Mô tả WEB": "Cổng thông tin Điện lực", "Link truy cập": "https://evnspc.vn", "Ghi chú": "Tra cứu quy định & chỉ đạo"},
    {"Mô tả WEB": "Quản lý an toàn", "Link truy cập": "https://ecp.evn.com.vn", "Ghi chú": "Quản lý an toàn"},
    {"Mô tả WEB": "Lịch tuần", "Link truy cập": "https://lichtuan.evnspc.vn", "Ghi chú": "Công ty Điện lực Tây Ninh"},
    {"Mô tả WEB": "Hệ thống HRMS", "Link truy cập": "https://hrms.evn.com.vn", "Ghi chú": "Quản lý lao động tiền lương"},
    {"Mô tả WEB": "Hệ thống PMIS", "Link truy cập": "https://pmis.evn.com.vn", "Ghi chú": "Quản lý vận hành thiết bị & lưới điện"},
    {"Mô tả WEB": "Hệ thống E-Learning", "Link truy cập": "https://elearning.evn.com.vn", "Ghi chú": "Huấn luyện an toàn & thi trực tuyến"},
    {"Mô tả WEB": "Thư viện Quy chuẩn - Quy định An toàn", "Link truy cập": "https://drive.google.com", "Ghi chú": "Tra cứu tài liệu an toàn PCCC & ĐT"},
    {"Mô tả WEB": "Cổng Dịch vụ công Quốc gia", "Link truy cập": "https://dichvucong.gov.vn", "Ghi chú": "Thực hiện thủ tục hành chính PCCC/ĐTXD"},
    {"Mô tả WEB": "Cổng Thông tin Bộ Công Thương", "Link truy cập": "https://moit.gov.vn", "Ghi chú": "Theo dõi văn bản quy phạm kỹ thuật"},
    {"Mô tả WEB": "Cổng Báo cáo Phòng chống thiên tai", "Link truy cập": "https://pctt.evn.com.vn", "Ghi chú": "Cập nhật tình hình PCTT & TKCN"},
    {"Mô tả WEB": "Hệ thống Quản lý Đầu tư Xây dựng (IMIS)", "Link truy cập": "https://imis.evn.com.vn", "Ghi chú": "Theo dõi an toàn dự án ĐTXD"},
    {"Mô tả WEB": "Hệ thống Thông tin Báo cáo EVN", "Link truy cập": "https://baocao.evn.com.vn", "Ghi chú": "Tổng hợp chỉ tiêu an toàn - kỹ thuật"},
    {"Mô tả WEB": "Lưu trữ Hồ sơ / Biểu mẫu TMT", "Link truy cập": "https://drive.google.com", "Ghi chú": "Kho lưu trữ dữ liệu dùng chung TMT"}
]

DEFAULT_GSHEETS = [
    {"Mô tả Gsheet": "Bảng tổng hợp Báo cáo An toàn Lưới điện", "Link truy cập": "https://docs.google.com/spreadsheets", "Ghi chú": "Dữ liệu trực tuyến dùng chung"},
    {"Mô tả Gsheet": "Theo dõi tiến độ Khắc phục Kiến nghị Kiểm tra", "Link truy cập": "https://docs.google.com/spreadsheets", "Ghi chú": "Cập nhật tuần"}
]

def reindex_df(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.dropna(how='all').reset_index(drop=True)
    if "STT" in df.columns:
        df = df.drop(columns=["STT"])
    df.insert(0, "STT", range(1, len(df) + 1))
    return df

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

def normalize_bc_df(df):
    df = df.dropna(how='all')
    df.columns = [str(c).strip() for c in df.columns]
    mapping = {}
    for col in df.columns:
        col_lower = col.lower()
        if "tên" in col_lower or "báo cáo" in col_lower or "công việc" in col_lower:
            mapping[col] = "Tên Báo Cáo / Công Việc"
        elif "tần suất" in col_lower or "tan suat" in col_lower:
            mapping[col] = "Tần suất"
        elif "đơn vị" in col_lower or "don vi" in col_lower or "nhận" in col_lower:
            mapping[col] = "Đơn vị nhận"
        elif "link" in col_lower or "biểu mẫu" in col_lower or "bieu mau" in col_lower:
            mapping[col] = "Link biểu mẫu"
        elif "ghi chú" in col_lower or "ghi chu" in col_lower:
            mapping[col] = "Ghi chú"
    df = df.rename(columns=mapping)
    required_cols = ["Tên Báo Cáo / Công Việc", "Tần suất", "Đơn vị nhận", "Link biểu mẫu", "Ghi chú"]
    for c in required_cols:
        if c not in df.columns:
            df[c] = ""
    return reindex_df(df[required_cols])

# ----------------- SIDEBAR -----------------
st.sidebar.markdown("👤 **Xin chào:** `tmt`")
col_sb1, col_sb2 = st.sidebar.columns(2)
with col_sb1:
    if st.button("🚪 Đăng xuất", use_container_width=True):
        st.info("Đã đăng xuất!")
with col_sb2:
    if st.button("🧹 Xóa Cache", use_container_width=True):
        st.cache_data.clear()
        st.sidebar.success("Đã xóa cache!")

st.sidebar.markdown("---")
st.sidebar.header("📂 PHÂN MỤC CHÍNH")
main_menu = st.sidebar.radio(
    "Chọn phân mục làm việc:",
    [
        "1 🌐 DS WEBsites_CV", 
        "2 📋 DM QL Files_CV", 
        "3 📊 DS BCdinhky_CV",
        "4 🟢 DS Gsheet_CV"
    ]
)

st.sidebar.markdown("---")
st.sidebar.header("⚙️ Cấu Hình Thư Mục Lưu File")
st.sidebar.text_input("Đường dẫn thư mục lưu file Excel:", value=EXCEL_DIR, disabled=True)

if st.sidebar.button("📂 Mở thư mục lưu trữ"):
    try:
        os.startfile(EXCEL_DIR)
    except Exception:
        st.sidebar.error("Không mở được thư mục!")

# ----------------- DỮ LIỆU SESSION STATE -----------------
if "web_tools_df" not in st.session_state:
    st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))

if "gsheet_df" not in st.session_state:
    st.session_state.gsheet_df = reindex_df(pd.DataFrame(DEFAULT_GSHEETS))

if "data_store" not in st.session_state:
    st.session_state.data_store = {}
    for cat in CATEGORIES:
        st.session_state.data_store[cat] = reindex_df(pd.DataFrame([
            {"Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link xem": "https://drive.google.com", "Ghi chú": "Cập nhật định kỳ"}
        ]))

if "bc_dinhky_df" not in st.session_state:
    st.session_state.bc_dinhky_df = reindex_df(pd.DataFrame([
        {"Tên Báo Cáo / Công Việc": "Báo cáo công tác An toàn định kỳ Quý", "Tần suất": "Hàng Quý", "Đơn vị nhận": "Công ty Điện lực", "Link biểu mẫu": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 20 cuối quý"},
        {"Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link biểu mẫu": "https://drive.google.com", "Ghi chú": "Nộp trước ngày 25 hàng tháng"}
    ]))

# ==========================================
# DIALOGS
# ==========================================
@st.dialog("➕ Thêm mới Website_CV")
def add_web_dialog():
    mota = st.text_input("Mô tả WEB:")
    link = st.text_input("Link truy cập:")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if mota:
            new_row = pd.DataFrame([{"Mô tả WEB": mota, "Link truy cập": link, "Ghi chú": ghichu}])
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
    link = st.text_input("Link truy cập:", value=row["Link truy cập"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.web_tools_df.at[idx, "Mô tả WEB"] = mota
        st.session_state.web_tools_df.at[idx, "Link truy cập"] = link
        st.session_state.web_tools_df.at[idx, "Ghi chú"] = ghichu
        st.session_state.web_tools_df = reindex_df(st.session_state.web_tools_df)
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Gsheet_CV")
def add_gsheet_dialog():
    mota = st.text_input("Mô tả Gsheet:")
    link = st.text_input("Link truy cập (Google Sheets):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if mota:
            new_row = pd.DataFrame([{"Mô tả Gsheet": mota, "Link truy cập": link, "Ghi chú": ghichu}])
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
    link = st.text_input("Link truy cập:", value=row["Link truy cập"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.gsheet_df.at[idx, "Mô tả Gsheet"] = mota
        st.session_state.gsheet_df.at[idx, "Link truy cập"] = link
        st.session_state.gsheet_df.at[idx, "Ghi chú"] = ghichu
        st.session_state.gsheet_df = reindex_df(st.session_state.gsheet_df)
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Hồ Sơ / Thư Mục")
def add_file_dialog(selected_cat):
    hoso = st.text_input("Tên Thư mục / Hồ sơ:")
    link = st.text_input("Link xem (Google Drive):")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if hoso:
            new_row = pd.DataFrame([{"Thư mục / Hồ sơ": hoso, "Link xem": link, "Ghi chú": ghichu}])
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
    link = st.text_input("Link xem:", value=row["Link xem"])
    ghichu = st.text_input("Ghi chú:", value=row["Ghi chú"])
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.data_store[selected_cat].at[idx, "Thư mục / Hồ sơ"] = hoso
        st.session_state.data_store[selected_cat].at[idx, "Link xem"] = link
        st.session_state.data_store[selected_cat].at[idx, "Ghi chú"] = ghichu
        st.session_state.data_store[selected_cat] = reindex_df(st.session_state.data_store[selected_cat])
        st.success("Đã cập nhật!")
        st.rerun()

@st.dialog("➕ Thêm mới Báo Cáo / Công Việc")
def add_bc_dialog():
    ten_bc = st.text_input("Tên Báo Cáo / Công Việc:")
    tan_suat = st.selectbox("Tần suất:", ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"])
    don_vi = st.text_input("Đơn vị nhận:")
    link = st.text_input("Link biểu mẫu:")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if ten_bc:
            new_row = pd.DataFrame([{"Tên Báo Cáo / Công Việc": ten_bc, "Tần suất": tan_suat, "Đơn vị nhận": don_vi, "Link biểu mẫu": link, "Ghi chú": ghichu}])
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
    link = st.text_input("Link biểu mẫu:", value=str(row.get("Link biểu mẫu", "")))
    ghichu = st.text_input("Ghi chú:", value=str(row.get("Ghi chú", "")))
    if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
        st.session_state.bc_dinhky_df.at[idx, "Tên Báo Cáo / Công Việc"] = ten_bc
        st.session_state.bc_dinhky_df.at[idx, "Tần suất"] = tan_suat
        st.session_state.bc_dinhky_df.at[idx, "Đơn vị nhận"] = don_vi
        st.session_state.bc_dinhky_df.at[idx, "Link biểu mẫu"] = link
        st.session_state.bc_dinhky_df.at[idx, "Ghi chú"] = ghichu
        st.session_state.bc_dinhky_df = reindex_df(st.session_state.bc_dinhky_df)
        st.success("Đã cập nhật!")
        st.rerun()

# ----------------- GIAO DIỆN CHÍNH -----------------
st.title("🛡️ Hệ Thống Quản Lý An Toàn & Công Tác Chuyên Môn TMT")
st.caption("📌 Phiên bản hệ thống đã hiệu chỉnh ngày: 26/09/2026")
st.markdown("---")

# ==========================================
# 1. DS WEBsites_CV
# ==========================================
if main_menu == "1 🌐 DS WEBsites_CV":
    st.subheader("🌐 Bảng Danh Sách WEBsites_CV")
    
    col_btn1, col_btn2 = st.columns([2, 1])
    with col_btn1:
        if st.button("➕ Thêm mới Website_CV", type="primary"):
            add_web_dialog()
    with col_btn2:
        if st.button("🔄 Khôi phục 14 Web mặc định"):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
            st.success("Đã khôi phục 14 Web mặc định!")
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    h1, h2, h3, h4, h5 = st.columns([1, 3, 2, 3, 2])
    h1.markdown("**STT**")
    h2.markdown("**Mô tả WEB**")
    h3.markdown("**Link truy cập**")
    h4.markdown("**Ghi chú**")
    h5.markdown("<span style='color: #4CAF50; font-weight: bold;'>Thao tác</span>", unsafe_allow_html=True)
    st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

    df_web = st.session_state.web_tools_df
    for idx, row in df_web.iterrows():
        c1, c2, c3, c4, c5 = st.columns([1, 3, 2, 3, 2])
        c1.write(f"**{row['STT']}**")
        c2.write(row["Mô tả WEB"])
        c3.markdown(f"[Link 1]({row['Link truy cập']})")
        c4.write(row["Ghi chú"])
        
        btn_edit, btn_del = c5.columns(2)
        if btn_edit.button("✏️", key=f"edit_w_{idx}"):
            edit_web_dialog(idx)
        if btn_del.button("🗑️", key=f"del_w_{idx}"):
            st.session_state.web_tools_df = reindex_df(df_web.drop(idx))
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel WEBsites_CV")
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        if st.button("📥 Xuất toàn bộ Excel WEBsites_CV", type="primary", use_container_width=True):
            try:
                out_path_web = os.path.join(EXCEL_DIR, f"1 DS WEBsites_CV out_{DATE_STR}.xlsx")
                out_df = reindex_df(st.session_state.web_tools_df)
                with pd.ExcelWriter(out_path_web, engine='openpyxl') as writer:
                    out_df.to_excel(writer, sheet_name="WEBSITES", index=False)
                    auto_fit_columns(writer.book)
                st.balloons()
                st.success(f"Đã xuất file thành công tại:\n`{out_path_web}`")
            except Exception as e:
                st.error(f"Lỗi: {e}")

    with col_w2:
        up_w = st.file_uploader("Nhập file Excel WEBsites để cập nhật:", type=["xlsx", "xls"], key="up_w")
        if up_w is not None:
            if st.button("🚀 Tiến hành cập nhật WEBsites từ File", type="primary", use_container_width=True):
                try:
                    df_u = pd.read_excel(up_w)
                    st.session_state.web_tools_df = reindex_df(df_u)
                    st.balloons()
                    st.toast("🎉 Đã cập nhật thành công!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")

# ==========================================
# 2. DM QL Files_CV
# ==========================================
elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.sidebar.radio("Chọn mảng công việc:", CATEGORIES)
    st.subheader(f"📂 Quản Lý Hồ Sơ: {selected_cat}")

    if st.button(f"➕ Thêm mới Hồ Sơ ({selected_cat})", type="primary"):
        add_file_dialog(selected_cat)

    st.markdown("<br>", unsafe_allow_html=True)

    q1, q2, q3, q4, q5 = st.columns([1, 4, 2, 3, 2])
    q1.markdown("**STT**")
    q2.markdown("**Thư mục / Hồ sơ**")
    q3.markdown("**Link xem**")
    q4.markdown("**Ghi chú**")
    q5.markdown("<span style='color: #4CAF50; font-weight: bold;'>Thao tác</span>", unsafe_allow_html=True)
    st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

    df_cat = st.session_state.data_store[selected_cat]
    for idx, row in df_cat.iterrows():
        c1, c2, c3, c4, c5 = st.columns([1, 4, 2, 3, 2])
        c1.write(f"**{row['STT']}**")
        c2.write(row["Thư mục / Hồ sơ"])
        c3.markdown(f"[Mở xem]({row['Link xem']})")
        c4.write(row["Ghi chú"])
        
        btn_e, btn_d = c5.columns(2)
        if btn_e.button("✏️", key=f"e_q_{idx}"):
            edit_file_dialog(selected_cat, idx)
        if btn_d.button("🗑️", key=f"d_q_{idx}"):
            st.session_state.data_store[selected_cat] = reindex_df(df_cat.drop(idx))
            st.rerun()

    st.markdown("---")
    st.subheader(f"📊 Xuất / Nhập Excel Hồ Sơ: {selected_cat}")
    col_q1, col_q2 = st.columns(2)
    with col_q1:
        if st.button(f"📥 Xuất Excel mục {selected_cat}", type="primary", use_container_width=True):
            try:
                safe_name_cat = selected_cat.replace(" ", "_")
                out_path_cat = os.path.join(EXCEL_DIR, f"2 DM QL Files_CV_{safe_name_cat}_out_{DATE_STR}.xlsx")
                out_df = reindex_df(st.session_state.data_store[selected_cat])
                with pd.ExcelWriter(out_path_cat, engine='openpyxl') as writer:
                    out_df.to_excel(writer, sheet_name="HOSO", index=False)
                    auto_fit_columns(writer.book)
                st.balloons()
                st.success(f"Đã xuất file thành công tại:\n`{out_path_cat}`")
            except Exception as e:
                st.error(f"Lỗi: {e}")

    with col_q2:
        up_q = st.file_uploader(f"Nhập file Excel cho mục {selected_cat}:", type=["xlsx", "xls"], key=f"up_q_{selected_cat}")
        if up_q is not None:
            if st.button(f"🚀 Tiến hành cập nhật Hồ Sơ {selected_cat}", type="primary", use_container_width=True):
                try:
                    df_qu = pd.read_excel(up_q)
                    st.session_state.data_store[selected_cat] = reindex_df(df_qu)
                    st.balloons()
                    st.toast("🎉 Đã cập nhật thành công!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")

# ==========================================
# 3. DS BCdinhky_CV
# ==========================================
elif main_menu == "3 📊 DS BCdinhky_CV":
    st.subheader("📊 Quản Lý Báo Cáo Định Kỳ")
    
    if st.button("➕ Thêm mới Báo Cáo / Công Việc", type="primary"):
        add_bc_dialog()

    st.markdown("<br>", unsafe_allow_html=True)

    b1, b2, b3, b4, b5, b6, b7 = st.columns([1, 3, 2, 2, 2, 2, 2])
    b1.markdown("**STT**")
    b2.markdown("**Tên Báo Cáo / Công Việc**")
    b3.markdown("**Tần suất**")
    b4.markdown("**Đơn vị nhận**")
    b5.markdown("**Link biểu mẫu**")
    b6.markdown("**Ghi chú**")
    b7.markdown("<span style='color: #4CAF50; font-weight: bold;'>Thao tác</span>", unsafe_allow_html=True)
    st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

    df_bc = st.session_state.bc_dinhky_df
    for idx, row in df_bc.iterrows():
        c1, c2, c3, c4, c5, c6, c7 = st.columns([1, 3, 2, 2, 2, 2, 2])
        c1.write(f"**{row['STT']}**")
        c2.write(str(row.get("Tên Báo Cáo / Công Việc", "")))
        c3.write(str(row.get("Tần suất", "")))
        c4.write(str(row.get("Đơn vị nhận", "")))
        
        link_val = str(row.get("Link biểu mẫu", ""))
        if link_val.startswith("http"):
            c5.markdown(f"[Biểu mẫu]({link_val})")
        else:
            c5.write(link_val if link_val else "-")
            
        c6.write(str(row.get("Ghi chú", "")))
        
        btn_e, btn_d = c7.columns(2)
        if btn_e.button("✏️", key=f"e_bc_{idx}"):
            edit_bc_dialog(idx)
        if btn_d.button("🗑️", key=f"d_bc_{idx}"):
            st.session_state.bc_dinhky_df = reindex_df(df_bc.drop(idx))
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel Báo Cáo Định Kỳ")
    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("📥 Xuất toàn bộ Excel Báo Cáo Định Kỳ", type="primary", use_container_width=True):
            try:
                out_path_bc = os.path.join(EXCEL_DIR, f"3 DS BCdinhky_CV out_{DATE_STR}.xlsx")
                out_df = reindex_df(st.session_state.bc_dinhky_df)
                with pd.ExcelWriter(out_path_bc, engine='openpyxl') as writer:
                    out_df.to_excel(writer, sheet_name="BAOCAO", index=False)
                    auto_fit_columns(writer.book)
                st.balloons()
                st.success(f"Đã xuất file thành công tại:\n`{out_path_bc}`")
            except Exception as e:
                st.error(f"Lỗi: {e}")

    with col_b2:
        up_b = st.file_uploader("Nhập file Excel Báo cáo định kỳ để cập nhật:", type=["xlsx", "xls"], key="up_b")
        if up_b is not None:
            if st.button("🚀 Tiến hành cập nhật Báo Cáo Định Kỳ từ File", type="primary", use_container_width=True):
                try:
                    df_raw = pd.read_excel(up_b)
                    st.session_state.bc_dinhky_df = normalize_bc_df(df_raw)
                    st.balloons()
                    st.toast("🎉 Đã cập nhật thành công đủ các dòng dữ liệu!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi đọc file Excel: {e}")

# ==========================================
# 4. DS Gsheet_CV
# ==========================================
elif main_menu == "4 🟢 DS Gsheet_CV":
    st.subheader("🟢 Bảng Danh Sách Gsheet_CV")
    
    col_g1, col_g2 = st.columns([2, 1])
    with col_g1:
        if st.button("➕ Thêm mới Gsheet_CV", type="primary"):
            add_gsheet_dialog()
    with col_g2:
        if st.button("🔄 Khôi phục Gsheet mặc định"):
            st.session_state.gsheet_df = reindex_df(pd.DataFrame(DEFAULT_GSHEETS))
            st.success("Đã khôi phục Gsheet mặc định!")
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    gh1, gh2, gh3, gh4, gh5 = st.columns([1, 3, 2, 3, 2])
    gh1.markdown("**STT**")
    gh2.markdown("**Mô tả Gsheet**")
    gh3.markdown("**Link truy cập**")
    gh4.markdown("**Ghi chú**")
    gh5.markdown("<span style='color: #4CAF50; font-weight: bold;'>Thao tác</span>", unsafe_allow_html=True)
    st.markdown("<hr style='margin: 5px 0 15px 0;'>", unsafe_allow_html=True)

    df_gsheet = st.session_state.gsheet_df
    for idx, row in df_gsheet.iterrows():
        c1, c2, c3, c4, c5 = st.columns([1, 3, 2, 3, 2])
        c1.write(f"**{row['STT']}**")
        c2.write(row["Mô tả Gsheet"])
        c3.markdown(f"[Link Gsheet]({row['Link truy cập']})")
        c4.write(row["Ghi chú"])
        
        btn_edit, btn_del = c5.columns(2)
        if btn_edit.button("✏️", key=f"edit_g_{idx}"):
            edit_gsheet_dialog(idx)
        if btn_del.button("🗑️", key=f"del_g_{idx}"):
            st.session_state.gsheet_df = reindex_df(df_gsheet.drop(idx))
            st.rerun()

    st.markdown("---")
    st.subheader("📊 Xuất / Nhập Excel Gsheet_CV")
    col_gx1, col_gx2 = st.columns(2)
    with col_gx1:
        if st.button("📥 Xuất toàn bộ Excel Gsheet_CV", type="primary", use_container_width=True):
            try:
                out_path_gsheet = os.path.join(EXCEL_DIR, f"4 DS Gsheet_CV out_{DATE_STR}.xlsx")
                out_df = reindex_df(st.session_state.gsheet_df)
                with pd.ExcelWriter(out_path_gsheet, engine='openpyxl') as writer:
                    out_df.to_excel(writer, sheet_name="GSHEETS", index=False)
                    auto_fit_columns(writer.book)
                st.balloons()
                st.success(f"Đã xuất file thành công tại:\n`{out_path_gsheet}`")
            except Exception as e:
                st.error(f"Lỗi: {e}")

    with col_gx2:
        up_g = st.file_uploader("Nhập file Excel Gsheets để cập nhật:", type=["xlsx", "xls"], key="up_g")
        if up_g is not None:
            if st.button("🚀 Tiến hành cập nhật Gsheets từ File", type="primary", use_container_width=True):
                try:
                    df_gu = pd.read_excel(up_g)
                    st.session_state.gsheet_df = reindex_df(df_gu)
                    st.balloons()
                    st.toast("🎉 Đã cập nhật danh sách Gsheet thành công!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi: {e}")