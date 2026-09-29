import streamlit as st
import pandas as pd
import os
import openpyxl
from datetime import datetime

# 1. Cấu hình trang Dashboard
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS để giao diện mượt mà hơn
st.markdown("""
    <style>
    .st-emotion-cache-1y4p8pa { padding-top: 2rem; }
    .table-header { font-weight: bold; color: #1E88E5; }
    .row-text { font-size: 14px; }
    </style>
""", unsafe_allow_html=True)

# Thư mục làm việc & Lưu trữ
EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(EXCEL_DIR):
    EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

DATE_STR = datetime.now().strftime("%Y%m%d")

# Danh sách 27 mảng công việc chuyên môn
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

# --- HELPER FUNCTIONS ---
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

# --- DỮ LIỆU SESSION STATE ---
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
# DIALOGS (Giữ nguyên logic của bạn)
# ==========================================
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


# ----------------- SIDEBAR -----------------
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

    st.divider()
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

    st.divider()
    st.header("⚙️ Cấu Hình Thư Mục")
    st.text_input("Thư mục lưu file Excel:", value=EXCEL_DIR, disabled=True)
    if st.button("📂 Mở thư mục lưu trữ", use_container_width=True):
        try:
            os.startfile(EXCEL_DIR)
        except Exception:
            st.error("Không mở được thư mục!")


# ----------------- MAIN LAYOUT -----------------
st.title("🛡️ Quản Lý An Toàn TTM")
st.caption("📌 Phiên bản hệ thống hiệu chỉnh ngày: 28/09/2026")

# Hàng thông số tổng quan (Dashboard metrics)
m1, m2, m3 = st.columns(3)
m1.metric("🌐 Tổng số Websites", len(st.session_state.web_tools_df))
m2.metric("📊 Tổng số BC Định kỳ", len(st.session_state.bc_dinhky_df))
m3.metric("🟢 Tổng số GSheets", len(st.session_state.gsheet_df))
st.divider()

# ==========================================
# 1. DS WEBsites_CV
# ==========================================
if main_menu == "1 🌐 DS WEBsites_CV":
    col_title, col_action1, col_action2 = st.columns([5, 2, 2])
    col_title.subheader("🌐 Danh Sách Các Website Hỗ Trợ CV")
    with col_action1:
        if st.button("➕ Thêm mới Website", type="primary", use_container_width=True):
            add_web_dialog()
    with col_action2:
        if st.button("🔄 Khôi phục mặc định", use_container_width=True):
            st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
            st.rerun()

    # Bảng dữ liệu có viền (Card UI)
    with st.container(border=True):
        # Header của bảng
        h1, h2, h3, h4, h5 = st.columns([0.7, 3, 3, 2.5, 1.5])
        h1.markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h2.markdown("<span class='table-header'>Mô tả WEB</span>", unsafe_allow_html=True)
        h3.markdown("<span class='table-header'>Links truy cập</span>", unsafe_allow_html=True)
        h4.markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        h5.markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        # Dữ liệu
        df_web = st.session_state.web_tools_df
        for idx, row in df_web.iterrows():
            c1, c2, c3, c4, c5 = st.columns([0.7, 3, 3, 2.5, 1.5], vertical_alignment="center")
            c1.markdown(f"**{row['STT']}**")
            c2.write(row["Mô tả WEB"])
            c3.markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c4.write(row["Ghi chú"])
            
            btn_edit, btn_del = c5.columns(2)
            if btn_edit.button("✏️", key=f"ew_{idx}", help="Chỉnh sửa"):
                edit_web_dialog(idx)
            if btn_del.button("🗑️", key=f"dw_{idx}", help="Xóa"):
                st.session_state.web_tools_df = reindex_df(df_web.drop(idx))
                st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

    # Nút Xuất / Nhập thu gọn
    with st.expander("⚙️ Quản lý Nhập / Xuất Excel (WEBsites)", expanded=False):
        col_w1, col_w2 = st.columns(2)
        with col_w1:
            st.markdown("#### Xuất dữ liệu")
            if st.button("📥 Xuất toàn bộ ra Excel", type="primary", use_container_width=True):
                try:
                    out_path = os.path.join(EXCEL_DIR, f"1 DS WEBsites_CV out_{DATE_STR}.xlsx")
                    out_df = reindex_df(st.session_state.web_tools_df)
                    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                        out_df.to_excel(writer, sheet_name="WEBSITES", index=False)
                        auto_fit_columns(writer.book)
                    st.success(f"Đã xuất thành công: {out_path}")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
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

# ==========================================
# 2. DM QL Files_CV
# ==========================================
elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.selectbox("📌 Chọn Mảng Công Việc:", CATEGORIES)
    
    col_title, col_action = st.columns([7, 3])
    col_title.subheader(f"📂 Hồ Sơ: {selected_cat}")
    with col_action:
        if st.button("➕ Thêm Hồ Sơ Mới", type="primary", use_container_width=True):
            add_file_dialog(selected_cat)

    with st.container(border=True):
        h1, h2, h3, h4, h5 = st.columns([0.7, 3.5, 3, 2.5, 1.5])
        h1.markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        h2.markdown("<span class='table-header'>Thư mục / Hồ sơ</span>", unsafe_allow_html=True)
        h3.markdown("<span class='table-header'>Links xem</span>", unsafe_allow_html=True)
        h4.markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        h5.markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_cat = st.session_state.data_store[selected_cat]
        for idx, row in df_cat.iterrows():
            c1, c2, c3, c4, c5 = st.columns([0.7, 3.5, 3, 2.5, 1.5], vertical_alignment="center")
            c1.markdown(f"**{row['STT']}**")
            c2.write(row["Thư mục / Hồ sơ"])
            c3.markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c4.write(row["Ghi chú"])
            
            btn_e, btn_d = c5.columns(2)
            if btn_e.button("✏️", key=f"eq_{idx}"):
                edit_file_dialog(selected_cat, idx)
            if btn_d.button("🗑️", key=f"dq_{idx}"):
                st.session_state.data_store[selected_cat] = reindex_df(df_cat.drop(idx))
                st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

    with st.expander(f"⚙️ Quản lý Nhập / Xuất Excel ({selected_cat})", expanded=False):
        col_q1, col_q2 = st.columns(2)
        with col_q1:
            st.markdown("#### Xuất dữ liệu")
            if st.button("📥 Xuất Excel mục này", type="primary", use_container_width=True):
                try:
                    safe_name = selected_cat.replace(" ", "_")
                    out_path = os.path.join(EXCEL_DIR, f"2 DM QL Files_CV_{safe_name}_out_{DATE_STR}.xlsx")
                    out_df = reindex_df(st.session_state.data_store[selected_cat])
                    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                        out_df.to_excel(writer, sheet_name="HOSO", index=False)
                        auto_fit_columns(writer.book)
                    st.success(f"Đã xuất file: {out_path}")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
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

# ==========================================
# 3. DS BCdinhky_CV
# ==========================================
elif main_menu == "3 📊 DS BCdinhky_CV":
    col_title, col_action = st.columns([7, 3])
    col_title.subheader("📊 Quản Lý Báo Cáo Định Kỳ")
    with col_action:
        if st.button("➕ Thêm Báo Cáo Mới", type="primary", use_container_width=True):
            add_bc_dialog()

    with st.container(border=True):
        # Tỷ lệ cột cho bảng Báo cáo (cần nhiều chi tiết hơn)
        b1, b2, b3, b4, b5, b6, b7 = st.columns([0.6, 2.5, 1.5, 1.5, 2, 2, 1.2])
        b1.markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        b2.markdown("<span class='table-header'>Tên Báo Cáo / CV</span>", unsafe_allow_html=True)
        b3.markdown("<span class='table-header'>Tần suất</span>", unsafe_allow_html=True)
        b4.markdown("<span class='table-header'>Đơn vị nhận</span>", unsafe_allow_html=True)
        b5.markdown("<span class='table-header'>Links biểu mẫu</span>", unsafe_allow_html=True)
        b6.markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        b7.markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_bc = st.session_state.bc_dinhky_df
        for idx, row in df_bc.iterrows():
            c1, c2, c3, c4, c5, c6, c7 = st.columns([0.6, 2.5, 1.5, 1.5, 2, 2, 1.2], vertical_alignment="center")
            c1.markdown(f"**{row['STT']}**")
            c2.write(str(row.get("Tên Báo Cáo / Công Việc", "")))
            c3.write(str(row.get("Tần suất", "")))
            c4.write(str(row.get("Đơn vị nhận", "")))
            c5.markdown(render_links_cell(row.get("Link 1", ""), row.get("Link 2", ""), row.get("Link 3", "")))
            c6.write(str(row.get("Ghi chú", "")))
            
            btn_e, btn_d = c7.columns(2)
            if btn_e.button("✏️", key=f"ebc_{idx}"):
                edit_bc_dialog(idx)
            if btn_d.button("🗑️", key=f"dbc_{idx}"):
                st.session_state.bc_dinhky_df = reindex_df(df_bc.drop(idx))
                st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

    with st.expander("⚙️ Quản lý Nhập / Xuất Excel (Báo Cáo)", expanded=False):
        col_b1, col_b2 = st.columns(2)
        with col_b1:
            st.markdown("#### Xuất dữ liệu")
            if st.button("📥 Xuất toàn bộ ra Excel", type="primary", use_container_width=True):
                try:
                    out_path = os.path.join(EXCEL_DIR, f"3 DS BCdinhky_CV out_{DATE_STR}.xlsx")
                    out_df = reindex_df(st.session_state.bc_dinhky_df)
                    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                        out_df.to_excel(writer, sheet_name="BAOCAO", index=False)
                        auto_fit_columns(writer.book)
                    st.success(f"Đã xuất thành công: {out_path}")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
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

# ==========================================
# 4. DS Gsheet_CV
# ==========================================
elif main_menu == "4 🟢 DS Gsheet_CV":
    col_title, col_action1, col_action2 = st.columns([5, 2, 2])
    col_title.subheader("🟢 Danh Sách Gsheet Hỗ Trợ CV")
    with col_action1:
        if st.button("➕ Thêm mới Gsheet", type="primary", use_container_width=True):
            add_gsheet_dialog()
    with col_action2:
        if st.button("🔄 Khôi phục mặc định", use_container_width=True):
            st.session_state.gsheet_df = reindex_df(pd.DataFrame([
                {"Mô tả Gsheet": "Bảng Theo Dõi Công Việc Theo Tuần", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Dùng chung phòng An Toàn"},
                {"Mô tả Gsheet": "Theo Dõi Kiến Nghị Kiểm Tra", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật trực tuyến"}
            ]))
            st.rerun()

    with st.container(border=True):
        gh1, gh2, gh3, gh4, gh5 = st.columns([0.7, 3, 3, 2.5, 1.5])
        gh1.markdown("<span class='table-header'>STT</span>", unsafe_allow_html=True)
        gh2.markdown("<span class='table-header'>Mô tả Gsheet</span>", unsafe_allow_html=True)
        gh3.markdown("<span class='table-header'>Links truy cập</span>", unsafe_allow_html=True)
        gh4.markdown("<span class='table-header'>Ghi chú</span>", unsafe_allow_html=True)
        gh5.markdown("<span class='table-header'>Thao tác</span>", unsafe_allow_html=True)
        st.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

        df_gsheet = st.session_state.gsheet_df
        for idx, row in df_gsheet.iterrows():
            c1, c2, c3, c4, c5 = st.columns([0.7, 3, 3, 2.5, 1.5], vertical_alignment="center")
            c1.markdown(f"**{row['STT']}**")
            c2.write(row["Mô tả Gsheet"])
            c3.markdown(render_links_cell(row["Link 1"], row["Link 2"], row["Link 3"]))
            c4.write(row["Ghi chú"])
            
            btn_edit, btn_del = c5.columns(2)
            if btn_edit.button("✏️", key=f"eg_{idx}"):
                edit_gsheet_dialog(idx)
            if btn_del.button("🗑️", key=f"dg_{idx}"):
                st.session_state.gsheet_df = reindex_df(df_gsheet.drop(idx))
                st.rerun()
            st.markdown("<hr style='margin: 4px 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

    with st.expander("⚙️ Quản lý Nhập / Xuất Excel (Gsheet)", expanded=False):
        col_gx1, col_gx2 = st.columns(2)
        with col_gx1:
            st.markdown("#### Xuất dữ liệu")
            if st.button("📥 Xuất toàn bộ ra Excel", type="primary", use_container_width=True):
                try:
                    out_path = os.path.join(EXCEL_DIR, f"4 DS Gsheet_CV out_{DATE_STR}.xlsx")
                    out_df = reindex_df(st.session_state.gsheet_df)
                    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                        out_df.to_excel(writer, sheet_name="GSHEETS", index=False)
                        auto_fit_columns(writer.book)
                    st.success(f"Đã xuất thành công: {out_path}")
                except Exception as e:
                    st.error(f"Lỗi: {e}")
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
