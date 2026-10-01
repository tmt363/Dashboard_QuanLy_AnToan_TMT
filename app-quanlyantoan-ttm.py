import io
import os
import sys
from datetime import datetime
from typing import Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG VÀ HẰNG SỐ
# ==========================================
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")
DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))
TODAY_STR = datetime.now().strftime("%d/%m/%Y")
DATE_STR = datetime.now().strftime("%Y%m%d")

CATEGORIES = [
    "DTTU_01 AT", "DTTU_01 AT 01 Bao cao", "DTTU_01 AT 01 Bao cao 2026", "DTTU_02 PCTT",
    "DTTU_03 PCCC", "DTTU_04 HL", "DTTU_05 ATDTXD", "DTTU_ATGT", "DTTU_CNTT", "DTTU_DCAT",
    "DTTU_DCNN va Cac loai xe", "DTTU_DGRR", "DTTU_HNTH cac loai", "DTTU_KIEM TRA",
    "DTTU_KIEM TRA-Thuc hien Kien Nghi", "DTTU_UCKC", "DTTU_UCKC dien tap cac loai",
    "DTTU_khac 01 PHOI HOP CAC TO", "DTTU_khac 02 XEM DE BIET CTY", "DTTU_khac 03 ATD dia phuong",
    "DTTU_khac 03 XEM DE BIET dia phuong", "Quy dinh 0000 Discussion", "Quy dinh GOV",
    "Quy dinh PCTN", "Quy dinh PCTN file tham khao cac Doi", "Quy dinh SPC va EVN",
    "Quy dinh trao doi EVN-SPC-PCTN"
]

DEFAULT_14_WEBS = [
    {"Mô tả WEB": "Hệ thống D-Office", "Link 1": "https://doffice.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý văn bản điều hành"},
    {"Mô tả WEB": "Cổng thông tin Điện lực (EVN SPC)", "Link 1": "https://evnspc.vn", "Link 2": "https://www.congcuweb.net/", "Link 3": "", "Ghi chú": "Tra cứu quy định & chỉ đạo"},
    {"Mô tả WEB": "Quản lý An toàn (ECP / ATLD)", "Link 1": "https://giamsatantoan.evnspc.vn/Home/Index", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý giám sát an toàn SPC"},
    {"Mô tả WEB": "Lịch công tác / Lịch tuần", "Link 1": "https://lichtuan.evnspc.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Công ty Điện lực Tây Ninh"},
    {"Mô tả WEB": "Hệ thống PMIS", "Link 1": "https://pmis.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý vận hành thiết bị & lưới điện"},
    {"Mô tả WEB": "Tritm.la Dashboard", "Link 1": "https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit#gid=964445540", "Link 2": "", "Link 3": "", "Ghi chú": "Google sheet CV"},
]

# ==========================================
# 2. HÀM BỔ TRỢ & QUẢN LÝ DỮ LIỆU
# ==========================================
def reindex_df(df: pd.DataFrame) -> pd.DataFrame:
    """Tối ưu đánh lại số thứ tự và định dạng chuẩn DataFrame."""
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.dropna(how="all").reset_index(drop=True)
    if "STT" in df.columns:
        df = df.drop(columns=["STT"])
    df.insert(0, "STT", range(1, len(df) + 1))
    for col in ["Link 1", "Link 2", "Link 3"]:
        if col not in df.columns:
            df[col] = ""
        else:
            df[col] = df[col].fillna("").astype(str)
    return df

def init_session_state():
    """Khởi tạo trạng thái ứng dụng một lần duy nhất."""
    defaults = {
        "is_admin": False,
        "excel_dir": DEFAULT_EXCEL_DIR,
        "lan_hieuchinh": "009",
        "main_menu": "1 🌐 DS WEBsites_CV",
        "web_tools_df": reindex_df(pd.DataFrame(DEFAULT_14_WEBS)),
        "data_store": {
            cat: reindex_df(pd.DataFrame([{
                "Thư mục / Hồ sơ": f"Hồ sơ {cat}",
                "Link 1": "https://drive.google.com",
                "Link 2": "",
                "Link 3": "",
                "Ghi chú": "Cập nhật định kỳ"
            }])) for cat in CATEGORIES
        },
        "bc_dinhky_df": reindex_df(pd.DataFrame([{
            "Tên Báo Cáo / Công Việc": "Báo cáo PCCC",
            "Tần suất": "Hàng Tháng",
            "Đơn vị nhận": "Phòng An toàn",
            "Link 1": "https://drive.google.com",
            "Link 2": "",
            "Link 3": "",
            "Ghi chú": "Nộp trước ngày 25"
        }])),
        "gsheet_df": reindex_df(pd.DataFrame([{
            "Mô tả Gsheet": "Bảng Theo Dõi CV",
            "Link 1": "https://docs.google.com/spreadsheets",
            "Link 2": "",
            "Link 3": "",
            "Ghi chú": "Dùng chung"
        }]))
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

init_session_state()

def generate_excel_download(df: pd.DataFrame, sheet_name: str) -> bytes:
    """Xuất file Excel chất lượng cao với định dạng đẹp mắt."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color="1E88E5", end_color="1E88E5", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        for col in worksheet.columns:
            max_len = max((len(str(cell.value or '')) for cell in col), default=0)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 60)

    return output.getvalue()

def open_local_folder(path: str):
    """Mở thư mục an toàn đa nền tảng (Windows/macOS/Linux)."""
    if not os.path.exists(path):
        st.error("❌ Đường dẫn không tồn tại trên hệ thống!")
        return
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')
        st.toast("📂 Đã mở thư mục!", icon="✅")
    except Exception as e:
        st.error(f"Không thể mở thư mục: {e}")

# ==========================================
# 3. HIỂN THỊ BẢNG NATIVE VỚI LINK COLUMN
# ==========================================
def render_interactive_table(df: pd.DataFrame):
    """Sử dụng Native Dataframe với Link Columns tiên tiến."""
    if df.empty:
        st.info("Chưa có dữ liệu.")
        return

    column_config = {
        "STT": st.column_config.NumberColumn("STT", width="small"),
        "Link 1": st.column_config.LinkColumn("Link 1 🔗", display_text="Truy cập 1"),
        "Link 2": st.column_config.LinkColumn("Link 2 🔗", display_text="Truy cập 2"),
        "Link 3": st.column_config.LinkColumn("Link 3 🔗", display_text="Truy cập 3"),
    }

    st.dataframe(
        df,
        column_config=column_config,
        use_container_width=True,
        hide_index=True,
        height=450
    )

# ==========================================
# 4. HỘP THOẠI DIALOG
# ==========================================
@st.dialog("🔐 Đăng Nhập Quản Trị Viên")
def admin_login_dialog():
    st.write("Nhập mật khẩu quản trị để kích hoạt tính năng chỉnh sửa:")
    pwd = st.text_input("Mật khẩu:", type="password")
    if st.button("Xác nhận", type="primary", use_container_width=True):
        if pwd == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.success("Đăng nhập thành công!")
            st.rerun()
        else:
            st.error("Mật khẩu không chính xác!")

@st.dialog("➕ Thêm Mới Dữ Liệu Website")
def add_web_dialog():
    mota = st.text_input("Mô tả WEB (*):")
    l1 = st.text_input("Link 1 (*):")
    l2 = st.text_input("Link 2:")
    l3 = st.text_input("Link 3:")
    ghichu = st.text_input("Ghi chú:")
    if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
        if mota:
            new_row = pd.DataFrame([{"Mô tả WEB": mota, "Link 1": l1, "Link 2": l2, "Link 3": l3, "Ghi chú": ghichu}])
            st.session_state.web_tools_df = reindex_df(pd.concat([st.session_state.web_tools_df, new_row], ignore_index=True))
            st.toast("Thêm thành công!", icon="✅")
            st.rerun()
        else:
            st.warning("Vui lòng nhập mô tả WEB!")

# ==========================================
# 5. THANH ĐIỀU HƯỚNG (SIDEBAR)
# ==========================================
with st.sidebar:
    st.title("🛡️ Quản Lý An Toàn TTM")
    st.caption(f"📌 Cập nhật: {TODAY_STR} | V{st.session_state.lan_hieuchinh}")
    st.divider()

    st.markdown("### 📂 PHÂN VÙNG LÀM VIỆC")
    menu_options = [
        "1 🌐 DS WEBsites_CV", 
        "2 📋 DM QL Files_CV", 
        "3 📊 DS BCdinhky_CV", 
        "4 🟢 DS Gsheet_CV"
    ]
    for opt in menu_options:
        if st.button(
            opt, 
            use_container_width=True, 
            type="primary" if st.session_state.main_menu == opt else "secondary"
        ):
            st.session_state.main_menu = opt
            st.rerun()

    main_menu = st.session_state.main_menu

    if st.session_state.is_admin:
        st.divider()
        st.markdown("### ⚙️ Cấu Hình Thư Mục")
        new_dir = st.text_input("Đường dẫn lưu file:", value=st.session_state.excel_dir)
        d1, d2 = st.columns(2)
        if d1.button("💾 Lưu", type="primary", use_container_width=True):
            st.session_state.excel_dir = new_dir
            st.toast("Đã cập nhật đường dẫn!", icon="✅")
        if d2.button("📂 Mở Thư Mục", use_container_width=True):
            open_local_folder(st.session_state.excel_dir)

    st.divider()
    st.markdown("### 👤 Trạng thái Hệ thống")
    if st.session_state.is_admin:
        st.success("Quyền hiện tại: ADMIN")
        if st.button("🔓 Thoát Admin", use_container_width=True):
            st.session_state.is_admin = False
            st.rerun()
    else:
        st.info("Quyền hiện tại: XEM (GUEST)")
        if st.button("🔐 Đăng nhập Admin", use_container_width=True):
            admin_login_dialog()

# ==========================================
# 6. BẢNG ĐIỀU KHUYỂN ADMIN & QUẢN LÝ FILE
# ==========================================
def render_admin_panel(df_key: str, export_file_name: str, cat_key: Optional[str] = None, title: str = "Dữ liệu"):
    st.markdown("---")
    st.markdown(f"### ⚙️ Bảng Điều Khiển Admin ({title})")
    
    df_target = st.session_state[df_key] if not cat_key else st.session_state[df_key][cat_key]

    col_del, col_tools = st.columns([2, 1])

    with col_del:
        with st.container(border=True):
            st.caption("📌 **Xóa Dòng Theo Số Thứ Tự (STT)**")
            stt_list = df_target["STT"].tolist() if not df_target.empty else []
            if stt_list:
                c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
                selected_stt = c1.selectbox("Chọn STT cần xóa:", stt_list, key=f"sel_{title}")
                if c2.button("🗑 Xóa dòng", use_container_width=True, key=f"del_{title}"):
                    idx = df_target[df_target["STT"] == selected_stt].index[0]
                    updated_df = reindex_df(df_target.drop(idx))
                    if cat_key:
                        st.session_state[df_key][cat_key] = updated_df
                    else:
                        st.session_state[df_key] = updated_df
                    st.toast("Đã xóa thành công!", icon="🗑️")
                    st.rerun()
            else:
                st.info("Bảng dữ liệu đang trống.")

    with col_tools:
        if st.button("🔄 Làm mới STT", use_container_width=True, key=f"re_{title}"):
            if cat_key:
                st.session_state[df_key][cat_key] = reindex_df(df_target)
            else:
                st.session_state[df_key] = reindex_df(df_target)
            st.rerun()

    with st.expander(f"📥 / 🚀 Tải lên & Tải xuống Excel ({title})"):
        c_down, c_up = st.columns(2)
        with c_down:
            st.markdown("##### 📥 Tải dữ liệu về máy")
            excel_data = generate_excel_download(df_target, "DATA")
            st.download_button(
                "📥 Tải File Excel",
                data=excel_data,
                file_name=export_file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with c_up:
            st.markdown("##### 🚀 Cập nhật từ File Excel")
            up_file = st.file_uploader("Chọn file Excel ghi đè:", type=["xlsx", "xls"], key=f"up_{title}")
            if up_file and st.button("🚀 Ghi đè dữ liệu", type="primary", use_container_width=True, key=f"up_btn_{title}"):
                try:
                    new_up_df = reindex_df(pd.read_excel(up_file))
                    if cat_key:
                        st.session_state[df_key][cat_key] = new_up_df
                    else:
                        st.session_state[df_key] = new_up_df
                    st.toast("Cập nhật file thành công!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Lỗi đọc file: {e}")

# ==========================================
# 7. NỘI DUNG CHÍNH (MAIN CONTENT)
# ==========================================
st.title(f"📌 {main_menu}")

if main_menu == "1 🌐 DS WEBsites_CV":
    render_interactive_table(st.session_state.web_tools_df)
    if st.session_state.is_admin:
        if st.button("➕ Thêm Mới Website", type="primary"):
            add_web_dialog()
        render_admin_panel("web_tools_df", export_file_name=f"DS_Websites_{DATE_STR}.xlsx", title="WEBSITE")

elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.selectbox("📌 Chọn Mảng Công Việc:", CATEGORIES)
    render_interactive_table(st.session_state.data_store[selected_cat])
    if st.session_state.is_admin:
        safe_name = selected_cat.replace(" ", "_")
        render_admin_panel(
            "data_store", 
            export_file_name=f"DM_Files_{safe_name}_{DATE_STR}.xlsx", 
            cat_key=selected_cat, 
            title=f"HOSO_{safe_name}"
        )

elif main_menu == "3 📊 DS BCdinhky_CV":
    render_interactive_table(st.session_state.bc_dinhky_df)
    if st.session_state.is_admin:
        render_admin_panel("bc_dinhky_df", export_file_name=f"DS_BaoCao_{DATE_STR}.xlsx", title="BAOCAO")

elif main_menu == "4 🟢 DS Gsheet_CV":
    render_interactive_table(st.session_state.gsheet_df)
    if st.session_state.is_admin:
        render_admin_panel("gsheet_df", export_file_name=f"DS_Gsheet_{DATE_STR}.xlsx", title="GSHEET")
