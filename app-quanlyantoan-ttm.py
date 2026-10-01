import io
from datetime import datetime
import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG VÀ SESSION STATE
# ==========================================
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thư mục mặc định ban đầu
DEFAULT_EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(DEFAULT_EXCEL_DIR):
    DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False
if "excel_dir" not in st.session_state:
    st.session_state.excel_dir = DEFAULT_EXCEL_DIR
if "lan_hieuchinh" not in st.session_state:
    st.session_state.lan_hieuchinh = "006"
if "main_menu" not in st.session_state:
    st.session_state.main_menu = "1 🌐 DS WEBsites_CV"
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "Dark"

# ==========================================
# 2. BỘ MÃ CSS TỐI ƯU HÓA
# ==========================================
is_dark = st.session_state.theme_mode == "Dark"

bg_app = "#0B0F19" if is_dark else "#F1F5F9"
text_app = "#F8FAFC" if is_dark else "#0F172A"
sidebar_bg = "#111827" if is_dark else "#FFFFFF"
sidebar_text = "#F8FAFC" if is_dark else "#0F172A"
sidebar_caption = "#94A3B8" if is_dark else "#475569"
sidebar_border = "rgba(255, 255, 255, 0.1)" if is_dark else "#CBD5E1"
btn_bg = "rgba(30, 136, 229, 0.1)" if is_dark else "rgba(30, 136, 229, 0.05)"
btn_text = "#90CAF9" if is_dark else "#1E88E5"
btn_border = "#1E88E5"

css_style = f"""
<style>
.block-container {{ 
    padding-top: 3.8rem !important; 
    padding-bottom: 1rem !important; 
}}
.stApp {{ 
    background-color: {bg_app} !important; 
    color: {text_app} !important; 
}}
.stApp p, .stApp span, .stApp div, .stApp label {{
    color: {text_app};
}}
header[data-testid="stHeader"] {{
    background: {'rgba(11, 15, 25, 0.85)' if is_dark else 'rgba(241, 245, 249, 0.85)'} !important;
    backdrop-filter: blur(10px) !important;
    border-bottom: 1px solid {sidebar_border} !important;
    z-index: 99999 !important;
}}

/* Trả lại thiết kế gốc ổn định cho nút Menu Mở rộng/Thu gọn của Streamlit */
[data-testid="collapsedControl"], [data-testid="stSidebarCollapseButton"] {{
    transition: transform 0.2s ease !important;
}}
[data-testid="collapsedControl"]:hover, [data-testid="stSidebarCollapseButton"]:hover {{
    transform: scale(1.1) !important;
}}

/* CẤU HÌNH SIDEBAR VÀ BẢNG */
[data-testid="stSidebar"] {{
    background-color: {sidebar_bg} !important;
    border-right: 1px solid {sidebar_border} !important;
}}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h3, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] div {{
    color: {sidebar_text} !important;
}}

[data-testid="stSidebar"] div.stButton > button {{
    width: 100% !important; text-align: left !important; justify-content: flex-start !important;
    padding: 12px 16px !important; font-weight: 600 !important; border-radius: 10px !important;
    transition: all 0.3s ease !important;
}}

[data-testid="stSidebar"] div.stButton > button[kind="secondary"] {{
    border: 1px solid {btn_border} !important; 
    background: {btn_bg} !important; 
    color: {btn_text} !important;
}}
[data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {{
    background: rgba(30, 136, 229, 0.2) !important;
    transform: translateX(4px) !important;
}}

[data-testid="stSidebar"] div.stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important; 
    color: #FFFFFF !important;
    border: 1px solid #42A5F5 !important; 
}}

/* Nút Link trong Bảng */
.slide-link-btn {{
    display: inline-flex !important; align-items: center !important; justify-content: space-between !important;
    gap: 8px !important; padding: 6px 14px !important; margin: 3px 4px !important;
    background: linear-gradient(135deg, rgba(30, 136, 229, 0.15) 0%, rgba(21, 101, 192, 0.3) 100%) !important;
    border: 1px solid rgba(66, 165, 245, 0.5) !important; border-radius: 20px !important;
    color: {'#E3F2FD' if is_dark else '#1565C0'} !important; text-decoration: none !important;
    font-size: 13px !important; font-weight: 600 !important;
}}
.slide-link-btn:hover {{
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important; color: #FFFFFF !important;
}}
.slide-link-btn .btn-knob {{
    background: rgba(255, 255, 255, 0.25) !important; border-radius: 50% !important; padding: 2px 6px; font-size: 10px !important;
}}
</style>
"""
st.markdown(css_style, unsafe_allow_html=True)


# ==========================================
# 3. DỮ LIỆU VÀ HÀM TRỢ GIÚP (TÍCH HỢP XUẤT EXCEL KIỂU 2)
# ==========================================
TODAY_STR = datetime.now().strftime("%d/%m/%Y")
DATE_STR = datetime.now().strftime("%Y%m%d")

CATEGORIES = [
    "DTTU_01 AT", "DTTU_01 AT 01 Bao cao", "DTTU_01 AT 01 Bao cao 2026", "DTTU_02 PCTT", "DTTU_03 PCCC", "DTTU_04 HL", "DTTU_05 ATDTXD",
    "DTTU_ATGT", "DTTU_CNTT", "DTTU_DCAT", "DTTU_DCNN va Cac loai xe", "DTTU_DGRR", "DTTU_HNTH cac loai", "DTTU_KIEM TRA",
    "DTTU_KIEM TRA-Thuc hien Kien Nghi", "DTTU_UCKC", "DTTU_UCKC dien tap cac loai", "DTTU_khac 01 PHOI HOP CAC TO",
    "DTTU_khac 02 XEM DE BIET CTY", "DTTU_khac 03 ATD dia phuong", "DTTU_khac 03 XEM DE BIET dia phuong", "Quy dinh 0000 Discussion",
    "Quy dinh GOV", "Quy dinh PCTN", "Quy dinh PCTN file tham khao cac Doi", "Quy dinh SPC va EVN", "Quy dinh trao doi EVN-SPC-PCTN"
]

DEFAULT_14_WEBS = [
    {"Mô tả WEB": "Hệ thống D-Office", "Link 1": "https://doffice.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý văn bản điều hành"},
    {"Mô tả WEB": "Cổng thông tin Điện lực (EVN SPC)", "Link 1": "https://evnspc.vn", "Link 2": "https://www.congcuweb.net/", "Link 3": "", "Ghi chú": "Tra cứu quy định & chỉ đạo"},
    {"Mô tả WEB": "Quản lý An toàn (ECP / ATLD)", "Link 1": "https://giamsatantoan.evnspc.vn/Home/Index", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý giám sát an toàn SPC"},
    {"Mô tả WEB": "Lịch công tác / Lịch tuần", "Link 1": "https://lichtuan.evnspc.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Công ty Điện lực Tây Ninh"},
    {"Mô tả WEB": "Hệ thống PMIS", "Link 1": "https://pmis.evn.com.vn", "Link 2": "", "Link 3": "", "Ghi chú": "Quản lý vận hành thiết bị & lưới điện"},
    {"Mô tả WEB": "Tritm.la Dashboard", "Link 1": "https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit#gid=964445540", "Link 2": "", "Link 3": "", "Ghi chú": "Google sheet CV"},
]

def reindex_df(df):
    if df is None or df.empty: return pd.DataFrame()
    df = df.dropna(how="all").reset_index(drop=True)
    if "STT" in df.columns: df = df.drop(columns=["STT"])
    df.insert(0, "STT", range(1, len(df) + 1))
    for col in ["Link 1", "Link 2", "Link 3"]:
        if col not in df.columns: df[col] = ""
        else: df[col] = df[col].fillna("")
    return df

# Cập nhật hàm xuất Excel chuẩn Kiểu 2 (Màu xanh, tách cột)
def generate_excel_download(df, sheet_name):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Xuất dữ liệu thô (Các cột STT, Mô tả, Link 1, Link 2, Link 3, Ghi chú)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        workbook = writer.book
        worksheet = workbook[sheet_name]
        
        # Style Header: Nền Xanh lá cây (Kiểu 2), Chữ đen đậm
        header_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        header_font = Font(bold=True, color="000000")
        header_alignment = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]: # Dòng 1 là Header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment

        # Căn chỉnh tự động độ rộng cột
        for col in worksheet.columns:
            max_len = max((len(str(cell.value)) for cell in col if cell.value), default=0)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 60)

    return output.getvalue()

def create_responsive_table(df, main_col_name, extra_cols=None):
    if extra_cols is None: extra_cols = []
    dark = st.session_state.theme_mode == "Dark"
    tb_bg = "#111827" if dark else "#FFFFFF"
    tb_text = "#F8FAFC" if dark else "#0F172A"
    th_bg = "#1F2937" if dark else "#F1F5F9"
    th_text = "#60A5FA" if dark else "#1D4ED8"
    row_border = "rgba(255,255,255,0.06)" if dark else "#F1F5F9"

    html = f'<div style="max-height: 580px; overflow-y: auto; overflow-x: auto; margin-bottom: 20px; border-radius: 12px; background-color: {tb_bg}; box-shadow: 0 8px 24px rgba(0,0,0,0.08);">'
    html += f'<table style="width: 100%; border-collapse: separate; border-spacing: 0; font-size: 15px; text-align: left; color: {tb_text};">'
    th_style = f"padding: 14px 16px; color: {th_text}; position: sticky; top: 0; background-color: {th_bg}; z-index: 10; border-bottom: 2px solid #1E88E5; font-weight: 700;"
    
    html += f'<thead><tr><th style="{th_style}">STT</th><th style="{th_style} min-width: 260px;">{main_col_name}</th>'
    for ec in extra_cols: html += f'<th style="{th_style}">{ec}</th>'
    html += f'<th style="{th_style} min-width: 220px;">Links truy cập</th><th style="{th_style} min-width: 200px;">Ghi chú</th></tr></thead><tbody>'

    if df.empty:
        html += '<tr><td colspan="10" style="padding: 20px; text-align: center; color: #64748B;">Chưa có dữ liệu</td></tr>'
    else:
        for idx, row in df.iterrows():
            html += f'<tr style="border-bottom: 1px solid {row_border};">'
            html += f'<td style="padding: 12px 16px; font-weight: bold; border-bottom: 1px solid {row_border};">{row.get("STT", "")}</td>'
            html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get(main_col_name, "")}</td>'
            for ec in extra_cols: html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get(ec, "")}</td>'

            links, l1, l2, l3 = [], row.get("Link 1", ""), row.get("Link 2", ""), row.get("Link 3", "")
            if l1 and str(l1).strip().startswith("http"): links.append(f'<a href="{str(l1).strip()}" target="_blank" class="slide-link-btn"><span>Link 1</span><span class="btn-knob">❯</span></a>')
            if l2 and str(l2).strip().startswith("http"): links.append(f'<a href="{str(l2).strip()}" target="_blank" class="slide-link-btn"><span>Link 2</span><span class="btn-knob">❯</span></a>')
            if l3 and str(l3).strip().startswith("http"): links.append(f'<a href="{str(l3).strip()}" target="_blank" class="slide-link-btn"><span>Link 3</span><span class="btn-knob">❯</span></a>')
            
            html += f'<td style="padding: 10px 16px; border-bottom: 1px solid {row_border};">{"".join(links) if links else "-"}</td>'
            html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get("Ghi chú", "")}</td></tr>'
    html += "</tbody></table></div>"
    return html

# Khởi tạo Session Storage
if "web_tools_df" not in st.session_state: st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
if "data_store" not in st.session_state:
    st.session_state.data_store = {cat: reindex_df(pd.DataFrame([{"Thư mục / Hồ sơ": f"Hồ sơ {cat}", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Cập nhật định kỳ"}])) for cat in CATEGORIES}
if "bc_dinhky_df" not in st.session_state:
    st.session_state.bc_dinhky_df = reindex_df(pd.DataFrame([{"Tên Báo Cáo / Công Việc": "Báo cáo PCCC", "Tần suất": "Hàng Tháng", "Đơn vị nhận": "Phòng An toàn", "Link 1": "https://drive.google.com", "Link 2": "", "Link 3": "", "Ghi chú": "Nộp trước ngày 25"}]))
if "gsheet_df" not in st.session_state:
    st.session_state.gsheet_df = reindex_df(pd.DataFrame([{"Mô tả Gsheet": "Bảng Theo Dõi CV", "Link 1": "https://docs.google.com/spreadsheets", "Link 2": "", "Link 3": "", "Ghi chú": "Dùng chung"}]))

# ==========================================
# 4. DIALOGS (HỘP THOẠI ADMIN)
# ==========================================
@st.dialog("🔐 Đăng Nhập Quản Trị Viên")
def admin_login_dialog():
    st.write("Nhập mật khẩu để kích hoạt tính năng thêm/sửa/xóa.")
    if st.button("Xác nhận", type="primary", use_container_width=True) if (pwd := st.text_input("Mật khẩu:", type="password")) else False:
        if pwd == "admin123":
            st.session_state.is_admin = True
            st.success("Đăng nhập thành công!")
            st.rerun()
        else: st.error("Mật khẩu sai!")

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

# ==========================================
# 5. SIDEBAR
# ==========================================
with st.sidebar:
    st.title("🛡️ Quản Lý An Toàn TTM")
    st.caption(f"📌 Cập nhật: {TODAY_STR} _ V{st.session_state.lan_hieuchinh}")
    st.divider()

    st.markdown("### 🎨 CHẾ ĐỘ GIAO DIỆN")
    c1, c2 = st.columns(2)
    if c1.button("☀️ Sáng", use_container_width=True, type="primary" if st.session_state.theme_mode == "Light" else "secondary"): st.session_state.theme_mode = "Light"; st.rerun()
    if c2.button("🌙 Tối", use_container_width=True, type="primary" if st.session_state.theme_mode == "Dark" else "secondary"): st.session_state.theme_mode = "Dark"; st.rerun()
    st.divider()

    st.markdown("### 📂 PHÂN VÙNG LÀM VIỆC")
    for item in ["1 🌐 DS WEBsites_CV", "2 📋 DM QL Files_CV", "3 📊 DS BCdinhky_CV", "4 🟢 DS Gsheet_CV"]:
        if st.button(item, use_container_width=True, type="primary" if st.session_state.main_menu == item else "secondary"): st.session_state.main_menu = item; st.rerun()
    
    main_menu = st.session_state.main_menu

    if st.session_state.is_admin:
        st.divider()
        st.markdown("### ⚙️ Cấu Hình Thư Mục")
        new_dir = st.text_input("Đường dẫn lưu file cục bộ:", value=st.session_state.excel_dir)
        d1, d2 = st.columns(2)
        if d1.button("💾 Xác nhận", type="primary", use_container_width=True): st.session_state.excel_dir = new_dir; st.toast("Đã lưu!", icon="✅")
        if d2.button("📂 Mở thư mục", use_container_width=True):
            if os.path.exists(st.session_state.excel_dir): os.startfile(st.session_state.excel_dir)
            else: st.error("Đường dẫn không tồn tại!")

    st.divider()
    st.markdown("### 👤 Người dùng: `ttm`")
    if st.session_state.is_admin: st.success("Quyền: ADMIN")
    
    s1, s2 = st.columns(2)
    if s1.button("🚪 Đăng xuất", use_container_width=True): st.info("Đã đăng xuất!")
    if s2.button("🧹 Xóa Cache", use_container_width=True): st.cache_data.clear(); st.success("Đã xóa!")

    if st.session_state.is_admin:
        if st.button("🔓 Thoát Admin", use_container_width=True): st.session_state.is_admin = False; st.rerun()
    else:
        if st.button("🔐 Đăng nhập Quản trị", use_container_width=True): admin_login_dialog()

# ==========================================
# 6. MAIN LAYOUT (RENDER DATA & ADMIN PANELS)
# ==========================================

def render_admin_panel(df_key, export_file_name, cat_key=None, title="Dữ liệu"):
    st.markdown("---")
    st.markdown(f"### ⚙ Bảng Điều Khiển Admin ({title})")
    col_add, col_edit = st.columns([1, 2])
    
    df_target = st.session_state[df_key] if not cat_key else st.session_state[df_key][cat_key]

    with col_add:
        st.caption("📌 **Quản lý Hệ thống**")
        if st.button(f"🔄 Làm mới STT", use_container_width=True, key=f"re_{title}"):
            if cat_key: st.session_state[df_key][cat_key] = reindex_df(df_target)
            else: st.session_state[df_key] = reindex_df(df_target)
            st.rerun()

    with col_edit:
        with st.container(border=True):
            st.caption("📌 **Xóa dữ liệu theo STT** (Vui lòng Nhập File để Sửa/Thêm hàng loạt)")
            stt_list = df_target["STT"].tolist() if not df_target.empty else []
            if stt_list:
                c1, c2 = st.columns([3, 1], vertical_alignment="bottom")
                selected_stt = c1.selectbox("Chọn STT cần Xóa:", stt_list, key=f"sel_{title}")
                if c2.button("🗑 Xóa dòng này", use_container_width=True, key=f"del_{title}"):
                    idx = df_target[df_target["STT"] == selected_stt].index[0]
                    new_df = reindex_df(df_target.drop(idx))
                    if cat_key: st.session_state[df_key][cat_key] = new_df
                    else: st.session_state[df_key] = new_df
                    st.rerun()
            else: st.info("Bảng trống.")

    with st.expander(f"⚙️ Tải xuống / Tải lên Excel ({title})", expanded=False):
        c_up1, c_up2 = st.columns(2)
        with c_up1:
            st.markdown("#### 📥 Tải xuống (Backup Kiểu 2)")
            excel_data = generate_excel_download(df_target, "DATA")
            st.download_button("📥 Click để Tải File Excel", data=excel_data, file_name=export_file_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
        with c_up2:
            st.markdown("#### 🚀 Tải lên (Update)")
            up_file = st.file_uploader("Chọn file Excel để ghi đè dữ liệu:", type=["xlsx", "xls"], key=f"up_{title}")
            if up_file and st.button("🚀 Cập nhật từ File", type="primary", use_container_width=True, key=f"up_btn_{title}"):
                try:
                    new_up_df = reindex_df(pd.read_excel(up_file))
                    if cat_key: st.session_state[df_key][cat_key] = new_up_df
                    else: st.session_state[df_key] = new_up_df
                    st.toast("🎉 Thành công!", icon="✅")
                    st.rerun()
                except Exception as e: st.error(f"Lỗi: {e}")

# --- HIỂN THỊ THEO MENU ---
if main_menu == "1 🌐 DS WEBsites_CV":
    st.markdown(create_responsive_table(st.session_state.web_tools_df, "Mô tả WEB"), unsafe_allow_html=True)
    if st.session_state.is_admin: 
        c1, c2 = st.columns(2)
        if c1.button("➕ Thêm mới Website", type="primary"): add_web_dialog()
        render_admin_panel("web_tools_df", export_file_name=f"1 DS WEBsites_CV out_{DATE_STR}.xlsx", title="WEBSITES")

elif main_menu == "2 📋 DM QL Files_CV":
    selected_cat = st.selectbox("📌 Chọn Mảng Công Việc:", CATEGORIES)
    st.markdown(create_responsive_table(st.session_state.data_store[selected_cat], "Thư mục / Hồ sơ"), unsafe_allow_html=True)
    if st.session_state.is_admin: 
        safe_name = selected_cat.replace(" ", "_")
        render_admin_panel("data_store", export_file_name=f"2 DM QL Files_CV_{safe_name}_out_{DATE_STR}.xlsx", cat_key=selected_cat, title=f"HOSO_{safe_name}")

elif main_menu == "3 📊 DS BCdinhky_CV":
    st.markdown(create_responsive_table(st.session_state.bc_dinhky_df, "Tên Báo Cáo / Công Việc", ["Tần suất", "Đơn vị nhận"]), unsafe_allow_html=True)
    if st.session_state.is_admin: 
        render_admin_panel("bc_dinhky_df", export_file_name=f"3 DS BCdinhky_CV out_{DATE_STR}.xlsx", title="BAOCAO")

elif main_menu == "4 🟢 DS Gsheet_CV":
    st.markdown(create_responsive_table(st.session_state.gsheet_df, "Mô tả Gsheet"), unsafe_allow_html=True)
    if st.session_state.is_admin: 
        render_admin_panel("gsheet_df", export_file_name=f"4 DS Gsheet_CV out_{DATE_STR}.xlsx", title="GSHEET")
