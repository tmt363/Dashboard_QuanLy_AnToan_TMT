import io
from datetime import datetime
import os
import openpyxl
import pandas as pd
import streamlit as st

# ==========================================
# 1. CẤU HÌNH TRANG VÀ SESSION STATE
# ==========================================
st.set_page_config(
    page_title="Quản Lý An Toàn TTM",
    page_icon="🛡️️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thư mục mặc định ban đầu
DEFAULT_EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(DEFAULT_EXCEL_DIR):
  DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

# Khởi tạo các biến session state
if "is_admin" not in st.session_state:
  st.session_state.is_admin = False
if "excel_dir" not in st.session_state:
  st.session_state.excel_dir = DEFAULT_EXCEL_DIR
if "lan_hieuchinh" not in st.session_state:
  st.session_state.lan_hieuchinh = "001"
if "main_menu" not in st.session_state:
  st.session_state.main_menu = "1 🌐 DS WEBsites_CV"
if "theme_mode" not in st.session_state:
  st.session_state.theme_mode = "Dark"

# ==========================================
# 2. BỘ MÃ CSS TỐI ƯU GIAO DIỆN VÀ TƯƠNG PHẢN
# ==========================================
is_dark = st.session_state.theme_mode == "Dark"

# Bảng màu tương phản chuẩn UI/UX
bg_app = "#0B0F19" if is_dark else "#F1F5F9"
text_app = "#F8FAFC" if is_dark else "#0F172A"

sidebar_bg = "#111827" if is_dark else "#FFFFFF"
sidebar_text = "#F8FAFC" if is_dark else "#0F172A"  # Chữ cực rõ trên nền sáng/tối
sidebar_caption = "#94A3B8" if is_dark else "#475569"
sidebar_border = "rgba(255, 255, 255, 0.1)" if is_dark else "#CBD5E1"

btn_bg = "rgba(255, 255, 255, 0.05)" if is_dark else "#F8FAFC"
btn_text = "#E2E8F0" if is_dark else "#1E293B"
btn_border = "rgba(255, 255, 255, 0.12)" if is_dark else "#94A3B8"

css_style = f"""
<style>
/* Khoảng cách chính tránh che Header */
.block-container {{ 
    padding-top: 3.8rem !important; 
    padding-bottom: 1rem !important; 
}}

/* Dynamic App Theme Background & Text */
.stApp {{ 
    background-color: {bg_app} !important; 
    color: {text_app} !important; 
}}

.stApp p, .stApp span, .stApp div, .stApp label {{
    color: {text_app};
}}

/* ---------------------------------------------------- */
/* 🌟 1. NÚT MỞ SIDEBAR ( >> ) NỔI BẬT CHỐNG ẨN        */
/* ---------------------------------------------------- */
[data-testid="stSidebarCollapsedControl"] {{
    position: fixed !important;
    top: 12px !important;
    left: 12px !important;
    z-index: 99999 !important;
}}

[data-testid="stSidebarCollapsedControl"] button {{
    background: linear-gradient(135deg, #00C6FF 0%, #0072FF 100%) !important;
    border: 2px solid #FFFFFF !important;
    border-radius: 10px !important;
    box-shadow: 0 0 15px rgba(0, 198, 255, 0.8) !important;
    padding: 6px 10px !important;
    transition: all 0.3s ease !important;
}}

[data-testid="stSidebarCollapsedControl"] button:hover {{
    transform: scale(1.1) !important;
    box-shadow: 0 0 22px rgba(0, 198, 255, 1) !important;
}}

/* Ép biểu tượng mũi tên >> bên trong nút thành màu trắng rõ nét */
[data-testid="stSidebarCollapsedControl"] button svg {{
    fill: #FFFFFF !important;
    color: #FFFFFF !important;
    stroke: #FFFFFF !important;
    width: 22px !important;
    height: 22px !important;
}}

/* ---------------------------------------------------- */
/* ⚡ 2. NÚT ĐÓNG SIDEBAR ( << ) BÊN TRONG SIDEBAR      */
/* ---------------------------------------------------- */
[data-testid="stSidebarCollapseButton"] button,
button[aria-label="Close sidebar"],
button[aria-label="Collapse sidebar"] {{
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important;
    color: #FFFFFF !important;
    border-radius: 50% !important;
    border: 1px solid #42A5F5 !important;
    box-shadow: 0 0 10px rgba(30, 136, 229, 0.5) !important;
    transition: all 0.3s ease !important;
}}

[data-testid="stSidebarCollapseButton"] button svg {{
    fill: #FFFFFF !important;
    color: #FFFFFF !important;
}}

[data-testid="stSidebarCollapseButton"] button:hover {{
    transform: scale(1.1) rotate(-90deg) !important;
    box-shadow: 0 0 18px rgba(66, 165, 245, 0.9) !important;
}}

/* ---------------------------------------------------- */
/* 🎨 3. ÉP MÀU CHỮ SIDEBAR CHUẨN HIỂN THỊ CHỐNG LÓA    */
/* ---------------------------------------------------- */
[data-testid="stSidebar"] {{
    background-color: {sidebar_bg} !important;
    border-right: 1px solid {sidebar_border} !important;
}}

[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3, 
[data-testid="stSidebar"] p, 
[data-testid="stSidebar"] span, 
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div {{
    color: {sidebar_text} !important;
}}

[data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
    color: {sidebar_caption} !important;
}}

[data-testid="stSidebar"] div.stButton > button {{
    width: 100% !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 12px 16px !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    border-radius: 10px !important;
    margin-bottom: 6px !important;
    transition: all 0.3s ease !important;
    border: 1px solid {btn_border} !important;
    background: {btn_bg} !important;
    color: {btn_text} !important;
}}

[data-testid="stSidebar"] div.stButton > button:hover {{
    background: rgba(30, 136, 229, 0.15) !important;
    color: #1E88E5 !important;
    border-color: rgba(30, 136, 229, 0.6) !important;
    transform: translateX(4px) !important;
}}

[data-testid="stSidebar"] div.stButton > button[kind="primary"],
[data-testid="stSidebar"] div.stButton > button[data-testid="baseButton-primary"] {{
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important;
    color: #FFFFFF !important;
    border: 1px solid #42A5F5 !important;
    box-shadow: 0 4px 15px rgba(30, 136, 229, 0.4) !important;
    font-weight: 700 !important;
}}

/* ---------------------------------------------------- */
/* 🛠️ 4. HEADER VÀ NÚT LINK                              */
/* ---------------------------------------------------- */
header[data-testid="stHeader"] {{
    background: {'rgba(11, 15, 25, 0.85)' if is_dark else 'rgba(241, 245, 249, 0.85)'} !important;
    backdrop-filter: blur(10px) !important;
    border-bottom: 1px solid {sidebar_border} !important;
}}

.slide-link-btn {{
    display: inline-flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 8px !important;
    padding: 6px 14px !important;
    margin: 3px 4px !important;
    background: linear-gradient(135deg, rgba(30, 136, 229, 0.15) 0%, rgba(21, 101, 192, 0.3) 100%) !important;
    border: 1px solid rgba(66, 165, 245, 0.5) !important;
    border-radius: 20px !important;
    color: {'#E3F2FD' if is_dark else '#1565C0'} !important;
    text-decoration: none !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1) !important;
    transition: all 0.3s ease !important;
}}

.slide-link-btn:hover {{
    background: linear-gradient(135deg, #1E88E5 0%, #1565C0 100%) !important;
    color: #FFFFFF !important;
    border-color: #64B5F6 !important;
    box-shadow: 0 4px 14px rgba(30, 136, 229, 0.5) !important;
    transform: translateX(4px) scale(1.02) !important;
}}

.slide-link-btn .btn-knob {{
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    width: 18px !important;
    height: 18px !important;
    background: rgba(255, 255, 255, 0.25) !important;
    border-radius: 50% !important;
    font-size: 10px !important;
}}
</style>
"""
st.markdown(css_style, unsafe_allow_html=True)


# ==========================================
# 3. DỮ LIỆU VÀ HÀM TRỢ GIÚP
# ==========================================
TODAY_STR = datetime.now().strftime("%d/%m/%Y")

CATEGORIES = [
    "DTTU_01 AT",
    "DTTU_01 AT 01 Bao cao",
    "DTTU_01 AT 01 Bao cao 2026",
    "DTTU_02 PCTT",
    "DTTU_03 PCCC",
    "DTTU_04 HL",
    "DTTU_05 ATDTXD",
    "DTTU_ATGT",
    "DTTU_CNTT",
    "DTTU_DCAT",
    "DTTU_DCNN va Cac loai xe",
    "DTTU_DGRR",
    "DTTU_HNTH cac loai",
    "DTTU_KIEM TRA",
    "DTTU_KIEM TRA-Thuc hien Kien Nghi",
    "DTTU_UCKC",
    "DTTU_UCKC dien tap cac loai",
    "DTTU_khac 01 PHOI HOP CAC TO",
    "DTTU_khac 02 XEM DE BIET CTY",
    "DTTU_khac 03 ATD dia phuong",
    "DTTU_khac 03 XEM DE BIET dia phuong",
    "Quy dinh 0000 Discussion",
    "Quy dinh GOV",
    "Quy dinh PCTN",
    "Quy dinh PCTN file tham khao cac Doi",
    "Quy dinh SPC va EVN",
    "Quy dinh trao doi EVN-SPC-PCTN",
]

DEFAULT_14_WEBS = [
    {
        "Mô tả WEB": "Hệ thống D-Office",
        "Link 1": "https://doffice.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Quản lý văn bản điều hành",
    },
    {
        "Mô tả WEB": "Cổng thông tin Điện lực (EVN SPC)",
        "Link 1": "https://evnspc.vn",
        "Link 2": "https://www.congcuweb.net/",
        "Link 3": "",
        "Ghi chú": "Tra cứu quy định & chỉ đạo",
    },
    {
        "Mô tả WEB": "Quản lý An toàn (ECP / ATLD)",
        "Link 1": "https://giamsatantoan.evnspc.vn/Home/Index",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Quản lý giám sát an toàn SPC",
    },
    {
        "Mô tả WEB": "Lịch công tác / Lịch tuần",
        "Link 1": "https://lichtuan.evnspc.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Công ty Điện lực Tây Ninh",
    },
    {
        "Mô tả WEB": "Hệ thống PMIS",
        "Link 1": "https://pmis.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Quản lý vận hành thiết bị & lưới điện",
    },
    {
        "Mô tả WEB": "Tritm.la Dashboard 2026 DTTU",
        "Link 1": (
            "https://docs.google.com/spreadsheets/d/1gVAroFIytWwrBMCScYuXWbzlS1ZNXrPY4Pcgb__Dv-c/edit?gid=964445540#gid=964445540"
        ),
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Google sheet CV",
    },
    {
        "Mô tả WEB": "Hệ thống Giám sát Thiên tai Việt Nam",
        "Link 1": "https://vndms.gov.vn/",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Cảnh báo và phòng chống thiên tai",
    },
    {
        "Mô tả WEB": "Hệ thống HRMS",
        "Link 1": "https://hrms.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Quản lý lao động tiền lương",
    },
    {
        "Mô tả WEB": "Hệ thống E-Learning",
        "Link 1": "https://elearning.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Huấn luyện an toàn & thi trực tuyến",
    },
    {
        "Mô tả WEB": "Cổng Dịch vụ công Quốc gia",
        "Link 1": "https://dichvucong.gov.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Thực hiện thủ tục hành chính PCCC/ĐTXD",
    },
]


def reindex_df(df):
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
      df[col] = df[col].fillna("")
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


def generate_excel_download(df, sheet_name):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    auto_fit_columns(writer.book)
  return output.getvalue()


# BẢNG DỮ LIỆU CẢI TIẾN TƯƠNG PHẢN ĐẸP MẮT
def create_responsive_table(df, main_col_name, extra_cols=None):
  if extra_cols is None:
    extra_cols = []

  dark = st.session_state.theme_mode == "Dark"
  tb_bg = "#111827" if dark else "#FFFFFF"
  tb_text = "#F8FAFC" if dark else "#0F172A"
  th_bg = "#1F2937" if dark else "#F1F5F9"
  th_text = "#60A5FA" if dark else "#1D4ED8"
  border_col = "rgba(255,255,255,0.1)" if dark else "#E2E8F0"
  row_border = "rgba(255,255,255,0.06)" if dark else "#F1F5F9"

  html = f"""
    <div style="max-height: 580px; overflow-y: auto; overflow-x: auto; margin-bottom: 20px; border-radius: 12px; border: 1px solid {border_col}; background-color: {tb_bg}; box-shadow: 0 8px 24px rgba(0,0,0,0.08);">
    <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-size: 15px; text-align: left; color: {tb_text};">
    """

  th_style = f"padding: 14px 16px; color: {th_text}; position: sticky; top: 0; background-color: {th_bg}; z-index: 10; border-bottom: 2px solid #1E88E5; box-shadow: 0 2px 4px rgba(0,0,0,0.05); font-weight: 700; font-size: 15px; vertical-align: middle;"

  html += "<thead><tr>"
  html += f'<th style="{th_style} white-space: nowrap;">STT</th>'
  html += f'<th style="{th_style} min-width: 260px;">{main_col_name}</th>'
  for ec in extra_cols:
    html += f'<th style="{th_style} white-space: nowrap;">{ec}</th>'
  html += f'<th style="{th_style} min-width: 220px;">Links truy cập</th>'
  html += f'<th style="{th_style} min-width: 200px;">Ghi chú</th>'
  html += "</tr></thead><tbody>"

  if df.empty:
    html += f'<tr><td colspan="10" style="padding: 20px; text-align: center; color: #64748B;">Chưa có dữ liệu</td></tr>'
  else:
    for idx, row in df.iterrows():
      html += f'<tr style="border-bottom: 1px solid {row_border};">'
      html += f'<td style="padding: 12px 16px; font-weight: bold; border-bottom: 1px solid {row_border};">{row.get("STT", "")}</td>'
      html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get(main_col_name, "")}</td>'
      for ec in extra_cols:
        html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get(ec, "")}</td>'

      links = []
      l1, l2, l3 = (
          row.get("Link 1", ""),
          row.get("Link 2", ""),
          row.get("Link 3", ""),
      )
      if l1 and str(l1).strip().startswith("http"):
        links.append(
            f'<a href="{str(l1).strip()}" target="_blank"'
            ' class="slide-link-btn"><span>Link 1</span><span'
            ' class="btn-knob">❯</span></a>'
        )
      if l2 and str(l2).strip().startswith("http"):
        links.append(
            f'<a href="{str(l2).strip()}" target="_blank"'
            ' class="slide-link-btn"><span>Link 2</span><span'
            ' class="btn-knob">❯</span></a>'
        )
      if l3 and str(l3).strip().startswith("http"):
        links.append(
            f'<a href="{str(l3).strip()}" target="_blank"'
            ' class="slide-link-btn"><span>Link 3</span><span'
            ' class="btn-knob">❯</span></a>'
        )

      links_str = (
          "".join(links)
          if links
          else '<span style="color: #94A3B8; font-style: italic;">-</span>'
      )

      html += f'<td style="padding: 10px 16px; border-bottom: 1px solid {row_border};">{links_str}</td>'
      html += f'<td style="padding: 12px 16px; border-bottom: 1px solid {row_border};">{row.get("Ghi chú", "")}</td>'
      html += "</tr>"
  html += "</tbody></table></div>"
  return html


# Kho dữ liệu tĩnh ban đầu
if "web_tools_df" not in st.session_state:
  st.session_state.web_tools_df = reindex_df(pd.DataFrame(DEFAULT_14_WEBS))
if "data_store" not in st.session_state:
  st.session_state.data_store = {}
  for cat in CATEGORIES:
    st.session_state.data_store[cat] = reindex_df(
        pd.DataFrame([{
            "Thư mục / Hồ sơ": f"Hồ sơ {cat}",
            "Link 1": "https://drive.google.com",
            "Link 2": "",
            "Link 3": "",
            "Ghi chú": "Cập nhật định kỳ",
        }])
    )
if "bc_dinhky_df" not in st.session_state:
  st.session_state.bc_dinhky_df = reindex_df(
      pd.DataFrame([
          {
              "Tên Báo Cáo / Công Việc": (
                  "Báo cáo công tác An toàn định kỳ Quý"
              ),
              "Tần suất": "Hàng Quý",
              "Đơn vị nhận": "Công ty Điện lực",
              "Link 1": "https://drive.google.com",
              "Link 2": "",
              "Link 3": "",
              "Ghi chú": "Nộp trước ngày 20 cuối quý",
          },
          {
              "Tên Báo Cáo / Công Việc": "Báo cáo công tác PCCC & CNCH",
              "Tần suất": "Hàng Tháng",
              "Đơn vị nhận": "Phòng An toàn",
              "Link 1": "https://drive.google.com",
              "Link 2": "",
              "Link 3": "",
              "Ghi chú": "Nộp trước ngày 25 hàng tháng",
          },
      ])
  )
if "gsheet_df" not in st.session_state:
  st.session_state.gsheet_df = reindex_df(
      pd.DataFrame([
          {
              "Mô tả Gsheet": "Bảng Theo Dõi Công Việc Theo Tuần",
              "Link 1": "https://docs.google.com/spreadsheets",
              "Link 2": "",
              "Link 3": "",
              "Ghi chú": "Dùng chung phòng An Toàn",
          },
          {
              "Mô tả Gsheet": "Theo Dõi Kiến Nghị Kiểm Tra",
              "Link 1": "https://docs.google.com/spreadsheets",
              "Link 2": "",
              "Link 3": "",
              "Ghi chú": "Cập nhật trực tuyến",
          },
      ])
  )


# ==========================================
# 4. DIALOGS (HỘP THOẠI ADMIN & CHỈNH SỬA)
# ==========================================
@st.dialog("🔐 Đăng Nhập Quản Trị Viên")
def admin_login_dialog():
  st.write("Vui lòng nhập mật khẩu để kích hoạt tính năng admin.")
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
      new_row = pd.DataFrame([{
          "Mô tả WEB": mota,
          "Link 1": l1,
          "Link 2": l2,
          "Link 3": l3,
          "Ghi chú": ghichu,
      }])
      st.session_state.web_tools_df = reindex_df(
          pd.concat([st.session_state.web_tools_df, new_row], ignore_index=True)
      )
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
    st.session_state.web_tools_df.loc[
        idx, ["Mô tả WEB", "Link 1", "Link 2", "Link 3", "Ghi chú"]
    ] = [mota, l1, l2, l3, ghichu]
    st.session_state.web_tools_df = reindex_df(st.session_state.web_tools_df)
    st.success("Đã cập nhật!")
    st.rerun()


# ==========================================
# 5. SIDEBAR (THANH ĐIỀU HƯỚNG BÊN TRÁI)
# ==========================================
with st.sidebar:
  st.title("🛡️ Quản Lý An Toàn TTM")
  st.caption(
      f"📌 Phiên bản hiệu chỉnh: {TODAY_STR} _ lần"
      f" {st.session_state.lan_hieuchinh}"
  )
  st.divider()

  # NÚT BẬT CHẾ ĐỘ SÁNG / TỐI
  st.markdown("### 🎨 CHẾ ĐỘ GIAO DIỆN")
  col_theme1, col_theme2 = st.columns(2)
  with col_theme1:
    btn_light_type = (
        "primary" if st.session_state.theme_mode == "Light" else "secondary"
    )
    if st.button("☀️ Sáng", use_container_width=True, type=btn_light_type):
      st.session_state.theme_mode = "Light"
      st.rerun()

  with col_theme2:
    btn_dark_type = (
        "primary" if st.session_state.theme_mode == "Dark" else "secondary"
    )
    if st.button("🌙 Tối", use_container_width=True, type=btn_dark_type):
      st.session_state.theme_mode = "Dark"
      st.rerun()

  st.divider()

  st.markdown("### 📂 PHÂN VÙNG LÀM VIỆC")

  menu_options = [
      "1 🌐 DS WEBsites_CV",
      "2 📋 DM QL Files_CV",
      "3 📊 DS BCdinhky_CV",
      "4 🟢 DS Gsheet_CV",
  ]

  for item in menu_options:
    is_active = st.session_state.main_menu == item
    btn_type = "primary" if is_active else "secondary"

    if st.button(
        item, key=f"nav_btn_{item}", use_container_width=True, type=btn_type
    ):
      st.session_state.main_menu = item
      st.rerun()

  main_menu = st.session_state.main_menu

  # Cấu hình Admin
  if st.session_state.get("is_admin"):
    st.divider()
    st.markdown("### ⚙️ Cấu Hình Thư Mục")
    new_dir = st.text_input(
        "Đường dẫn lưu file cục bộ:", value=st.session_state.excel_dir
    )

    col_dir1, col_dir2 = st.columns(2)
    with col_dir1:
      if st.button("💾 Xác nhận", type="primary", use_container_width=True):
        st.session_state.excel_dir = new_dir
        st.toast("🎉 Đã cập nhật đường dẫn!", icon="✅")
    with col_dir2:
      if st.button("📂 Mở thư mục", use_container_width=True):
        if os.path.exists(st.session_state.excel_dir):
          try:
            os.startfile(st.session_state.excel_dir)
          except Exception:
            st.error("Không hỗ trợ trên Web Cloud!")
        else:
          st.error("Đường dẫn không tồn tại!")

  st.divider()

  st.markdown("### 👤 Người dùng: `ttm`")
  if st.session_state.get("is_admin"):
    st.success("Quyền hiện tại: ADMIN")

  col_sb1, col_sb2 = st.columns(2)
  with col_sb1:
    if st.button("🚪 Đăng xuất", use_container_width=True):
      st.info("Đã đăng xuất!")
  with col_sb2:
    if st.button("🧹 Xóa Cache", use_container_width=True):
      st.cache_data.clear()
      st.success("Đã xóa cache!")

  if st.session_state.get("is_admin"):
    if st.button("🔓 Thoát Admin", use_container_width=True):
      st.session_state.is_admin = False
      st.rerun()
  else:
    if st.button("🔐 Đăng nhập Quản trị", use_container_width=True):
      admin_login_dialog()


# ==========================================
# 6. MAIN LAYOUT (GIAO DIỆN CHÍNH)
# ==========================================
if main_menu == "1 🌐 DS WEBsites_CV":
  st.markdown(
      create_responsive_table(st.session_state.web_tools_df, "Mô tả WEB"),
      unsafe_allow_html=True,
  )

  if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("### ⚙ Bảng Điều Khiển Admin")
    col_add, col_edit = st.columns([1, 2])

    with col_add:
      st.caption("📌 **Thêm Dữ Liệu**")
      if st.button(
          "➕ Thêm mới Website", type="primary", use_container_width=True
      ):
        add_web_dialog()

    with col_edit:
      with st.container(border=True):
        st.caption("📌 **Sửa / Xóa dữ liệu theo STT**")
        df_web = st.session_state.web_tools_df
        stt_list = df_web["STT"].tolist() if not df_web.empty else []
        if stt_list:
          c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")
          selected_stt = c1.selectbox(
              "Chọn STT cần sửa/xóa:", stt_list, key="sel_web"
          )
          if c2.button("✏ Sửa", use_container_width=True, key="edit_web"):
            idx = df_web[df_web["STT"] == selected_stt].index[0]
            edit_web_dialog(idx)
          if c3.button("🗑 Xóa", use_container_width=True, key="del_web"):
            idx = df_web[df_web["STT"] == selected_stt].index[0]
            st.session_state.web_tools_df = reindex_df(df_web.drop(idx))
            st.rerun()

elif main_menu == "2 📋 DM QL Files_CV":
  selected_cat = st.selectbox("📌 Chọn Mảng Công Việc:", CATEGORIES)
  st.markdown(
      create_responsive_table(
          st.session_state.data_store[selected_cat], "Thư mục / Hồ sơ"
      ),
      unsafe_allow_html=True,
  )

elif main_menu == "3 📊 DS BCdinhky_CV":
  st.markdown(
      create_responsive_table(
          st.session_state.bc_dinhky_df,
          "Tên Báo Cáo / Công Việc",
          ["Tần suất", "Đơn vị nhận"],
      ),
      unsafe_allow_html=True,
  )

elif main_menu == "4 🟢 DS Gsheet_CV":
  st.markdown(
      create_responsive_table(st.session_state.gsheet_df, "Mô tả Gsheet"),
      unsafe_allow_html=True,
  )
