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
    page_icon="🛡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Thư mục mặc định ban đầu
DEFAULT_EXCEL_DIR = r"D:\0 2025 0 LUU OFFICE drive\0000 chua luu\0 0 0 app\000TmT_VBA_source\Dashboard_AnToan"
if not os.path.exists(DEFAULT_EXCEL_DIR):
  DEFAULT_EXCEL_DIR = os.path.dirname(os.path.abspath(__file__))

# Khởi tạo biến lưu trạng thái
if "is_admin" not in st.session_state:
  st.session_state.is_admin = False
if "dark_mode" not in st.session_state:
  st.session_state.dark_mode = False
if "excel_dir" not in st.session_state:
  st.session_state.excel_dir = DEFAULT_EXCEL_DIR
if "lan_hieuchinh" not in st.session_state:
  st.session_state.lan_hieuchinh = "001"
if "main_menu" not in st.session_state:
  st.session_state.main_menu = "1 🌐 DS WEBsites_CV"

# Xử lý CSS Giao diện (Đã điều chỉnh padding-top lên 3.5rem để tránh bị Header đè)
css_style = """
<style>
.block-container { 
    padding-top: 3.5rem !important; 
    padding-bottom: 1rem !important; 
}

/* Tùy chỉnh nút Menu ở Sidebar */
[data-testid="stSidebar"] div.stButton > button {
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 10px 14px !important;
    font-weight: 500 !important;
    border-radius: 8px !important;
    margin-bottom: 2px !important;
}
</style>
"""
if st.session_state.get("dark_mode"):
  css_style += """
<style>
.stApp { background-color: #0E1117 !important; color: #FFFFFF !important; }
.stSidebar { background-color: #161B22 !important; }
h1, h2, h3, h4, h5, h6, p, span, div, strong { color: #E0E0E0 !important; }
</style>
"""
st.markdown(css_style, unsafe_allow_html=True)


# ==========================================
# 2. BIẾN VÀ HÀM HỖ TRỢ
# ==========================================
TODAY_STR = datetime.now().strftime("%d/%m/%Y")
DATE_STR = datetime.now().strftime("%Y%m%d")

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
    {
        "Mô tả WEB": "Cổng Thông tin Bộ Công Thương",
        "Link 1": "https://moit.gov.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Theo dõi văn bản quy phạm kỹ thuật",
    },
    {
        "Mô tả WEB": "Cổng Báo cáo Phòng chống thiên tai",
        "Link 1": "https://pctt.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Cập nhật tình hình PCTT & TKCN",
    },
    {
        "Mô tả WEB": "Hệ thống Quản lý Đầu tư Xây dựng (IMIS)",
        "Link 1": "https://imis.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Theo dõi an toàn dự án ĐTXD",
    },
    {
        "Mô tả WEB": "Hệ thống Thông tin Báo cáo EVN",
        "Link 1": "https://baocao.evn.com.vn",
        "Link 2": "",
        "Link 3": "",
        "Ghi chú": "Tổng hợp chỉ tiêu an toàn - kỹ thuật",
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


# BẢNG HTML TỰ ĐỘNG CỐ ĐỊNH TIÊU ĐỀ + CHỮ 17PX
def create_responsive_table(df, main_col_name, extra_cols=None):
  if extra_cols is None:
    extra_cols = []

  html = (
      '<div style="max-height: 580px; overflow-y: auto; overflow-x: auto;'
      ' margin-bottom: 20px; border-radius: 8px; border: 1px solid #333333;'
      ' background-color: #0E1117;">'
  )
  html += (
      '<table style="width: 100%; border-collapse: separate; border-spacing: 0;'
      ' font-size: 17px; text-align: left;">'
  )

  th_style = (
      "padding: 12px 12px; color: #1E88E5; position: sticky; top: 0;"
      " background-color: #161B22; z-index: 10; border-bottom: 2px solid"
      " #1E88E5; box-shadow: 0 2px 4px rgba(0,0,0,0.5); font-weight: bold;"
      " font-size: 17px; vertical-align: middle;"
  )

  html += "<thead><tr>"
  html += f'<th style="{th_style} white-space: nowrap;">STT</th>'
  html += f'<th style="{th_style} min-width: 280px;">{main_col_name}</th>'
  for ec in extra_cols:
    html += f'<th style="{th_style} white-space: nowrap;">{ec}</th>'
  html += f'<th style="{th_style} min-width: 170px;">Links truy cập</th>'
  html += f'<th style="{th_style} min-width: 220px;">Ghi chú</th>'
  html += "</tr></thead><tbody>"

  if df.empty:
    html += (
        '<tr><td colspan="10" style="padding: 15px; text-align: center; color:'
        ' #888;">Chưa có dữ liệu</td></tr>'
    )
  else:
    for idx, row in df.iterrows():
      html += '<tr style="border-bottom: 1px solid #222;">'
      html += (
          '<td style="padding: 14px 12px; font-weight: bold; border-bottom: 1px'
          f' solid #262730;">{row.get("STT", "")}</td>'
      )
      html += (
          '<td style="padding: 14px 12px; border-bottom: 1px solid'
          f' #262730;">{row.get(main_col_name, "")}</td>'
      )
      for ec in extra_cols:
        html += (
            '<td style="padding: 14px 12px; border-bottom: 1px solid'
            f' #262730;">{row.get(ec, "")}</td>'
        )

      links = []
      l1, l2, l3 = (
          row.get("Link 1", ""),
          row.get("Link 2", ""),
          row.get("Link 3", ""),
      )
      if l1 and str(l1).strip().startswith("http"):
        links.append(
            f'<a href="{str(l1).strip()}" target="_blank" style="color:'
            ' #64B5F6; text-decoration: none; font-weight: bold;">Link 1</a>'
        )
      if l2 and str(l2).strip().startswith("http"):
        links.append(
            f'<a href="{str(l2).strip()}" target="_blank" style="color:'
            ' #64B5F6; text-decoration: none; font-weight: bold;">Link 2</a>'
        )
      if l3 and str(l3).strip().startswith("http"):
        links.append(
            f'<a href="{str(l3).strip()}" target="_blank" style="color:'
            ' #64B5F6; text-decoration: none; font-weight: bold;">Link 3</a>'
        )
      links_str = " | ".join(links) if links else "-"

      html += (
          '<td style="padding: 14px 12px; border-bottom: 1px solid'
          f' #262730;">{links_str}</td>'
      )
      html += (
          '<td style="padding: 14px 12px; border-bottom: 1px solid'
          f' #262730;">{row.get("Ghi chú", "")}</td>'
      )
      html += "</tr>"
  html += "</tbody></table></div>"
  return html


# Khởi tạo dữ liệu session
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
# 3. DIALOGS (HỘP THOẠI)
# ==========================================
@st.dialog("🔐 Đăng Nhập Quản Trị Viên")
def admin_login_dialog():
  st.write("Vui lòng nhập mật khẩu để kích hoạt tính năng thêm/sửa/xóa.")
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


@st.dialog("➕ Thêm mới Gsheet_CV")
def add_gsheet_dialog():
  mota = st.text_input("Mô tả Gsheet (*):")
  l1 = st.text_input("Link 1 (*):")
  l2 = st.text_input("Link 2 (bổ sung):")
  l3 = st.text_input("Link 3 (bổ sung):")
  ghichu = st.text_input("Ghi chú:")
  if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
    if mota:
      new_row = pd.DataFrame([{
          "Mô tả Gsheet": mota,
          "Link 1": l1,
          "Link 2": l2,
          "Link 3": l3,
          "Ghi chú": ghichu,
      }])
      st.session_state.gsheet_df = reindex_df(
          pd.concat([st.session_state.gsheet_df, new_row], ignore_index=True)
      )
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
    st.session_state.gsheet_df.loc[
        idx, ["Mô tả Gsheet", "Link 1", "Link 2", "Link 3", "Ghi chú"]
    ] = [mota, l1, l2, l3, ghichu]
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
      new_row = pd.DataFrame([{
          "Thư mục / Hồ sơ": hoso,
          "Link 1": l1,
          "Link 2": l2,
          "Link 3": l3,
          "Ghi chú": ghichu,
      }])
      st.session_state.data_store[selected_cat] = reindex_df(
          pd.concat(
              [st.session_state.data_store[selected_cat], new_row],
              ignore_index=True,
          )
      )
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
    st.session_state.data_store[selected_cat].loc[
        idx, ["Thư mục / Hồ sơ", "Link 1", "Link 2", "Link 3", "Ghi chú"]
    ] = [hoso, l1, l2, l3, ghichu]
    st.session_state.data_store[selected_cat] = reindex_df(
        st.session_state.data_store[selected_cat]
    )
    st.success("Đã cập nhật!")
    st.rerun()


@st.dialog("➕ Thêm mới Báo Cáo / Công Việc")
def add_bc_dialog():
  ten_bc = st.text_input("Tên Báo Cáo / Công Việc (*):")
  tan_suat = st.selectbox(
      "Tần suất:", ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"]
  )
  don_vi = st.text_input("Đơn vị nhận:")
  l1 = st.text_input("Link 1 (*):")
  l2 = st.text_input("Link 2 (bổ sung):")
  l3 = st.text_input("Link 3 (bổ sung):")
  ghichu = st.text_input("Ghi chú:")
  if st.button("💾 Lưu Mới", type="primary", use_container_width=True):
    if ten_bc:
      new_row = pd.DataFrame([{
          "Tên Báo Cáo / Công Việc": ten_bc,
          "Tần suất": tan_suat,
          "Đơn vị nhận": don_vi,
          "Link 1": l1,
          "Link 2": l2,
          "Link 3": l3,
          "Ghi chú": ghichu,
      }])
      st.session_state.bc_dinhky_df = reindex_df(
          pd.concat([st.session_state.bc_dinhky_df, new_row], ignore_index=True)
      )
      st.success("Đã thêm báo cáo!")
      st.rerun()
    else:
      st.warning("Vui lòng nhập tên báo cáo!")


@st.dialog("✏️ Chỉnh sửa Báo Cáo / Công Việc")
def edit_bc_dialog(idx):
  df = st.session_state.bc_dinhky_df
  row = df.loc[idx]
  ten_bc = st.text_input(
      "Tên Báo Cáo / Công Việc:",
      value=str(row.get("Tên Báo Cáo / Công Việc", "")),
  )
  tan_suat = st.selectbox(
      "Tần suất:",
      ["Hàng Tuần", "Hàng Tháng", "Hàng Quý", "Hàng Năm", "Đột xuất"],
      index=0,
  )
  don_vi = st.text_input("Đơn vị nhận:", value=str(row.get("Đơn vị nhận", "")))
  l1 = st.text_input("Link 1:", value=str(row.get("Link 1", "")))
  l2 = st.text_input("Link 2:", value=str(row.get("Link 2", "")))
  l3 = st.text_input("Link 3:", value=str(row.get("Link 3", "")))
  ghichu = st.text_input("Ghi chú:", value=str(row.get("Ghi chú", "")))
  if st.button("💾 Cập Nhật", type="primary", use_container_width=True):
    st.session_state.bc_dinhky_df.loc[idx, [
        "Tên Báo Cáo / Công Việc",
        "Tần suất",
        "Đơn vị nhận",
        "Link 1",
        "Link 2",
        "Link 3",
        "Ghi chú",
    ]] = [ten_bc, tan_suat, don_vi, l1, l2, l3, ghichu]
    st.session_state.bc_dinhky_df = reindex_df(st.session_state.bc_dinhky_df)
    st.success("Đã cập nhật!")
    st.rerun()


# ==========================================
# 4. SIDEBAR (THANH ĐIỀU HƯỚNG BÊN TRÁI)
# ==========================================
with st.sidebar:
  st.title("🛡️ Quản Lý An Toàn TTM")
  st.caption(
      f"📌 Phiên bản hiệu chỉnh: {TODAY_STR} _ lần"
      f" {st.session_state.lan_hieuchinh}"
  )
  st.divider()

  st.header("📂 PHÂN VÙNG LÀM VIỆC")

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

  if st.session_state.get("is_admin"):
    st.divider()
    st.header("⚙️ Cấu Hình Thư Mục")
    st.caption(
        "*(Lưu ý: Mở thư mục chỉ hoạt động khi chạy trên máy tính cá nhân)*"
    )

    new_dir = st.text_input(
        "Đường dẫn lưu file cục bộ:", value=st.session_state.excel_dir
    )

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
            st.error(
                "Tính năng này không khả dụng khi chạy trên Web Cloud!"
            )
        else:
          st.error("Đường dẫn thư mục không tồn tại!")

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

  theme_label = "☀️ Sáng" if st.session_state.get("dark_mode") else "🌙 Tối"
  if st.button(theme_label, use_container_width=True):
    st.session_state.dark_mode = not st.session_state.dark_mode
    st.rerun()

  if st.session_state.get("is_admin"):
    if st.button("🔓 Thoát chế độ Admin", use_container_width=True):
      st.session_state.is_admin = False
      st.rerun()
  else:
    if st.button("🔐 Đăng nhập Quản trị", use_container_width=True):
      admin_login_dialog()


# ==========================================
# 5. MAIN LAYOUT (GIAO DIỆN CHÍNH)
# ==========================================

# ------------------------------------------
# PHẦN 1: DS WEBsites_CV
# ------------------------------------------
if main_menu == "1 🌐 DS WEBsites_CV":
  st.markdown(
      create_responsive_table(st.session_state.web_tools_df, "Mô tả WEB"),
      unsafe_allow_html=True,
  )

  if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("### ⚙️️ Bảng Điều Khiển Admin")
    col_add, col_edit = st.columns([1, 2])

    with col_add:
      st.caption("📌 **Thêm Dữ Liệu**")
      if st.button(
          "➕ Thêm mới Website", type="primary", use_container_width=True
      ):
        add_web_dialog()
      if st.button("🔄 Khôi phục mặc định", use_container_width=True):
        st.session_state.web_tools_df = reindex_df(
            pd.DataFrame(DEFAULT_14_WEBS)
        )
        st.rerun()

    with col_edit:
      with st.container(border=True):
        st.caption("📌 **Sửa / Xóa dữ liệu theo STT**")
        df_web = st.session_state.web_tools_df
        stt_list = df_web["STT"].tolist() if not df_web.empty else []
        if stt_list:
          c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")
          selected_stt = c1.selectbox(
              "Nhìn bảng và chọn số STT cần thao tác:",
              stt_list,
              key="sel_web",
          )
          if c2.button("✏️️ Sửa", use_container_width=True, key="edit_web"):
            idx = df_web[df_web["STT"] == selected_stt].index[0]
            edit_web_dialog(idx)
          if c3.button("🗑️ Xóa", use_container_width=True, key="del_web"):
            idx = df_web[df_web["STT"] == selected_stt].index[0]
            st.session_state.web_tools_df = reindex_df(df_web.drop(idx))
            st.rerun()
        else:
          st.info("Bảng đang trống.")

    with st.expander("⚙ Quản lý Nhập / Xuất Excel (WEBsites)", expanded=False):
      col_w1, col_w2 = st.columns(2)
      with col_w1:
        st.markdown("#### Tải dữ liệu xuống máy")
        out_df = reindex_df(st.session_state.web_tools_df)
        excel_data = generate_excel_download(out_df, "WEBSITES")
        st.download_button(
            label="📥 Tải file Excel",
            data=excel_data,
            file_name=f"1_DS_WEBsites_CV_{DATE_STR}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
      with col_w2:
        st.markdown("#### Nhập dữ liệu")
        up_w = st.file_uploader(
            "Chọn file Excel:",
            type=["xlsx", "xls"],
            key="up_w",
            label_visibility="collapsed",
        )
        if up_w and st.button(
            "🚀 Cập nhật từ File", type="primary", use_container_width=True
        ):
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

  st.markdown(
      create_responsive_table(
          st.session_state.data_store[selected_cat], "Thư mục / Hồ sơ"
      ),
      unsafe_allow_html=True,
  )

  if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("### ⚙️ Bảng Điều Khiển Admin")
    col_add, col_edit = st.columns([1, 2])

    with col_add:
      st.caption("📌 **Thêm Dữ Liệu**")
      if st.button(
          "➕ Thêm Hồ Sơ Mới", type="primary", use_container_width=True
      ):
        add_file_dialog(selected_cat)

    with col_edit:
      with st.container(border=True):
        st.caption("📌 **Sửa / Xóa dữ liệu theo STT**")
        df_cat = st.session_state.data_store[selected_cat]
        stt_list = df_cat["STT"].tolist() if not df_cat.empty else []
        if stt_list:
          c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")
          safe_key = selected_cat.replace(" ", "_")
          selected_stt = c1.selectbox(
              "Nhìn bảng và chọn số STT cần thao tác:",
              stt_list,
              key=f"sel_cat_{safe_key}",
          )
          if c2.button(
              "✏️ Sửa", use_container_width=True, key=f"edit_cat_{safe_key}"
          ):
            idx = df_cat[df_cat["STT"] == selected_stt].index[0]
            edit_file_dialog(selected_cat, idx)
          if c3.button(
              "🗑️ Xóa", use_container_width=True, key=f"del_cat_{safe_key}"
          ):
            idx = df_cat[df_cat["STT"] == selected_stt].index[0]
            st.session_state.data_store[selected_cat] = reindex_df(
                df_cat.drop(idx)
            )
            st.rerun()
        else:
          st.info("Bảng đang trống.")

    with st.expander(
        f"⚙️ Quản lý Nhập / Xuất Excel ({selected_cat})", expanded=False
    ):
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
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
      with col_q2:
        st.markdown("#### Nhập dữ liệu")
        up_q = st.file_uploader(
            "Chọn file Excel:",
            type=["xlsx", "xls"],
            key=f"up_q_{selected_cat}",
            label_visibility="collapsed",
        )
        if up_q and st.button(
            "🚀 Cập nhật từ File", type="primary", use_container_width=True
        ):
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
  st.markdown(
      create_responsive_table(
          st.session_state.bc_dinhky_df,
          "Tên Báo Cáo / Công Việc",
          ["Tần suất", "Đơn vị nhận"],
      ),
      unsafe_allow_html=True,
  )

  if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("### ⚙️ Bảng Điều Khiển Admin")
    col_add, col_edit = st.columns([1, 2])

    with col_add:
      st.caption("📌 **Thêm Dữ Liệu**")
      if st.button(
          "➕ Thêm Báo Cáo Mới", type="primary", use_container_width=True
      ):
        add_bc_dialog()

    with col_edit:
      with st.container(border=True):
        st.caption("📌 **Sửa / Xóa dữ liệu theo STT**")
        df_bc = st.session_state.bc_dinhky_df
        stt_list = df_bc["STT"].tolist() if not df_bc.empty else []
        if stt_list:
          c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")
          selected_stt = c1.selectbox(
              "Nhìn bảng và chọn số STT cần thao tác:",
              stt_list,
              key="sel_bc",
          )
          if c2.button("✏️ Sửa", use_container_width=True, key="edit_bc"):
            idx = df_bc[df_bc["STT"] == selected_stt].index[0]
            edit_bc_dialog(idx)
          if c3.button("🗑️ Xóa", use_container_width=True, key="del_bc"):
            idx = df_bc[df_bc["STT"] == selected_stt].index[0]
            st.session_state.bc_dinhky_df = reindex_df(df_bc.drop(idx))
            st.rerun()
        else:
          st.info("Bảng đang trống.")

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
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
      with col_b2:
        st.markdown("#### Nhập dữ liệu")
        up_b = st.file_uploader(
            "Chọn file Excel:",
            type=["xlsx", "xls"],
            key="up_b",
            label_visibility="collapsed",
        )
        if up_b and st.button(
            "🚀 Cập nhật từ File", type="primary", use_container_width=True
        ):
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
  st.markdown(
      create_responsive_table(
          st.session_state.gsheet_df, "Mô tả Gsheet"
      ),
      unsafe_allow_html=True,
  )

  if st.session_state.is_admin:
    st.markdown("---")
    st.markdown("### ⚙️ Bảng Điều Khiển Admin")
    col_add, col_edit = st.columns([1, 2])

    with col_add:
      st.caption("📌 **Thêm Dữ Liệu**")
      if st.button(
          "➕ Thêm mới Gsheet", type="primary", use_container_width=True
      ):
        add_gsheet_dialog()
      if st.button("🔄 Khôi phục mặc định", use_container_width=True):
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
        st.rerun()

    with col_edit:
      with st.container(border=True):
        st.caption("📌 **Sửa / Xóa dữ liệu theo STT**")
        df_gsheet = st.session_state.gsheet_df
        stt_list = df_gsheet["STT"].tolist() if not df_gsheet.empty else []
        if stt_list:
          c1, c2, c3 = st.columns([2, 1, 1], vertical_alignment="bottom")
          selected_stt = c1.selectbox(
              "Nhìn bảng và chọn số STT cần thao tác:",
              stt_list,
              key="sel_gs",
          )
          if c2.button("✏️ Sửa", use_container_width=True, key="edit_gs"):
            idx = df_gsheet[df_gsheet["STT"] == selected_stt].index[0]
            edit_gsheet_dialog(idx)
          if c3.button("🗑️ Xóa", use_container_width=True, key="del_gs"):
            idx = df_gsheet[df_gsheet["STT"] == selected_stt].index[0]
            st.session_state.gsheet_df = reindex_df(df_gsheet.drop(idx))
            st.rerun()
        else:
          st.info("Bảng đang trống.")

    with st.expander("⚙ Quản lý Nhập / Xuất Excel (Gsheet)", expanded=False):
      col_g1, col_g2 = st.columns(2)
      with col_g1:
        st.markdown("#### Tải dữ liệu xuống máy")
        out_df = reindex_df(st.session_state.gsheet_df)
        excel_data = generate_excel_download(out_df, "GSHEET")
        st.download_button(
            label="📥 Tải file Excel",
            data=excel_data,
            file_name=f"4_DS_Gsheet_CV_{DATE_STR}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )
      with col_g2:
        st.markdown("#### Nhập dữ liệu")
        up_g = st.file_uploader(
            "Chọn file Excel:",
            type=["xlsx", "xls"],
            key="up_g",
            label_visibility="collapsed",
        )
        if up_g and st.button(
            "🚀 Cập nhật từ File",
            type="primary",
            use_container_width=True,
            key="btn_up_g",
        ):
          try:
            df_raw = pd.read_excel(up_g)
            st.session_state.gsheet_df = reindex_df(df_raw)
            st.toast("🎉 Đã cập nhật thành công!", icon="✅")
            st.rerun()
          except Exception as e:
            st.error(f"Lỗi: {e}")
