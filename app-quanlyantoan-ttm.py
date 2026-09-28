import streamlit as st

# Khởi tạo session state
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "show_admin_panel" not in st.session_state:
    st.session_state.show_admin_panel = False

# ==========================================
# 1. TRANG ĐĂNG NHẬP (Đã cập nhật tiêu đề)
# ==========================================
if not st.session_state.logged_in:
    st.markdown("<h2 style='text-align: center;'>Quản lý an toàn TTM TMT</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: gray;'>Version 1.0</p>", unsafe_allow_html=True)
    
    with st.form("login_form"):
        st.subheader("🔐 Đăng Nhập Tài Khoản")
        username = st.text_input("Tên đăng nhập:")
        password = st.text_input("Mật khẩu:", type="password")
        submit = st.form_submit_button("🔑 ĐĂNG NHẬP", use_container_width=True)
        
        if submit:
            if username == "tmt" and password == "123456":  # Thay bằng logic xác thực của bạn
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error("Tên đăng nhập hoặc mật khẩu không đúng!")

# ==========================================
# 2. GIAO DIỆN CHÍNH SAU KHI ĐĂNG NHẬP
# ==========================================
else:
    # --- SIDEBAR (Đã bỏ phần hiển thị thông tin user) ---
    with st.sidebar:
        # Nút bật/tắt Quản lý tài khoản
        if st.button("⚙️ Quản lý tài khoản", use_container_width=True):
            st.session_state.show_admin_panel = not st.session_state.show_admin_panel
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            if st.button("🚪 Thoát", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.show_admin_panel = False
                st.rerun()
        with col_btn2:
            if st.button("🧹 Clear C...", use_container_width=True):
                st.cache_data.clear()
                st.success("Đã xóa cache!")

        st.markdown("---")
        st.markdown("**📌 Thư mục Excel:**")
        st.caption("`/mount/src/app-quanlyantoan-ttm`")
        
        st.markdown("---")
        st.markdown("**🎨 CHỌN GIAO DIỆN (THEME)**")
        # Code chọn theme của bạn...

        st.markdown("---")
        st.markdown("**📁 MỤC LÀM VIỆC**")
        selected_menu = st.radio(
            "Chọn mục:",
            ["1 🌐 DS WEBsites_CV", "2 📋 DM QL Files_CV", "3 📊 DS BCdinhky_CV", "4 🟢 DS Gsheet_CV"],
            label_visibility="collapsed"
        )

    # --- NỘI DUNG CHÍNH ---
    
    # CHI HIỂN THỊ KHU VỰC THÊM/TẢI FILE KHI NHẤN "QUẢN LÝ TÀI KHOẢN"
    if st.session_state.show_admin_panel:
        st.info("⚙️ **BẢNG QUẢN TRỊ TÀI KHOẢN & DỮ LIỆU**")
        if st.button("➕ Thêm mới dòng", type="primary"):
            st.write("Mở form thêm mới dòng...")
        
        col_upload, col_export = st.columns(2)
        with col_upload:
            st.subheader("📤 Tải lên / Thay thế dữ liệu từ file Excel (.xlsx)")
            uploaded_file = st.file_uploader("Chọn file Excel [DanhMuc_CongCu_WEB_TMT]", type=["xlsx", "xls"])
        
        with col_export:
            st.subheader("📥 Xuất dữ liệu ra file Excel")
            st.button("💾 Tải file Excel về máy", use_container_width=True)
        st.markdown("---")

    # HIỂN THỊ BẢNG DỮ LIỆU CHÍNH (Đảm bảo hiển thị đầy đủ cột STT và Link)
    st.subheader(selected_menu)
    
    # Ví dụ hiển thị bảng danh sách giữ nguyên cột STT & Link mặc định
    # Bạn kiểm tra lại dataframe của mình đảm bảo không bị drop/pop cột STT nhé!
