import streamlit as st

# --- CẤU HÌNH SIDEBAR ---
with st.sidebar:
    # 1. TIÊU ĐỀ HỆ THỐNG (Số 1)
    st.markdown("### 🛡️ HỆ THỐNG QUẢN LÝ AN TOÀN - TTM")
    st.markdown("---")

    # 2. QUẢN LÝ TÀI KHOẢN (Chứa Thư mục Excel - Số 2)
    with st.expander("⚙️ Quản lý tài khoản", expanded=False):
        st.caption("Thông tin người dùng & Cấu hình thư mục")
        
        # Đưa Thư mục Excel vào bên trong
        st.markdown("📌 **Thư mục Excel:**")
        st.code("/mount/src/app-quanlyantoan-ttm", language="text")

    # 3. NÚT THOÁT VÀ CLEAR CACHE
    col_out, col_clear = st.columns(2)
    with col_out:
        if st.button("🚪 Thoát", use_container_width=True):
            st.info("Đã đăng xuất")
    with col_clear:
        if st.button("🧹 Clear C...", use_container_width=True):
            st.cache_data.clear()
            st.success("Đã xóa Cache!")

    st.markdown("---")

    # 4. CHỌN GIAO DIỆN (THEME)
    st.markdown("🎨 **CHỌN GIAO DIỆN (THEME)**")
    theme = st.radio(
        "Lựa chọn giao diện",
        options=["☀️ Light Mode", "🌙 Dark Mode"],
        index=0,
        label_visibility="collapsed"
    )

    st.markdown("---")

    # 5. MỤC LÀM VIỆC
    st.markdown("📁 **MỤC LÀM VIỆC**")
    
    st.button("1 🌐 DS WEBsites_CV", type="primary", use_container_width=True)
    st.button("2 📋 DM QL Files_CV", use_container_width=True)
    st.button("3 📊 DS BCdinhky_CV", use_container_width=True)
    st.button("4 🟢 DS Gsheet_CV", use_container_width=True)
