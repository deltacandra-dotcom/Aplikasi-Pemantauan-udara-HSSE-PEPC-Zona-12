import streamlit as st
import pandas as pd
import os
from datetime import datetime
from gui_input import tampilkan_halaman_input
from gui_output import tampilkan_halaman_output

st.set_page_config(
    page_title="Aplikasi Pemantauan Udara Sesaat HSSE - PEPC Zona 12",
    page_icon="logo_icon.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

FILE_USER = "master_users.xlsx"
FILE_DB = "database_monitoring_jtb.xlsx"
FILE_ALAT = "database_inspeksi_alat.xlsx"

# ====================== LOGIN AMAN ======================
def cek_login():
    """Cek apakah user sudah login"""
    return st.session_state.get("login", False)

def tampilkan_halaman_login():
    st.markdown("""
    <style>
        .login-box {
            max-width: 400px;
            margin: 80px auto;
            padding: 30px;
            background: white;
            border-radius: 12px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.1);
        }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("### 🔐 Login Aplikasi")
        st.caption("Pemantauan Udara Sesaat - PEPC Zona 12")
        
        username = st.text_input("Username", key="login_user")
        password = st.text_input("Password", type="password", key="login_pass")
        
        if st.button("Login", type="primary", use_container_width=True):
            try:
                # Ambil password dari Streamlit Secrets
                users = st.secrets["users"]
                
                if username in users and users[username] == password:
                    st.session_state["login"] = True
                    st.session_state["username"] = username
                    
                    # Nama tampilan
                    if username == "admin":
                        st.session_state["nama_user"] = "Admin"
                    elif username == "delta":
                        st.session_state["nama_user"] = "Delta Candra"
                    else:
                        st.session_state["nama_user"] = username
                        
                    st.success("Login berhasil!")
                    st.rerun()
                else:
                    st.error("Username atau password salah")
            except Exception as e:
                st.error("Konfigurasi secrets belum benar. Hubungi admin.")
                st.caption(str(e))

# ====================== MAIN APP ======================
def main():
    # Cek login dulu
    if not cek_login():
        tampilkan_halaman_login()
        return
    
    # Jika sudah login, tampilkan aplikasi
    st.sidebar.title(f"Halo, {st.session_state.get('nama_user', 'User')}")
    
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()
    
    # Menu
    menu = st.sidebar.radio(
        "Menu",
        ["Dashboard", "Input Data", "Review Data", "Download Laporan", 
         "Inspeksi Alat", "Kelola Master Data", "Kelola User"]
    )
    
    # ====================== CSS ======================
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        html, body, [class*="css"] {
            font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
        }
        
        .stApp { background-color: #F0F4F8; }
        
        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0A3D91 0%, #0D47A1 50%, #1565C0 100%);
            padding-top: 1.2rem;
        }
        [data-testid="stSidebar"] * { color: #FFFFFF !important; }
        
        [data-testid="stSidebar"] .stButton > button {
            background-color: rgba(255,255,255,0.08);
            color: #FFFFFF !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # Routing menu
    if menu == "Dashboard":
        st.title("Dashboard Pemantauan Udara Sesaat")
        st.info("Selamat datang di sistem monitoring.")
        
        # Contoh ringkasan data hari ini
        if os.path.exists(FILE_DB):
            df = pd.read_excel(FILE_DB)
            hari_ini = datetime.now().strftime("%Y-%m-%d")
            df_hari = df[df["Tanggal"] == hari_ini] if "Tanggal" in df.columns else pd.DataFrame()
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Data Hari Ini", len(df_hari))
            with col2:
                aman = len(df_hari[df_hari["Status"] == "Aman"]) if not df_hari.empty else 0
                st.metric("Status Aman", aman)
    
    elif menu == "Input Data":
        tampilkan_halaman_input(st.session_state.get("nama_user", ""))
    
    else:
        mapping = {
            "Review Data": "2. Review Data",
            "Download Laporan": "3. Download Laporan",
            "Inspeksi Alat": "4. Inspeksi Alat",
            "Kelola Master Data": "5. Kelola Master Data",
            "Kelola User": "6. Kelola User"
        }
        menu_lama = mapping.get(menu, menu)
        tampilkan_halaman_output(menu_lama, st.session_state.get("nama_user", ""), st.session_state.get("username", ""))

# Jalankan aplikasi
if __name__ == "__main__":
    main()