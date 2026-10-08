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

# ====================== CSS ======================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', Arial, sans-serif;
    }
    
    .stApp { background-color: #F0F4F8; }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0A3D91 0%, #0D47A1 50%, #1565C0 100%);
        padding-top: 1.2rem;
    }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    
    [data-testid="stSidebar"] .stButton > button {
        background-color: rgba(255,255,255,0.08);
        color: #FFFFFF !important;
        border: 1px solid rgba(255,255,255,0.18);
        border-radius: 8px;
        font-weight: 500;
        height: 42px;
        width: 100%;
        margin-bottom: 6px;
        text-align: left;
        padding-left: 16px;
        font-size: 14px;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background-color: rgba(255,255,255,0.20);
        border-color: rgba(255,255,255,0.45);
        color: #FFFFFF !important;
    }
    
    /* Tombol utama */
    .stButton > button {
        background-color: #0D47A1;
        color: white;
        border-radius: 8px;
        font-weight: 600;
        border: none;
        height: 42px;
        font-size: 14px;
    }
    .stButton > button:hover {
        background-color: #1565C0;
        color: white;
    }
    
    h1 {
        color: #0A3D91 !important;
        font-weight: 700 !important;
        letter-spacing: -0.3px;
    }
    h2, h3, h4 {
        color: #0D47A1 !important;
        font-weight: 600 !important;
    }
    
    /* ===== INPUT FIELD LEBIH KONTRAS ===== */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input,
    .stSelectbox > div > div,
    .stDateInput > div > div > input,
    .stTextArea > div > div > textarea {
        background-color: #FFFFFF !important;
        border: 1.5px solid #90A4AE !important;
        border-radius: 8px !important;
        color: #1A237E !important;
        font-weight: 500 !important;
    }
    
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus,
    .stSelectbox > div > div:focus,
    .stDateInput > div > div > input:focus,
    .stTextArea > div > div > textarea:focus {
        border: 1.5px solid #0D47A1 !important;
        box-shadow: 0 0 0 2px rgba(13,71,161,0.15) !important;
    }
    
    /* Label input */
    .stTextInput label, .stNumberInput label, .stSelectbox label,
    .stDateInput label, .stTextArea label, .stRadio label {
        color: #37474F !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
    }
    
    /* Emergency */
    .emergency-box {
        background-color: #B71C1C;
        color: white;
        padding: 12px 14px;
        border-radius: 8px;
        font-weight: 600;
        text-align: center;
        margin: 12px 0 10px 0;
        font-size: 13px;
        line-height: 1.45;
    }
    
    /* Backup button */
    [data-testid="stSidebar"] .stDownloadButton > button {
        background-color: #2E7D32 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        height: 42px !important;
        width: 100% !important;
        margin-bottom: 6px !important;
    }
    
    .footer {
        position: fixed;
        left: 0; bottom: 0; width: 100%;
        background-color: #0A3D91;
        color: white;
        text-align: center;
        padding: 8px 0;
        font-size: 12.5px;
        z-index: 999;
    }
    
    .main .block-container {
        padding-bottom: 65px;
        padding-top: 1rem;
    }
    
    .user-info {
        font-size: 14px;
        font-weight: 600;
        margin: 0 0 12px 0;
        padding-bottom: 10px;
        border-bottom: 1px solid rgba(255,255,255,0.2);
    }
    
    /* Kartu Dashboard */
    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 18px 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        text-align: center;
    }
    .metric-label {
        font-size: 13px;
        color: #64748B;
        font-weight: 500;
        margin-bottom: 6px;
    }
    .metric-value {
        font-size: 26px;
        font-weight: 700;
        color: #0A3D91;
    }
    .metric-sub {
        font-size: 12px;
        margin-top: 4px;
    }
</style>
""", unsafe_allow_html=True)

# ====================== SESSION ======================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.nama_user = ""

# ====================== LOGIN PAGE (AMAN - pakai Secrets) ======================
if not st.session_state.logged_in:
    st.markdown("<br><br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1.3, 1.4, 1.3])
    
    with col2:
        st.markdown("""
            <div style='text-align:center; margin-bottom: 28px;'>
                <div style='font-size:26px; font-weight:700; color:#0A3D91; letter-spacing:-0.3px;'>
                    PERTAMINA EP Cepu Zona 12
                </div>
                <div style='font-size:14px; color:#64748B; margin-top:6px;'>
                    Selamat Datang di Sistem Pemantauan Udara Sesaat
                </div>
            </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Masukkan username")
            password = st.text_input("Password", type="password", placeholder="Masukkan password")
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            submit = st.form_submit_button("Masuk", type="primary", use_container_width=True)
            
            if submit:
                try:
                    # Ambil password dari Streamlit Secrets (lebih aman)
                    users = st.secrets["users"]
                    
                    if username.strip() in users and users[username.strip()] == password:
                        st.session_state.logged_in = True
                        st.session_state.username = username.strip()
                        
                        # Nama tampilan
                        if username.strip() == "admin":
                            st.session_state.nama_user = "Admin"
                        elif username.strip() == "delta":
                            st.session_state.nama_user = "Delta Candra"
                        else:
                            st.session_state.nama_user = username.strip()
                        
                        st.session_state.menu = "Dashboard"
                        st.rerun()
                    else:
                        st.error("Username atau Password salah")
                except Exception:
                    st.error("Konfigurasi secrets belum benar. Pastikan file .streamlit/secrets.toml sudah ada.")
    
    st.stop()

# ====================== HEADER ======================
col_logo, col_title = st.columns([0.7, 5.5])

with col_logo:
    if os.path.exists("logo_pertamina.png"):
        st.image("logo_pertamina.png", width=100)
    else:
        st.markdown("""
            <div style='background:#0A3D91;color:white;padding:10px 6px;text-align:center;
                        font-weight:700;border-radius:6px;font-size:13px;'>PERTAMINA</div>
        """, unsafe_allow_html=True)

with col_title:
    st.markdown("""
        <h1 style='margin:0;padding-top:6px;font-size:24px;'>PERTAMINA EP CEPU ZONA 12</h1>
        <p style='margin:2px 0 0 0;color:#64748B;font-size:13.5px;font-weight:500;'>
            Sistem Informasi Manajemen Pemantauan Udara Sesaat (GPF JTB)
        </p>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin:10px 0 18px 0;border:none;border-top:2px solid #BBDEFB;'>", unsafe_allow_html=True)

# ====================== SIDEBAR ======================
with st.sidebar:
    st.markdown(f"<div class='user-info'>{st.session_state.nama_user}</div>", unsafe_allow_html=True)
    
    if "menu" not in st.session_state:
        st.session_state.menu = "Dashboard"
    
    menus = [
        "Dashboard",
        "Input Data",
        "Review Data",
        "Download Laporan",
        "Inspeksi Alat",
        "Kelola Master Data"
    ]
    
    if st.session_state.username == "admin":
        menus.append("Kelola User")
    
    for m in menus:
        if st.button(m, key=f"btn_{m}", use_container_width=True):
            st.session_state.menu = m
            st.rerun()
    
    st.markdown("""
    <div class="emergency-box">
        Kontak Emergency Response<br>
        <span style="font-size:15px;letter-spacing:0.5px;">0811-9046-777</span>
    </div>
    """, unsafe_allow_html=True)
    
    if os.path.exists(FILE_DB):
        with open(FILE_DB, "rb") as f:
            st.download_button(
                label="Backup Database",
                data=f,
                file_name=f"backup_database_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="btn_backup"
            )
    
    if st.button("Logout", use_container_width=True, key="btn_logout"):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()

# ====================== ROUTING ======================
menu = st.session_state.menu

if menu == "Dashboard":
    st.markdown("### Dashboard Pemantauan")
    st.caption(f"Data kondisi hari ini • {datetime.now().strftime('%d %B %Y')}")
    
    df_db = pd.read_excel(FILE_DB) if os.path.exists(FILE_DB) else pd.DataFrame()
    df_alat = pd.read_excel(FILE_ALAT) if os.path.exists(FILE_ALAT) else pd.DataFrame()
    
    hari_ini = datetime.now().strftime("%Y-%m-%d")
    df_hari = df_db[df_db["Tanggal"] == hari_ini] if not df_db.empty else pd.DataFrame()
    
    total_data = len(df_hari)
    ada_terpapar = "Terpapar" in df_hari["Status"].values if not df_hari.empty else False
    titik_sudah = df_hari["Kode_Titik"].nunique() if not df_hari.empty else 0
    
    dari_file = pd.read_excel("master_daftar_alat.xlsx") if os.path.exists("master_daftar_alat.xlsx") else pd.DataFrame({"Nama_Alat": ["Gas Detector 1","Gas Detector 2","SO2 Detector 1","SO2 Detector 2"]})
    total_alat = len(dari_file)
    siap = 0
    if not df_alat.empty:
        for nama in dari_file["Nama_Alat"]:
            last = df_alat[df_alat["Nama_Alat"] == nama].tail(1)
            if len(last) > 0 and last.iloc[0]["Status"] == "Siap Pakai":
                siap += 1
    
    c1, c2, c3, c4 = st.columns(4)
    
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Data Hari Ini</div>
            <div class="metric-value">{total_data}</div>
            <div class="metric-sub" style="color:#64748B;">data tercatat</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c2:
        if ada_terpapar:
            status_text = "TEROBAR"
            status_color = "#C62828"
            status_sub = "Perlu perhatian"
        else:
            status_text = "AMAN"
            status_color = "#2E7D32"
            status_sub = "Kondisi normal"
        
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Status</div>
            <div class="metric-value" style="color:{status_color};">{status_text}</div>
            <div class="metric-sub" style="color:{status_color};">{status_sub}</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Titik Terukur</div>
            <div class="metric-value">{titik_sudah}</div>
            <div class="metric-sub" style="color:#64748B;">titik pantau</div>
        </div>
        """, unsafe_allow_html=True)
    
    with c4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Alat Siap Pakai</div>
            <div class="metric-value">{siap}<span style="font-size:16px;color:#64748B;"> / {total_alat}</span></div>
            <div class="metric-sub" style="color:#64748B;">unit alat</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
    
    col_kiri, col_kanan = st.columns(2)
    
    with col_kiri:
        st.markdown("#### Data Pemantauan Hari Ini")
        if not df_hari.empty:
            st.dataframe(
                df_hari[["Waktu", "Shift", "Kode_Titik", "Lokasi", "H2S", "O2", "SO2", "Status"]],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Belum ada data pemantauan hari ini.")
    
    with col_kanan:
        st.markdown("#### Status Alat")
        if not df_alat.empty:
            status_terakhir = df_alat.sort_values("Tanggal_Inspeksi").groupby("Nama_Alat").tail(1)
            st.dataframe(
                status_terakhir[["Nama_Alat", "Status", "Tanggal_Inspeksi", "Tanggal_Berikutnya"]],
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Belum ada data inspeksi alat.")
    
    st.markdown("---")
    st.caption("Data Realtime Sistem Pemantauan Udara Sesaat")

elif menu == "Input Data":
    tampilkan_halaman_input(st.session_state.nama_user)
else:
    mapping = {
        "Review Data": "2. Review Data",
        "Download Laporan": "3. Download Laporan",
        "Inspeksi Alat": "4. Inspeksi Alat",
        "Kelola Master Data": "5. Kelola Master Data",
        "Kelola User": "6. Kelola User"
    }
    menu_lama = mapping.get(menu, menu)
    tampilkan_halaman_output(menu_lama, st.session_state.nama_user, st.session_state.username)

# Footer
st.markdown("""
<div class="footer">
    Digitalisasi Monitoring Kualitas Udara • Delta Candra • KP Unigoro 2026
</div>
""", unsafe_allow_html=True)