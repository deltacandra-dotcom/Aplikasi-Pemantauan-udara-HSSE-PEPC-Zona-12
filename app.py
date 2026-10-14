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

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_ALAT = "database_inspeksi_alat.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"

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
    h2, h3, h4, h5 {
        color: #0D47A1 !important;
        font-weight: 600 !important;
    }
    
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
    
    .stTextInput label, .stNumberInput label, .stSelectbox label,
    .stDateInput label, .stTextArea label, .stRadio label {
        color: #37474F !important;
        font-weight: 600 !important;
        font-size: 13.5px !important;
    }
    
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
    
    .metric-card {
        background: linear-gradient(145deg, #FFFFFF 0%, #F8FAFC 100%);
        border-radius: 14px;
        padding: 20px 16px;
        border: 1px solid #E2E8F0;
        box-shadow: 0 4px 12px rgba(10, 61, 145, 0.08);
        text-align: center;
        height: 100%;
    }
    .metric-label {
        font-size: 12px;
        color: #64748B;
        font-weight: 600;
        margin-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #0A3D91;
        line-height: 1.1;
    }
    .metric-sub {
        font-size: 12px;
        margin-top: 6px;
        font-weight: 500;
    }
    
    .alert-bahaya {
        background: linear-gradient(135deg, #FFEBEE 0%, #FFCDD2 100%);
        border-left: 5px solid #C62828;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 14px;
    }
    
    .alert-waspada {
        background: linear-gradient(135deg, #FFF8E1 0%, #FFECB3 100%);
        border-left: 5px solid #F9A825;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 14px;
    }
    
    .info-hari {
        background: linear-gradient(135deg, #E3F2FD 0%, #BBDEFB 100%);
        border-left: 5px solid #0D47A1;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 16px;
    }
</style>
""", unsafe_allow_html=True)

# ====================== SESSION ======================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.username = ""
    st.session_state.nama_user = ""
    st.session_state.role = "petugas"

# ====================== LOGIN DARI SECRETS ======================
def cek_login(username, password):
    try:
        users = st.secrets["users"]
        if username in users:
            user_data = users[username]
            if str(user_data["password"]) == str(password):
                nama = user_data.get("nama", username)
                role = str(user_data.get("role", "petugas")).strip().lower()
                if role not in ["petugas", "diperiksa", "disetujui", "admin"]:
                    role = "petugas"
                return True, nama, role
        return False, "", ""
    except Exception:
        return False, "", ""

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
                ok, nama, role = cek_login(username.strip(), password)
                if ok:
                    st.session_state.logged_in = True
                    st.session_state.username = username.strip()
                    st.session_state.nama_user = nama
                    st.session_state.role = role
                    st.session_state.menu = "Dashboard"
                    st.rerun()
                else:
                    st.error("Username atau Password salah")
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
    role = st.session_state.get("role", "petugas")
    role_label = {
        "petugas": "Petugas",
        "diperiksa": "Environment Officer",
        "disetujui": "Superintendent HSSE",
        "admin": "Admin"
    }.get(role, role)
    
    st.markdown(
        "<div class='user-info'>" + st.session_state.nama_user
        + "<br><span style='font-size:12px;opacity:0.85;'>" + role_label + "</span></div>",
        unsafe_allow_html=True
    )
    
    if "menu" not in st.session_state:
        st.session_state.menu = "Dashboard"
    
    menus = [
        "Dashboard",
        "Input Data",
        "Review Data",
        "Download Laporan"
    ]
    if role in ["diperiksa", "disetujui", "admin"]:
        menus.append("Persetujuan Laporan")
    menus.extend(["Cek Kesiapan Alat", "Kelola Master Data"])
    if role == "admin" or st.session_state.username == "admin":
        menus.append("Kelola User")
    
    for m in menus:
        if st.button(m, key="btn_" + m, use_container_width=True):
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
                file_name="backup_database_" + datetime.now().strftime("%Y%m%d") + ".xlsx",
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
    
    st.markdown("##### Filter Periode")
    col_f1, col_f2, col_f3 = st.columns([1.8, 1.5, 2])
    with col_f1:
        mode_periode = st.radio("Pilih Periode", ["Harian", "Bulanan", "Tahunan"], horizontal=True, key="mode_dash")
    with col_f2:
        if mode_periode == "Harian":
            tgl_dash = st.date_input("Tanggal", datetime.now(), key="tgl_dash")
        elif mode_periode == "Bulanan":
            bln_dash = st.selectbox(
                "Bulan",
                ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                 "Juli", "Agustus", "September", "Oktober", "November", "Desember"],
                index=datetime.now().month - 1,
                key="bln_dash"
            )
            thn_dash = st.number_input("Tahun", 2024, 2030, datetime.now().year, key="thn_bln")
        else:
            thn_dash = st.number_input("Tahun", 2024, 2030, datetime.now().year, key="thn_thn")
    
    df_db = pd.read_excel(FILE_DB) if os.path.exists(FILE_DB) else pd.DataFrame()
    df_alat = pd.read_excel(FILE_ALAT) if os.path.exists(FILE_ALAT) else pd.DataFrame()
    
    if os.path.exists(FILE_TITIK):
        list_titik = pd.read_excel(FILE_TITIK)["Kode"].astype(str).tolist()
    else:
        list_titik = ["TP 1", "TP 2", "TP 3", "TP 4", "TP 5", "TP 6"]
    
    df_filter = pd.DataFrame()
    label_periode = ""
    tgl_str_harian = None
    
    if not df_db.empty and "Tanggal" in df_db.columns:
        df_db["Tanggal_dt"] = pd.to_datetime(df_db["Tanggal"], errors="coerce")
        if mode_periode == "Harian":
            tgl_str_harian = tgl_dash.strftime("%Y-%m-%d")
            df_filter = df_db[df_db["Tanggal"] == tgl_str_harian].copy()
            label_periode = tgl_dash.strftime("%d %B %Y")
        elif mode_periode == "Bulanan":
            bulan_map = {
                "Januari": 1, "Februari": 2, "Maret": 3, "April": 4,
                "Mei": 5, "Juni": 6, "Juli": 7, "Agustus": 8,
                "September": 9, "Oktober": 10, "November": 11, "Desember": 12
            }
            bln_angka = bulan_map.get(bln_dash, datetime.now().month)
            df_filter = df_db[
                (df_db["Tanggal_dt"].dt.month == bln_angka) &
                (df_db["Tanggal_dt"].dt.year == thn_dash)
            ].copy()
            label_periode = bln_dash + " " + str(thn_dash)
        else:
            df_filter = df_db[df_db["Tanggal_dt"].dt.year == thn_dash].copy()
            label_periode = "Tahun " + str(thn_dash)
    
    st.caption("Menampilkan data periode: **" + label_periode + "**")
    
    if not df_filter.empty and "Status" in df_filter.columns:
        df_bahaya = df_filter[df_filter["Status"].isin(["Bahaya", "Terpapar"])]
        if not df_bahaya.empty:
            baris_bahaya = []
            for _, r in df_bahaya.iterrows():
                kode = r.get("Kode_Titik", "-")
                lokasi = r.get("Lokasi", "")
                detail = []
                try:
                    if float(r.get("H2S", 0) or 0) > 0:
                        detail.append("H2S " + str(r["H2S"]) + " ppm")
                    if float(r.get("SO2", 0) or 0) > 0:
                        detail.append("SO2 " + str(r["SO2"]) + " ppm")
                    if float(r.get("CO", 0) or 0) > 0:
                        detail.append("CO " + str(r["CO"]) + " ppm")
                    o2v = float(r.get("O2", 20.9) or 20.9)
                    if o2v < 19.5 or o2v > 23.5:
                        detail.append("O2 " + str(o2v) + "%")
                except Exception:
                    pass
                det = " — " + ", ".join(detail) if detail else ""
                baris_bahaya.append("• <b>" + str(kode) + "</b> (" + str(lokasi) + ")" + det)
            st.markdown(
                '<div class="alert-bahaya"><div style="font-weight:700;color:#C62828;margin-bottom:6px;">STATUS BAHAYA</div>'
                + '<div style="color:#B71C1C;font-size:13.5px;line-height:1.6;">'
                + "<br>".join(baris_bahaya) + "</div></div>",
                unsafe_allow_html=True
            )
        
        df_waspada = df_filter[df_filter["Status"].isin(["Waspada", "O2 Tidak Normal"])]
        if not df_waspada.empty:
            baris_waspada = []
            for _, r in df_waspada.iterrows():
                kode = r.get("Kode_Titik", "-")
                lokasi = r.get("Lokasi", "")
                detail = []
                try:
                    if float(r.get("H2S", 0) or 0) > 0:
                        detail.append("H2S " + str(r["H2S"]) + " ppm")
                    if float(r.get("SO2", 0) or 0) > 0:
                        detail.append("SO2 " + str(r["SO2"]) + " ppm")
                    if float(r.get("CO", 0) or 0) > 0:
                        detail.append("CO " + str(r["CO"]) + " ppm")
                    if r.get("Noise") is not None and float(r.get("Noise", 0) or 0) > 55:
                        detail.append("Noise " + str(r["Noise"]) + " dB")
                except Exception:
                    pass
                det = " — " + ", ".join(detail) if detail else ""
                baris_waspada.append("• <b>" + str(kode) + "</b> (" + str(lokasi) + ")" + det)
            st.markdown(
                '<div class="alert-waspada"><div style="font-weight:700;color:#F57F17;margin-bottom:6px;">STATUS WASPADA</div>'
                + '<div style="color:#E65100;font-size:13.5px;line-height:1.6;">'
                + "<br>".join(baris_waspada) + "</div></div>",
                unsafe_allow_html=True
            )
    
    if mode_periode == "Harian" and tgl_str_harian:
        def cek_shift(df, shift_name):
            if df.empty:
                return 0, list_titik
            ds = df[df["Shift"] == shift_name] if "Shift" in df.columns else pd.DataFrame()
            sudah = ds["Kode_Titik"].astype(str).unique().tolist() if not ds.empty and "Kode_Titik" in ds.columns else []
            belum = [t for t in list_titik if t not in sudah]
            return len(sudah), belum
        
        n_ds, belum_ds = cek_shift(df_filter, "Day Shift")
        n_ns, belum_ns = cek_shift(df_filter, "Night Shift")
        total_tp = len(list_titik)
        teks_ds = "Day Shift &nbsp;<b>" + str(n_ds) + "/" + str(total_tp) + "</b> titik"
        teks_ds += (" &nbsp;|&nbsp; Belum: " + ", ".join(belum_ds)) if belum_ds else " &nbsp;|&nbsp; Lengkap"
        teks_ns = "Night Shift &nbsp;<b>" + str(n_ns) + "/" + str(total_tp) + "</b> titik"
        teks_ns += (" &nbsp;|&nbsp; Belum: " + ", ".join(belum_ns)) if belum_ns else " &nbsp;|&nbsp; Lengkap"
        st.markdown(
            '<div class="info-hari"><div style="font-weight:700;color:#0D47A1;margin-bottom:6px;">Data hari ini: '
            + label_periode + '</div><div style="color:#1565C0;font-size:13.5px;line-height:1.7;">'
            + teks_ds + "<br>" + teks_ns + "</div></div>",
            unsafe_allow_html=True
        )
    
    total_data = len(df_filter)
    jml_aman = jml_waspada = jml_bahaya = 0
    if not df_filter.empty and "Status" in df_filter.columns:
        jml_aman = len(df_filter[df_filter["Status"] == "Aman"])
        jml_waspada = len(df_filter[df_filter["Status"].isin(["Waspada", "O2 Tidak Normal"])])
        jml_bahaya = len(df_filter[df_filter["Status"].isin(["Bahaya", "Terpapar"])])
    pct_aman = (jml_aman / total_data * 100) if total_data > 0 else 0
    pct_waspada = (jml_waspada / total_data * 100) if total_data > 0 else 0
    pct_bahaya = (jml_bahaya / total_data * 100) if total_data > 0 else 0
    jml_h2s = len(df_filter[df_filter["H2S"] > 0]) if not df_filter.empty and "H2S" in df_filter.columns else 0
    jml_so2 = len(df_filter[df_filter["SO2"] > 0]) if not df_filter.empty and "SO2" in df_filter.columns else 0
    titik_sudah = df_filter["Kode_Titik"].nunique() if not df_filter.empty and "Kode_Titik" in df_filter.columns else 0
    
    dari_file = (
        pd.read_excel("master_daftar_alat.xlsx")
        if os.path.exists("master_daftar_alat.xlsx")
        else pd.DataFrame({"Nama_Alat": ["Gas Detector 1", "Gas Detector 2", "SO2 Detector 1", "SO2 Detector 2"]})
    )
    total_alat = len(dari_file)
    siap = 0
    if not df_alat.empty:
        for nama in dari_file["Nama_Alat"]:
            last = df_alat[df_alat["Nama_Alat"] == nama].tail(1)
            if len(last) > 0 and last.iloc[0]["Status"] == "Siap Pakai":
                siap += 1
    
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Total Data</div>'
            '<div class="metric-value">' + str(total_data) + '</div>'
            '<div class="metric-sub" style="color:#64748B;">pengukuran</div></div>',
            unsafe_allow_html=True
        )
    with c2:
        warna_aman = "#2E7D32" if pct_aman >= 90 else "#F9A825" if pct_aman >= 70 else "#C62828"
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Status Aman</div>'
            '<div class="metric-value" style="color:' + warna_aman + ';">' + str(jml_aman) + '</div>'
            '<div class="metric-sub" style="color:' + warna_aman + ';">' + "{:.1f}".format(pct_aman) + '%</div></div>',
            unsafe_allow_html=True
        )
    with c3:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Waspada</div>'
            '<div class="metric-value" style="color:#F9A825;">' + str(jml_waspada) + '</div>'
            '<div class="metric-sub" style="color:#F9A825;">' + "{:.1f}".format(pct_waspada) + '%</div></div>',
            unsafe_allow_html=True
        )
    with c4:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Bahaya</div>'
            '<div class="metric-value" style="color:#C62828;">' + str(jml_bahaya) + '</div>'
            '<div class="metric-sub" style="color:#C62828;">' + "{:.1f}".format(pct_bahaya) + '%</div></div>',
            unsafe_allow_html=True
        )
    
    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">H2S Terdeteksi</div>'
            '<div class="metric-value">' + str(jml_h2s) + '</div>'
            '<div class="metric-sub" style="color:#64748B;">kali</div></div>',
            unsafe_allow_html=True
        )
    with c6:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">SO2 Terdeteksi</div>'
            '<div class="metric-value">' + str(jml_so2) + '</div>'
            '<div class="metric-sub" style="color:#64748B;">kali</div></div>',
            unsafe_allow_html=True
        )
    with c7:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Titik Terukur</div>'
            '<div class="metric-value">' + str(titik_sudah) + '</div>'
            '<div class="metric-sub" style="color:#64748B;">titik pantau</div></div>',
            unsafe_allow_html=True
        )
    with c8:
        st.markdown(
            '<div class="metric-card"><div class="metric-label">Alat Siap Pakai</div>'
            '<div class="metric-value">' + str(siap)
            + '<span style="font-size:16px;color:#64748B;"> / ' + str(total_alat) + '</span></div>'
            '<div class="metric-sub" style="color:#64748B;">unit alat</div></div>',
            unsafe_allow_html=True
        )
    
    st.markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
    col_kiri, col_kanan = st.columns(2)
    with col_kiri:
        st.markdown("#### Data Pemantauan (" + label_periode + ")")
        if not df_filter.empty:
            kolom_tampil = [c for c in ["Waktu", "Shift", "Kode_Titik", "Lokasi", "H2S", "O2", "SO2", "Status"] if c in df_filter.columns]
            st.dataframe(df_filter[kolom_tampil], use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data pada filter ini.")
    with col_kanan:
        st.markdown("#### Status Alat (Terkini)")
        if not df_alat.empty:
            status_terakhir = df_alat.sort_values("Tanggal_Inspeksi").groupby("Nama_Alat").tail(1)
            kolom_alat = [c for c in ["Nama_Alat", "Status", "Tanggal_Inspeksi", "Tanggal_Berikutnya"] if c in status_terakhir.columns]
            st.dataframe(status_terakhir[kolom_alat], use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data pada filter ini.")
    
    st.markdown("---")
    st.caption("Data Realtime Sistem Pemantauan Udara Sesaat • PEPC Zona 12")

elif menu == "Input Data":
    tampilkan_halaman_input(st.session_state.nama_user)

else:
    mapping = {
        "Review Data": "2. Review Data",
        "Download Laporan": "3. Download Laporan",
        "Persetujuan Laporan": "3b. Persetujuan Laporan",
        "Cek Kesiapan Alat": "4. Cek Kesiapan Alat",
        "Kelola Master Data": "5. Kelola Master Data",
        "Kelola User": "6. Kelola User"
    }
    menu_lama = mapping.get(menu, menu)
    tampilkan_halaman_output(
        menu_lama,
        st.session_state.nama_user,
        st.session_state.username,
        st.session_state.get("role", "petugas")
    )

st.markdown("""
<div class="footer">
    Digitalisasi Monitoring Kualitas Udara • Delta Candra • KP Unigoro 2026
</div>
""", unsafe_allow_html=True)