import streamlit as st
import pandas as pd
import os
from datetime import datetime
from gui_input import tampilkan_halaman_input
from gui_output import tampilkan_halaman_output
from supabase import create_client

st.set_page_config(
    page_title="Aplikasi Pemantauan Udara Sesaat HSSE - PEPC Zona 12",
    page_icon="logo_icon.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

FILE_USER = "master_users.xlsx"
FILE_DB = "database_monitoring_jtb.xlsx"
FILE_ALAT = "database_inspeksi_alat.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"


def get_supabase():
    return create_client(st.secrets["supabase"]["url"], st.secrets["supabase"]["key"])


def load_monitoring_dashboard():
    try:
        sb = get_supabase()
        res = sb.table("monitoring").select("*").order("id", desc=False).execute()
        rows = res.data or []
        if not rows:
            return pd.DataFrame()
        df = pd.DataFrame(rows)
        rename = {
            "tanggal": "Tanggal",
            "bulan": "Bulan",
            "shift": "Shift",
            "waktu": "Waktu",
            "petugas": "Petugas",
            "kode_titik": "Kode_Titik",
            "lokasi": "Lokasi",
            "h2s": "H2S",
            "o2": "O2",
            "co": "CO",
            "lel": "LEL",
            "so2": "SO2",
            "noise": "Noise",
            "kesimpulan": "Kesimpulan",
            "path_foto": "Path_Foto",
            "path_foto2": "Path_Foto2",
            "status": "Status",
        }
        return df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    except Exception:
        if os.path.exists(FILE_DB):
            return pd.read_excel(FILE_DB)
        return pd.DataFrame()


def load_cek_alat_dashboard():
    try:
        sb = get_supabase()
        res = sb.table("cek_kesiapan_alat").select("*").order("id", desc=False).execute()
        rows = res.data or []
        if not rows:
            return pd.DataFrame()
        df = pd.DataFrame(rows)
        rename = {
            "nama_alat": "Nama_Alat",
            "tanggal_inspeksi": "Tanggal_Inspeksi",
            "tanggal_berikutnya": "Tanggal_Berikutnya",
            "indikator": "Indikator",
            "status": "Status",
            "petugas": "Petugas",
            "keterangan": "Keterangan",
        }
        return df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    except Exception:
        if os.path.exists(FILE_ALAT):
            return pd.read_excel(FILE_ALAT)
        return pd.DataFrame()


def load_titik_kode():
    try:
        sb = get_supabase()
        res = sb.table("titik_pantau").select("kode").order("id", desc=False).execute()
        if res.data:
            return [str(r["kode"]) for r in res.data if r.get("kode")]
    except Exception:
        pass
    if os.path.exists(FILE_TITIK):
        df = pd.read_excel(FILE_TITIK)
        if "Kode" in df.columns:
            return df["Kode"].astype(str).tolist()
    return ["TP 1", "TP 2", "TP 3", "TP 4", "TP 5", "TP 6"]


def login_supabase(username, password):
    try:
        sb = get_supabase()
        res = (
            sb.table("users")
            .select("username, password, nama, role")
            .eq("username", username.strip())
            .limit(1)
            .execute()
        )
        if not res.data:
            return None
        row = res.data[0]
        if str(row.get("password", "")) == str(password):
            return {
                "username": str(row.get("username", "")),
                "nama": str(row.get("nama", "") or row.get("username", "")),
                "role": str(row.get("role", "petugas") or "petugas"),
            }
        return None
    except Exception:
        return None


st.markdown(
    """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', Arial, sans-serif; }
    .stApp { background-color: #F0F4F8; }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0A3D91 0%, #0D47A1 50%, #1565C0 100%);
        padding-top: 1.2rem;
    }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    [data-testid="stSidebar"] .stButton > button {
        background-color: rgba(255,255,255,0.08);
        color: #FFFFFF !important;
        border: 1px solid rgba(255,255,255,0.15);
        border-radius: 8px;
        width: 100%;
        text-align: left;
        font-weight: 500;
    }
    [data-testid="stSidebar"] .stButton > button:hover {
        background-color: rgba(255,255,255,0.18);
        border-color: rgba(255,255,255,0.35);
    }
    .footer {
        text-align: center;
        color: #64748B;
        font-size: 0.85rem;
        padding: 1.5rem 0 0.5rem 0;
    }
</style>
""",
    unsafe_allow_html=True,
)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""
if "nama_user" not in st.session_state:
    st.session_state.nama_user = ""
if "role" not in st.session_state:
    st.session_state.role = "petugas"

if not st.session_state.logged_in:
    st.markdown("## Aplikasi Pemantauan Udara Sesaat")
    st.caption("HSSE – PEPC Zona 12")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Masuk", type="primary")
        if submit:
            user = login_supabase(username, password)
            if user:
                st.session_state.logged_in = True
                st.session_state.username = user["username"]
                st.session_state.nama_user = user["nama"]
                st.session_state.role = user["role"]
                st.rerun()
            else:
                st.error("Username atau password salah.")
    st.stop()

with st.sidebar:
    st.markdown("### Menu")
    st.caption(st.session_state.nama_user + " (" + st.session_state.role + ")")
    role = st.session_state.role

    menu_items = ["Dashboard", "Input Data", "Review Data", "Download Laporan"]
    if role in ["diperiksa", "disetujui", "admin"]:
        menu_items.append("Persetujuan Laporan")
    menu_items.extend(["Cek Kesiapan Alat", "Kelola Master Data"])
    if role == "admin":
        menu_items.append("Kelola User")

    menu = st.radio("Navigasi", menu_items, label_visibility="collapsed")
    if st.button("Keluar"):
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.nama_user = ""
        st.session_state.role = "petugas"
        st.rerun()

if menu == "Dashboard":
    st.markdown("### Dashboard Pemantauan Udara Sesaat")
    st.caption("Data Realtime • PEPC Zona 12")

    df_db = load_monitoring_dashboard()
    df_alat = load_cek_alat_dashboard()
    list_tp = load_titik_kode()
    total_tp = len(list_tp) if list_tp else 6

    c_f1, c_f2, c_f3 = st.columns(3)
    with c_f1:
        mode_filter = st.selectbox("Periode", ["Harian", "Bulanan", "Tahunan"], index=0)
    with c_f2:
        if mode_filter == "Harian":
            tgl_filter = st.date_input("Tanggal", datetime.now())
            bulan_filter = None
            tahun_filter = tgl_filter.year
        elif mode_filter == "Bulanan":
            bulan_filter = st.selectbox(
                "Bulan",
                [
                    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
                ],
                index=datetime.now().month - 1,
            )
            tgl_filter = None
            tahun_filter = st.number_input("Tahun", 2024, 2030, datetime.now().year, key="thn_bln")
        else:
            tgl_filter = None
            bulan_filter = None
            tahun_filter = st.number_input("Tahun", 2024, 2030, datetime.now().year, key="thn_thn")
    with c_f3:
        st.write("")
        st.write("")

    df_filter = df_db.copy() if not df_db.empty else pd.DataFrame()
    if not df_filter.empty and "Tanggal" in df_filter.columns:
        try:
            df_filter["_dt"] = pd.to_datetime(df_filter["Tanggal"], errors="coerce")
            if mode_filter == "Harian" and tgl_filter is not None:
                df_filter = df_filter[df_filter["Tanggal"].astype(str) == tgl_filter.strftime("%Y-%m-%d")]
            elif mode_filter == "Bulanan" and bulan_filter:
                bulan_map = {
                    "Januari": 1, "Februari": 2, "Maret": 3, "April": 4, "Mei": 5, "Juni": 6,
                    "Juli": 7, "Agustus": 8, "September": 9, "Oktober": 10, "November": 11, "Desember": 12,
                }
                m = bulan_map.get(bulan_filter, datetime.now().month)
                df_filter = df_filter[
                    (df_filter["_dt"].dt.month == m) & (df_filter["_dt"].dt.year == int(tahun_filter))
                ]
            elif mode_filter == "Tahunan":
                df_filter = df_filter[df_filter["_dt"].dt.year == int(tahun_filter)]
        except Exception:
            pass

    jml_data = len(df_filter) if not df_filter.empty else 0
    jml_bahaya = 0
    jml_waspada = 0
    jml_aman = 0
    if not df_filter.empty and "Status" in df_filter.columns:
        jml_bahaya = int((df_filter["Status"].astype(str) == "Bahaya").sum())
        jml_waspada = int((df_filter["Status"].astype(str) == "Waspada").sum())
        jml_aman = int((df_filter["Status"].astype(str) == "Aman").sum())

    jml_siap = 0
    jml_tidak_siap = 0
    jml_belum = 0
    status_terakhir = pd.DataFrame()
    if not df_alat.empty and "Nama_Alat" in df_alat.columns:
        sort_col = "Tanggal_Inspeksi" if "Tanggal_Inspeksi" in df_alat.columns else "id"
        try:
            tmp = df_alat.copy()
            if sort_col in tmp.columns:
                tmp["_sort"] = pd.to_datetime(tmp[sort_col], errors="coerce")
                status_terakhir = tmp.sort_values("_sort").groupby("Nama_Alat", as_index=False).tail(1)
            else:
                status_terakhir = tmp.groupby("Nama_Alat", as_index=False).tail(1)
        except Exception:
            status_terakhir = df_alat.groupby("Nama_Alat", as_index=False).tail(1)

        if "Status" in status_terakhir.columns:
            jml_siap = int((status_terakhir["Status"].astype(str) == "Siap Pakai").sum())
            jml_tidak_siap = int((status_terakhir["Status"].astype(str) == "Tidak Siap").sum())
        jml_belum = max(0, len(status_terakhir) - jml_siap - jml_tidak_siap)

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1:
        st.metric("Total Data", jml_data)
    with m2:
        st.metric("Aman", jml_aman)
    with m3:
        st.metric("Waspada", jml_waspada)
    with m4:
        st.metric("Bahaya", jml_bahaya)
    with m5:
        st.metric("Alat Siap Pakai", jml_siap)

    if not df_filter.empty and "Status" in df_filter.columns and jml_bahaya > 0:
        bahaya = df_filter[df_filter["Status"].astype(str) == "Bahaya"]
        tp_bahaya = (
            bahaya["Kode_Titik"].astype(str).unique().tolist()
            if "Kode_Titik" in bahaya.columns
            else []
        )
        st.error("**ALERT BAHAYA** — Titik: " + (", ".join(tp_bahaya) if tp_bahaya else "-"))
    if not df_filter.empty and "Status" in df_filter.columns and jml_waspada > 0:
        wasp = df_filter[df_filter["Status"].astype(str) == "Waspada"]
        tp_w = wasp["Kode_Titik"].astype(str).unique().tolist() if "Kode_Titik" in wasp.columns else []
        st.warning("**WASPADA** — Titik: " + (", ".join(tp_w) if tp_w else "-"))

    if mode_filter == "Harian" and tgl_filter is not None:
        st.markdown("#### Data hari ini:")
        df_hari = df_filter.copy()
        for shift_name in ["Day Shift", "Night Shift"]:
            if df_hari.empty or "Shift" not in df_hari.columns:
                sudah = []
            else:
                d = df_hari[df_hari["Shift"] == shift_name]
                sudah = (
                    d["Kode_Titik"].astype(str).unique().tolist()
                    if not d.empty and "Kode_Titik" in d.columns
                    else []
                )
            belum = [t for t in list_tp if t not in sudah]
            if belum:
                st.info(
                    shift_name
                    + " — belum input: "
                    + ", ".join(belum)
                    + " ("
                    + str(len(sudah))
                    + "/"
                    + str(total_tp)
                    + ")"
                )
            else:
                st.success(shift_name + " — lengkap (" + str(total_tp) + "/" + str(total_tp) + ")")

    st.markdown("---")
    col_kiri, col_kanan = st.columns(2)
    with col_kiri:
        st.markdown("#### Data Pemantauan")
        if not df_filter.empty:
            kolom_tampil = [
                c
                for c in [
                    "Tanggal", "Waktu", "Shift", "Kode_Titik", "Lokasi",
                    "H2S", "O2", "SO2", "Status",
                ]
                if c in df_filter.columns
            ]
            st.dataframe(df_filter[kolom_tampil], use_container_width=True, hide_index=True)
        else:
            st.info("Belum ada data pada filter ini.")
    with col_kanan:
        st.markdown("#### Status Alat (Terkini)")
        if not status_terakhir.empty:
            kolom_alat = [
                c
                for c in ["Nama_Alat", "Status", "Tanggal_Inspeksi", "Tanggal_Berikutnya"]
                if c in status_terakhir.columns
            ]
            st.dataframe(status_terakhir[kolom_alat], use_container_width=True, hide_index=True)
            st.caption(
                "Siap: "
                + str(jml_siap)
                + " | Tidak siap: "
                + str(jml_tidak_siap)
                + " | Lainnya: "
                + str(jml_belum)
            )
        else:
            st.info("Belum ada data cek kesiapan alat.")

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
        "Kelola User": "6. Kelola User",
    }
    menu_lama = mapping.get(menu, menu)
    tampilkan_halaman_output(
        menu_lama,
        st.session_state.nama_user,
        st.session_state.username,
        st.session_state.role,
    )

st.markdown(
    """
<div class="footer">
    Digitalisasi Monitoring Kualitas Udara • Delta Candra • KP Unigoro 2026
</div>
""",
    unsafe_allow_html=True,
)