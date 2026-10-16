import streamlit as st
import pandas as pd
import os
from datetime import datetime
from pathlib import Path
import pytz
from supabase import create_client

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"
FILE_KESIMPULAN = "master_kesimpulan.xlsx"
FOLDER_FOTO = "foto_monitoring"

DEFAULT_TITIK = {
    "TP 1": {"Lokasi": "Pemukiman Ngasem", "Lat": "-7.1234", "Long": "111.6789", "LatNS": "-7.1235", "LongNS": "111.6790"},
    "TP 2": {"Lokasi": "Pemukiman Bandungrejo", "Lat": "-7.1245", "Long": "111.6791", "LatNS": "-7.1246", "LongNS": "111.6792"},
    "TP 3": {"Lokasi": "Petak 28", "Lat": "-7.1256", "Long": "111.6801", "LatNS": "-7.1257", "LongNS": "111.6803"},
    "TP 4": {"Lokasi": "Crossing Gayam", "Lat": "-7.1267", "Long": "111.6812", "LatNS": "-7.1268", "LongNS": "111.6814"},
    "TP 5": {"Lokasi": "Selatan Dormitory", "Lat": "-7.1278", "Long": "111.6823", "LatNS": "-7.1279", "LongNS": "111.6825"},
    "TP 6": {"Lokasi": "Dusun Besaran Ngasem", "Lat": "-7.1289", "Long": "111.6834", "LatNS": "-7.1290", "LongNS": "111.6836"},
}

DEFAULT_KESIMPULAN = [
    "Tidak terdeteksi adanya Gas H2S dan SO2 di seluruh lokasi pemantauan",
    "Terdeteksi Gas H2S di bawah baku mutu",
    "Terdeteksi Gas SO2, perlu investigasi lanjutan",
    "Terdeteksi Gas H2S dan SO2, segera tindak lanjut",
]


def get_supabase():
    url = st.secrets["supabase"]["url"]
    key = st.secrets["supabase"]["key"]
    return create_client(url, key)


def load_titik():
    try:
        sb = get_supabase()
        res = sb.table("titik_pantau").select("*").order("id", desc=False).execute()
        if res.data:
            titik = {}
            for row in res.data:
                kode = str(row.get("kode", "") or "")
                if not kode:
                    continue
                titik[kode] = {
                    "Lokasi": str(row.get("lokasi", "") or ""),
                    "Lat": str(row.get("latitude", "") or ""),
                    "Long": str(row.get("longitude", "") or ""),
                    "LatNS": str(row.get("latitude_ns", "") or row.get("latitude", "") or ""),
                    "LongNS": str(row.get("longitude_ns", "") or row.get("longitude", "") or ""),
                }
            if titik:
                return titik
    except Exception:
        pass
    if os.path.exists(FILE_TITIK):
        df = pd.read_excel(FILE_TITIK)
        titik = {}
        for _, row in df.iterrows():
            titik[str(row["Kode"])] = {
                "Lokasi": str(row["Lokasi"]),
                "Lat": str(row["Latitude"]),
                "Long": str(row["Longitude"]),
                "LatNS": str(row["Latitude_NS"]),
                "LongNS": str(row["Longitude_NS"]),
            }
        return titik
    return DEFAULT_TITIK


def load_kesimpulan_list():
    try:
        sb = get_supabase()
        res = sb.table("kesimpulan").select("teks").order("id", desc=False).execute()
        if res.data:
            lst = [str(r["teks"]) for r in res.data if r.get("teks")]
            if lst:
                return lst
    except Exception:
        pass
    if os.path.exists(FILE_KESIMPULAN):
        df = pd.read_excel(FILE_KESIMPULAN)
        if "Kesimpulan" in df.columns:
            return [str(x) for x in df["Kesimpulan"].tolist() if str(x).strip()]
    return list(DEFAULT_KESIMPULAN)


def load_monitoring_for_cek():
    try:
        sb = get_supabase()
        res = sb.table("monitoring").select("tanggal,shift,kode_titik").execute()
        if res.data:
            df = pd.DataFrame(res.data)
            return df.rename(columns={
                "tanggal": "Tanggal",
                "shift": "Shift",
                "kode_titik": "Kode_Titik",
            })
    except Exception:
        pass
    if os.path.exists(FILE_DB):
        return pd.read_excel(FILE_DB)
    return pd.DataFrame()


def tampilkan_halaman_input(nama_user=""):
    st.markdown("### 1. Input Data Pemantauan Udara Sesaat")
    st.caption("Isi data hasil pengukuran. Noise opsional. Foto 1 wajib, foto 2 opsional.")

    titik_data = load_titik()
    list_kesimpulan = load_kesimpulan_list()
    if not list_kesimpulan:
        list_kesimpulan = list(DEFAULT_KESIMPULAN)

    tz = pytz.timezone("Asia/Jakarta")
    now = datetime.now(tz)

    c1, c2, c3 = st.columns(3)
    with c1:
        tanggal_in = st.date_input("Tanggal *", now.date())
        bulan_nama = tanggal_in.strftime("%B")
        bulan_id = {
            "January": "Januari", "February": "Februari", "March": "Maret", "April": "April",
            "May": "Mei", "June": "Juni", "July": "Juli", "August": "Agustus",
            "September": "September", "October": "Oktober", "November": "November", "December": "Desember",
        }
        bulan_in = bulan_id.get(bulan_nama, bulan_nama)
    with c2:
        shift_in = st.selectbox("Shift *", ["Day Shift", "Night Shift"])
    with c3:
        waktu_in = st.time_input("Waktu *", now.time())

    c4, c5, c6 = st.columns(3)
    with c4:
        petugas_in = st.text_input("Petugas *", value=nama_user or "")
    with c5:
        tp_in = st.selectbox("Titik Pantau (TP) *", list(titik_data.keys()))
    with c6:
        lokasi_tampil = titik_data.get(tp_in, {}).get("Lokasi", "-")
        st.text_input("Lokasi", value=lokasi_tampil, disabled=True)

    st.markdown("---")
    st.markdown("#### Hasil Pengukuran")

    g1, g2, g3 = st.columns(3)
    with g1:
        h2s = st.number_input("H2S (ppm) *", min_value=0.0, value=0.0, step=0.1, format="%.2f")
        o2 = st.number_input("O2 (%) *", min_value=0.0, value=20.9, step=0.1, format="%.1f")
    with g2:
        co = st.number_input("CO (ppm) *", min_value=0.0, value=0.0, step=0.1, format="%.1f")
        lel = st.number_input("LEL (%) *", min_value=0.0, value=0.0, step=0.1, format="%.1f")
    with g3:
        so2 = st.number_input("SO2 (ppm) *", min_value=0.0, value=0.0, step=0.01, format="%.2f")
        noise = st.number_input("Noise dB(A) (opsional)", min_value=0.0, value=0.0, step=0.1, format="%.1f")

    st.markdown("---")
    opsi_kes = list_kesimpulan + ["Lainnya (ketik manual)"]
    pilih_kes = st.selectbox("Pilih Kesimpulan *", opsi_kes)
    if pilih_kes == "Lainnya (ketik manual)":
        kesimpulan_in = st.text_area("Tulis kesimpulan manual *", height=80, placeholder="Ketik kesimpulan...")
    else:
        kesimpulan_in = pilih_kes

    st.markdown("---")
    st.markdown("#### Dokumentasi Foto")
    f1, f2 = st.columns(2)
    with f1:
        foto1 = st.file_uploader("Foto 1 (wajib) *", type=["jpg", "jpeg", "png"], key="foto1_in")
        if foto1:
            st.image(foto1, width=220)
    with f2:
        foto2 = st.file_uploader("Foto 2 (opsional)", type=["jpg", "jpeg", "png"], key="foto2_in")
        if foto2:
            st.image(foto2, width=220)

    st.markdown("---")

    if "konfirmasi_dobel" not in st.session_state:
        st.session_state.konfirmasi_dobel = False

    if st.button("Simpan Data", type="primary"):
        kesalahan = []
        if not petugas_in.strip():
            kesalahan.append("Petugas")
        if foto1 is None:
            kesalahan.append("Foto 1")
        if not str(kesimpulan_in or "").strip():
            kesalahan.append("Kesimpulan")

        if kesalahan:
            st.error("Masih ada data yang belum valid: **" + ", ".join(kesalahan) + "**")
            return

        tgl_str = tanggal_in.strftime("%Y-%m-%d")
        df_cek = load_monitoring_for_cek()
        sudah_ada = False
        if not df_cek.empty:
            try:
                mask = (
                    (df_cek["Tanggal"].astype(str) == tgl_str)
                    & (df_cek["Shift"].astype(str) == shift_in)
                    & (df_cek["Kode_Titik"].astype(str) == tp_in)
                )
                sudah_ada = bool(mask.any())
            except Exception:
                sudah_ada = False

        if sudah_ada and not st.session_state.konfirmasi_dobel:
            st.session_state.konfirmasi_dobel = True
            st.warning(
                "Data **" + tp_in + " | " + shift_in + " | " + tgl_str
                + "** sudah pernah diinput. Klik **Simpan Data** sekali lagi untuk tetap menyimpan."
            )
            return

        detail_status = []
        status_gas = "Aman"

        if h2s > 1.0:
            detail_status.append("H2S {:.2f} > baku mutu 1 ppm".format(h2s))
            status_gas = "Bahaya"
        elif h2s > 0:
            detail_status.append("H2S {:.2f} terdeteksi (≤ 1 ppm)".format(h2s))
            if status_gas == "Aman":
                status_gas = "Waspada"

        if so2 > 0.057:
            detail_status.append("SO2 {:.3f} > baku mutu 0.057 ppm".format(so2))
            status_gas = "Bahaya"
        elif so2 > 0:
            detail_status.append("SO2 {:.3f} terdeteksi (≤ 0.057 ppm)".format(so2))
            if status_gas == "Aman":
                status_gas = "Waspada"

        if o2 < 19.5 or o2 > 23.5:
            detail_status.append("O2 {:.1f}% tidak normal (harus 19.5–23.5)".format(o2))
            status_gas = "Bahaya"

        if co > 8.72:
            detail_status.append("CO {:.1f} > baku mutu 8.72 ppm".format(co))
            status_gas = "Bahaya"

        if noise and noise > 55:
            detail_status.append("Noise {:.1f} dB melebihi 55 dB".format(noise))
            if status_gas == "Aman":
                status_gas = "Waspada"

        path_foto1, path_foto2 = "", ""
        shift_folder = "Day_Shift" if shift_in == "Day Shift" else "Night_Shift"
        folder = os.path.join(FOLDER_FOTO, tgl_str, shift_folder)
        os.makedirs(folder, exist_ok=True)
        kode_file = tp_in.replace(" ", "_")

        if foto1 is not None:
            ext = Path(foto1.name).suffix.lower() or ".jpg"
            path_foto1 = os.path.join(folder, kode_file + "_1" + ext)
            with open(path_foto1, "wb") as f:
                f.write(foto1.getbuffer())

        if foto2 is not None:
            ext = Path(foto2.name).suffix.lower() or ".jpg"
            path_foto2 = os.path.join(folder, kode_file + "_2" + ext)
            with open(path_foto2, "wb") as f:
                f.write(foto2.getbuffer())

        waktu_str = str(waktu_in)
        if len(waktu_str) >= 5:
            waktu_str = waktu_str[:5]

        payload = {
            "tanggal": tgl_str,
            "bulan": bulan_in,
            "shift": shift_in,
            "waktu": waktu_str,
            "petugas": petugas_in.strip(),
            "kode_titik": tp_in,
            "lokasi": titik_data[tp_in]["Lokasi"],
            "h2s": float(h2s),
            "o2": float(o2),
            "co": float(co),
            "lel": float(lel),
            "so2": float(so2),
            "noise": float(noise) if noise and float(noise) > 0 else None,
            "kesimpulan": str(kesimpulan_in).strip(),
            "path_foto": path_foto1 if path_foto1 else "",
            "path_foto2": path_foto2 if path_foto2 else "",
            "status": status_gas,
        }

        try:
            get_supabase().table("monitoring").insert(payload).execute()
        except Exception as e:
            st.error("Gagal simpan ke Supabase: " + str(e))
            return

        st.session_state.konfirmasi_dobel = False

        if status_gas == "Bahaya":
            st.error("STATUS **BAHAYA**: " + " | ".join(detail_status))
        elif status_gas == "Waspada":
            st.warning("STATUS **WASPADA**: " + " | ".join(detail_status))
        else:
            st.success("STATUS **AMAN** — semua parameter dalam batas normal")

        st.success("Data berhasil disimpan!")
        st.balloons()