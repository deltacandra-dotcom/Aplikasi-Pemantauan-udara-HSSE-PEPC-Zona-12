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


def simpan_monitoring_supabase(row: dict):
    try:
        sb = get_supabase()
        payload = {
            "tanggal": str(row.get("Tanggal", "")),
            "bulan": str(row.get("Bulan", "")),
            "shift": str(row.get("Shift", "")),
            "waktu": str(row.get("Waktu", "")),
            "petugas": str(row.get("Petugas", "")),
            "kode_titik": str(row.get("Kode_Titik", "")),
            "lokasi": str(row.get("Lokasi", "")),
            "h2s": float(row["H2S"]) if row.get("H2S") is not None else None,
            "o2": float(row["O2"]) if row.get("O2") is not None else None,
            "co": float(row["CO"]) if row.get("CO") is not None else None,
            "lel": float(row["LEL"]) if row.get("LEL") is not None else None,
            "so2": float(row["SO2"]) if row.get("SO2") is not None else None,
            "noise": float(row["Noise"]) if row.get("Noise") is not None else None,
            "kesimpulan": str(row.get("Kesimpulan", "") or ""),
            "path_foto": str(row.get("Path_Foto", "") or ""),
            "path_foto2": str(row.get("Path_Foto2", "") or ""),
            "status": str(row.get("Status", "")),
        }
        sb.table("monitoring").insert(payload).execute()
        return True, "Tersimpan di Supabase"
    except Exception as e:
        return False, str(e)


def load_titik():
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
    data = []
    for kode, v in DEFAULT_TITIK.items():
        data.append({
            "Kode": kode,
            "Lokasi": v["Lokasi"],
            "Latitude": v["Lat"],
            "Longitude": v["Long"],
            "Latitude_NS": v["LatNS"],
            "Longitude_NS": v["LongNS"],
        })
    pd.DataFrame(data).to_excel(FILE_TITIK, index=False)
    return DEFAULT_TITIK


def load_kesimpulan():
    if os.path.exists(FILE_KESIMPULAN):
        df = pd.read_excel(FILE_KESIMPULAN)
        if "Kesimpulan" in df.columns:
            return [str(x) for x in df["Kesimpulan"].tolist() if str(x).strip()]
    pd.DataFrame({"Kesimpulan": DEFAULT_KESIMPULAN}).to_excel(FILE_KESIMPULAN, index=False)
    return list(DEFAULT_KESIMPULAN)


def hitung_status(h2s, o2, co, lel, so2, noise):
    """Aman / Waspada / Bahaya sesuai baku mutu."""
    bahaya = False
    waspada = False

    if h2s is not None and h2s > 1.0:
        bahaya = True
    elif h2s is not None and h2s > 0:
        waspada = True

    if so2 is not None and so2 > 0.057:
        bahaya = True
    elif so2 is not None and so2 > 0:
        waspada = True

    if co is not None and co > 8.72:
        bahaya = True
    elif co is not None and co > 0:
        waspada = True

    if o2 is not None and (o2 < 19.5 or o2 > 23.5):
        bahaya = True

    if noise is not None and noise > 55:
        waspada = True

    if bahaya:
        return "Bahaya"
    if waspada:
        return "Waspada"
    return "Aman"


def tampilkan_halaman_input(nama_user=""):
    st.markdown("### 1. Input Data Pemantauan Udara Sesaat")
    st.caption("Isi data pengukuran per titik. Noise opsional. Foto 1 wajib untuk laporan PDF, foto 2 opsional.")

    titik_data = load_titik()
    list_kesimpulan = load_kesimpulan()

    tz = pytz.timezone("Asia/Jakarta")
    now = datetime.now(tz)

    col_a, col_b = st.columns(2)
    with col_a:
        tanggal_in = st.date_input("Tanggal *", now.date())
        shift_in = st.selectbox("Shift *", ["Day Shift", "Night Shift"])
        petugas_in = st.text_input("Petugas *", value=nama_user)
        tp_in = st.selectbox("Titik Pantau *", list(titik_data.keys()))
    with col_b:
        waktu_in = st.text_input("Waktu *", value=now.strftime("%H:%M:%S"))
        bulan_list = [
            "Januari", "Februari", "Maret", "April", "Mei", "Juni",
            "Juli", "Agustus", "September", "Oktober", "November", "Desember",
        ]
        bulan_in = st.selectbox("Bulan *", bulan_list, index=now.month - 1)
        st.text_input("Lokasi (otomatis)", value=titik_data[tp_in]["Lokasi"], disabled=True)

    st.markdown("---")
    st.markdown("##### Hasil Pengukuran")

    c1, c2, c3 = st.columns(3)
    with c1:
        h2s = st.number_input("H2S (ppm) *", min_value=0.0, value=0.0, step=0.1, format="%.2f")
        o2 = st.number_input("O2 (%) *", min_value=0.0, value=20.9, step=0.1, format="%.1f")
    with c2:
        co = st.number_input("CO (ppm) *", min_value=0.0, value=0.0, step=0.1, format="%.1f")
        lel = st.number_input("LEL (%) *", min_value=0.0, value=0.0, step=0.1, format="%.1f")
    with c3:
        so2 = st.number_input("SO2 (ppm) *", min_value=0.0, value=0.0, step=0.01, format="%.3f")
        noise_raw = st.number_input("Noise dB(A) (opsional)", min_value=0.0, value=0.0, step=0.1, format="%.1f")
        noise = None if noise_raw == 0.0 else noise_raw

    if o2 < 19.5 or o2 > 23.5:
        st.warning("O2 di luar rentang normal 19,5% – 23,5%. Status akan Bahaya, data tetap bisa disimpan.")

    st.markdown("---")
    st.markdown("##### Dokumentasi Foto")
    f1, f2 = st.columns(2)
    with f1:
        foto1 = st.file_uploader("Foto 1 (wajib untuk PDF) *", type=["jpg", "jpeg", "png"], key="foto1")
    with f2:
        foto2 = st.file_uploader("Foto 2 (opsional)", type=["jpg", "jpeg", "png"], key="foto2")

    st.markdown("---")
    kesimpulan_in = st.selectbox("Pilih Kesimpulan *", list_kesimpulan)

    # Cek dobel (peringatan 2x klik)
    tgl_str = tanggal_in.strftime("%Y-%m-%d")
    key_dobel = f"dobel_{tgl_str}_{shift_in}_{tp_in}"

    sudah_ada = False
    try:
        sb = get_supabase()
        cek = (
            sb.table("monitoring")
            .select("id")
            .eq("tanggal", tgl_str)
            .eq("shift", shift_in)
            .eq("kode_titik", tp_in)
            .limit(1)
            .execute()
        )
        sudah_ada = bool(cek.data)
    except Exception:
        if os.path.exists(FILE_DB):
            try:
                df_cek = pd.read_excel(FILE_DB)
                sudah_ada = not df_cek[
                    (df_cek["Tanggal"].astype(str) == tgl_str)
                    & (df_cek["Shift"] == shift_in)
                    & (df_cek["Kode_Titik"] == tp_in)
                ].empty
            except Exception:
                pass

    if sudah_ada:
        st.warning(
            f"Data **{tp_in}** ({shift_in}) tanggal **{tgl_str}** sudah pernah diinput. "
            "Klik **Simpan** sekali lagi untuk konfirmasi input ulang."
        )

    label_btn = "Konfirmasi Simpan Ulang" if (
        sudah_ada and st.session_state.get(key_dobel)
    ) else "Simpan Data"

    if st.button(label_btn, type="primary", use_container_width=True):
        # Validasi
        kesalahan = []
        if not petugas_in.strip():
            kesalahan.append("Petugas")
        if not str(waktu_in).strip():
            kesalahan.append("Waktu")
        if h2s is None:
            kesalahan.append("H2S")
        if o2 is None:
            kesalahan.append("O2")
        if co is None:
            kesalahan.append("CO")
        if lel is None:
            kesalahan.append("LEL")
        if so2 is None:
            kesalahan.append("SO2")
        if not kesimpulan_in:
            kesalahan.append("Kesimpulan")

        if kesalahan:
            st.error("Masih ada data wajib: **" + ", ".join(kesalahan) + "**")
            return

        # Dobel: butuh 2x klik
        if sudah_ada and not st.session_state.get(key_dobel):
            st.session_state[key_dobel] = True
            st.warning("Klik **Konfirmasi Simpan Ulang** sekali lagi untuk menyimpan.")
            st.rerun()

        # Simpan foto lokal (jika ada)
        path_foto1 = ""
        path_foto2 = ""
        shift_folder = "Day_Shift" if shift_in == "Day Shift" else "Night_Shift"
        folder = os.path.join(FOLDER_FOTO, tgl_str, shift_folder)
        Path(folder).mkdir(parents=True, exist_ok=True)
        base = tp_in.replace(" ", "_")

        if foto1 is not None:
            ext = foto1.name.split(".")[-1].lower()
            path_foto1 = os.path.join(folder, f"{base}_1.{ext}")
            with open(path_foto1, "wb") as f:
                f.write(foto1.getbuffer())
        if foto2 is not None:
            ext = foto2.name.split(".")[-1].lower()
            path_foto2 = os.path.join(folder, f"{base}_2.{ext}")
            with open(path_foto2, "wb") as f:
                f.write(foto2.getbuffer())

        status_gas = hitung_status(h2s, o2, co, lel, so2, noise)

        data_baru = {
            "Tanggal": [tgl_str],
            "Bulan": [bulan_in],
            "Shift": [shift_in],
            "Waktu": [str(waktu_in)],
            "Petugas": [petugas_in.strip()],
            "Kode_Titik": [tp_in],
            "Lokasi": [titik_data[tp_in]["Lokasi"]],
            "H2S": [h2s],
            "O2": [o2],
            "CO": [co],
            "LEL": [lel],
            "SO2": [so2],
            "Noise": [noise if noise is not None else None],
            "Kesimpulan": [kesimpulan_in],
            "Path_Foto": [path_foto1],
            "Path_Foto2": [path_foto2],
            "Status": [status_gas],
        }

        row_dict = {k: (v[0] if isinstance(v, list) else v) for k, v in data_baru.items()}
        ok_sb, msg_sb = simpan_monitoring_supabase(row_dict)
        if not ok_sb:
            st.error("Gagal simpan ke Supabase: " + msg_sb)
            return

        # Cadangan Excel lokal
        try:
            df_baru = pd.DataFrame(data_baru)
            if os.path.exists(FILE_DB):
                df_lama = pd.read_excel(FILE_DB)
                for col in ["Kesimpulan", "Path_Foto", "Path_Foto2"]:
                    if col not in df_lama.columns:
                        df_lama[col] = ""
                df_total = pd.concat([df_lama, df_baru], ignore_index=True)
            else:
                df_total = df_baru
            df_total.to_excel(FILE_DB, index=False)
        except Exception:
            pass

        if key_dobel in st.session_state:
            del st.session_state[key_dobel]

        st.success("✅ Data berhasil disimpan! (" + msg_sb + ") Status: **" + status_gas + "**")
        st.balloons()