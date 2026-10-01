import streamlit as st
import pandas as pd
import os
from datetime import datetime

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"

DEFAULT_TITIK = {
    "TP 1": {"Lokasi": "Pemukiman Ngasem", "Lat": "-7.1234", "Long": "111.6789", "LatNS": "-7.1235", "LongNS": "111.6790"},
    "TP 2": {"Lokasi": "Pemukiman Bandungrejo", "Lat": "-7.1245", "Long": "111.6791", "LatNS": "-7.1246", "LongNS": "111.6792"},
    "TP 3": {"Lokasi": "Petak 28", "Lat": "-7.1256", "Long": "111.6801", "LatNS": "-7.1257", "LongNS": "111.6803"},
    "TP 4": {"Lokasi": "Crossing Gayam", "Lat": "-7.1267", "Long": "111.6812", "LatNS": "-7.1268", "LongNS": "111.6814"},
    "TP 5": {"Lokasi": "Selatan Dormitory", "Lat": "-7.1278", "Long": "111.6823", "LatNS": "-7.1279", "LongNS": "111.6825"},
    "TP 6": {"Lokasi": "Dusun Besaran Ngasem", "Lat": "-7.1289", "Long": "111.6834", "LatNS": "-7.1290", "LongNS": "111.6836"}
}

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
                "LongNS": str(row["Longitude_NS"])
            }
        return titik
    else:
        data = []
        for kode, v in DEFAULT_TITIK.items():
            data.append({
                "Kode": kode, "Lokasi": v["Lokasi"],
                "Latitude": v["Lat"], "Longitude": v["Long"],
                "Latitude_NS": v["LatNS"], "Longitude_NS": v["LongNS"]
            })
        pd.DataFrame(data).to_excel(FILE_TITIK, index=False)
        return DEFAULT_TITIK

def tampilkan_halaman_input(nama_user=""):
    st.markdown("### 1. Form Registrasi Data Udara Sesaat")
    st.write("Silahkan Isi Data di Bawah ini. Semua kolom wajib diisi!")

    titik_data = load_titik()
    list_kode = list(titik_data.keys())

    col_kiri, _, col_kanan = st.columns([2, 0.2, 2])

    with col_kiri:
        st.markdown("#### Metadata Sesi")
        tanggal_in = st.date_input("Tanggal Pengukuran", datetime.now())
        bulan_in = tanggal_in.strftime("%B")
        shift_in = st.selectbox("Shift Kerja", ["Day Shift", "Night Shift"])
        
        jam_asli = datetime.now().strftime("%H:%M")
        waktu_in = st.text_input("Waktu Sampling", value=jam_asli)
        
        petugas_in = st.text_input("Nama Petugas", value=nama_user)
        
        tp_in = st.selectbox("Kode Titik Pantau", list_kode)
        
        info = titik_data[tp_in]
        st.info(f"**Area:** {info['Lokasi']}\n\n**Koordinat:** Lat {info['Lat']} | Long {info['Long']}")

    with col_kanan:
        st.markdown("#### Hasil Deteksi Alat")
        h2s = st.number_input("Gas H2S (ppm)", min_value=0.0, value=None, step=0.1, format="%.2f", placeholder="Wajib diisi")
        o2 = st.number_input("Gas O2 (%)", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="19.5 – 23.5", help="Rentang normal: 19.5 – 23.5 %")
        co = st.number_input("Gas CO (ppm)", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Wajib diisi")
        lel = st.number_input("Gas LEL (%)", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Wajib diisi")
        so2 = st.number_input("Gas SO2 (ppm)", min_value=0.0, value=None, step=0.1, format="%.2f", placeholder="Wajib diisi")
        noise = st.number_input("Noise Level (dB(A))", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Wajib diisi")

    st.markdown("---")

    if st.button("Simpan Data", type="primary", use_container_width=True):
        kesalahan = []
        
        if not petugas_in.strip():
            kesalahan.append("Nama Petugas")
        if not waktu_in.strip():
            kesalahan.append("Waktu Sampling")
        if h2s is None:
            kesalahan.append("H2S")
        if o2 is None:
            kesalahan.append("O2")
        elif o2 < 19.5 or o2 > 23.5:
            kesalahan.append("O2 (harus antara 19.5 – 23.5)")
        if co is None:
            kesalahan.append("CO")
        if lel is None:
            kesalahan.append("LEL")
        if so2 is None:
            kesalahan.append("SO2")
        if noise is None:
            kesalahan.append("Noise")
        
        if kesalahan:
            st.error(f"Masih ada data yang belum valid: **{', '.join(kesalahan)}**")
            return
        
        status_gas = "Terpapar" if (h2s > 0.0 or so2 > 0.0) else "Aman"
        
        data_baru = {
            "Tanggal": [tanggal_in.strftime("%Y-%m-%d")],
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
            "Noise": [noise],
            "Status": [status_gas]
        }
        
        df_baru = pd.DataFrame(data_baru)
        if os.path.exists(FILE_DB):
            df_lama = pd.read_excel(FILE_DB)
            df_total = pd.concat([df_lama, df_baru], ignore_index=True)
        else:
            df_total = df_baru
            
        df_total.to_excel(FILE_DB, index=False)
        st.success("Data berhasil disimpan!")
        st.balloons()