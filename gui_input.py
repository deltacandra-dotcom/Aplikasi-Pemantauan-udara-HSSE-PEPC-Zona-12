import streamlit as st
import pandas as pd
import os
from datetime import datetime
from pathlib import Path
import pytz

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
    "TP 6": {"Lokasi": "Dusun Besaran Ngasem", "Lat": "-7.1289", "Long": "111.6834", "LatNS": "-7.1290", "LongNS": "111.6836"}
}

DEFAULT_KESIMPULAN = [
    "Tidak terdeteksi adanya Gas H2S dan SO2 di seluruh lokasi pemantauan",
    "Terdeteksi Gas H2S di bawah baku mutu",
    "Terdeteksi Gas SO2, perlu investigasi lanjutan",
    "Terdeteksi Gas H2S dan SO2, segera tindak lanjut"
]

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

def load_kesimpulan():
    if os.path.exists(FILE_KESIMPULAN):
        df = pd.read_excel(FILE_KESIMPULAN)
        return df["Kesimpulan"].tolist()
    else:
        pd.DataFrame({"Kesimpulan": DEFAULT_KESIMPULAN}).to_excel(FILE_KESIMPULAN, index=False)
        return DEFAULT_KESIMPULAN

def simpan_foto(uploaded_file, tanggal, shift, kode_titik, nomor=1):
    if uploaded_file is None:
        return None
    
    folder_path = Path(FOLDER_FOTO) / tanggal / shift.replace(" ", "_")
    folder_path.mkdir(parents=True, exist_ok=True)
    
    ekstensi = uploaded_file.name.split(".")[-1].lower()
    if ekstensi not in ["jpg", "jpeg", "png"]:
        ekstensi = "jpg"
    
    nama_file = f"{kode_titik.replace(' ', '_')}_{nomor}.{ekstensi}"
    path_lengkap = folder_path / nama_file
    
    with open(path_lengkap, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return str(path_lengkap)

def tampilkan_halaman_input(nama_user=""):
    st.markdown("### 1. Form Registrasi Data Udara Sesaat")
    st.caption("Silahkan isi data di bawah ini. Field bertanda * wajib diisi. Noise bersifat opsional.")

    titik_data = load_titik()
    list_kode = list(titik_data.keys())
    list_kesimpulan = load_kesimpulan()

    col_kiri, _, col_kanan = st.columns([2, 0.2, 2])

    with col_kiri:
        st.markdown("#### Metadata Sesi")
        
        tanggal_in = st.date_input(
            "Tanggal Pengukuran *",
            datetime.now(),
            help="Tanggal pelaksanaan pemantauan"
        )
        
        bulan_in = tanggal_in.strftime("%B")
        
        shift_in = st.selectbox(
            "Shift Kerja *",
            ["Day Shift", "Night Shift"],
            help="Pilih shift sesuai waktu pengukuran"
        )
        
        tz_wib = pytz.timezone("Asia/Jakarta")
        jam_sekarang = datetime.now(tz_wib).strftime("%H:%M")
        
        waktu_in = st.text_input(
            "Waktu Sampling *",
            value=jam_sekarang,
            help="Format: JJ:MM (contoh 10:35). Otomatis terisi jam WIB sekarang."
        )
        
        petugas_in = st.text_input(
            "Nama Petugas *",
            value=nama_user,
            help="Nama petugas yang melakukan pengukuran"
        )
        
        tp_in = st.selectbox(
            "Kode Titik Pantau *",
            list_kode,
            help="Pilih titik pantau sesuai lokasi pengukuran"
        )
        
        info = titik_data[tp_in]
        st.info(f"**Area:** {info['Lokasi']}\n\n**Koordinat:** Lat {info['Lat']} | Long {info['Long']}")

    with col_kanan:
        st.markdown("#### Hasil Deteksi Alat")
        
        h2s = st.number_input("Gas H2S (ppm) *", min_value=0.0, value=None, step=0.1, format="%.2f", placeholder="Contoh: 0.00", help="Wajib diisi. Nilai 0 jika tidak terdeteksi.")
        o2 = st.number_input("Gas O2 (%) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="19.5 – 23.5", help="Rentang normal: 19.5 – 23.5 %. Tetap bisa disimpan jika di luar range (akan diberi peringatan).")
        co = st.number_input("Gas CO (ppm) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Contoh: 0.0", help="Wajib diisi.")
        lel = st.number_input("Gas LEL (%) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Contoh: 0.0", help="Wajib diisi.")
        so2 = st.number_input("Gas SO2 (ppm) *", min_value=0.0, value=None, step=0.1, format="%.2f", placeholder="Contoh: 0.00", help="Wajib diisi. Nilai 0 jika tidak terdeteksi.")
        noise = st.number_input("Noise Level / Kebisingan (dB(A))", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Opsional", help="Tidak wajib diisi.")

    st.markdown("---")
    st.markdown("#### Dokumentasi Foto")
    st.caption("Upload 1–2 foto alat di lokasi titik pantau. Foto 1 wajib. Foto 2 opsional. Kedua foto akan muncul berdampingan di laporan PDF.")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        foto1 = st.file_uploader("Foto 1 * (Wajib)", type=["jpg", "jpeg", "png"], key="foto1")
        if foto1:
            st.image(foto1, caption="Preview Foto 1", width=250)
    with col_f2:
        foto2 = st.file_uploader("Foto 2 (Opsional)", type=["jpg", "jpeg", "png"], key="foto2")
        if foto2:
            st.image(foto2, caption="Preview Foto 2", width=250)

    st.markdown("---")
    
    st.markdown("#### Kesimpulan")
    kesimpulan_in = st.selectbox(
        "Pilih Kesimpulan *",
        options=list_kesimpulan,
        help="Wajib dipilih. Satu hari hanya 1 kesimpulan yang dipakai (yang terakhir disimpan)."
    )

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
        if co is None:
            kesalahan.append("CO")
        if lel is None:
            kesalahan.append("LEL")
        if so2 is None:
            kesalahan.append("SO2")
        if not kesimpulan_in:
            kesalahan.append("Kesimpulan")
        if foto1 is None:
            kesalahan.append("Foto 1 (wajib)")
        
        if kesalahan:
            st.error(f"Masih ada data yang belum valid: **{', '.join(kesalahan)}**")
            return
        
        o2_tidak_normal = False
        if o2 < 19.5 or o2 > 23.5:
            o2_tidak_normal = True
            st.warning(f"⚠️ **Peringatan Oksigen:** Nilai O2 ({o2}%) berada di luar rentang normal (19.5 – 23.5). Data tetap disimpan.")
        
        tgl_str = tanggal_in.strftime("%Y-%m-%d")
        path_foto1 = simpan_foto(foto1, tgl_str, shift_in, tp_in, nomor=1)
        path_foto2 = simpan_foto(foto2, tgl_str, shift_in, tp_in, nomor=2) if foto2 is not None else ""
        
        if h2s > 0.0 or so2 > 0.0:
            status_gas = "Terpapar"
        elif o2_tidak_normal:
            status_gas = "O2 Tidak Normal"
        else:
            status_gas = "Aman"
        
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
            "Path_Foto": [path_foto1 if path_foto1 else ""],
            "Path_Foto2": [path_foto2 if path_foto2 else ""],
            "Status": [status_gas]
        }
        
        df_baru = pd.DataFrame(data_baru)
        
        if os.path.exists(FILE_DB):
            df_lama = pd.read_excel(FILE_DB)
            if "Kesimpulan" not in df_lama.columns:
                df_lama["Kesimpulan"] = ""
            if "Path_Foto" not in df_lama.columns:
                df_lama["Path_Foto"] = ""
            if "Path_Foto2" not in df_lama.columns:
                df_lama["Path_Foto2"] = ""
            df_total = pd.concat([df_lama, df_baru], ignore_index=True)
        else:
            df_total = df_baru
            
        df_total.to_excel(FILE_DB, index=False)
        
        st.success("✅ Data berhasil disimpan!")
        st.balloons()