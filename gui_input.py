import streamlit as st
import pandas as pd
import os
from datetime import datetime
from pathlib import Path

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"
FOLDER_FOTO = "foto_monitoring"   # Folder penyimpanan foto

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

def simpan_foto(uploaded_file, tanggal, shift, kode_titik):
    """Simpan foto ke folder terstruktur: foto_monitoring/YYYY-MM-DD/Shift/Kode.jpg"""
    if uploaded_file is None:
        return None
    
    # Buat folder
    folder_path = Path(FOLDER_FOTO) / tanggal / shift.replace(" ", "_")
    folder_path.mkdir(parents=True, exist_ok=True)
    
    # Nama file
    ekstensi = uploaded_file.name.split(".")[-1].lower()
    if ekstensi not in ["jpg", "jpeg", "png"]:
        ekstensi = "jpg"
    
    nama_file = f"{kode_titik.replace(' ', '_')}.{ekstensi}"
    path_lengkap = folder_path / nama_file
    
    # Simpan
    with open(path_lengkap, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    return str(path_lengkap)

def tampilkan_halaman_input(nama_user=""):
    st.markdown("### 1. Form Registrasi Data Udara Sesaat")
    st.caption("Silahkan isi data di bawah ini. Field bertanda * wajib diisi. Noise dan Keterangan bersifat opsional.")

    titik_data = load_titik()
    list_kode = list(titik_data.keys())

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
        
        jam_asli = datetime.now().strftime("%H:%M")
        waktu_in = st.text_input(
            "Waktu Sampling *",
            value=jam_asli,
            help="Format: JJ:MM (contoh 10:35). Otomatis terisi jam sekarang, bisa diubah."
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
        
        h2s = st.number_input(
            "Gas H₂S (ppm) *",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.2f",
            placeholder="Contoh: 0.00",
            help="Wajib diisi. Nilai 0 jika tidak terdeteksi."
        )
        
        o2 = st.number_input(
            "Gas O₂ (%) *",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.1f",
            placeholder="19.5 – 23.5",
            help="Rentang normal: 19.5 – 23.5 %. Wajib diisi."
        )
        
        co = st.number_input(
            "Gas CO (ppm) *",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.1f",
            placeholder="Contoh: 0.0",
            help="Wajib diisi."
        )
        
        lel = st.number_input(
            "Gas LEL (%) *",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.1f",
            placeholder="Contoh: 0.0",
            help="Wajib diisi."
        )
        
        so2 = st.number_input(
            "Gas SO₂ (ppm) *",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.2f",
            placeholder="Contoh: 0.00",
            help="Wajib diisi. Nilai 0 jika tidak terdeteksi."
        )
        
        noise = st.number_input(
            "Noise Level (dB(A))",
            min_value=0.0,
            value=None,
            step=0.1,
            format="%.1f",
            placeholder="Opsional",
            help="Tidak wajib diisi."
        )

    # ==================== UPLOAD FOTO ====================
    st.markdown("---")
    st.markdown("#### Dokumentasi Foto")
    st.caption("Upload 1 foto alat di lokasi titik pantau. Foto ini wajib jika ingin generate PDF Laporan.")
    
    foto_upload = st.file_uploader(
        "Upload Foto Titik Pantau * (untuk laporan PDF)",
        type=["jpg", "jpeg", "png"],
        help="Foto deteksi alat di lokasi. Akan otomatis masuk ke laporan PDF sesuai tanggal & shift."
    )
    
    if foto_upload:
        st.image(foto_upload, caption=f"Preview: {tp_in} - {shift_in}", width=300)

    # Field Keterangan (opsional)
    st.markdown("---")
    keterangan_in = st.text_area(
        "Keterangan (Opsional)",
        placeholder="Tuliskan catatan tambahan jika diperlukan (contoh: kondisi cuaca, gangguan, dll). Field ini hanya muncul di Dashboard & Review, tidak masuk Excel.",
        height=80,
        help="Tidak wajib diisi. Hanya untuk catatan internal."
    )

    st.markdown("---")

    if st.button("Simpan Data", type="primary", use_container_width=True):
        kesalahan = []
        
        if not petugas_in.strip():
            kesalahan.append("Nama Petugas")
        if not waktu_in.strip():
            kesalahan.append("Waktu Sampling")
        if h2s is None:
            kesalahan.append("H₂S")
        if o2 is None:
            kesalahan.append("O₂")
        elif o2 < 19.5 or o2 > 23.5:
            kesalahan.append("O₂ (harus antara 19.5 – 23.5)")
        if co is None:
            kesalahan.append("CO")
        if lel is None:
            kesalahan.append("LEL")
        if so2 is None:
            kesalahan.append("SO₂")
        # Noise dan Keterangan TIDAK divalidasi (opsional)
        # Foto juga tidak wajib di tahap simpan data (hanya wajib saat generate PDF)
        
        if kesalahan:
            st.error(f"Masih ada data yang belum valid: **{', '.join(kesalahan)}**")
            return
        
        # Simpan foto jika ada
        path_foto = None
        if foto_upload is not None:
            path_foto = simpan_foto(
                foto_upload,
                tanggal_in.strftime("%Y-%m-%d"),
                shift_in,
                tp_in
            )
        
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
            "Noise": [noise if noise is not None else None],
            "Keterangan": [keterangan_in.strip() if keterangan_in else ""],
            "Path_Foto": [path_foto if path_foto else ""],
            "Status": [status_gas]
        }
        
        df_baru = pd.DataFrame(data_baru)
        
        if os.path.exists(FILE_DB):
            df_lama = pd.read_excel(FILE_DB)
            # Pastikan kolom baru ada
            if "Keterangan" not in df_lama.columns:
                df_lama["Keterangan"] = ""
            if "Path_Foto" not in df_lama.columns:
                df_lama["Path_Foto"] = ""
            df_total = pd.concat([df_lama, df_baru], ignore_index=True)
        else:
            df_total = df_baru
            
        df_total.to_excel(FILE_DB, index=False)
        
        if path_foto:
            st.success(f"Data berhasil disimpan! Foto tersimpan di: `{path_foto}`")
        else:
            st.success("Data berhasil disimpan! (Tanpa foto)")
        st.balloons()