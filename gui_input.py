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
    else:
        data = []
        for kode, v in DEFAULT_TITIK.items():
            data.append({
                "Kode": kode, "Lokasi": v["Lokasi"],
                "Latitude": v["Lat"], "Longitude": v["Long"],
                "Latitude_NS": v["LatNS"], "Longitude_NS": v["LongNS"],
            })
        try:
            pd.DataFrame(data).to_excel(FILE_TITIK, index=False)
        except Exception:
            pass
        return DEFAULT_TITIK


def load_kesimpulan():
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
            return df["Kesimpulan"].tolist()
    else:
        try:
            pd.DataFrame({"Kesimpulan": DEFAULT_KESIMPULAN}).to_excel(FILE_KESIMPULAN, index=False)
        except Exception:
            pass
    return list(DEFAULT_KESIMPULAN)


def simpan_foto(uploaded_file, tanggal, shift, kode_titik, nomor=1):
    """Upload ke Supabase Storage; fallback lokal jika gagal. Return path storage atau lokal."""
    if uploaded_file is None:
        return None

    ekstensi = uploaded_file.name.split(".")[-1].lower()
    if ekstensi not in ["jpg", "jpeg", "png"]:
        ekstensi = "jpg"

    shift_folder = shift.replace(" ", "_")
    nama_file = f"{kode_titik.replace(' ', '_')}_{nomor}.{ekstensi}"
    storage_path = f"{tanggal}/{shift_folder}/{nama_file}"

    try:
        sb = get_supabase()
        file_bytes = uploaded_file.getvalue()
        sb.storage.from_("foto-monitoring").upload(
            path=storage_path,
            file=file_bytes,
            file_options={
                "content-type": uploaded_file.type or "image/jpeg",
                "upsert": "true",
            },
        )
        return storage_path
    except Exception as e:
        st.warning("Upload Storage gagal, simpan lokal: " + str(e))

    folder_path = Path(FOLDER_FOTO) / tanggal / shift_folder
    folder_path.mkdir(parents=True, exist_ok=True)
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
            help="Tanggal pelaksanaan pemantauan",
        )

        bulan_in = tanggal_in.strftime("%B")

        shift_in = st.selectbox(
            "Shift Kerja *",
            ["Day Shift", "Night Shift"],
            help="Pilih shift sesuai waktu pengukuran",
        )

        tz_wib = pytz.timezone("Asia/Jakarta")
        jam_sekarang = datetime.now(tz_wib).strftime("%H:%M")

        waktu_in = st.text_input(
            "Waktu Sampling *",
            value=jam_sekarang,
            help="Format: JJ:MM (contoh 10:35). Otomatis terisi jam WIB sekarang.",
        )

        petugas_in = st.text_input(
            "Nama Petugas *",
            value=nama_user,
            help="Nama petugas yang melakukan pengukuran",
        )

        tp_in = st.selectbox(
            "Kode Titik Pantau *",
            list_kode,
            help="Pilih titik pantau sesuai lokasi pengukuran",
        )

        info = titik_data[tp_in]
        st.info(f"**Area:** {info['Lokasi']}\n\n**Koordinat:** Lat {info['Lat']} | Long {info['Long']}")

    with col_kanan:
        st.markdown("#### Hasil Deteksi Alat")
        st.caption("Baku mutu: H2S=1 | SO2=0,057 | CO=8,72 | O2=19,5–23,5 | Noise=55")

        h2s = st.number_input("Gas H2S (ppm) *", min_value=0.0, value=None, step=0.1, format="%.2f", placeholder="Contoh: 0.00", help="Wajib diisi. Nilai 0 jika tidak terdeteksi.")
        o2 = st.number_input("Gas O2 (%) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="19.5 – 23.5", help="Rentang normal: 19.5 – 23.5 %. Tetap bisa disimpan jika di luar range.")
        co = st.number_input("Gas CO (ppm) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Contoh: 0.0", help="Wajib diisi.")
        lel = st.number_input("Gas LEL (%) *", min_value=0.0, value=None, step=0.1, format="%.1f", placeholder="Contoh: 0.0", help="Wajib diisi.")
        so2 = st.number_input("Gas SO2 (ppm) *", min_value=0.0, value=None, step=0.001, format="%.3f", placeholder="Contoh: 0.00", help="Wajib diisi. Nilai 0 jika tidak terdeteksi.")
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
    opsi_kes = list(list_kesimpulan) + ["Lainnya (ketik manual)"]
    pilih_kes = st.selectbox(
        "Pilih Kesimpulan *",
        options=opsi_kes,
        help="Wajib dipilih. Satu hari hanya 1 kesimpulan yang dipakai (yang terakhir disimpan). Pilih Lainnya untuk menulis manual.",
    )
    if pilih_kes == "Lainnya (ketik manual)":
        kesimpulan_in = st.text_area(
            "Tulis kesimpulan manual *",
            height=80,
            placeholder="Ketik kesimpulan di sini...",
        )
    else:
        kesimpulan_in = pilih_kes

    st.markdown("---")

    if "konfirmasi_dobel" not in st.session_state:
        st.session_state.konfirmasi_dobel = False

    col_btn1, col_btn2 = st.columns([3, 1])

    with col_btn1:
        label_tombol = "Konfirmasi & Simpan Data" if st.session_state.konfirmasi_dobel else "Simpan Data"
        klik = st.button(label_tombol, type="primary", use_container_width=True)

    with col_btn2:
        if st.session_state.konfirmasi_dobel:
            if st.button("Batal", use_container_width=True):
                st.session_state.konfirmasi_dobel = False
                st.rerun()

    if klik:
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
        if not str(kesimpulan_in or "").strip():
            kesalahan.append("Kesimpulan")
        if foto1 is None:
            kesalahan.append("Foto 1 (wajib)")

        if kesalahan:
            st.error(f"Masih ada data yang belum valid: **{', '.join(kesalahan)}**")
            st.session_state.konfirmasi_dobel = False
            return

        tgl_str = tanggal_in.strftime("%Y-%m-%d")

        sudah_ada = False
        try:
            sb = get_supabase()
            cek = (
                sb.table("monitoring")
                .select("id")
                .eq("tanggal", tgl_str)
                .eq("shift", shift_in)
                .eq("kode_titik", str(tp_in))
                .limit(1)
                .execute()
            )
            sudah_ada = bool(cek.data)
        except Exception:
            if os.path.exists(FILE_DB):
                df_cek = pd.read_excel(FILE_DB)
                if not df_cek.empty and all(c in df_cek.columns for c in ["Tanggal", "Shift", "Kode_Titik"]):
                    match_dobel = df_cek[
                        (df_cek["Tanggal"].astype(str) == tgl_str)
                        & (df_cek["Shift"] == shift_in)
                        & (df_cek["Kode_Titik"].astype(str) == str(tp_in))
                    ]
                    if len(match_dobel) > 0:
                        sudah_ada = True

        if sudah_ada and not st.session_state.konfirmasi_dobel:
            st.session_state.konfirmasi_dobel = True
            st.error(
                f"⚠️ Data untuk **{tp_in}** ({shift_in}, {tgl_str}) sudah pernah diinput. "
                f"Yakin ingin menginputnya kembali?"
            )
            st.warning("Klik tombol **Konfirmasi & Simpan Data** sekali lagi untuk menyimpan, atau klik **Batal**.")
            return

        ada_bahaya = False
        ada_waspada = False
        detail_status = []

        if h2s > 0:
            if h2s > 1.0:
                ada_bahaya = True
                detail_status.append(f"Terdeteksi H2S {h2s} ppm (> baku mutu 1 ppm)")
            else:
                ada_waspada = True
                detail_status.append(f"Terdeteksi H2S {h2s} ppm (≤ baku mutu 1 ppm)")

        if so2 > 0:
            if so2 > 0.057:
                ada_bahaya = True
                detail_status.append(f"Terdeteksi SO2 {so2} ppm (> baku mutu 0,057 ppm)")
            else:
                ada_waspada = True
                detail_status.append(f"Terdeteksi SO2 {so2} ppm (≤ baku mutu 0,057 ppm)")

        if co > 0:
            if co > 8.72:
                ada_bahaya = True
                detail_status.append(f"Terdeteksi CO {co} ppm (> baku mutu 8,72 ppm)")
            else:
                ada_waspada = True
                detail_status.append(f"Terdeteksi CO {co} ppm (≤ baku mutu 8,72 ppm)")

        if o2 < 19.5:
            ada_bahaya = True
            detail_status.append(f"O2 tidak normal: {o2}% (< 19,5%)")
        elif o2 > 23.5:
            ada_bahaya = True
            detail_status.append(f"O2 tidak normal: {o2}% (> 23,5%)")

        if noise is not None and noise > 55:
            ada_waspada = True
            detail_status.append(f"Kebisingan {noise} dB (melebihi baku mutu 55 dB)")

        if ada_bahaya:
            status_gas = "Bahaya"
        elif ada_waspada:
            status_gas = "Waspada"
        else:
            status_gas = "Aman"

        path_foto1 = simpan_foto(foto1, tgl_str, shift_in, tp_in, nomor=1)
        path_foto2 = simpan_foto(foto2, tgl_str, shift_in, tp_in, nomor=2) if foto2 is not None else ""

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
            "Kesimpulan": [str(kesimpulan_in).strip()],
            "Path_Foto": [path_foto1 if path_foto1 else ""],
            "Path_Foto2": [path_foto2 if path_foto2 else ""],
            "Status": [status_gas],
        }

        row_dict = {k: (v[0] if isinstance(v, list) else v) for k, v in data_baru.items()}
        ok_sb, msg_sb = simpan_monitoring_supabase(row_dict)
        if not ok_sb:
            st.error("Gagal simpan ke Supabase: " + msg_sb)
            return

        try:
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
        except Exception:
            pass

        st.session_state.konfirmasi_dobel = False

        if status_gas == "Bahaya":
            st.error("⚠️ STATUS **BAHAYA**: " + " | ".join(detail_status))
        elif status_gas == "Waspada":
            st.warning("⚡ STATUS **WASPADA**: " + " | ".join(detail_status))
        else:
            st.success("✅ STATUS **AMAN** — semua parameter dalam batas normal")

        st.success("✅ Data berhasil disimpan! (" + msg_sb + ")")
        st.balloons()