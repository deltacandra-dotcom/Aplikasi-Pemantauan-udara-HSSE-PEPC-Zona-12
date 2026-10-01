import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from calendar import monthrange

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_ALAT = "database_inspeksi_alat.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"
FILE_MASTER_ALAT = "master_daftar_alat.xlsx"
FILE_USER = "master_users.xlsx"

DEFAULT_TITIK = {
    "TP 1": {"Lokasi": "Pemukiman Ngasem", "Lat": "-7.1234", "Long": "111.6789", "LatNS": "-7.1234", "LongNS": "111.6789"},
    "TP 2": {"Lokasi": "Pemukiman Bandungrejo", "Lat": "-7.1245", "Long": "111.6791", "LatNS": "-7.1245", "LongNS": "111.6791"},
    "TP 3": {"Lokasi": "Petak 28", "Lat": "-7.1256", "Long": "111.6801", "LatNS": "-7.1256", "LongNS": "111.6801"},
    "TP 4": {"Lokasi": "Crossing Gayam", "Lat": "-7.1267", "Long": "111.6812", "LatNS": "-7.1267", "LongNS": "111.6812"},
    "TP 5": {"Lokasi": "Selatan Dormitory", "Lat": "-7.1278", "Long": "111.6823", "LatNS": "-7.1278", "LongNS": "111.6823"},
    "TP 6": {"Lokasi": "Dusun Besaran Ngasem", "Lat": "-7.1289", "Long": "111.6834", "LatNS": "-7.1289", "LongNS": "111.6834"}
}

DEFAULT_ALAT = ["Gas Detector 1", "Gas Detector 2", "SO2 Detector 1", "SO2 Detector 2"]

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

def load_daftar_alat():
    if os.path.exists(FILE_MASTER_ALAT):
        df = pd.read_excel(FILE_MASTER_ALAT)
        return df["Nama_Alat"].tolist()
    else:
        pd.DataFrame({"Nama_Alat": DEFAULT_ALAT}).to_excel(FILE_MASTER_ALAT, index=False)
        return DEFAULT_ALAT

def buat_excel_harian_resmi(df_db_all, tgl_target):
    titik_data = load_titik()
    wb = Workbook()
    ws = wb.active
    ws.title = "Laporan Harian"
    
    font_title = Font(name='Arial', size=12, bold=True)
    font_header = Font(name='Arial', size=10, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="0288D1", end_color="0288D1", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center')
    
    ws.merge_cells('A1:I1')
    ws.cell(row=1, column=1, value="LAPORAN PEMANTAUAN UDARA SESAAT - GPF JTB").font = font_title
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='center')
    
    ws.merge_cells('A2:I2')
    ws.cell(row=2, column=1, value=f"Tanggal: {tgl_target.strftime('%d %B %Y')} | Pertamina EP Cepu Zona 12").font = Font(name='Arial', size=10, italic=True)
    ws.cell(row=2, column=1).alignment = Alignment(horizontal='center')
    
    ws.cell(row=4, column=1, value="1. Daftar Titik Pantau dan Koordinat").font = font_title
    headers_koord = ["No.", "Kode", "Lokasi", "Latitude", "Longitude", "Latitude NS", "Longitude NS"]
    for col, text in enumerate(headers_koord, 1):
        cell = ws.cell(row=5, column=col, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center
        cell.border = thin_border
    
    for i, (kode, info) in enumerate(titik_data.items(), 1):
        row = 5 + i
        data = [i, kode, info["Lokasi"], info["Lat"], info["Long"], info["LatNS"], info["LongNS"]]
        for col, val in enumerate(data, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = thin_border
            cell.alignment = left if col == 3 else center
    
    df_f = pd.DataFrame()
    if not df_db_all.empty:
        try:
            df_f = df_db_all[df_db_all["Tanggal"] == tgl_target.strftime("%Y-%m-%d")].reset_index(drop=True)
        except:
            pass
    
    headers_ukur = ["No.", "Kode", "Lokasi", "Waktu", "H2S (ppm)", "O2 (%)", "CO (ppm)", "LEL (%)", "SO2 (ppm)"]
    
    start_day = 14
    ws.cell(row=start_day-1, column=1, value="2. Hasil Pengukuran - Day Shift").font = font_title
    for col, text in enumerate(headers_ukur, 1):
        cell = ws.cell(row=start_day, column=col, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center
        cell.border = thin_border
    
    df_day = df_f[df_f["Shift"] == "Day Shift"] if not df_f.empty else pd.DataFrame()
    for i, (kode, info) in enumerate(titik_data.items(), 1):
        row = start_day + i
        match = df_day[df_day["Kode_Titik"] == kode] if not df_day.empty else pd.DataFrame()
        if len(match) > 0:
            m = match.iloc[0]
            vals = [i, kode, info["Lokasi"], str(m.get("Waktu","")),
                    f"{float(m.get('H2S',0)):.2f}", f"{float(m.get('O2',0)):.1f}",
                    f"{float(m.get('CO',0)):.1f}", f"{float(m.get('LEL',0)):.1f}",
                    f"{float(m.get('SO2',0)):.2f}"]
        else:
            vals = [i, kode, info["Lokasi"], "-", "-", "-", "-", "-", "-"]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = thin_border
            cell.alignment = left if col == 3 else center
    
    start_night = 23
    ws.cell(row=start_night-1, column=1, value="3. Hasil Pengukuran - Night Shift").font = font_title
    for col, text in enumerate(headers_ukur, 1):
        cell = ws.cell(row=start_night, column=col, value=text)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center
        cell.border = thin_border
    
    df_night = df_f[df_f["Shift"] == "Night Shift"] if not df_f.empty else pd.DataFrame()
    for i, (kode, info) in enumerate(titik_data.items(), 1):
        row = start_night + i
        match = df_night[df_night["Kode_Titik"] == kode] if not df_night.empty else pd.DataFrame()
        if len(match) > 0:
            m = match.iloc[0]
            vals = [i, kode, info["Lokasi"], str(m.get("Waktu","")),
                    f"{float(m.get('H2S',0)):.2f}", f"{float(m.get('O2',0)):.1f}",
                    f"{float(m.get('CO',0)):.1f}", f"{float(m.get('LEL',0)):.1f}",
                    f"{float(m.get('SO2',0)):.2f}"]
        else:
            vals = [i, kode, info["Lokasi"], "-", "-", "-", "-", "-", "-"]
        for col, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=col, value=val)
            cell.border = thin_border
            cell.alignment = left if col == 3 else center
    
    ws.cell(row=32, column=1, value="Catatan: Berdasarkan hasil pemantauan udara sesaat menggunakan Multiple Gas Detector dan SO2 Detector.").font = Font(name='Arial', size=9, italic=True)
    
    widths = [6, 10, 22, 12, 12, 10, 12, 10, 12]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

def buat_excel_bulanan(df_db_all, bulan_target, tahun_target):
    titik_data = load_titik()
    wb = Workbook()
    ws = wb.active
    ws.title = f"{bulan_target[:3]}-{str(tahun_target)[-2:]}"
    
    font_title = Font(name='Arial', size=14, bold=True)
    font_header = Font(name='Arial', size=8, bold=True, color="FFFFFF")
    font_sub = Font(name='Arial', size=7, bold=True)
    fill_header = PatternFill(start_color="0288D1", end_color="0288D1", fill_type="solid")
    fill_ds = PatternFill(start_color="BBDEFB", end_color="BBDEFB", fill_type="solid")
    fill_ns = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    ws.merge_cells('A1:Z1')
    ws.cell(row=1, column=1, value="Pemantauan Udara Sesaat").font = font_title
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='center')
    
    col = 2
    list_tp = list(titik_data.items())
    
    ws.cell(row=3, column=1, value=bulan_target).font = Font(name='Arial', size=10, bold=True)
    
    for kode, info in list_tp:
        start_col = col
        end_col = col + 13
        ws.merge_cells(start_row=3, start_column=start_col, end_row=3, end_column=end_col)
        cell = ws.cell(row=3, column=start_col, value=kode)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = center
        for c in range(start_col, end_col+1):
            ws.cell(row=3, column=c).border = thin_border
            ws.cell(row=3, column=c).fill = fill_header
        col = end_col + 1
    
    col = 2
    for kode, info in list_tp:
        start_col = col
        end_col = col + 13
        ws.merge_cells(start_row=4, start_column=start_col, end_row=4, end_column=end_col)
        cell = ws.cell(row=4, column=start_col, value=info["Lokasi"])
        cell.font = Font(name='Arial', size=8, bold=True)
        cell.alignment = center
        for c in range(start_col, end_col+1):
            ws.cell(row=4, column=c).border = thin_border
        col = end_col + 1
    
    col = 2
    for _ in list_tp:
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col+6)
        cell = ws.cell(row=5, column=col, value="Ds/")
        cell.font = font_sub
        cell.fill = fill_ds
        cell.alignment = center
        for c in range(col, col+7):
            ws.cell(row=5, column=c).border = thin_border
            ws.cell(row=5, column=c).fill = fill_ds
        
        ws.merge_cells(start_row=5, start_column=col+7, end_row=5, end_column=col+13)
        cell = ws.cell(row=5, column=col+7, value="Ns/")
        cell.font = font_sub
        cell.fill = fill_ns
        cell.alignment = center
        for c in range(col+7, col+14):
            ws.cell(row=5, column=c).border = thin_border
            ws.cell(row=5, column=c).fill = fill_ns
        col += 14
    
    params = ["Waktu", "H2S\n(ppm)", "O2\n(%)", "CO\n(ppm)", "LEL\n(%)", "SO2\n(ppm)", "Noise\ndB(A)"]
    col = 2
    for _ in list_tp:
        for p in params:
            cell = ws.cell(row=6, column=col, value=p)
            cell.font = Font(name='Arial', size=7, bold=True)
            cell.fill = fill_ds
            cell.alignment = center
            cell.border = thin_border
            col += 1
        for p in params:
            cell = ws.cell(row=6, column=col, value=p)
            cell.font = Font(name='Arial', size=7, bold=True)
            cell.fill = fill_ns
            cell.alignment = center
            cell.border = thin_border
            col += 1
    
    df_bulan = pd.DataFrame()
    if not df_db_all.empty:
        try:
            df_db_all["Tanggal_dt"] = pd.to_datetime(df_db_all["Tanggal"])
            df_bulan = df_db_all[
                (df_db_all["Tanggal_dt"].dt.month_name() == bulan_target) &
                (df_db_all["Tanggal_dt"].dt.year == tahun_target)
            ].copy()
        except:
            pass
    
    bulan_angka = datetime.strptime(bulan_target, "%B").month
    jumlah_hari = monthrange(tahun_target, bulan_angka)[1]
    
    for hari in range(1, jumlah_hari + 1):
        row = 6 + hari
        ws.cell(row=row, column=1, value=hari).border = thin_border
        ws.cell(row=row, column=1).alignment = center
        
        tgl_str = f"{tahun_target}-{bulan_angka:02d}-{hari:02d}"
        
        col = 2
        for kode, info in list_tp:
            for shift, fill in [("Day Shift", fill_ds), ("Night Shift", fill_ns)]:
                match = pd.DataFrame()
                if not df_bulan.empty:
                    match = df_bulan[(df_bulan["Tanggal"] == tgl_str) & 
                                    (df_bulan["Kode_Titik"] == kode) & 
                                    (df_bulan["Shift"] == shift)]
                
                if len(match) > 0:
                    m = match.iloc[0]
                    vals = [
                        str(m.get("Waktu", "")),
                        f"{float(m.get('H2S', 0)):.1f}",
                        f"{float(m.get('O2', 0)):.1f}",
                        f"{float(m.get('CO', 0)):.1f}",
                        f"{float(m.get('LEL', 0)):.1f}",
                        f"{float(m.get('SO2', 0)):.1f}",
                        f"{float(m.get('Noise', 0)):.1f}" if pd.notna(m.get("Noise")) else ""
                    ]
                else:
                    vals = ["", "0", "0", "0", "0", "0", ""]
                
                for v in vals:
                    cell = ws.cell(row=row, column=col, value=v)
                    cell.border = thin_border
                    cell.alignment = center
                    cell.fill = fill
                    col += 1
    
    note_row = 6 + jumlah_hari + 2
    ws.cell(row=note_row, column=1, value="Note : Berdasarkan hasil pemantauan udara sesaat dengan menggunakan Multiple Gas Detector dan SO2 Detector di beberapa titik.").font = Font(name='Arial', size=8, italic=True)
    
    ws.column_dimensions['A'].width = 5
    for i in range(2, 2 + len(list_tp)*14):
        ws.column_dimensions[get_column_letter(i)].width = 6
    
    bio = BytesIO()
    wb.save(bio)
    return bio.getvalue()

def tampilkan_halaman_output(pilihan_menu, nama_user="", username=""):
    titik_data = load_titik()
    daftar_alat = load_daftar_alat()
    
    if pilihan_menu == "2. Review Data":
        st.markdown("### 2. Portal Pencarian & Review Data")
        
        if not os.path.exists(FILE_DB):
            st.info("Database belum terbentuk.")
            return
        
        df_db = pd.read_excel(FILE_DB)
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            cari_tgl = st.date_input("Tanggal", datetime.now())
        with c2:
            cari_shift = st.selectbox("Shift", ["Semua", "Day Shift", "Night Shift"])
        with c3:
            cari_tp = st.selectbox("Titik", ["Semua"] + list(titik_data.keys()))
        with c4:
            cari_status = st.selectbox("Status", ["Semua", "Aman", "Terpapar"])
        
        df_hasil = df_db[df_db["Tanggal"] == cari_tgl.strftime("%Y-%m-%d")].copy()
        if cari_shift != "Semua":
            df_hasil = df_hasil[df_hasil["Shift"] == cari_shift]
        if cari_tp != "Semua":
            df_hasil = df_hasil[df_hasil["Kode_Titik"] == cari_tp]
        if cari_status != "Semua":
            df_hasil = df_hasil[df_hasil["Status"] == cari_status]
        
        if not df_hasil.empty:
            if "Terpapar" in df_hasil["Status"].values:
                st.error("TERDETEKSI PAPARAN GAS")
            else:
                st.success("SEMUA TITIK AMAN")
            
            def style_row(row):
                return ['background-color:#FFEBEE' if row.Status=='Terpapar' else 'background-color:#E8F5E9' for _ in row]
            
            st.dataframe(df_hasil.style.apply(style_row, axis=1), use_container_width=True)
            
            st.markdown("---")
            idx = st.selectbox("Pilih indeks untuk dihapus:", df_hasil.index)
            if st.button("Hapus Permanen"):
                df_baru = df_db.drop(idx).reset_index(drop=True)
                df_baru.to_excel(FILE_DB, index=False)
                st.success("Data dihapus")
                st.rerun()
        else:
            st.warning("Tidak ada data.")
    
    elif pilihan_menu == "3. Download Laporan":
        st.markdown("### 3. Panel Cetak & Download Excel")
        
        if not os.path.exists(FILE_DB):
            st.info("Database belum terbentuk.")
            return
        
        df_db = pd.read_excel(FILE_DB)
        
        tab1, tab2 = st.tabs(["Excel Harian", "Excel Bulanan"])
        
        with tab1:
            st.caption("Silahkan Unduh Excel Harian di Bawah ini:")
            tgl = st.date_input("Pilih Tanggal", datetime.now(), key="tgl_harian")
            if st.button("Generate Excel Harian", type="primary"):
                data = buat_excel_harian_resmi(df_db, tgl)
                st.download_button("Download", data=data,
                    file_name=f"Laporan_Harian_{tgl.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        with tab2:
            st.caption("Silahkan Unduh Excel Bulanan di Bawah ini:")
            c1, c2 = st.columns(2)
            with c1:
                bln = st.selectbox("Bulan", 
                    ["Januari","Februari","Maret","April","Mei","Juni",
                     "Juli","Augustus","September","Oktober","November","Desember"],
                    index=datetime.now().month-1)
            with c2:
                thn = st.number_input("Tahun", 2024, 2030, datetime.now().year)
            
            if st.button("Generate Excel Bulanan", type="primary"):
                data = buat_excel_bulanan(df_db, bln, thn)
                st.download_button("Download", data=data,
                    file_name=f"Laporan_Bulanan_{bln}_{thn}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    elif pilihan_menu == "4. Inspeksi Alat":
        st.markdown("### 4. Inspeksi & Status Kesiapan Alat")
        
        if os.path.exists(FILE_ALAT):
            df_alat = pd.read_excel(FILE_ALAT)
        else:
            df_alat = pd.DataFrame(columns=["Nama_Alat","Tanggal_Inspeksi","Tanggal_Berikutnya","Sudah_Kalibrasi","Status","Petugas","Keterangan"])
        
        hari_ini = datetime.now().date()
        perlu_update = False
        if not df_alat.empty:
            for idx, row in df_alat.iterrows():
                try:
                    tgl_next = pd.to_datetime(row["Tanggal_Berikutnya"]).date()
                    if tgl_next < hari_ini and row["Status"] == "Siap Pakai":
                        df_alat.at[idx, "Status"] = "Belum diinspeksi"
                        perlu_update = True
                except:
                    pass
            if perlu_update:
                df_alat.to_excel(FILE_ALAT, index=False)
        
        st.markdown("#### Status Terkini")
        cols = st.columns(min(4, len(daftar_alat)) or 1)
        for i, nama in enumerate(daftar_alat):
            with cols[i % 4]:
                last = df_alat[df_alat["Nama_Alat"] == nama].tail(1)
                if len(last) > 0:
                    status = last.iloc[0]["Status"]
                    tgl = last.iloc[0]["Tanggal_Inspeksi"]
                    tgl_next = last.iloc[0].get("Tanggal_Berikutnya", "-")
                    if status == "Siap Pakai":
                        st.success(f"**{nama}**\n\nSiap Pakai\n\n{tgl} → {tgl_next}")
                    elif status == "Tidak Siap Pakai":
                        st.error(f"**{nama}**\n\nTidak Siap\n\n{tgl}")
                    else:
                        st.warning(f"**{nama}**\n\nBelum diinspeksi")
                else:
                    st.warning(f"**{nama}**\n\nBelum diinspeksi")
        
        st.markdown("---")
        with st.form("form_inspeksi"):
            alat = st.selectbox("Pilih Alat", daftar_alat)
            tgl_insp = st.date_input("Tanggal Inspeksi", datetime.now())
            interval = st.number_input("Interval (hari)", 1, 365, 30)
            kalibrasi = st.radio("Sudah dikalibrasi?", ["Ya, sudah", "Belum"])
            petugas = st.text_input("Petugas", value=nama_user)
            ket = st.text_area("Keterangan")
            
            if st.form_submit_button("Simpan", type="primary"):
                if not petugas.strip():
                    st.error("Petugas wajib diisi")
                else:
                    status = "Siap Pakai" if kalibrasi == "Ya, sudah" else "Tidak Siap Pakai"
                    tgl_next = (tgl_insp + timedelta(days=interval)).strftime("%Y-%m-%d")
                    baru = pd.DataFrame([{
                        "Nama_Alat": alat,
                        "Tanggal_Inspeksi": tgl_insp.strftime("%Y-%m-%d"),
                        "Tanggal_Berikutnya": tgl_next,
                        "Sudah_Kalibrasi": kalibrasi,
                        "Status": status,
                        "Petugas": petugas.strip(),
                        "Keterangan": ket
                    }])
                    df_alat = pd.concat([df_alat, baru], ignore_index=True)
                    df_alat.to_excel(FILE_ALAT, index=False)
                    st.success(f"Tersimpan. Status: {status}")
                    st.rerun()
        
        if not df_alat.empty:
            st.markdown("#### Riwayat")
            st.dataframe(df_alat.sort_values("Tanggal_Inspeksi", ascending=False), use_container_width=True)
    
    elif pilihan_menu == "5. Kelola Master Data":
        st.markdown("### 5. Kelola Master Data")
        
        tab1, tab2 = st.tabs(["Kelola Titik Pantau", "Kelola Daftar Alat"])
        
        with tab1:
            df_titik = pd.read_excel(FILE_TITIK) if os.path.exists(FILE_TITIK) else pd.DataFrame()
            if not df_titik.empty:
                st.dataframe(df_titik, use_container_width=True)
            
            st.markdown("---")
            with st.form("form_titik"):
                kode = st.text_input("Kode Titik (contoh: TP 7)")
                lokasi = st.text_input("Nama Lokasi")
                lat = st.text_input("Latitude")
                long = st.text_input("Longitude")
                lat_ns = st.text_input("Latitude NS")
                long_ns = st.text_input("Longitude NS")
                
                c1, c2 = st.columns(2)
                with c1:
                    simpan = st.form_submit_button("Simpan Titik", type="primary")
                with c2:
                    hapus = st.form_submit_button("Hapus Titik")
                
                if simpan:
                    if not kode.strip() or not lokasi.strip():
                        st.error("Kode dan Lokasi wajib")
                    else:
                        df = pd.read_excel(FILE_TITIK) if os.path.exists(FILE_TITIK) else pd.DataFrame(columns=["Kode","Lokasi","Latitude","Longitude","Latitude_NS","Longitude_NS"])
                        if kode.strip() in df["Kode"].values:
                            df.loc[df["Kode"]==kode.strip(), ["Lokasi","Latitude","Longitude","Latitude_NS","Longitude_NS"]] = [lokasi, lat, long, lat_ns, long_ns]
                            st.success("Diperbarui")
                        else:
                            baru = pd.DataFrame([{"Kode":kode.strip(),"Lokasi":lokasi.strip(),"Latitude":lat,"Longitude":long,"Latitude_NS":lat_ns,"Longitude_NS":long_ns}])
                            df = pd.concat([df, baru], ignore_index=True)
                            st.success("Ditambahkan")
                        df.to_excel(FILE_TITIK, index=False)
                        st.rerun()
                
                if hapus:
                    if kode.strip():
                        df = pd.read_excel(FILE_TITIK)
                        if kode.strip() in df["Kode"].values:
                            df = df[df["Kode"] != kode.strip()]
                            df.to_excel(FILE_TITIK, index=False)
                            st.success("Dihapus")
                            st.rerun()
        
        with tab2:
            df_master = pd.read_excel(FILE_MASTER_ALAT) if os.path.exists(FILE_MASTER_ALAT) else pd.DataFrame({"Nama_Alat": DEFAULT_ALAT})
            st.dataframe(df_master, use_container_width=True)
            
            st.markdown("---")
            with st.form("form_alat"):
                nama_alat = st.text_input("Nama Alat")
                c1, c2 = st.columns(2)
                with c1:
                    tambah = st.form_submit_button("Tambah", type="primary")
                with c2:
                    hapus_a = st.form_submit_button("Hapus")
                
                if tambah and nama_alat.strip():
                    df = pd.read_excel(FILE_MASTER_ALAT) if os.path.exists(FILE_MASTER_ALAT) else pd.DataFrame(columns=["Nama_Alat"])
                    if nama_alat.strip() not in df["Nama_Alat"].values:
                        df = pd.concat([df, pd.DataFrame([{"Nama_Alat": nama_alat.strip()}])], ignore_index=True)
                        df.to_excel(FILE_MASTER_ALAT, index=False)
                        st.success("Ditambahkan")
                        st.rerun()
                
                if hapus_a and nama_alat.strip():
                    df = pd.read_excel(FILE_MASTER_ALAT)
                    if nama_alat.strip() in df["Nama_Alat"].values:
                        df = df[df["Nama_Alat"] != nama_alat.strip()]
                        df.to_excel(FILE_MASTER_ALAT, index=False)
                        st.success("Dihapus")
                        st.rerun()
    
    elif pilihan_menu == "6. Kelola User":
        if username != "admin":
            st.warning("Anda tidak memiliki akses ke menu ini.")
            return
        
        st.markdown("### 6. Kelola User & Password")
        
        df_user = pd.read_excel(FILE_USER)
        st.dataframe(df_user, use_container_width=True)
        
        st.markdown("---")
        with st.form("form_user"):
            username_in = st.text_input("Username")
            password = st.text_input("Password")
            nama = st.text_input("Nama Tampilan")
            
            c1, c2 = st.columns(2)
            with c1:
                simpan = st.form_submit_button("Tambah / Update User", type="primary")
            with c2:
                hapus = st.form_submit_button("Hapus User")
            
            if simpan:
                if not username_in.strip() or not password.strip() or not nama.strip():
                    st.error("Semua field wajib diisi")
                else:
                    if username_in.strip() in df_user["Username"].values:
                        df_user.loc[df_user["Username"]==username_in.strip(), ["Password","Nama"]] = [password, nama]
                        st.success("User diperbarui")
                    else:
                        baru = pd.DataFrame([{"Username": username_in.strip(), "Password": password, "Nama": nama}])
                        df_user = pd.concat([df_user, baru], ignore_index=True)
                        st.success("User ditambahkan")
                    df_user.to_excel(FILE_USER, index=False)
                    st.rerun()
            
            if hapus:
                if username_in.strip() in df_user["Username"].values:
                    if username_in.strip() == "admin":
                        st.error("User admin tidak boleh dihapus")
                    else:
                        df_user = df_user[df_user["Username"] != username_in.strip()]
                        df_user.to_excel(FILE_USER, index=False)
                        st.success("User dihapus")
                        st.rerun()
                else:
                    st.warning("Username tidak ditemukan")