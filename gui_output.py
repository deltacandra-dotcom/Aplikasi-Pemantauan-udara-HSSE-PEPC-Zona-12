import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from calendar import monthrange
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak, KeepTogether
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT

FILE_DB = "database_monitoring_jtb.xlsx"
FILE_ALAT = "database_inspeksi_alat.xlsx"
FILE_TITIK = "master_titik_pantau.xlsx"
FILE_MASTER_ALAT = "master_daftar_alat.xlsx"
FILE_USER = "master_users.xlsx"
FILE_PENGATURAN = "pengaturan_laporan.xlsx"
FOLDER_PETA = "peta_monitoring"

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

def load_pengaturan_laporan():
    default = {
        "Nama_Diperiksa": "Ian T.M",
        "Jabatan_Diperiksa": "Environment Officer",
        "Nama_Disetujui": "Ahmad Afifuddin",
        "Jabatan_Disetujui": "Superintendent Field HSSE",
        "Kesimpulan": "Berdasarkan hasil pemantauan udara sesaat dengan menggunakan Multiple Gas Detector dan SO2 Detector di beberapa titik, tidak terdeteksi adanya Gas H2S dan SO2 di seluruh lokasi pemantauan."
    }
    if os.path.exists(FILE_PENGATURAN):
        try:
            df = pd.read_excel(FILE_PENGATURAN)
            hasil = {
                "Nama_Diperiksa": str(df.loc[0, "Nama_Diperiksa"]) if "Nama_Diperiksa" in df.columns else default["Nama_Diperiksa"],
                "Jabatan_Diperiksa": str(df.loc[0, "Jabatan_Diperiksa"]) if "Jabatan_Diperiksa" in df.columns else default["Jabatan_Diperiksa"],
                "Nama_Disetujui": str(df.loc[0, "Nama_Disetujui"]) if "Nama_Disetujui" in df.columns else default["Nama_Disetujui"],
                "Jabatan_Disetujui": str(df.loc[0, "Jabatan_Disetujui"]) if "Jabatan_Disetujui" in df.columns else default["Jabatan_Disetujui"],
                "Kesimpulan": str(df.loc[0, "Kesimpulan"]) if "Kesimpulan" in df.columns else default["Kesimpulan"]
            }
            return hasil
        except:
            return default
    return default

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
    bulan_map = {
        "Januari": 1, "Februari": 2, "Maret": 3, "April": 4,
        "Mei": 5, "Juni": 6, "Juli": 7, "Agustus": 8,
        "September": 9, "Oktober": 10, "November": 11, "Desember": 12
    }
    
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
            bulan_angka = bulan_map.get(bulan_target, datetime.now().month)
            df_bulan = df_db_all[
                (df_db_all["Tanggal_dt"].dt.month == bulan_angka) &
                (df_db_all["Tanggal_dt"].dt.year == tahun_target)
            ].copy()
        except:
            pass
    
    bulan_angka = bulan_map.get(bulan_target, datetime.now().month)
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

def buat_pdf_laporan(df_db, tgl_target):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.5*cm, leftMargin=1.5*cm, topMargin=1.2*cm, bottomMargin=1.2*cm)
    
    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('TitleCustom', parent=styles['Heading1'], fontSize=13, alignment=TA_CENTER, spaceAfter=2, fontName='Helvetica-Bold', textColor=colors.HexColor('#0A3D91'))
    style_subtitle = ParagraphStyle('Subtitle', parent=styles['Normal'], fontSize=11, alignment=TA_CENTER, spaceAfter=6, fontName='Helvetica-Bold')
    style_heading = ParagraphStyle('HeadingCustom', parent=styles['Heading2'], fontSize=10, spaceBefore=8, spaceAfter=3, fontName='Helvetica-Bold')
    style_normal = ParagraphStyle('NormalCustom', parent=styles['Normal'], fontSize=9, leading=11, alignment=TA_JUSTIFY)
    style_small = ParagraphStyle('Small', parent=styles['Normal'], fontSize=8, leading=10, alignment=TA_CENTER)
    style_center = ParagraphStyle('Center', parent=styles['Normal'], fontSize=9, alignment=TA_CENTER)
    style_caption = ParagraphStyle('Caption', parent=styles['Normal'], fontSize=8, leading=10, alignment=TA_CENTER, fontName='Helvetica')
    
    story = []
    titik_data = load_titik()
    
    # ========== HEADER LOGO (tidak miring) ==========
    logo_skk = "logo_skkmigas.png" if os.path.exists("logo_skkmigas.png") else None
    logo_pepc = "logo_pertamina.png" if os.path.exists("logo_pertamina.png") else None
    
    if logo_skk or logo_pepc:
        row_logo = []
        if logo_skk:
            img_skk = Image(logo_skk, width=2.6*cm, height=1.0*cm)
            img_skk.hAlign = 'LEFT'
            row_logo.append(img_skk)
        else:
            row_logo.append("")
        row_logo.append("")
        if logo_pepc:
            img_pepc = Image(logo_pepc, width=3.0*cm, height=1.0*cm)
            img_pepc.hAlign = 'RIGHT'
            row_logo.append(img_pepc)
        else:
            row_logo.append("")
        
        t_header = Table([row_logo], colWidths=[6*cm, 5*cm, 6*cm])
        t_header.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),
            ('ALIGN', (2, 0), (2, 0), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        story.append(t_header)
    
    story.append(Spacer(1, 4))
    story.append(Paragraph("LAPORAN PEMANTAUAN UDARA SESAAT", style_title))
    story.append(Paragraph("PADA KEGIATAN PRODUKSI FASE OPERASI", style_subtitle))
    story.append(Spacer(1, 3))
    
    intro = "Berikut disampaikan laporan kegiatan pemantauan sesaat dengan menggunakan <i>Multiple Gas Detector</i> dan SO<sub>2</sub> detector di Area GPF dan sekitar GPF pada saat kegiatan Produksi Fase Operasi."
    story.append(Paragraph(intro, style_normal))
    story.append(Spacer(1, 4))
    
    # 1. Pelaksanaan
    story.append(Paragraph("1. Pelaksanaan :", style_heading))
    
    df_hari = df_db[df_db["Tanggal"] == tgl_target.strftime("%Y-%m-%d")].copy() if not df_db.empty else pd.DataFrame()
    
    petugas_ds = "-"
    petugas_ns = "-"
    jam_ds_awal, jam_ds_akhir = "-", "-"
    jam_ns_awal, jam_ns_akhir = "-", "-"
    
    if not df_hari.empty:
        ds = df_hari[df_hari["Shift"] == "Day Shift"]
        ns = df_hari[df_hari["Shift"] == "Night Shift"]
        if not ds.empty:
            petugas_ds = str(ds.iloc[0]["Petugas"])
            waktu_ds = sorted(ds["Waktu"].astype(str).tolist())
            if waktu_ds:
                jam_ds_awal = waktu_ds[0]
                jam_ds_akhir = waktu_ds[-1]
        if not ns.empty:
            petugas_ns = str(ns.iloc[0]["Petugas"])
            waktu_ns = sorted(ns["Waktu"].astype(str).tolist())
            if waktu_ns:
                jam_ns_awal = waktu_ns[0]
                jam_ns_akhir = waktu_ns[-1]
    
    hari_list = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    bulan_list = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    tgl_str = f"{hari_list[tgl_target.weekday()]}, {tgl_target.day} {bulan_list[tgl_target.month]} {tgl_target.year}"
    
    pel_data = [
        ["Hari, Tanggal", ":", tgl_str],
        ["Jam", ":", f"Day Shift {jam_ds_awal} WIB s/d {jam_ds_akhir} WIB"],
        ["", "", f"Night Shift {jam_ns_awal} WIB s/d {jam_ns_akhir} WIB"],
        ["Tempat", ":", "Luar Area GPF"],
        ["Dilaporkan oleh", ":", ""],
        ["Petugas Sampling", ":", "Environment (DS) & ERCM (NS)"],
        ["Siang dilakukan pengukuran oleh", ":", petugas_ds],
        ["Malam dilakukan pengukuran oleh", ":", petugas_ns],
    ]
    
    t_pel = Table(pel_data, colWidths=[5.5*cm, 0.4*cm, 11*cm])
    t_pel.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    story.append(t_pel)
    
    # 2. Tujuan
    story.append(Paragraph("2. Tujuan :", style_heading))
    story.append(Paragraph("a. Memastikan dispersi dari kegiatan flaring JTB GPF aman.", style_normal))
    story.append(Paragraph("b. Sebagai salah satu data rona lingkungan pada Tahap Operasi berlangsung.", style_normal))
    
    # 3. Titik Pengukuran
    story.append(Paragraph("3. Titik Pengukuran", style_heading))
    story.append(Paragraph("Titik monitoring dilakukan di luar area GPF, penentuan titik pemantauan berdasarkan 4 arah mata angin serta kawasan permukiman terdekat:", style_normal))
    story.append(Spacer(1, 3))
    
    header_titik = ["No.", "Kode", "Lokasi", "Latitude", "Longitude", "Latitude NS", "Longitude NS"]
    data_titik = [header_titik]
    for i, (kode, info) in enumerate(titik_data.items(), 1):
        data_titik.append([str(i), kode, info["Lokasi"], info["Lat"], info["Long"], info["LatNS"], info["LongNS"]])
    
    t_titik = Table(data_titik, colWidths=[1*cm, 1.5*cm, 4*cm, 2.5*cm, 2.5*cm, 2.5*cm, 2.5*cm])
    t_titik.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0288D1')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(t_titik)
    
    # 4. Hasil Pengukuran
    story.append(Paragraph("4. Hasil Pengukuran", style_heading))
    story.append(Paragraph("Hasil monitoring pada titik pantau menggunakan <i>Multiple Gas Detector</i> adalah sebagai berikut:", style_normal))
    
    def buat_tabel_hasil(df_shift, judul):
        elems = [Paragraph(f"<b>{judul}</b>", style_small)]
        header = ["No.", "Kode", "Lokasi", "Waktu", "H₂S (ppm)", "O₂ (%)", "CO (ppm)", "LEL (%)", "SO₂ (ppm)"]
        data = [header]
        for i, (kode, info) in enumerate(titik_data.items(), 1):
            match = df_shift[df_shift["Kode_Titik"] == kode] if not df_shift.empty else pd.DataFrame()
            if len(match) > 0:
                m = match.iloc[0]
                row = [str(i), kode, info["Lokasi"], str(m.get("Waktu", "")),
                       f"{float(m.get('H2S', 0)):.1f}", f"{float(m.get('O2', 0)):.1f}",
                       f"{float(m.get('CO', 0)):.1f}", f"{float(m.get('LEL', 0)):.1f}",
                       f"{float(m.get('SO2', 0)):.1f}"]
            else:
                row = [str(i), kode, info["Lokasi"], "-", "0", "0", "0", "0", "0"]
            data.append(row)
        
        t = Table(data, colWidths=[0.8*cm, 1.3*cm, 3.5*cm, 1.8*cm, 1.8*cm, 1.5*cm, 1.6*cm, 1.5*cm, 1.7*cm])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0288D1')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        elems.append(t)
        return elems
    
    df_day = df_hari[df_hari["Shift"] == "Day Shift"] if not df_hari.empty else pd.DataFrame()
    df_night = df_hari[df_hari["Shift"] == "Night Shift"] if not df_hari.empty else pd.DataFrame()
    
    story.append(Spacer(1, 3))
    story.extend(buat_tabel_hasil(df_day, "Day Shift"))
    story.append(Spacer(1, 5))
    story.extend(buat_tabel_hasil(df_night, "Night Shift"))
    
    # ========== 5. DOKUMENTASI DAY SHIFT (dengan tabel border) ==========
    story.append(PageBreak())
    story.append(Paragraph("5. Dokumentasi, Pemantauan Day Shift", style_heading))
    
    for kode, info in titik_data.items():
        match = df_day[df_day["Kode_Titik"] == kode] if not df_day.empty else pd.DataFrame()
        waktu = str(match.iloc[0]["Waktu"]) if len(match) > 0 else "..."
        caption = f"Pemantauan di {kode} ({info['Lokasi']}) pukul {waktu} WIB"
        
        path_foto = None
        if len(match) > 0 and "Path_Foto" in match.columns:
            path_foto = match.iloc[0].get("Path_Foto", "")
        
        if not path_foto or not os.path.exists(str(path_foto)):
            folder = f"foto_monitoring/{tgl_target.strftime('%Y-%m-%d')}/Day_Shift"
            for ext in ["jpg", "jpeg", "png"]:
                candidate = f"{folder}/{kode.replace(' ', '_')}.{ext}"
                if os.path.exists(candidate):
                    path_foto = candidate
                    break
        
        # Buat konten dalam tabel berborder
        if path_foto and os.path.exists(str(path_foto)):
            try:
                img = Image(str(path_foto), width=7.5*cm, height=5.5*cm)
                img.hAlign = 'CENTER'
                content = [[Paragraph(caption, style_caption)], [img]]
            except:
                content = [[Paragraph(caption, style_caption)], [Paragraph("[Foto tidak dapat dimuat]", style_center)]]
        else:
            content = [[Paragraph(caption, style_caption)], [Paragraph("[Belum ada foto]", style_center)]]
        
        t_foto = Table(content, colWidths=[16*cm])
        t_foto.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.8, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F5F5F5')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t_foto)
        story.append(Spacer(1, 6))
    
    # ========== 6. DOKUMENTASI NIGHT SHIFT + PETA ==========
    story.append(PageBreak())
    
    path_peta1 = os.path.join(FOLDER_PETA, "peta_1.jpg")
    if os.path.exists(path_peta1):
        try:
            img_peta = Image(path_peta1, width=16*cm, height=9*cm)
            img_peta.hAlign = 'CENTER'
            story.append(img_peta)
            story.append(Spacer(1, 6))
        except:
            pass
    
    story.append(Paragraph("6. Dokumentasi, pemantauan Night Shift", style_heading))
    
    for kode, info in titik_data.items():
        match = df_night[df_night["Kode_Titik"] == kode] if not df_night.empty else pd.DataFrame()
        waktu = str(match.iloc[0]["Waktu"]) if len(match) > 0 else "..."
        caption = f"Pemantauan di {kode} ({info['Lokasi']}) pukul {waktu} WIB"
        
        path_foto = None
        if len(match) > 0 and "Path_Foto" in match.columns:
            path_foto = match.iloc[0].get("Path_Foto", "")
        
        if not path_foto or not os.path.exists(str(path_foto)):
            folder = f"foto_monitoring/{tgl_target.strftime('%Y-%m-%d')}/Night_Shift"
            for ext in ["jpg", "jpeg", "png"]:
                candidate = f"{folder}/{kode.replace(' ', '_')}.{ext}"
                if os.path.exists(candidate):
                    path_foto = candidate
                    break
        
        if path_foto and os.path.exists(str(path_foto)):
            try:
                img = Image(str(path_foto), width=7.5*cm, height=5.5*cm)
                img.hAlign = 'CENTER'
                content = [[Paragraph(caption, style_caption)], [img]]
            except:
                content = [[Paragraph(caption, style_caption)], [Paragraph("[Foto tidak dapat dimuat]", style_center)]]
        else:
            content = [[Paragraph(caption, style_caption)], [Paragraph("[Belum ada foto]", style_center)]]
        
        t_foto = Table(content, colWidths=[16*cm])
        t_foto.setStyle(TableStyle([
            ('BOX', (0, 0), (-1, -1), 0.8, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F5F5F5')),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t_foto)
        story.append(Spacer(1, 6))
    
    path_peta2 = os.path.join(FOLDER_PETA, "peta_2.jpg")
    if os.path.exists(path_peta2):
        try:
            img_peta2 = Image(path_peta2, width=16*cm, height=9*cm)
            img_peta2.hAlign = 'CENTER'
            story.append(img_peta2)
        except:
            pass
    
    # ========== 7. KESIMPULAN + TANDA TANGAN ==========
    story.append(PageBreak())
    story.append(Paragraph("7. Kesimpulan :", style_heading))
    
    peng = load_pengaturan_laporan()
    story.append(Paragraph(peng["Kesimpulan"], style_normal))
    story.append(Spacer(1, 10))
    
    story.append(Paragraph(f"Bojonegoro, {tgl_target.day} {bulan_list[tgl_target.month]} {tgl_target.year}", style_center))
    story.append(Spacer(1, 14))
    
    # Tabel tanda tangan 4 kolom (lebih mirip contoh)
    ttd_data = [
        ["Dilaporkan oleh", "", "Diperiksa oleh", "Disetujui Oleh"],
        ["Day Shift", "Night Shift", "", ""],
        ["", "", "", ""],
        ["", "", "", ""],
        ["", "", "", ""],
        [petugas_ds, petugas_ns, peng["Nama_Diperiksa"], peng["Nama_Disetujui"]],
        ["", "", peng["Jabatan_Diperiksa"], peng["Jabatan_Disetujui"]],
    ]
    
    t_ttd = Table(ttd_data, colWidths=[4.2*cm, 4.2*cm, 4.2*cm, 4.2*cm])
    t_ttd.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, 1), (1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.6, colors.black),
        ('SPAN', (0, 0), (1, 0)),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_ttd)
    
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

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
        st.markdown("### 3. Panel Cetak & Download")
        
        if not os.path.exists(FILE_DB):
            st.info("Database belum terbentuk.")
            return
        
        df_db = pd.read_excel(FILE_DB)
        
        tab1, tab2, tab3 = st.tabs(["Excel Harian", "Excel Bulanan", "PDF Laporan Resmi"])
        
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
                     "Juli","Agustus","September","Oktober","November","Desember"],
                    index=datetime.now().month-1)
            with c2:
                thn = st.number_input("Tahun", 2024, 2030, datetime.now().year)
            
            if st.button("Generate Excel Bulanan", type="primary"):
                data = buat_excel_bulanan(df_db, bln, thn)
                st.download_button("Download", data=data,
                    file_name=f"Laporan_Bulanan_{bln}_{thn}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        with tab3:
            st.caption("Generate PDF Laporan Resmi (format mirip contoh perusahaan)")
            tgl_pdf = st.date_input("Pilih Tanggal untuk PDF", datetime.now(), key="tgl_pdf")
            st.info("Pastikan sudah upload foto di Input Data, upload peta, dan mengisi nama pejabat + kesimpulan di Kelola User.")
            
            if st.button("Generate PDF Laporan", type="primary"):
                with st.spinner("Sedang membuat PDF..."):
                    try:
                        pdf_data = buat_pdf_laporan(df_db, tgl_pdf)
                        st.download_button(
                            "Download PDF",
                            data=pdf_data,
                            file_name=f"Laporan_Pemantauan_Udara_{tgl_pdf.strftime('%Y%m%d')}.pdf",
                            mime="application/pdf"
                        )
                        st.success("PDF berhasil dibuat!")
                    except Exception as e:
                        st.error(f"Gagal membuat PDF: {e}")
    
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
        
        tab1, tab2, tab3 = st.tabs(["Kelola Titik Pantau", "Kelola Daftar Alat", "Upload Peta Laporan"])
        
        with tab1:
            st.markdown("#### Daftar Titik Pantau Saat Ini")
            
            if os.path.exists(FILE_TITIK):
                df_titik = pd.read_excel(FILE_TITIK)
            else:
                data = []
                for kode, v in DEFAULT_TITIK.items():
                    data.append({
                        "Kode": kode, "Lokasi": v["Lokasi"],
                        "Latitude": v["Lat"], "Longitude": v["Long"],
                        "Latitude_NS": v["LatNS"], "Longitude_NS": v["LongNS"]
                    })
                df_titik = pd.DataFrame(data)
                df_titik.to_excel(FILE_TITIK, index=False)
            
            st.dataframe(df_titik, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.markdown("#### Tambah / Update / Hapus Titik")
            
            mode_titik = st.radio("Pilih Aksi", ["Tambah Titik Baru", "Update Titik", "Hapus Titik"], horizontal=True, key="mode_titik")
            
            with st.form("form_titik_baru"):
                if mode_titik == "Tambah Titik Baru":
                    kode = st.text_input("Kode Titik (contoh: TP 7) *")
                    lokasi = st.text_input("Nama Lokasi *")
                    lat = st.text_input("Latitude *")
                    long = st.text_input("Longitude *")
                    lat_ns = st.text_input("Latitude NS")
                    long_ns = st.text_input("Longitude NS")
                elif mode_titik == "Update Titik":
                    list_kode = df_titik["Kode"].tolist() if not df_titik.empty else []
                    kode = st.selectbox("Pilih Kode Titik yang ingin diubah *", list_kode)
                    if kode and not df_titik.empty:
                        row_lama = df_titik[df_titik["Kode"] == kode].iloc[0]
                        lokasi = st.text_input("Nama Lokasi *", value=str(row_lama["Lokasi"]))
                        lat = st.text_input("Latitude *", value=str(row_lama["Latitude"]))
                        long = st.text_input("Longitude *", value=str(row_lama["Longitude"]))
                        lat_ns = st.text_input("Latitude NS", value=str(row_lama["Latitude_NS"]))
                        long_ns = st.text_input("Longitude NS", value=str(row_lama["Longitude_NS"]))
                    else:
                        lokasi = st.text_input("Nama Lokasi *")
                        lat = st.text_input("Latitude *")
                        long = st.text_input("Longitude *")
                        lat_ns = st.text_input("Latitude NS")
                        long_ns = st.text_input("Longitude NS")
                else:
                    list_kode = df_titik["Kode"].tolist() if not df_titik.empty else []
                    kode = st.selectbox("Pilih Kode Titik yang ingin dihapus *", list_kode)
                    lokasi = lat = long = lat_ns = long_ns = ""
                    st.warning(f"Titik **{kode}** akan dihapus permanen.")
                
                submitted = st.form_submit_button("Simpan Perubahan", type="primary", use_container_width=True)
                
                if submitted:
                    if mode_titik == "Tambah Titik Baru":
                        if not kode.strip() or not lokasi.strip() or not lat.strip() or not long.strip():
                            st.error("Kode, Lokasi, Latitude, dan Longitude wajib diisi!")
                        elif kode.strip() in df_titik["Kode"].values:
                            st.error(f"Kode **{kode}** sudah ada.")
                        else:
                            baru = pd.DataFrame([{"Kode": kode.strip(), "Lokasi": lokasi.strip(), "Latitude": lat.strip(), "Longitude": long.strip(),
                                                  "Latitude_NS": lat_ns.strip() if lat_ns.strip() else lat.strip(),
                                                  "Longitude_NS": long_ns.strip() if long_ns.strip() else long.strip()}])
                            df_titik = pd.concat([df_titik, baru], ignore_index=True)
                            df_titik.to_excel(FILE_TITIK, index=False)
                            st.success(f"Titik **{kode}** berhasil ditambahkan!")
                            st.rerun()
                    elif mode_titik == "Update Titik":
                        if not kode or not lokasi.strip() or not lat.strip() or not long.strip():
                            st.error("Lokasi, Latitude, dan Longitude wajib diisi!")
                        else:
                            df_titik.loc[df_titik["Kode"] == kode, ["Lokasi", "Latitude", "Longitude", "Latitude_NS", "Longitude_NS"]] = [
                                lokasi.strip(), lat.strip(), long.strip(),
                                lat_ns.strip() if lat_ns.strip() else lat.strip(),
                                long_ns.strip() if long_ns.strip() else long.strip()
                            ]
                            df_titik.to_excel(FILE_TITIK, index=False)
                            st.success(f"Titik **{kode}** berhasil diperbarui!")
                            st.rerun()
                    elif mode_titik == "Hapus Titik":
                        if kode:
                            df_titik = df_titik[df_titik["Kode"] != kode].reset_index(drop=True)
                            df_titik.to_excel(FILE_TITIK, index=False)
                            st.success(f"Titik **{kode}** berhasil dihapus!")
                            st.rerun()
        
        with tab2:
            st.markdown("#### Daftar Alat Saat Ini")
            if os.path.exists(FILE_MASTER_ALAT):
                df_master = pd.read_excel(FILE_MASTER_ALAT)
            else:
                df_master = pd.DataFrame({"Nama_Alat": DEFAULT_ALAT})
                df_master.to_excel(FILE_MASTER_ALAT, index=False)
            
            st.dataframe(df_master, use_container_width=True, hide_index=True)
            st.markdown("---")
            st.markdown("#### Tambah / Update / Hapus Alat")
            
            mode_alat = st.radio("Pilih Aksi", ["Tambah Alat Baru", "Update Nama Alat", "Hapus Alat"], horizontal=True, key="mode_alat")
            
            with st.form("form_alat_baru"):
                if mode_alat == "Tambah Alat Baru":
                    nama_alat = st.text_input("Nama Alat Baru *")
                    nama_lama = None
                elif mode_alat == "Update Nama Alat":
                    list_alat = df_master["Nama_Alat"].tolist() if not df_master.empty else []
                    nama_lama = st.selectbox("Pilih Alat yang ingin diubah *", list_alat)
                    nama_alat = st.text_input("Nama Baru *", value=nama_lama if nama_lama else "")
                else:
                    list_alat = df_master["Nama_Alat"].tolist() if not df_master.empty else []
                    nama_alat = st.selectbox("Pilih Alat yang ingin dihapus *", list_alat)
                    nama_lama = None
                    st.warning(f"Alat **{nama_alat}** akan dihapus permanen.")
                
                submitted_alat = st.form_submit_button("Simpan Perubahan", type="primary", use_container_width=True)
                
                if submitted_alat:
                    if mode_alat == "Tambah Alat Baru":
                        if not nama_alat.strip():
                            st.error("Nama Alat wajib diisi!")
                        elif nama_alat.strip() in df_master["Nama_Alat"].values:
                            st.error(f"Alat **{nama_alat}** sudah ada.")
                        else:
                            df_master = pd.concat([df_master, pd.DataFrame([{"Nama_Alat": nama_alat.strip()}])], ignore_index=True)
                            df_master.to_excel(FILE_MASTER_ALAT, index=False)
                            st.success(f"Alat **{nama_alat}** berhasil ditambahkan!")
                            st.rerun()
                    elif mode_alat == "Update Nama Alat":
                        if not nama_alat.strip() or not nama_lama:
                            st.error("Nama baru wajib diisi!")
                        else:
                            df_master.loc[df_master["Nama_Alat"] == nama_lama, "Nama_Alat"] = nama_alat.strip()
                            df_master.to_excel(FILE_MASTER_ALAT, index=False)
                            st.success(f"Nama alat berhasil diubah!")
                            st.rerun()
                    elif mode_alat == "Hapus Alat":
                        if nama_alat:
                            df_master = df_master[df_master["Nama_Alat"] != nama_alat].reset_index(drop=True)
                            df_master.to_excel(FILE_MASTER_ALAT, index=False)
                            st.success(f"Alat **{nama_alat}** berhasil dihapus!")
                            st.rerun()
        
        with tab3:
            st.markdown("#### Upload Peta Monitoring untuk Laporan PDF")
            st.caption("Upload 2 peta yang akan otomatis masuk ke laporan PDF.")
            
            os.makedirs(FOLDER_PETA, exist_ok=True)
            
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                st.markdown("**Peta 1**")
                peta1 = st.file_uploader("Upload Peta 1", type=["jpg", "jpeg", "png"], key="peta1")
                if peta1:
                    with open(os.path.join(FOLDER_PETA, "peta_1.jpg"), "wb") as f:
                        f.write(peta1.getbuffer())
                    st.image(peta1, width=250)
                    st.success("Peta 1 tersimpan")
                elif os.path.exists(os.path.join(FOLDER_PETA, "peta_1.jpg")):
                    st.image(os.path.join(FOLDER_PETA, "peta_1.jpg"), width=250)
                    st.caption("Peta 1 sudah ada")
            
            with col_p2:
                st.markdown("**Peta 2**")
                peta2 = st.file_uploader("Upload Peta 2", type=["jpg", "jpeg", "png"], key="peta2")
                if peta2:
                    with open(os.path.join(FOLDER_PETA, "peta_2.jpg"), "wb") as f:
                        f.write(peta2.getbuffer())
                    st.image(peta2, width=250)
                    st.success("Peta 2 tersimpan")
                elif os.path.exists(os.path.join(FOLDER_PETA, "peta_2.jpg")):
                    st.image(os.path.join(FOLDER_PETA, "peta_2.jpg"), width=250)
                    st.caption("Peta 2 sudah ada")
    
    elif pilihan_menu == "6. Kelola User":
        if username != "admin":
            st.warning("Anda tidak memiliki akses ke menu ini.")
            return
        
        st.markdown("### 6. Kelola User & Pengaturan Laporan")
        
        tab_user, tab_pejabat = st.tabs(["Kelola User", "Nama Pejabat & Kesimpulan"])
        
        with tab_user:
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
        
        with tab_pejabat:
            st.markdown("#### Pengaturan Nama Pejabat & Kesimpulan Laporan PDF")
            st.caption("Nama Day Shift & Night Shift diambil otomatis dari data input.")
            
            peng = load_pengaturan_laporan()
            
            with st.form("form_pejabat"):
                col1, col2 = st.columns(2)
                with col1:
                    nama_diperiksa = st.text_input("Nama Diperiksa oleh", value=peng["Nama_Diperiksa"])
                    jabatan_diperiksa = st.text_input("Jabatan Diperiksa oleh", value=peng["Jabatan_Diperiksa"])
                with col2:
                    nama_disetujui = st.text_input("Nama Disetujui oleh", value=peng["Nama_Disetujui"])
                    jabatan_disetujui = st.text_input("Jabatan Disetujui oleh", value=peng["Jabatan_Disetujui"])
                
                st.markdown("---")
                kesimpulan = st.text_area("Teks Kesimpulan Laporan", value=peng["Kesimpulan"], height=100)
                
                if st.form_submit_button("Simpan Pengaturan", type="primary"):
                    df_simpan = pd.DataFrame([{
                        "Nama_Diperiksa": nama_diperiksa,
                        "Jabatan_Diperiksa": jabatan_diperiksa,
                        "Nama_Disetujui": nama_disetujui,
                        "Jabatan_Disetujui": jabatan_disetujui,
                        "Kesimpulan": kesimpulan
                    }])
                    df_simpan.to_excel(FILE_PENGATURAN, index=False)
                    st.success("Pengaturan berhasil disimpan!")
                    st.rerun()