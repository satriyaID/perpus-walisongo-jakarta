import os
import pandas as pd
import urllib.parse
import urllib.request
from io import BytesIO

# Import dari ReportLab dan Pillow
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont

def gabung_kartu_proporsional_a3plus(excel_path, template_png_path, font_custom_path, output_pdf="Cetak_Massal_Kartu_A3Plus.pdf"):
    print("Membaca data siswa dari Excel...")
    # dtype=str memastikan angka 0 di depan NISN tidak hilang
    df = pd.read_excel(excel_path, dtype=str)
    df = df.dropna(subset=['NISN', 'Nama'])
    
    # Konversi kertas A3+ ke satuan poin ReportLab (1 mm = 2.83465 points)
    mm = 2.83465
    paper_width = 320 * mm
    paper_height = 480 * mm
    
    # Konfigurasi ukuran kartu fisik saat dicetak (Standard ID Card: 85mm x 54mm)
    card_w = 85 * mm
    card_h = 54 * mm
    
    # Pengaturan Jarak/Grid di Kertas A3+
    gap_x = 8 * mm          # Jarak antar kartu horizontal
    gap_y = 4 * mm          # Jarak antar kartu vertikal
    margin_left = 22 * mm   # Margin kiri agar seimbang di tengah
    margin_top = 18 * mm    # Margin atas
    
    # Setup canvas PDF A3+
    c = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
    
    # =====================================================================
    # PENGATURAN FONT PROPORSIONAL (NAMA LEBIH BESAR)
    # =====================================================================
    try:
        # Menggunakan Chau Philomene Ukuran 32 untuk Nama agar tampak menonjol & profesional
        font_nama = ImageFont.truetype(font_custom_path, 32)
        print(f"Sukses memuat font Nama: {font_custom_path} (Ukuran 32)")
        # Font untuk data informasi lainnya (NISN, Jurusan, Rombel) menggunakan Arial Bold
        font_data = ImageFont.truetype("arialbd.ttf", 22)
    except IOError:
        font_nama = ImageFont.load_default()
        font_data = ImageFont.load_default()
        print("Peringatan: Gagal memuat font khusus, menggunakan font standar sistem.")

    col_idx = 0
    row_idx = 0
    total_processed = 0
    
    for index, row in df.iterrows():
        nama_siswa = str(row['Nama']).upper()
        nisn = str(row['NISN']).strip() # Mengambil NISN murni (aman dari leading zero hilang)
        jurusan = str(row['Jurusan']).strip()
        rombel = str(row['Rombel']).strip()
        
        print(f"Memproses [{total_processed + 1}]: {nama_siswa} - NISN: {nisn}")
        
        # 1. PROSES KARTU INDIVIDU DI MEMORI (PILLOW)
        card_img = Image.open(template_png_path).convert("RGBA")
        draw = ImageDraw.Draw(card_img)
        
        # Penulisan Teks Dinamis
        # Koordinat Y untuk Nama diturunkan sedikit ke 282 agar huruf besar Chau Philomene tetap simetris
        draw.text((310, 282), nama_siswa, font=font_nama, fill=(0, 0, 0))
        
        # Data administrasi lainnya tetap sejajar rapi
        draw.text((310, 332), nisn, font=font_data, fill=(0, 0, 0))
        draw.text((310, 377), jurusan, font=font_data, fill=(0, 0, 0))
        draw.text((310, 422), rombel, font=font_data, fill=(0, 0, 0))
        
        # Generate & Timpa QR Code berbasis NISN asli
        try:
            encoded_nisn = urllib.parse.quote(nisn)
            qr_url = f"https://quickchart.io/qr?text={encoded_nisn}&size=200&margin=1"
            req = urllib.request.Request(qr_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                qr_data = BytesIO(response.read())
                qr_image = Image.open(qr_data).convert("RGBA").resize((175, 175))
                card_img.alpha_composite(qr_image, (740, 280)) # Menimpa QR lama dengan ukuran pas 175x175
        except Exception as e:
            print(f"   [Gagal QR] {nama_siswa}: {e}")
            
        # Simpan olahan stiker kartu ke buffer memori sementara
        img_buffer = BytesIO()
        card_img.convert("RGB").save(img_buffer, format="JPEG", quality=95)
        img_buffer.seek(0)
        
        # 2. SUSUN KARTU KE GRID KERTAS A3+ (REPORTLAB)
        x = margin_left + (col_idx * (card_w + gap_x))
        y = paper_height - margin_top - card_h - (row_idx * (card_h + gap_y))
        
        # Menggambar gambar stiker kartu ke halaman PDF A3+
        reportlab_img = ImageReader(img_buffer)
        c.drawImage(reportlab_img, x, y, width=card_w, height=card_h)
        
        # Tambahkan garis tipis warna magenta (sebagai mal potong stiker kartu jika digabung)
        c.setStrokeColorRGB(1.0, 0.0, 1.0)
        c.setLineWidth(0.3)
        c.rect(x, y, card_w, card_h, stroke=1, fill=0)
        
        total_processed += 1
        col_idx += 1
        
        # Pengaturan Grid Maksimal 3 Kolom x 8 Baris (Total 24 Kartu per lembar)
        if col_idx >= 3:
            col_idx = 0
            row_idx += 1
            
        if row_idx >= 8:
            c.showPage() # Buka lembar halaman baru kertas A3+
            col_idx = 0
            row_idx = 0

    c.save()
    print(f"\n[SUKSES] {total_processed} Kartu Perpustakaan selesai digabungkan secara proporsional.")
    print(f"File PDF Siap Cetak disimpan di: {output_pdf}")

# =====================================================================
# BAGIAN EKSEKUSI PROGRAM
# =====================================================================
FILE_EXCEL = r"G:\2025-2026\WS\Data_Siswa_Perpus.xlsx"
TEMPLATE_GAMBAR = r"G:\2025-2026\WS\Kartu Perpustakaan SMK Walisongo.png"
FONT_CHAU = r"G:\2025-2026\WS\ChauPhilomeneOne-Regular.ttf"

# Pastikan nama file font .ttf di atas disesuaikan dengan nama file font yang Anda unduh
gabung_kartu_proporsional_a3plus(FILE_EXCEL, TEMPLATE_GAMBAR, FONT_CHAU)