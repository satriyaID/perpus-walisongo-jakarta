import os
import pandas as pd
import urllib.parse
import urllib.request
from io import BytesIO

# Import ReportLab & Pillow
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont

# =====================================================================
# UTILITIES & KONFIGURASI GLOBAL
# =====================================================================
MM_TO_POINT = 2.83465
FOLDER_KERJA = r"G:\2025-2026\WS"

# Jalur File Default
EXCEL_BUKU = os.path.join(FOLDER_KERJA, "QR_ID_BUKU.xlsx")
EXCEL_SISWA = os.path.join(FOLDER_KERJA, "Data_Siswa_Perpus.xlsx")
TEMPLATE_KARTU = os.path.join(FOLDER_KERJA, "Kartu Perpustakaan SMK Walisongo.png")
FONT_CHAU = os.path.join(FOLDER_KERJA, "ChauPhilomeneOne-Regular.ttf")

# =====================================================================
# MODUL 1: GENERATE & LAYOUT LABEL BUKU (A3+)
# =====================================================================
def modul_cetak_label_buku():
    print("\n" + "="*50)
    print(" [1] PROSES CETAK LABEL BUKU MASSAL (A3+)")
    print("="*50)
    
    output_pdf = os.path.join(FOLDER_KERJA, "Cetak_Label_Buku_A3Plus.pdf")
    output_mal = os.path.join(FOLDER_KERJA, "Mal_Kisscut_Label_A3Plus.pdf")
    
    if not os.path.exists(EXCEL_BUKU):
        print(f"[ERROR] File Excel Buku tidak ditemukan di: {EXCEL_BUKU}")
        return

    print("Membaca data dari Excel...")
    df = pd.read_excel(EXCEL_BUKU, header=None, dtype=str)
    df = df.dropna(subset=[3])  
    
    paper_width = 320 * MM_TO_POINT
    paper_height = 480 * MM_TO_POINT
    
    qr_size = 25 * MM_TO_POINT
    gap_x = 5 * MM_TO_POINT
    gap_y = 4 * MM_TO_POINT
    margin_left = 10 * MM_TO_POINT
    margin_top = 12 * MM_TO_POINT
    label_space = 2.5 * MM_TO_POINT
    
    # Buat Canvas untuk Desain dan Canvas untuk Mal
    c_design = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
    c_mal = canvas.Canvas(output_mal, pagesize=(paper_width, paper_height))
    
    col_idx = 0
    row_idx = 0
    total_processed = 0
    
    for index, row in df.iterrows():
        qr_text = str(row[1]) if pd.notna(row[1]) else str(row[3])
        label_text = str(row[3])
        
        if "label" in label_text.lower() or "qr code" in qr_text.lower():
            continue
            
        x = margin_left + (col_idx * (qr_size + gap_x))
        y = paper_height - margin_top - qr_size - (row_idx * (qr_size + gap_y))
        
        # Cek Kode Warna di Kolom E (Index 4)
        jurusan = str(row[4]).strip().upper() if pd.notna(row[4]) else ""
        if jurusan in ["TKJ", "TEKNO"]:
            bg_color = (0.88, 0.82, 0.98) # Lilac
        elif jurusan in ["MP", "BISMEN"]:
            bg_color = (1.0, 0.8, 0.85)  # Pink
        elif jurusan == "BD":
            bg_color = (0.75, 0.93, 0.75) # Hijau
        elif jurusan == "AKL":
            bg_color = (1.0, 0.96, 0.7)  # Kuning
        else:
            bg_color = (1.0, 1.0, 1.0)   # Putih
            
        # Gambar di File Desain Utama
        c_design.setFillColorRGB(*bg_color)
        c_design.setStrokeColorRGB(0.86, 0.86, 0.86)
        c_design.setLineWidth(0.5)
        c_design.rect(x, y, qr_size, qr_size, stroke=1, fill=1)
        
        # Gambar di File Mal Kisscut (Hanya garis potong Magenta)
        c_mal.setStrokeColorRGB(1.0, 0.0, 1.0)
        c_mal.setLineWidth(0.3)
        c_mal.rect(x, y, qr_size, qr_size, stroke=1, fill=0)
        
        # Ambil & Gambar QR Code
        try:
            encoded_text = urllib.parse.quote(qr_text)
            qr_url = f"https://quickchart.io/qr?text={encoded_text}&size=150&margin=0"
            req = urllib.request.Request(qr_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                img_data = BytesIO(response.read())
                img = ImageReader(img_data)
                c_design.drawImage(img, x, y, width=qr_size, height=qr_size)
        except Exception as e:
            c_design.line(x, y, x + qr_size, y + qr_size)
            
        # Cetak Teks Label Buku
        c_design.setFillColorRGB(0.1, 0.1, 0.1)
        c_design.setFont("Helvetica-Bold", 6.5)
        c_design.drawCentredString(x + (qr_size / 2), y - label_space, label_text)
        
        total_processed += 1
        col_idx += 1
        if col_idx >= 10:
            col_idx = 0
            row_idx += 1
        if row_idx >= 15:
            c_design.showPage()
            c_mal.showPage()
            col_idx = 0
            row_idx = 0

    c_design.save()
    c_mal.save()
    print(f"\n[SUKSES] Selesai memproses {total_processed} stiker label buku.")
    print(f"-> File Desain: {output_pdf}")
    print(f"-> File Mal Kisscut: {output_mal}")

# =====================================================================
# MODUL 2: GENERATE & LAYOUT KARTU PERPUSTAKAAN (A3+)
# =====================================================================
def modul_cetak_kartu_perpus():
    print("\n" + "="*50)
    print(" [2] PROSES CETAK KARTU PERPUS MASSAL (A3+)")
    print("="*50)
    
    output_pdf = os.path.join(FOLDER_KERJA, "Cetak_Massal_Kartu_A3Plus.pdf")
    output_mal = os.path.join(FOLDER_KERJA, "Mal_Diecut_Kartu_A3Plus.pdf")
    
    if not os.path.exists(EXCEL_SISWA) or not os.path.exists(TEMPLATE_KARTU):
        print("[ERROR] File Excel Siswa atau Template Gambar PNG tidak ditemukan!")
        return

    print("Membaca data siswa...")
    df = pd.read_excel(EXCEL_SISWA, dtype=str)
    df = df.dropna(subset=['NISN', 'Nama'])
    
    paper_width = 320 * MM_TO_POINT
    paper_height = 480 * MM_TO_POINT
    card_w = 85 * MM_TO_POINT
    card_h = 54 * MM_TO_POINT
    
    gap_x = 8 * MM_TO_POINT
    gap_y = 4 * MM_TO_POINT
    margin_left = 22 * MM_TO_POINT
    margin_top = 18 * MM_TO_POINT
    
    c_design = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
    c_mal = canvas.Canvas(output_mal, pagesize=(paper_width, paper_height))
    
    try:
        font_nama = ImageFont.truetype(FONT_CHAU, 32)
        font_data = ImageFont.truetype("arialbd.ttf", 22)
    except IOError:
        font_nama = ImageFont.load_default()
        font_data = ImageFont.load_default()
        print("[Peringatan] Gagal memuat font khusus, beralih ke font default.")

    col_idx = 0
    row_idx = 0
    total_processed = 0
    
    for index, row in df.iterrows():
        nama_siswa = str(row['Nama']).upper()
        nisn = str(row['NISN']).strip()
        jurusan = str(row['Jurusan']).strip()
        rombel = str(row['Rombel']).strip()
        
        print(f"Memproses Kartu [{total_processed + 1}]: {nama_siswa}")
        
        # Olah Gambar Kartu Mandiri di Memori via Pillow
        card_img = Image.open(TEMPLATE_KARTU).convert("RGBA")
        draw = ImageDraw.Draw(card_img)
        
        # Gambar Teks Proporsional
        draw.text((310, 282), nama_siswa, font=font_nama, fill=(0, 0, 0))
        draw.text((310, 332), nisn, font=font_data, fill=(0, 0, 0))
        draw.text((310, 377), jurusan, font=font_data, fill=(0, 0, 0))
        draw.text((310, 422), rombel, font=font_data, fill=(0, 0, 0))
        
        # Tempel QR Code NISN
        try:
            encoded_nisn = urllib.parse.quote(nisn)
            qr_url = f"https://quickchart.io/qr?text={encoded_nisn}&size=200&margin=1"
            req = urllib.request.Request(qr_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                qr_data = BytesIO(response.read())
                qr_image = Image.open(qr_data).convert("RGBA").resize((175, 175))
                card_img.alpha_composite(qr_image, (740, 280))
        except Exception as e:
            print(f"   [Gagal QR] {nama_siswa}")
            
        img_buffer = BytesIO()
        card_img.convert("RGB").save(img_buffer, format="JPEG", quality=95)
        img_buffer.seek(0)
        
        # Tempatkan ke Grid Canvas Lembar A3+
        x = margin_left + (col_idx * (card_w + gap_x))
        y = paper_height - margin_top - card_h - (row_idx * (card_h + gap_y))
        
        # Gambarkan ke Dokumen Utama
        reportlab_img = ImageReader(img_buffer)
        c_design.drawImage(reportlab_img, x, y, width=card_w, height=card_h)
        
        # Gambarkan Garis Mal Magenta di Kedua File (Desain dan Mal Terpisah)
        c_design.setStrokeColorRGB(1.0, 0.0, 1.0)
        c_design.setLineWidth(0.3)
        c_design.rect(x, y, card_w, card_h, stroke=1, fill=0)
        
        c_mal.setStrokeColorRGB(1.0, 0.0, 1.0)
        c_mal.setLineWidth(0.3)
        c_mal.rect(x, y, card_w, card_h, stroke=1, fill=0)
        
        total_processed += 1
        col_idx += 1
        if col_idx >= 3:
            col_idx = 0
            row_idx += 1
        if row_idx >= 8:
            c_design.showPage()
            c_mal.showPage()
            col_idx = 0
            row_idx = 0

    c_design.save()
    c_mal.save()
    print(f"\n[SUKSES] Selesai memproses {total_processed} kartu perpustakaan siswa.")
    print(f"-> File Desain Kartu: {output_pdf}")
    print(f"-> File Mal Potong Kartu: {output_mal}")

# =====================================================================
# INTERFACE MENU UTAMA (CLI)
# =====================================================================
def main_menu():
    while True:
        os.system('cls' if os.name == 'nt' else 'clear')
        print("="*60)
        print("     APLIKASI PETUGAS PERPUSTAKAAN - SMK WALISONGO JAKARTA ")
        print("="*60)
        print(" Silakan pilih modul administrasi cetak yang ingin dijalankan:")
        print(" [1] Generate & Layout Cetak STIKER LABEL BUKU (A3+)")
        print(" [2] Generate & Layout Cetak KARTU PERPUSTAKAAN SISWA (A3+)")
        print(" [3] Keluar dari Aplikasi")
        print("="*60)
        
        pilihan = input("Masukkan nomor pilihan Anda (1/2/3): ").strip()
        
        if pilihan == '1':
            modul_cetak_label_buku()
            input("\nTekan ENTER untuk kembali ke Menu Utama...")
        elif pilihan == '2':
            modul_cetak_kartu_perpus()
            input("\nTekan ENTER untuk kembali ke Menu Utama...")
        elif pilihan == '3':
            print("\nTerima kasih! Aplikasi ditutup. Selamat melanjutkan tugas administrasi.")
            break
        else:
            print("\n[Peringatan] Pilihan tidak valid! Masukkan angka 1, 2, atau 3.")
            input("Tekan ENTER untuk mencoba lagi...")

if __name__ == "__main__":
    main_menu()