import os
import pandas as pd
import urllib.parse
import urllib.request
from io import BytesIO
from PIL import Image, ImageDraw, ImageFont

def generate_kartu_walisongo(excel_path, template_png_path, output_folder="Hasil_Kartu_Perpus_WS"):
    # 1. Membuat folder tujuan jika belum ada
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        
    print("Membaca data siswa...")
    # dtype=str menjaga angka 0 di depan NISN agar tidak hilang
    df = pd.read_excel(excel_path, dtype=str)
    
    # Bersihkan baris yang data NISN atau Nama-nya kosong/NaN
    df = df.dropna(subset=['NISN', 'Nama'])
    
    # =====================================================================
    # PENGATURAN FONT KHUSUS
    # =====================================================================
    # Pastikan file 'ChauPhilomene-Regular.ttf' sudah ditaruh di folder yang sama
    # Ukuran font Nama dinaikkan menjadi 32 (Bisa disesuaikan lagi jika kurang besar)
    nama_font_file = "ChauPhilomeneOne-Regular.ttf" 
    
    try:
        font_nama = ImageFont.truetype(nama_font_file, 32)
        print(f"Sukses memuat font khusus: {nama_font_file} (Ukuran 32)")
    except IOError:
        # Jika file ttf font khusus tidak ditemukan, otomatis pakai Arial Bold bawaan Windows
        font_nama = ImageFont.truetype("arialbd.ttf", 26)
        print(f"[Peringatan] File {nama_font_file} tidak ditemukan di folder. Menggunakan Arial Bold.")
        
    # Font untuk data informasi lainnya (NISN, Jurusan, Rombel) tetap menggunakan Arial Bold standar
    try:
        font_data_lain = ImageFont.truetype("arialbd.ttf", 22)  
    except IOError:
        font_data_lain = ImageFont.load_default()

    total_cetak = 0
    
    # 3. Proses Looping Data Siswa
    for index, row in df.iterrows():
        nama_siswa = str(row['Nama']).upper()      # Ubah nama ke huruf kapital
        nisn = str(row['NISN']).strip()            # Menjaga NISN tetap utuh beserta angka 0 di depan
        jurusan = str(row['Jurusan']).strip()
        rombel = str(row['Rombel']).strip()
        
        print(f"Memproses Kartu [{total_cetak + 1}]: {nama_siswa} - NISN: {nisn}")
        
        # Buka berkas gambar template PNG Anda
        base_image = Image.open(template_png_path).convert("RGBA")
        draw = ImageDraw.Draw(base_image)
        
        # A. MENULIS DATA DINAMIS SISWA
        # Kolom Nama menggunakan font_nama (Chau Philomene) yang ukurannya lebih besar
        # Koordinat Y untuk Nama diturunkan sedikit (dari 287 ke 282) agar teks yang lebih besar tetap simetris di jalurnya
        draw.text((310, 282), nama_siswa, font=font_nama, fill=(0, 0, 0))
        
        # Kolom lainnya tetap menggunakan font standar
        draw.text((310, 332), nisn, font=font_data_lain, fill=(0, 0, 0))
        draw.text((310, 377), jurusan, font=font_data_lain, fill=(0, 0, 0))
        draw.text((310, 422), rombel, font=font_data_lain, fill=(0, 0, 0))
        
        # B. GENERATE & MENIMPA QR CODE
        try:
            encoded_nisn = urllib.parse.quote(nisn)
            qr_url = f"https://quickchart.io/qr?text={encoded_nisn}&size=200&margin=1"
            
            req = urllib.request.Request(qr_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                qr_data = BytesIO(response.read())
                qr_image = Image.open(qr_data).convert("RGBA")
                
                # Diubah ukurannya menjadi 175x175 piksel agar pas menutupi total QR Code contoh bawaan
                qr_image = qr_image.resize((175, 175))
                
                posisi_qr_x = 740
                posisi_qr_y = 280
                
                # Tempelkan QR Code baru menimpa QR contoh bawaan template gambar
                base_image.alpha_composite(qr_image, (posisi_qr_x, posisi_qr_y))
                
        except Exception as e:
            print(f"   [Gagal] Mengambil QR Code untuk {nama_siswa}: {e}")
            
        # C. MENYIMPAN HASIL KARTU PER SISWA
        filename = f"{nisn}_{nama_siswa.replace(' ', '_')}.png"
        output_path = os.path.join(output_folder, filename)
        
        final_card = base_image.convert("RGB")
        final_card.save(output_path, "PNG")
        total_cetak += 1

    print(f"\n[SUKSES] Berhasil membuat total {total_cetak} Kartu Perpustakaan.")
    print(f"Semua file disimpan dengan rapi di dalam folder: '{output_folder}'")

# =====================================================================
# BAGIAN EKSEKUSI PROGRAM
# =====================================================================
FILE_EXCEL = r"G:\2025-2026\WS\Data Kartu Perpus 12.xlsx"
TEMPLATE_GAMBAR = r"G:\2025-2026\WS\Kartu Perpustakaan SMK Walisongo.png"

generate_kartu_walisongo(FILE_EXCEL, TEMPLATE_GAMBAR)