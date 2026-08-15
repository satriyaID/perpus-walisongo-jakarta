import os
import pandas as pd
import urllib.parse
import urllib.request
from io import BytesIO

# Import dari ReportLab
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

def konversi_excel_ke_pdf_a3plus(excel_path, output_pdf="Cetak_Label_Buku_A3Plus.pdf"):
    print("Membaca data dari Excel...")
    # Membaca file Excel asli (.xlsx)
    # dtype=str memastikan angka/teks dibaca murni sebagai string
    df = pd.read_excel(excel_path, header=None, dtype=str)
    
    # Bersihkan baris yang kolom labelnya kosong (Kolom D / Index 3)
    df = df.dropna(subset=[3])  
    
    # Konversi ukuran kertas A3+ ke satuan poin ReportLab (1 mm = 2.83465 points)
    mm = 2.83465
    paper_width = 320 * mm
    paper_height = 480 * mm
    
    # Konfigurasi ukuran layout agar tidak terpotong di baris ke-15
    qr_size = 20 * mm        # Tetap presisi 2.0 cm
    gap_x = 5 * mm          # Jarak horizontal 5 mm
    gap_y = 5 * mm          # Jarak vertikal 5 mm
    margin_left = 10 * mm    # Margin kiri 1 cm
    margin_top = 12 * mm     # Margin atas 1.2 cm
    label_space = 2.5 * mm   # Jarak teks label di bawah kotak QR
    
    # Setup canvas PDF dengan ukuran custom A3+
    c = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
    
    col_idx = 0
    row_idx = 0
    total_processed = 0
    
    for index, row in df.iterrows():
        qr_text = str(row[1]) if pd.notna(row[1]) else str(row[3])
        label_text = str(row[3])
        
        # Lewati baris header jika terbaca teks judul di file Excel
        if "label" in label_text.lower() or "qr code" in qr_text.lower():
            continue
            
        print(f"Memproses [{total_processed + 1}]: {label_text}")
        
        # Kalkulasi Koordinat X & Y
        x = margin_left + (col_idx * (qr_size + gap_x))
        y = paper_height - margin_top - qr_size - (row_idx * (qr_size + gap_y))
        
        # =====================================================================
        # LOGIKA DETEKSI WARNA BERDASARKAN KOLOM E (INDEX 4) - DATA TERBARU
        # =====================================================================
        # row[4] adalah Kolom E di Excel. Kita gunakan .strip().upper() agar aman dari salah ketik spasi
        jurusan = str(row[4]).strip().upper() if pd.notna(row[4]) else ""
        
        if jurusan in ["TKJ", "TEKNO"]:
            # Lilac Lembut (Soft Lilac)
            bg_color = (0.88, 0.82, 0.98)
        elif jurusan in ["MP", "BISMEN"]:
            # Merah Muda Lembut (Soft Pink)
            bg_color = (1.0, 0.8, 0.85)
        elif jurusan == "BD":
            # Hijau Lembut (Soft Green)
            bg_color = (0.75, 0.93, 0.75) 
        elif jurusan == "AKL":
            # Kuning Lembut (Soft Yellow)
            bg_color = (1.0, 0.96, 0.7)
        else:
            # Putih jika tidak ada kode jurusan yang cocok / Umum
            bg_color = (1.0, 1.0, 1.0)
            
        # 1. Gambar Background Berwarna & Garis Bantu Potong Kotak QR
        c.setFillColorRGB(*bg_color)                 # Set warna background sesuai kondisi di atas
        c.setStrokeColorRGB(0.86, 0.86, 0.86)         # Set warna garis pembatas (Abu-abu)
        c.setLineWidth(0.5)
        
        # Menggambar kotak label (stroke=1 ada garis pinggir, fill=1 kotak diwarnai)
        c.rect(x, y, qr_size, qr_size, stroke=1, fill=1)
        
        # 2. Ambil gambar QR Code dari API secara realtime
        try:
            encoded_text = urllib.parse.quote(qr_text)
            qr_url = f"https://quickchart.io/qr?text={encoded_text}&size=150&margin=0"
            
            req = urllib.request.Request(qr_url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response:
                img_data = BytesIO(response.read())
                img = ImageReader(img_data)
                c.drawImage(img, x, y, width=qr_size, height=qr_size)
        except Exception as e:
            print(f"Gagal mengambil QR untuk {label_text}: {e}")
            c.line(x, y, x + qr_size, y + qr_size)
            c.line(x, y + qr_size, x + qr_size, y)

        # 3. Tulis Teks Label di bawah kotak QR Code
        c.setFillColorRGB(0.1, 0.1, 0.1) # Warna teks hitam/gelap
        c.setFont("Helvetica-Bold", 6.5)
        c.drawCentredString(x + (qr_size / 2), y - label_space, label_text)
        
        total_processed += 1
        col_idx += 1
        
        # Jika sudah mencapai 10 kolom, pindah ke baris berikutnya
        if col_idx >= 10:
            col_idx = 0
            row_idx += 1
            
        # Jika sudah mencapai 15 baris (Total 150 QR), buat halaman baru
        if row_idx >= 15:
            c.showPage()
            col_idx = 0
            row_idx = 0

    # Simpan file PDF secara final
    c.save()
    print(f"\nSukses! Total {total_processed} label buku selesai disusun dengan kode warna terbaru.")
    print(f"File PDF siap cetak disimpan di: {output_pdf}")

# Jalankan fungsi (Arahkan ke lokasi file Excel Anda di Drive G)
konversi_excel_ke_pdf_a3plus(r"G:\2025-2026\WS\QR_ID_BUKU.xlsx")