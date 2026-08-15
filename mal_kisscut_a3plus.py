import os
import pandas as pd

# Import dari ReportLab
from reportlab.pdfgen import canvas

def buat_mal_kisscut_a3plus(excel_path, output_pdf="Mal_Kisscut_Label_A3Plus.pdf"):
    print("Membuat Mal Kisscut untuk percetakan...")
    df = pd.read_excel(excel_path, header=None, dtype=str)
    df = df.dropna(subset=[3])  
    
    mm = 2.83465
    paper_width = 320 * mm
    paper_height = 480 * mm
    
    # Konfigurasi ukuran layout (Harus SAMA PERSIS dengan file desain)
    qr_size = 25 * mm        # Ukuran kotak stiker 2.5 cm
    gap_x = 5 * mm
    gap_y = 4 * mm
    margin_left = 10 * mm
    margin_top = 12 * mm
    
    c = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
    
    col_idx = 0
    row_idx = 0
    total_processed = 0
    
    for index, row in df.iterrows():
        label_text = str(row[3])
        if "label" in label_text.lower():
            continue
            
        x = margin_left + (col_idx * (qr_size + gap_x))
        y = paper_height - margin_top - qr_size - (row_idx * (qr_size + gap_y))
        
        # =====================================================================
        # WARNA MAL KISSCUT: MAGENTA MURNI (100% M) UNTUK SENSOR MESIN POTONG
        # =====================================================================
        c.setStrokeColorRGB(1.0, 0.0, 1.0) # Warna Magenta Cerah
        c.setLineWidth(0.5)                 # Garis tipis hairline
        
        # Menggambar kotak mal potong (stroke=1 (garis saja), fill=0 (kosong transparan))
        c.rect(x, y, qr_size, qr_size, stroke=1, fill=0)
        
        total_processed += 1
        col_idx += 1
        
        if col_idx >= 10:
            col_idx = 0
            row_idx += 1
            
        if row_idx >= 15:
            c.showPage()
            col_idx = 0
            row_idx = 0

    c.save()
    print(f"\nSukses! File Mal Kisscut berhasil dibuat: {output_pdf}")

# Jalankan fungsi ke file Excel Anda
buat_mal_kisscut_a3plus(r"G:\2025-2026\WS\QR_ID_BUKU.xlsx")