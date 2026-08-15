import os
import sys
import sqlite3
import pandas as pd
import urllib.parse
import urllib.request
import shutil
from io import BytesIO
from datetime import datetime
import threading

# Import Tkinter untuk GUI & Dialog Input
import tkinter as tk
from tkinter import messagebox, ttk, simpledialog

# Import ReportLab & Pillow untuk Cetak PDF
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont

# =====================================================================
# PATH CONFIGURATION (DYNAMIC, RELATIVE & SECURE "PERPUS" FOLDER)
# =====================================================================
if getattr(sys, 'frozen', False):
    FOLDER_KERJA = os.path.dirname(sys.executable)
else:
    FOLDER_KERJA = os.path.dirname(os.path.abspath(__file__))

# Folder Aktif Sesuai Analisis Keamanan & Efek Psikologis
DIR_PERPUS = os.path.join(FOLDER_KERJA, "DB_PERPUS")
os.makedirs(DIR_PERPUS, exist_ok=True)

FOLDER_BACKUP = os.path.join(DIR_PERPUS, "Backup_Data")
os.makedirs(FOLDER_BACKUP, exist_ok=True)

# Kunci Database Utama SQLite
DB_PATH = os.path.join(DIR_PERPUS, "perpustakaan.db")

# Aset Gambar & Font Pendukung (Tetap Relatif di Folder Kerja)
TEMPLATE_KARTU = os.path.join(FOLDER_KERJA, "Kartu Perpustakaan SMK Walisongo.png")
FONT_CHAU = os.path.join(FOLDER_KERJA, "ChauPhilomene-Regular.ttf")

# Referensi File Excel Lama di Folder Luar untuk Target Migrasi
EXCEL_BUKU_LAMA = os.path.join(FOLDER_KERJA, "QR_ID_BUKU.xlsx")
EXCEL_SISWA_LAMA = os.path.join(FOLDER_KERJA, "Data_Siswa_Perpus.xlsx")
EXCEL_LOG_LAMA = os.path.join(FOLDER_KERJA, "Data_Peminjaman_Buku.xlsx")
EXCEL_UMUM_LAMA = os.path.join(FOLDER_KERJA, "Data_Buku_Umum.xlsx")

MM_TO_POINT = 2.83465

class AplikasiPerpusTerintegrasi:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Informasi Perpustakaan - SMK Walisongo Jakarta")
        self.root.geometry("920x860")
        self.root.configure(bg="#f4f6f9")
        
        self.target_buku = {"10": 10, "11": 15, "12": 4}
        
        # State Operasional Aplikasi
        self.siswa_aktif = None
        self.daftar_buku_dijepit = []
        self.daftar_kembali_dijepit = []
        self.antrean_reprint = []
        self.dict_check_siswa = {} 
        self.siswa_terfilter_cetak = []
        
        # 1. Amankan Peringatan Teks Psikologis di Folder PERPUS
        with open(os.path.join(DIR_PERPUS, "⚠️ JANGAN DIHAPUS.txt"), "w", encoding="utf-8") as f:
            f.write("PENTING:\nFolder ini berisi database utama SQLite Aplikasi Perpustakaan SMK Walisongo.\n"
                    "Menghapus atau mengubah isi folder ini akan merusak sistem.")
            
        # 2. Inisialisasi Rumah Tabel SQLite & Jalankan Sedot Data Excel Lama (Migrasi)
        self.inisialisasi_database()
        self.migrasi_excel_ke_sqlite()
        self.jalankan_auto_backup_silent()
        
        # =====================================================================
        # BAGIAN INITIALISASI TEKS BERJALAN (MARQUEE WAKAF & COUNTER HARI)
        # =====================================================================
        tanggal_mulai = datetime(2026, 7, 15)
        tanggal_sekarang = datetime.now()
        selisih_hari = (tanggal_sekarang - tanggal_mulai).days
        if selisih_hari < 0: selisih_hari = 0
        
        self.teks_marquee = (
            f"Wakaf Aplikasi Perpustakaan SMKS Walisongo Jakarta oleh: Satriyana, S.Kom., M.Pd. "
            f"kepada Yayasan Pendidikan Islam Hj. Dardjah Amin. Juli 2026.     "
            f"★ Telah dimanfaatkan selama {selisih_hari} Hari terhitung sejak 15 Juli 2026.          "
        )
        
        self.lbl_marquee = tk.Label(
            self.root, 
            text=self.teks_marquee, 
            font=("Arial", 10, "bold"), 
            bg="#202124", 
            fg="#00ff00", 
            pady=6, 
            anchor=tk.W
        )
        self.lbl_marquee.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.animasi_teks_berjalan()
        self.tampilkan_menu_utama()

    # =====================================================================
    # ARSITEKTUR DATABASE ENGINE: SQLITE MANAGEMENT
    # =====================================================================
    def inisialisasi_database(self):
        """Membuat rumah berkas database dan tabel utama terstruktur."""
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # Tabel 1: Log Sirkulasi Transaksi
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS data_peminjaman_buku (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                tanggal_pinjam TEXT,
                tanggal_kembali TEXT DEFAULT '-',
                nisn TEXT,
                nama_siswa TEXT,
                rombel TEXT,
                kode_satuan_buku TEXT,
                status TEXT
            )
        ''')
        
        # Tabel 2: Master Buku Paket Sekolah (QR_ID_BUKU)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS qr_id_buku (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                kolom_0 TEXT,
                qr_code TEXT,
                kolom_2 TEXT,
                label_buku TEXT,
                jurusan TEXT
            )
        ''')
        
        # Tabel 3: Master Siswa Perpus (Dibuat fleksibel tanpa batasan rigid agar migrasi mulus)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS data_siswa_perpus (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nisn TEXT,
                nama TEXT,
                jurusan TEXT,
                rombel TEXT
            )
        ''')
        
        # Tabel 4: Master Buku Fiksi / ISBN Umum On-the-spot
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS data_buku_umum (
                isbn TEXT PRIMARY KEY,
                judul_buku TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def migrasi_excel_ke_sqlite(self):
        """
        Melakukan migrasi hanya SATU KALI saat database masih kosong.
        Jika tabel di SQLite sudah berisi data, proses migrasi dilewati agar
        transaksi harian tidak terhapus / ter-reset.
        """
        import datetime
        
        folder_arsip = os.path.join(DIR_PERPUS, "Arsip_Migrasi")
        
        mapping_migrasi = [
            (EXCEL_LOG_LAMA, "data_peminjaman_buku"),
            (EXCEL_BUKU_LAMA, "qr_id_buku"),
            (EXCEL_SISWA_LAMA, "data_siswa_perpus"),
            (EXCEL_UMUM_LAMA, "data_buku_umum")
        ]
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        for path_excel, nama_tabel in mapping_migrasi:
            # 1. Cek dulu apakah tabel di SQLite sudah ada isinya
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {nama_tabel}")
                jumlah_baris = cursor.fetchone()[0]
            except sqlite3.OperationalError:
                jumlah_baris = 0
            
            # PENTING: Jika tabel sudah ada isinya (> 0 baris), SKIP migrasi!
            # Ini mencegah data transaksi harian ter-reset ke kondisi awal.
            if jumlah_baris > 0:
                print(f"ℹ️ [MIGRATION] Tabel '{nama_tabel}' sudah terisi ({jumlah_baris} baris). Memproses data SQLite aktif...")
                continue
            
            # 2. Jika tabel masih kosong dan file Excel lama ditemukan, lakukan impor
            if os.path.exists(path_excel):
                file_name = os.path.basename(path_excel)
                print(f"🔄 Mengimpor data awal dari '{file_name}' ke tabel '{nama_tabel}'...")
                
                try:
                    if nama_tabel == "qr_id_buku":
                        df_lama = pd.read_excel(path_excel, header=None, dtype=str)
                        df_lama.columns = ["kolom_0", "qr_code", "kolom_2", "label_buku", "jurusan"][:len(df_lama.columns)]
                    else:
                        df_lama = pd.read_excel(path_excel, dtype=str)
                        df_lama = df_lama.loc[:, ~df_lama.columns.str.contains('^Unnamed')]
                        df_lama.columns = [c.strip().replace(" ", "_").lower() for c in df_lama.columns]
                    
                    df_lama = df_lama.dropna(how='all')
                    
                    if nama_tabel == "data_siswa_perpus":
                        df_lama = df_lama.dropna(subset=["nisn", "nama"], how='any')
                        df_lama = df_lama.drop_duplicates(subset=["nisn"], keep='first')
                    elif nama_tabel == "qr_id_buku":
                        df_lama = df_lama.dropna(subset=["label_buku"])
                        df_lama = df_lama.drop_duplicates(subset=["label_buku"], keep='first')
                    
                    # Gunakan append karena dipastikan hanya berjalan saat tabel kosong
                    df_lama.to_sql(nama_tabel, conn, if_exists="append", index=False)
                    print(f"✅ [MIGRATION] Impor awal {file_name} sukses!")
                    
                    # Pindahkan Excel lama ke Arsip agar tidak dibaca lagi
                    waktu_sekarang = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                    os.makedirs(folder_arsip, exist_ok=True)
                    shutil.move(path_excel, os.path.join(folder_arsip, f"Migrated_{waktu_sekarang}_{file_name}"))
                    
                except Exception as e:
                    print(f"❌ [MIGRATION ERROR] Gagal konversi {file_name}: {e}")
                    
        conn.commit()
        conn.close()

    def bersihkan_layar(self):
        for widget in self.root.winfo_children():
            if widget != self.lbl_marquee:
                widget.destroy()

    def animasi_teks_berjalan(self):
        try:
            self.teks_marquee = self.teks_marquee[1:] + self.teks_marquee[0]
            self.lbl_marquee.config(text=self.teks_marquee)
            self.root.after(150, self.animasi_teks_berjalan)
        except: pass

    def log_ke_gui(self, text_widget, message):
        text_widget.config(state=tk.NORMAL)
        text_widget.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        text_widget.see(tk.END)
        text_widget.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def jalankan_auto_backup_silent(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(DB_PATH, os.path.join(FOLDER_BACKUP, f"perpustakaan_backup_{timestamp}.db"))
        except: pass

    def pemicu_backup_manual(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(DB_PATH, os.path.join(FOLDER_BACKUP, f"MANUAL_{timestamp}_perpustakaan.db"))
            messagebox.showinfo("Backup Sukses", "Database SQLite sukses dikunci di folder Backup_Data!")
        except Exception as e:
            messagebox.showerror("Gagal", f"Gagal mencadangkan database: {e}")

    def pemicu_recovery_data(self):
        if os.path.exists(FOLDER_BACKUP): os.startfile(FOLDER_BACKUP)

    def gambar_crop_mark(self, canvas_obj, x, y, w, h, bleed, length=15):
        canvas_obj.setStrokeColorRGB(0.5, 0.5, 0.5)
        canvas_obj.setLineWidth(0.5)
        canvas_obj.line(x - bleed - length, y, x - bleed, y)
        canvas_obj.line(x, y - bleed - length, x, y - bleed)
        canvas_obj.line(x - bleed - length, y + h, x - bleed, y + h)
        canvas_obj.line(x, y + h + bleed, x, y + h + bleed + length)
        canvas_obj.line(x + w + bleed, y + h, x + w + bleed + length, y + h)
        canvas_obj.line(x + w, y + h + bleed, x + w, y + h + bleed + length)
        canvas_obj.line(x + w + bleed, y, x + w + bleed + length, y)
        canvas_obj.line(x + w, y - bleed - length, x + w, y - bleed)

    def dapatkan_judul_buku(self, kode_buku):
        """Mendeteksi judul buku paket (via nama convention) atau fiksi umum (via database SQLite ISBN)"""
        kode_str = str(kode_buku).upper().strip()
        if kode_str.startswith("978") or kode_str.startswith("979"):
            isbn_induk = kode_str.split('-')[0]
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query("SELECT judul_buku FROM data_buku_umum WHERE isbn = ?", conn, params=(isbn_induk,))
            conn.close()
            if not df.empty:
                return df.iloc[0]['judul_buku']
            return "Buku Fiksi / Umum Baru (Belum Bernama)"

        kamus_judul = {
            "TKJ-ASJ1": "Administrasi Sistem Jaringan V1", "TKJ-ASJ2": "Administrasi Sistem Jaringan V2",
            "AKL-AKJS": "Akuntansi Jasa", "AKL-AKKU": "Akuntansi Keuangan", "UMUM-INDO": "Bahasa Indonesia",
            "UMUM-INGG": "Bahasa Inggris", "BISMEN-INGB": "Bahasa Inggris Bismen", "TEKNO-INGT": "Bahasa Inggris Teknologi",
            "AKL-DSA1": "Dasar-dasar Akuntansi V1", "AKL-DSA2": "Dasar-dasar Akuntansi V2", "MP-DSM1": "Dasar-dasar MPLB V1",
            "MP-DSM2": "Dasar-dasar MPLB V2", "BD-DSP1": "Dasar-dasar Pemasaran V1", "BD-DSP2": "Dasar-dasar Pemasaran V2",
            "TKJ-DST1": "Dasar-dasar TJKT V1", "TKJ-DST2": "Dasar-dasar TJKT V2", "MP-HMAS": "Humas",
            "BISMEN-INFB": "Informatika Bismen", "TEKNO-INFT": "Informatika Teknologi", "BISMEN-IPAB": "IPAS Bismen",
            "TEKNO-IPAT": "IPAS Teknologi", "TKJ-KJA1": "Keamanan Jaringan V1", "TKJ-KJA2": "Keamanan Jaringan V2",
            "MP-ARSP": "Kearsipan", "UMUM-MATE": "Matematika", "BISMEN-MATB": "Matematika Bismen",
            "TEKNO-MATT": "Matematika Teknologi", "AKL-MYOB": "MYOB", "AKL-PJAK": "Perpajakan",
            "TKJ-JAR1": "Perangkat Jaringan V1", "TKJ-JAR2": "Perangkat Jaringan V2", "UMUM-PJOK": "PJOK",
            "UMUM-PPKN": "Pendidikan Pancasila", "UMUM-SJRH": "Sejarah", "TKJ-TKJR": "Teknik Komputer Jaringan"
        }
        for kunci, judul in kamus_judul.items():
            if kunci in kode_str: return judul
        return "Buku Paket Umum / Kejuruan"

    # =====================================================================
    # HALAMAN 1: MENU UTAMA GUI (DASHBOARD AREA)
    # =====================================================================
    def tampilkan_menu_utama(self):
        self.bersihkan_layar()
        
        banner = tk.Label(self.root, text="SISTEM PERPUSTAKAAN SMK WALISONGO JAKARTA", font=("Arial", 16, "bold"), bg="#1a73e8", fg="white", pady=15)
        banner.pack(fill=tk.X)
        
        lbl_path = tk.Label(self.root, text=f"Folder Aktif Sistem: {DIR_PERPUS}", font=("Consolas", 9), bg="#e8eaed", fg="#5f6368", pady=4)
        lbl_path.pack(fill=tk.X)
        
        frame_tombol = tk.Frame(self.root, bg="#f4f6f9")
        frame_tombol.pack(expand=True, pady=5)
        
        tk.Button(frame_tombol, text="[1] MODUL SIRKULASI TERINTEGRASI\n(Mendukung Scan Buku Paket & ISBN Fiksi On-The-Spot)", font=("Arial", 11, "bold"), bg="#34a853", fg="white", width=45, height=2, command=self.tampilkan_layaran_sirkulasi).grid(row=0, column=0, pady=4)
        tk.Button(frame_tombol, text="[2] CETAK STIKER LABEL BUKU MASSAL\n(Proses Latar Belakang A3+)", font=("Arial", 11, "bold"), bg="#1a73e8", fg="white", width=45, height=2, command=self.tampilkan_layar_proses_cetak_buku).grid(row=1, column=0, pady=4)
        tk.Button(frame_tombol, text="[3] CETAK KARTU PERPUS DENGAN CEKLIS\n(Pilih Rombel & Pilih Siswa Secara Selektif)", font=("Arial", 11, "bold"), bg="#ff9900", fg="white", width=45, height=2, command=self.tampilkan_layar_setup_cetak_kartu).grid(row=2, column=0, pady=4)
        tk.Button(frame_tombol, text="[4] CETAK ULANG KARTU (REPRINT KOLEKTIF)\n(Cari via NISN - Maks 9 Siswa di A4 Landscape)", font=("Arial", 11, "bold"), bg="#6f42c1", fg="white", width=45, height=2, command=self.tampilkan_layar_cetak_satuan).grid(row=3, column=0, pady=4)
        tk.Button(frame_tombol, text="[5] MODUL REKAP & LIVE ANALYTICS\n(Metrik Laporan Buku Paket & Buku Umum Fiksi)", font=("Arial", 11, "bold"), bg="#007afc", fg="white", width=45, height=2, command=self.tampilkan_layar_rekap).grid(row=4, column=0, pady=4)
        
        frame_utilitas = tk.LabelFrame(self.root, text=" Pengaman & Pemulihan Data Relasional ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_utilitas.pack(fill=tk.X, padx=30, pady=5)
        tk.Button(frame_utilitas, text="Amankan Cadangan SQLite (Backup)", font=("Arial", 9, "bold"), bg="#5f6368", fg="white", command=self.pemicu_backup_manual).pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)
        tk.Button(frame_utilitas, text="Buka Folder Cadangan (Recovery)", font=("Arial", 9, "bold"), bg="#7c7c7c", fg="white", command=self.pemicu_recovery_data).pack(side=tk.RIGHT, padx=10, expand=True, fill=tk.X)
        
        tk.Button(self.root, text="Keluar Aplikasi", font=("Arial", 10), bg="#d93025", fg="white", width=15, pady=5, command=self.root.quit).pack(pady=10)

    # =====================================================================
    # HALAMAN 2: MODUL SIRKULASI TERINTEGRASI (KUNCI RENDER HORIZONTAL BAWAH)
    # =====================================================================
    def tampilkan_layaran_sirkulasi(self):
        self.bersihkan_layar()
        
        conn = sqlite3.connect(DB_PATH)
        self.df_siswa = pd.read_sql_query("SELECT * FROM data_siswa_perpus", conn)
        conn.close()
        
        if self.df_siswa.empty:
            messagebox.showwarning("Database Kosong", "Master data siswa di SQLite tidak ditemukan.")
            self.tampilkan_menu_utama()
            return
            
        self.df_siswa['nisn'] = self.df_siswa['nisn'].str.strip()
        
        banner = tk.Label(self.root, text="MODUL SIRKULASI PERPUSTAKAAN DUA ARAH (SQLITE ENGINE)", font=("Arial", 14, "bold"), bg="#34a853", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        tabControl = ttk.Notebook(self.root)
        tab_pinjam = tk.Frame(tabControl, bg="#f4f6f9"); tab_kembali = tk.Frame(tabControl, bg="#f4f6f9")
        tabControl.add(tab_pinjam, text="  [+] LAYANAN PEMINJAMAN BUKU PAS  "); tabControl.add(tab_kembali, text="  [-] LAYANAN PENGEMBALIAN BUKU PAS  ")
        tabControl.pack(expand=1, fill="both", padx=10, pady=5)
        
        # --- TAB 1: PEMINJAMAN GRIDS ---
        frame_scan_siswa = tk.LabelFrame(tab_pinjam, text=" Langkah 1: Scan Kartu Perpustakaan Siswa ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_scan_siswa.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_scan_siswa, text="Scan NISN:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_nisn = tk.Entry(frame_scan_siswa, font=("Arial", 12), width=30); self.ent_nisn.grid(row=0, column=1, padx=10); self.ent_nisn.bind("<Return>", self.proses_scan_siswa); self.ent_nisn.focus()
        self.lbl_info_siswa = tk.Label(frame_scan_siswa, text="[Silakan scan kartu siswa untuk memulai]", font=("Arial", 11, "italic"), fg="#555", bg="#e8eaed", width=75, height=4, anchor=tk.W, justify=tk.LEFT, padx=10, pady=5); self.lbl_info_siswa.grid(row=1, column=0, columnspan=2, pady=10)
        
        self.frame_scan_buku = tk.LabelFrame(tab_pinjam, text=" Langkah 2: Scan Kode Buku Paket / Barcode ISBN Buku Umum ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); self.frame_scan_buku.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Label(self.frame_scan_buku, text="Scan Barcode:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_buku = tk.Entry(self.frame_scan_buku, font=("Arial", 12), width=30, state=tk.DISABLED); self.ent_buku.grid(row=0, column=1, padx=10); self.ent_buku.bind("<Return>", self.proses_scan_buku)
        
        self.tree_buku = ttk.Treeview(self.frame_scan_buku, columns=("No", "Kode Buku", "Nama Buku"), show="headings", height=6)
        self.tree_buku.heading("No", text="No"); self.tree_buku.heading("Kode Buku", text="ID Satuan / Eksemplar"); self.tree_buku.heading("Nama Buku", text="Judul Buku Terdeteksi")
        self.tree_buku.column("No", width=50, anchor=tk.CENTER); self.tree_buku.column("Kode Buku", width=200, anchor=tk.W); self.tree_buku.column("Nama Buku", width=450, anchor=tk.W)
        self.tree_buku.grid(row=1, column=0, columnspan=2, pady=5, sticky="nsew")
        
        self.lbl_counter = tk.Label(self.frame_scan_buku, text="Jumlah Buku Di-scan: 0", font=("Arial", 11, "bold"), bg="#f4f6f9", fg="#1a73e8"); self.lbl_counter.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        self.frame_scan_buku.grid_rowconfigure(1, weight=1); self.frame_scan_buku.grid_columnconfigure(1, weight=1)
        
        # --- TAB 2: PENGEMBALIAN GRIDS ---
        frame_scan_kembali = tk.LabelFrame(tab_kembali, text=" Masukkan / Scan Label Buku Yang Dikembalikan ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=10); frame_scan_kembali.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_scan_kembali, text="Scan Barcode:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_buku_kembali = tk.Entry(frame_scan_kembali, font=("Arial", 12), width=30); self.ent_buku_kembali.grid(row=0, column=1, padx=10, sticky=tk.W); self.ent_buku_kembali.bind("<Return>", self.proses_scan_pengembalian_buku)
        frame_tabel_kembali = tk.LabelFrame(tab_kembali, text=" Manifest Validasi Buku Yang Berhasil Dikembalikan ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_tabel_kembali.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_kembali = ttk.Treeview(frame_tabel_kembali, columns=("No", "Kode Buku", "Nama Siswa", "Rombel", "Status"), show="headings", height=8); self.tree_kembali.heading("No", text="No"); self.tree_kembali.heading("Kode Buku", text="Kode Seri Buku"); self.tree_kembali.heading("Nama Siswa", text="Nama Mantan Peminjam"); self.tree_kembali.heading("Rombel", text="Rombel"); self.tree_kembali.heading("Status", text="Status Akhir")
        self.tree_kembali.column("No", width=40, anchor=tk.CENTER); self.tree_kembali.column("Kode Buku", width=160, anchor=tk.W); self.tree_kembali.column("Nama Siswa", width=250, anchor=tk.W); self.tree_kembali.column("Rombel", width=90, anchor=tk.CENTER); self.tree_kembali.column("Status", width=90, anchor=tk.CENTER); self.tree_kembali.pack(fill=tk.BOTH, expand=True, pady=5)
        
        frame_aksi_kembali = tk.Frame(tab_kembali, bg="#f4f6f9", pady=5); frame_aksi_kembali.pack(fill=tk.X, padx=10)
        self.btn_simpan_kembali = tk.Button(frame_aksi_kembali, text="Proses & Bukukan Pengembalian", font=("Arial", 11, "bold"), bg="#1a73e8", fg="white", padx=15, pady=6, command=self.eksekusi_simpan_pengembalian, state=tk.DISABLED); self.btn_simpan_kembali.pack(side=tk.RIGHT)

        # =====================================================================
        # MASTERPIECE NAVIGASI HORIZONTAL LAYOUT - ANTI TERBALIK KUNCI 100%
        # =====================================================================
        self.lbl_marquee.pack_forget()
        
        # 1. TEMPEL TEXT BERJALAN DULUAN
        self.lbl_marquee.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 2. CONTAINER UTAMA TOMBOL (Pack side BOTTOM setelah marquee akan otomatis berada DI ATAS marquee)
        frame_navigasi_total = tk.Frame(self.root, bg="#f4f6f9", pady=10)
        
        frame_nav_kiri = tk.Frame(frame_navigasi_total, bg="#f4f6f9")
        frame_nav_kiri.pack(side=tk.LEFT, padx=15)
        btn_back_utama = tk.Button(frame_nav_kiri, text="◀ Kembali ke Menu Utama (Dashboard)", font=("Arial", 11, "bold"), bg="#d93025", fg="white", padx=20, pady=8, command=self.tampilkan_menu_utama)
        btn_back_utama.pack(anchor=tk.W)

        frame_aksi_sirkulasi = tk.Frame(frame_navigasi_total, bg="#f4f6f9")
        frame_aksi_sirkulasi.pack(side=tk.RIGHT, padx=15)
        self.btn_simpan = tk.Button(frame_aksi_sirkulasi, text="Simpan & Lanjut Siswa Lain ▶", font=("Arial", 11, "bold"), bg="#188038", fg="white", padx=20, pady=8, command=self.konfirmasi_dan_simpan, state=tk.DISABLED)
        self.btn_simpan.pack(anchor=tk.E)
        
        frame_navigasi_total.pack(fill=tk.X, side=tk.BOTTOM)

    def proses_scan_siswa(self, event=None):
        nisn_input = self.ent_nisn.get().strip()
        if not nisn_input: return
        hasil = self.df_siswa[self.df_siswa['nisn'] == nisn_input]
        if not hasil.empty:
            row = hasil.iloc[0]
            self.siswa_aktif = {
                "nisn": row['nisn'], "nama": row['nama'].upper(), "jurusan": row['jurusan'], "rombel": row['rombel'], "kelas": str(row['rombel']).split('-')[0]
            }
            kelas_digit = "".join(filter(str.isdigit, self.siswa_aktif["kelas"]))
            self.target_seharusnya = self.target_buku.get(kelas_digit, 10)
            info_text = f"Nama    : {self.siswa_aktif['nama']}\nNISN    : {self.siswa_aktif['nisn']}\nRombel  : {self.siswa_aktif['rombel']} | Jurusan: {self.siswa_aktif['jurusan']}\nTarget Wajib: {self.target_seharusnya} Buku Paket"
            self.lbl_info_siswa.config(text=info_text, fg="#155724", bg="#d4edda", font=("Arial", 11, "bold"))
            self.ent_buku.config(state=tk.NORMAL); self.ent_buku.focus(); self.ent_nisn.config(state=tk.DISABLED)
        else:
            messagebox.showwarning("Tidak Ditemukan", f"Siswa dengan NISN {nisn_input} tidak terdaftar di SQLite.")
            self.ent_nisn.delete(0, tk.END)

    def proses_scan_buku(self, event=None):
        input_raw = self.ent_buku.get().strip()
        self.ent_buku.delete(0, tk.END)
        if not input_raw: return
        
        id_final_buku = input_raw
        judul_terdeteksi = ""
        
        if len(input_raw) == 13 and (input_raw.startswith("978") or input_raw.startswith("979")):
            conn = sqlite3.connect(DB_PATH)
            df_umum = pd.read_sql_query("SELECT * FROM data_buku_umum WHERE isbn = ?", conn, params=(input_raw,))
            
            if df_umum.empty:
                judul_baru = simpledialog.askstring("Buku Fiksi/Umum Baru", f"Barcode ISBN {input_raw} BELUM TERCATAT.\n\nMasukkan Judul Buku / Novel:", parent=self.root)
                if not judul_baru: 
                    conn.close()
                    return
                judul_terdeteksi = judul_baru.upper().strip()
                cursor = conn.cursor()
                cursor.execute("INSERT INTO data_buku_umum (isbn, judul_buku) VALUES (?,?)", (input_raw, judul_terdeteksi))
                conn.commit()
                eksemplar_ke = 1
            else:
                judul_terdeteksi = df_umum.iloc[0]['judul_buku']
                df_log_cek = pd.read_sql_query("SELECT kode_satuan_buku FROM data_peminjaman_buku WHERE kode_satuan_buku LIKE ?", conn, params=(f"{input_raw}%",))
                eksemplar_ke = len(df_log_cek['kode_satuan_buku'].unique()) + 1
            
            conn.close()
            id_final_buku = f"{input_raw}-{str(eksemplar_ke).zfill(3)}"
        else:
            judul_terdeteksi = self.dapatkan_judul_buku(input_raw)

        if any(buku_id == id_final_buku for buku_id, _, _ in self.daftar_buku_dijepit):
            messagebox.showwarning("Duplikasi", "Buku tersebut sudah di-scan di lembar ini.")
            return

        self.daftar_buku_dijepit.append((id_final_buku, judul_terdeteksi, input_raw))
        no_urut = len(self.daftar_buku_dijepit)
        self.tree_buku.insert("", tk.END, values=(no_urut, id_final_buku, judul_terdeteksi))
        self.lbl_counter.config(text=f"Jumlah Buku Di-scan: {no_urut} Buku")
        self.btn_simpan.config(state=tk.NORMAL)

    def konfirmasi_dan_simpan(self):
        if messagebox.askyesno("Konfirmasi", "Simpan seluruh daftar peminjaman siswa ini ke SQLite?"):
            tgl_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                for id_buku, _, _ in self.daftar_buku_dijepit:
                    cursor.execute('''
                        INSERT INTO data_peminjaman_buku (tanggal_pinjam, nisn, nama_siswa, rombel, kode_satuan_buku, status)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (tgl_sekarang, self.siswa_aktif["nisn"], self.siswa_aktif["nama"], self.siswa_aktif["rombel"], id_buku, "DIPINJAM"))
                conn.commit()
                conn.close()
                self.jalankan_auto_backup_silent()
                messagebox.showinfo("Sukses", "Data peminjaman aman terkunci di SQLite!")
                self.tampilkan_layaran_sirkulasi()
            except Exception as e:
                messagebox.showerror("Error", f"Gagal membukukan transaksi: {e}")

    def proses_scan_pengembalian_buku(self, event=None):
        kode_input = self.ent_buku_kembali.get().strip()
        self.ent_buku_kembali.delete(0, tk.END)
        if not kode_input: return
        
        # Cek apakah buku sudah masuk di antrean layar pengembalian saat ini
        if any(b['Kode Buku'] == kode_input for b in self.daftar_kembali_dijepit):
            messagebox.showwarning("Duplikasi", "Buku tersebut sudah ada dalam daftar pengembalian saat ini.")
            return
            
        conn = sqlite3.connect(DB_PATH)
        
        # KUNCI SOLUSI: Jika pustakawan scan ISBN Fisik (13 digit angka 978/979 tanpa -001)
        if len(kode_input) == 13 and (kode_input.startswith("978") or kode_input.startswith("979")):
            # Cari di SQLite menggunakan Wildcard (LIKE '978...%')
            df_log = pd.read_sql_query(
                "SELECT * FROM data_peminjaman_buku WHERE kode_satuan_buku LIKE ? AND UPPER(status) LIKE '%PINJAM%'", 
                conn, 
                params=(f"{kode_input}%",)
            )
        else:
            # Pencarian presisi exact untuk Buku Paket Sekolah (misal: 12-TJKT-001)
            df_log = pd.read_sql_query(
                "SELECT * FROM data_peminjaman_buku WHERE kode_satuan_buku = ? AND UPPER(status) LIKE '%PINJAM%'", 
                conn, 
                params=(kode_input,)
            )
            
        conn.close()
        
        if not df_log.empty:
            row = df_log.iloc[-1] # Ambil transaksi aktif paling mutakhir
            id_db = str(row['id'])
            kode_real_db = str(row['kode_satuan_buku']) # Ambil ID lengkap asli dari DB (misal 9786232669383-001)
            nama_siswa = str(row['nama_siswa']).upper()
            rombel_siswa = str(row['rombel'])
            
            buku_kembali = {
                "id_pk": id_db, 
                "Kode Buku": kode_real_db, 
                "Nama Siswa": nama_siswa, 
                "Rombel": rombel_siswa, 
                "Status": "KEMBALI"
            }
            
            self.daftar_kembali_dijepit.append(buku_kembali)
            no_urut = len(self.daftar_kembali_dijepit)
            
            # Tampilkan di tabel manifest GUI pengembalian
            self.tree_kembali.insert("", tk.END, values=(no_urut, kode_real_db, nama_siswa, rombel_siswa, "VALIDASI ✓"))
            
            # AKTIFKAN TOMBOL PROSES
            self.btn_simpan_kembali.config(state=tk.NORMAL, bg="#1a73e8", fg="white")
        else:
            messagebox.showwarning("Tidak Ditemukan", f"Buku dengan kode '{kode_input}' TIDAK TERDETEKSI sedang dipinjam di sistem.")

    def eksekusi_simpan_pengembalian(self):
        if not self.daftar_kembali_dijepit:
            messagebox.showwarning("Kosong", "Tidak ada antrean buku untuk dikembalikan.")
            return
            
        if messagebox.askyesno("Konfirmasi Pengembalian", f"Proses pengembalian resmi untuk {len(self.daftar_kembali_dijepit)} buku ini?"):
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                tgl_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # Eksekusi Update Status ke SQLite untuk setiap buku di antrean
                for b in self.daftar_kembali_dijepit:
                    cursor.execute(
                        "UPDATE data_peminjaman_buku SET status = 'KEMBALI', tanggal_kembali = ? WHERE id = ?", 
                        (tgl_sekarang, b['id_pk'])
                    )
                
                conn.commit()
                conn.close()
                
                # Jalankan backup otomatis setelah status berhasil diubah
                self.jalankan_auto_backup_silent()
                
                messagebox.showinfo("Pengembalian Sukses", f"Berhasil mengembalikan {len(self.daftar_kembali_dijepit)} buku ke dalam database!")
                
                # Reset antrean pengembalian & muat ulang layar sirkulasi
                self.daftar_kembali_dijepit = []
                self.tampilkan_layaran_sirkulasi()
                
            except Exception as e:
                messagebox.showerror("Error Database", f"Gagal membukukan pengembalian: {e}")

    # =====================================================================
    # INTERFACE MODUL 3: PENCETAKAN KARTU SELEKTIF (CEKLIS DARI SQLITE)
    # =====================================================================
    def tampilkan_layar_setup_cetak_kartu(self):
        self.bersihkan_layar()
        conn = sqlite3.connect(DB_PATH)
        self.df_master_siswa = pd.read_sql_query("SELECT * FROM data_siswa_perpus WHERE nisn IS NOT NULL AND nama IS NOT NULL", conn)
        conn.close()
        
        if self.df_master_siswa.empty:
            messagebox.showerror("Error", "Master data siswa di SQLite kosong!"); self.tampilkan_menu_utama(); return
            
        daftar_rombel = sorted(self.df_master_siswa['rombel'].dropna().unique())
        
        banner = tk.Label(self.root, text="MODUL PENCETAKAN KARTU PERPUS DENGAN FITUR CEKLIS SISWA", font=("Arial", 14, "bold"), bg="#ff9900", fg="white", pady=10); banner.pack(fill=tk.X)
        frame_top = tk.LabelFrame(self.root, text=" Langkah 1: Pilih Rombel Kelas ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_top.pack(fill=tk.X, padx=15, pady=10)
        tk.Label(frame_top, text="Pilih Rombel Target:", font=("Arial", 11), bg="#f4f6f9").pack(side=tk.LEFT, padx=5)
        self.cb_cetak_rombel = ttk.Combobox(frame_top, values=daftar_rombel, state="readonly", font=("Arial", 11), width=20); self.cb_cetak_rombel.pack(side=tk.LEFT, padx=10); self.cb_cetak_rombel.bind("<<ComboboxSelected>>", self.muat_daftar_siswa_dengan_ceklis)
        
        frame_tabel_check = tk.LabelFrame(self.root, text=" Langkah 2: Berikan Ceklis Centang ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_tabel_check.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        self.canvas_scroll = tk.Canvas(frame_tabel_check, bg="white", highlightthickness=0); scrollbar = ttk.Scrollbar(frame_tabel_check, orient="vertical", command=self.canvas_scroll.yview); self.scrollable_frame = tk.Frame(self.canvas_scroll, bg="white")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas_scroll.configure(scrollregion=self.canvas_scroll.bbox("all"))); self.canvas_scroll.create_window((0, 0), window=self.scrollable_frame, anchor="nw"); self.canvas_scroll.configure(yscrollcommand=scrollbar.set); self.canvas_scroll.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        self.lbl_status_check = tk.Label(self.scrollable_frame, text="[Silakan pilih Rombel terlebih dahulu]", font=("Arial", 11, "italic"), fg="#555", bg="white", pady=20); self.lbl_status_check.pack(anchor=tk.W, padx=10)
        
        frame_opsi = tk.Frame(self.root, bg="#f4f6f9", pady=10); frame_opsi.pack(fill=tk.X, padx=15)
        self.btn_mulai_cetak_massal = tk.Button(frame_opsi, text="Mulai Proses Layout & Cetak", font=("Arial", 11, "bold"), bg="#34a853", fg="white", padx=15, pady=8, state=tk.DISABLED, command=self.pemicu_proses_cetak_selektif_background); self.btn_mulai_cetak_massal.pack(side=tk.RIGHT, padx=5)
        tk.Button(frame_opsi, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=8, command=self.tampilkan_menu_utama).pack(side=tk.LEFT)

    def muat_daftar_siswa_dengan_ceklis(self, event=None):
        rombel_target = self.cb_cetak_rombel.get()
        for widget in self.scrollable_frame.winfo_children(): widget.destroy()
        self.dict_check_siswa = {}
        df_filter = self.df_master_siswa[self.df_master_siswa['rombel'] == rombel_target]
        if df_filter.empty: self.btn_mulai_cetak_massal.config(state=tk.DISABLED); return
        
        f_head = tk.Frame(self.scrollable_frame, bg="#e8eaed", pady=5); f_head.pack(fill=tk.X, expand=True)
        tk.Label(f_head, text=" Status Cetak ", font=("Arial", 10, "bold"), bg="#e8eaed", width=15).pack(side=tk.LEFT)
        tk.Label(f_head, text=" NISN ", font=("Arial", 10, "bold"), bg="#e8eaed", width=15, anchor=tk.W).pack(side=tk.LEFT)
        tk.Label(f_head, text=" Nama Lengkap Siswa ", font=("Arial", 10, "bold"), bg="#e8eaed", width=50, anchor=tk.W).pack(side=tk.LEFT)
        
        f_ctrl = tk.Frame(self.scrollable_frame, bg="white", pady=5); f_ctrl.pack(fill=tk.X)
        def set_semua_centang(status):
            for var in self.dict_check_siswa.values(): var.set(status)
        tk.Button(f_ctrl, text="✓ Centang Semua", font=("Arial", 9), bg="#e8eaed", command=lambda: set_semua_centang(True)).pack(side=tk.LEFT, padx=5)
        tk.Button(f_ctrl, text="✕ Hapus Semua Centang", font=("Arial", 9), bg="#e8eaed", command=lambda: set_semua_centang(False)).pack(side=tk.LEFT, padx=5)
        
        for _, row in df_filter.iterrows():
            nisn = str(row['nisn']).strip(); nama = str(row['nama']).upper(); jurusan = str(row['jurusan']).strip(); rombel = str(row['rombel']).strip()
            f_row = tk.Frame(self.scrollable_frame, bg="white", pady=3); f_row.pack(fill=tk.X, expand=True)
            var_check = tk.BooleanVar(value=True); self.dict_check_siswa[nisn] = var_check
            chk = tk.Checkbutton(f_row, variable=var_check, bg="white", activebackground="white"); chk.pack(side=tk.LEFT, padx=35)
            tk.Label(f_row, text=nisn, font=("Consolas", 10), bg="white", width=15, anchor=tk.W).pack(side=tk.LEFT)
            tk.Label(f_row, text=nama, font=("Arial", 10), bg="white", width=50, anchor=tk.W).pack(side=tk.LEFT)
            var_check.metadata = {"NISN": nisn, "Nama": nama, "Jurusan": jurusan, "Rombel": rombel}
        self.btn_mulai_cetak_massal.config(state=tk.NORMAL)

    def pemicu_proses_cetak_selektif_background(self):
        self.siswa_terfilter_cetak = []
        for nisn, var in self.dict_check_siswa.items():
            if var.get() == True: self.siswa_terfilter_cetak.append(var.metadata)
        if not self.siswa_terfilter_cetak: messagebox.showwarning("Kosong", "Tidak ada siswa dipilih!"); return
        
        pilihan_kertas = tk.StringVar(value="A3+"); dialog = tk.Toplevel(self.root); dialog.title("Pilih Kertas"); dialog.geometry("350x180"); dialog.configure(bg="#f4f6f9"); dialog.transient(self.root); dialog.grab_set()
        tk.Label(dialog, text="Pilih Ukuran Kertas Cetak Kartu:", font=("Arial", 11, "bold"), bg="#f4f6f9").pack(pady=10)
        frame_radio = tk.Frame(dialog, bg="#f4f6f9"); frame_radio.pack(pady=5)
        tk.Radiobutton(frame_radio, text="A3+ Percetakan (320x480 mm)", variable=pilihan_kertas, value="A3+", font=("Arial", 10), bg="#f4f6f9").pack(anchor=tk.W)
        tk.Radiobutton(frame_radio, text="A4 Desktop Sekolah (297x210 mm)", variable=pilihan_kertas, value="A4", font=("Arial", 10), bg="#f4f6f9").pack(anchor=tk.W)
        def konfirmasi():
            self.kertas_terpilih = pilihan_kertas.get(); dialog.destroy()
            self.buat_layar_monitoring(f"PROSES CETAK KARTU SELEKTIF ({self.kertas_terpilih})", self.eksekusi_back_end_cetak_kartu_selektif)
        tk.Button(dialog, text="Mulai Cetak", font=("Arial", 10, "bold"), bg="#1a73e8", fg="white", command=konfirmasi).pack(pady=15)

    def eksekusi_back_end_cetak_kartu_selektif(self):
        jenis_kertas = self.kertas_terpilih; output_pdf = os.path.join(DIR_PERPUS, f"Cetak_Selektif_Kartu_{jenis_kertas}.pdf"); output_mal = os.path.join(DIR_PERPUS, f"Mal_Selektif_Kartu_{jenis_kertas}.pdf")
        if not os.path.exists(TEMPLATE_KARTU): return
        try:
            total_data = len(self.siswa_terfilter_cetak)
            bleed_size = 3 * MM_TO_POINT; card_w, card_h = (85 + 6) * MM_TO_POINT, (54 + 6) * MM_TO_POINT
            if jenis_kertas == "A3+":
                paper_width, paper_height = 320 * MM_TO_POINT, 480 * MM_TO_POINT; max_cols, max_rows = 3, 7; gap_x, gap_y = 4 * MM_TO_POINT, 4 * MM_TO_POINT; margin_left, margin_top = 15 * MM_TO_POINT, 18 * MM_TO_POINT
            else:
                paper_width, paper_height = 297 * MM_TO_POINT, 210 * MM_TO_POINT; max_cols, max_rows = 3, 3; gap_x, gap_y = 4 * MM_TO_POINT, 6 * MM_TO_POINT; margin_left, margin_top = 12 * MM_TO_POINT, 15 * MM_TO_POINT
            c_design = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height)); c_mal = canvas.Canvas(output_mal, pagesize=(paper_width, paper_height))
            try: font_nama = ImageFont.truetype(FONT_CHAU, 32); font_data = ImageFont.truetype("arialbd.ttf", 30)
            except: font_nama = font_data = ImageFont.load_default()
            
            col_idx, row_idx, total = 0, 0, 0
            for siswa in self.siswa_terfilter_cetak:
                nama_siswa = siswa['Nama']; nisn = siswa['NISN']; jurusan = siswa['Jurusan']; rombel = siswa['Rombel']
                card_img = Image.open(TEMPLATE_KARTU).convert("RGBA"); w_orig, h_orig = card_img.size; px_bleed_x, px_bleed_y = int((3 / 85) * w_orig), int((3 / 54) * h_orig)
                card_bleed_img = Image.new("RGBA", (w_orig + (px_bleed_x * 2), h_orig + (px_bleed_y * 2)), (255, 255, 255, 255)); card_bleed_img.paste(card_img, (px_bleed_x, px_bleed_y))
                draw = ImageDraw.Draw(card_bleed_img); draw.text((310 + px_bleed_x, 282 + px_bleed_y), nama_siswa, font=font_nama, fill=(0, 0, 0)); draw.text((310 + px_bleed_x, 332 + px_bleed_y), nisn, font=font_data, fill=(0, 0, 0)); draw.text((310 + px_bleed_x, 377 + px_bleed_y), jurusan, font=font_data, fill=(0, 0, 0)); draw.text((310 + px_bleed_x, 422 + px_bleed_y), rombel, font=font_data, fill=(0, 0, 0))
                try:
                    encoded = urllib.parse.quote(nisn); req = urllib.request.Request(f"https://quickchart.io/qr?text={encoded}&size=200&margin=1", headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=5) as response:
                        qr_img = Image.open(BytesIO(response.read())).convert("RGBA").resize((175, 175)); card_bleed_img.alpha_composite(qr_img, (740 + px_bleed_x, 280 + px_bleed_y))
                    self.log_ke_gui(self.txt_log, f"✓ QR Generated: {nama_siswa}")
                except: pass
                img_buffer = BytesIO(); card_bleed_img.convert("RGB").save(img_buffer, format="JPEG", quality=95); img_buffer.seek(0)
                x = margin_left + (col_idx * (card_w + gap_x)); y = paper_height - margin_top - card_h - (row_idx * (card_h + gap_y))
                c_design.drawImage(ImageReader(img_buffer), x, y, width=card_w, height=card_h); self.gambar_crop_mark(c_design, x + bleed_size, y + bleed_size, 85 * MM_TO_POINT, 54 * MM_TO_POINT, bleed_size)
                c_mal.setStrokeColorRGB(1.0, 0.0, 1.0); c_mal.setLineWidth(0.3); c_mal.rect(x + bleed_size, y + bleed_size, 85 * MM_TO_POINT, 54 * MM_TO_POINT, stroke=1, fill=0)
                total += 1; col_idx += 1; persentase = int((total / total_data) * 100); self.progress_bar['value'] = persentase
                if col_idx >= max_cols: col_idx = 0; row_idx += 1
                if row_idx >= max_rows: c_design.showPage(); c_mal.showPage(); col_idx, row_idx = 0, 0
            c_design.save(); c_mal.save(); messagebox.showinfo("Sukses", "PDF Berhasil dibuat."); os.startfile(output_pdf)
        except Exception as e: messagebox.showerror("Error", f"{e}")
        finally: self.btn_kembali_menu.config(state=tk.NORMAL, bg="#1a73e8")

    # =====================================================================
    # MODUL 4: REPRINT KOLEKTIF DARI NISN (MAKS 9 SISWA DI A4) - FIX UTUH
    # =====================================================================
    def tampilkan_layar_cetak_satuan(self):
        self.bersihkan_layar(); self.antrean_reprint = []
        conn = sqlite3.connect(DB_PATH)
        self.df_siswa = pd.read_sql_query("SELECT * FROM data_siswa_perpus", conn)
        conn.close()
        
        if self.df_siswa.empty: return
        self.df_siswa['nisn'] = self.df_siswa['nisn'].str.strip()
        
        banner = tk.Label(self.root, text="MODUL CETAK ULANG KARTU (REPRINT KOLEKTIF MAKS 9 SISWA)", font=("Arial", 14, "bold"), bg="#6f42c1", fg="white", pady=10); banner.pack(fill=tk.X)
        frame_cari = tk.LabelFrame(self.root, text=" Langkah 1: Cari & Tambah Siswa Kehilangan Kartu ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=15, pady=10); frame_cari.pack(fill=tk.X, padx=20, pady=10)
        
        tk.Label(frame_cari, text="Scan / Input NISN:", font=("Arial", 11), bg="#f4f6f9").pack(side=tk.LEFT, padx=5)
        self.ent_cari_nisn = tk.Entry(frame_cari, font=("Arial", 12), width=25); self.ent_cari_nisn.pack(side=tk.LEFT, padx=5); self.ent_cari_nisn.bind("<Return>", self.proses_cari_siswa_satuan)
        tk.Button(frame_cari, text="Cari & Tambah", font=("Arial", 10, "bold"), bg="#1a73e8", fg="white", command=self.proses_cari_siswa_satuan).pack(side=tk.LEFT, padx=10)
        
        self.lbl_kuota = tk.Label(frame_cari, text="Kuota Terisi: 0 / 9 Siswa", font=("Arial", 10, "bold"), bg="#f4f6f9", fg="#6f42c1"); self.lbl_kuota.pack(side=tk.RIGHT, padx=10)
        
        frame_tabel = tk.LabelFrame(self.root, text=" Langkah 2: Daftar Antrean Kartu Siswa ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=15, pady=10); frame_tabel.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        self.tree_reprint = ttk.Treeview(frame_tabel, columns=("No", "NISN", "Nama Siswa", "Rombel"), show="headings", height=8); self.tree_reprint.heading("No", text="No"); self.tree_reprint.heading("NISN", text="NISN"); self.tree_reprint.heading("Nama Siswa", text="Nama Siswa"); self.tree_reprint.heading("Rombel", text="Rombel")
        self.tree_reprint.column("No", width=50, anchor=tk.CENTER); self.tree_reprint.column("NISN", width=120, anchor=tk.CENTER); self.tree_reprint.column("Nama Siswa", width=400, anchor=tk.W); self.tree_reprint.column("Rombel", width=120, anchor=tk.CENTER); self.tree_reprint.pack(fill=tk.BOTH, expand=True, pady=5)
        
        frame_aksi = tk.Frame(self.root, bg="#f4f6f9", pady=15); frame_aksi.pack(fill=tk.X, padx=20)
        self.btn_eksekusi_satuan = tk.Button(frame_aksi, text="Cetak Lembar Kartu Kolektif (A4 Landscape)", font=("Arial", 11, "bold"), bg="#34a853", fg="white", padx=20, pady=10, state=tk.DISABLED, command=self.proses_pdf_kartu_satuan); self.btn_eksekusi_satuan.pack(side=tk.RIGHT)
        tk.Button(frame_aksi, text="Bersihkan Antrean", font=("Arial", 11), bg="#5f6368", fg="white", padx=15, pady=10, command=self.reset_antrean_reprint).pack(side=tk.RIGHT, padx=10)
        tk.Button(frame_aksi, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=10, command=self.tampilkan_menu_utama).pack(side=tk.LEFT)

    def proses_cari_siswa_satuan(self, event=None):
        nisn_input = self.ent_cari_nisn.get().strip(); self.ent_cari_nisn.delete(0, tk.END)
        if not nisn_input: return
        if len(self.antrean_reprint) >= 9: return
        if any(s['NISN'] == nisn_input for s in self.antrean_reprint): return
        
        hasil = self.df_siswa[self.df_siswa['nisn'] == nisn_input]
        if not hasil.empty:  
            row = hasil.iloc[0]; siswa = {"NISN": row['nisn'], "Nama": row['nama'].upper(), "Jurusan": row['jurusan'], "Rombel": row['rombel']}
            self.antrean_reprint.append(siswa); no_urut = len(self.antrean_reprint); self.tree_reprint.insert("", tk.END, values=(no_urut, siswa['NISN'], siswa['Nama'], siswa['Rombel']))
            self.lbl_kuota.config(text=f"Kuota Terisi: {no_urut} / 9 Siswa"); self.btn_eksekusi_satuan.config(state=tk.NORMAL)

    def reset_antrean_reprint(self):
        self.antrean_reprint = []; self.lbl_kuota.config(text="Kuota Terisi: 0 / 9 Siswa"); self.btn_eksekusi_satuan.config(state=tk.DISABLED)
        for item in self.tree_reprint.get_children(): self.tree_reprint.delete(item)

    def proses_pdf_kartu_satuan(self):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S"); output_pdf = os.path.join(DIR_PERPUS, f"Reprint_Kolektif_{timestamp}.pdf")
        if not os.path.exists(TEMPLATE_KARTU): return
        try:
            paper_width, paper_height = 297 * MM_TO_POINT, 210 * MM_TO_POINT; bleed_size = 3 * MM_TO_POINT; card_w, card_h = (85 + 6) * MM_TO_POINT, (54 + 6) * MM_TO_POINT; gap_x, gap_y = 4 * MM_TO_POINT, 6 * MM_TO_POINT; margin_left, margin_top = 12 * MM_TO_POINT, 15 * MM_TO_POINT
            c = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
            
            try: 
                font_nama = ImageFont.truetype(FONT_CHAU, 32)
                font_data = ImageFont.truetype("arialbd.ttf", 30)
            except: font_nama = font_data = ImageFont.load_default()
            
            col_idx, row_idx = 0, 0
            for siswa in self.antrean_reprint:
                card_img = Image.open(TEMPLATE_KARTU).convert("RGBA"); w_orig, h_orig = card_img.size; px_bleed_x, px_bleed_y = int((3 / 85) * w_orig), int((3 / 54) * h_orig)
                card_bleed_img = Image.new("RGBA", (w_orig + (px_bleed_x * 2), h_orig + (px_bleed_y * 2)), (255, 255, 255, 255)); card_bleed_img.paste(card_img, (px_bleed_x, px_bleed_y))
                
                draw = ImageDraw.Draw(card_bleed_img)
                draw.text((310 + px_bleed_x, 282 + px_bleed_y), siswa['Nama'], font=font_nama, fill=(0, 0, 0))
                draw.text((310 + px_bleed_x, 332 + px_bleed_y), siswa['NISN'], font=font_data, fill=(0, 0, 0))
                draw.text((310 + px_bleed_x, 377 + px_bleed_y), siswa['Jurusan'], font=font_data, fill=(0, 0, 0))
                draw.text((310 + px_bleed_x, 422 + px_bleed_y), siswa['Rombel'], font=font_data, fill=(0, 0, 0))
                
                try:
                    encoded = urllib.parse.quote(siswa['NISN']); req = urllib.request.Request(f"https://quickchart.io/qr?text={encoded}&size=200&margin=1", headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=5) as response: qr_img = Image.open(BytesIO(response.read())).convert("RGBA").resize((175, 175)); card_bleed_img.alpha_composite(qr_img, (740 + px_bleed_x, 280 + px_bleed_y))
                except: pass
                
                img_buffer = BytesIO(); card_bleed_img.convert("RGB").save(img_buffer, format="JPEG", quality=95); img_buffer.seek(0)
                x = margin_left + (col_idx * (card_w + gap_x)); y = paper_height - margin_top - card_h - (row_idx * (card_h + gap_y))
                
                c.drawImage(ImageReader(img_buffer), x, y, width=card_w, height=card_h); self.gambar_crop_mark(c, x + bleed_size, y + bleed_size, 85 * MM_TO_POINT, 54 * MM_TO_POINT, bleed_size)
                col_idx += 1
                if col_idx >= 3: col_idx = 0; row_idx += 1
                
            c.save(); messagebox.showinfo("Reprint Sukses", "Berhasil menyusun kartu siswa ke lembar A4.")
            os.startfile(output_pdf); self.reset_antrean_reprint()
        except Exception as e: messagebox.showerror("Error", f"{e}")

    # =====================================================================
    # HALAMAN 5: MAHA DASHBOARD LIVE METRIK KEDUA DATABASE (SQLITE ENGINE)
    # =====================================================================
    def tampilkan_layar_rekap(self):
        self.bersihkan_layar()
        
        banner = tk.Label(self.root, text="MODUL REKAP SIRKULASI & LIVE ANALYTICS MASTER BUKU", font=("Arial", 14, "bold"), bg="#007afc", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        # --- CLEANUP RIGID: RAPIKAN SPASI & HAPUS DUPLIKAT TERHUBUNG ---
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
        # 1. Rapikan spasi pada kolom kunci
        cursor.execute("UPDATE data_peminjaman_buku SET nisn = TRIM(nisn), kode_satuan_buku = TRIM(kode_satuan_buku), status = TRIM(status)")
        
        # 2. Hapus entri duplikat persis berbasis (nisn, kode_satuan_buku, status)
        cursor.execute('''
            DELETE FROM data_peminjaman_buku 
            WHERE rowid NOT IN (
                SELECT MAX(rowid) 
                FROM data_peminjaman_buku 
                GROUP BY nisn, kode_satuan_buku, UPPER(status)
            )
        ''')
        conn.commit()
        
        self.df_rekap_log = pd.read_sql_query("SELECT * FROM data_peminjaman_buku", conn)
        self.df_master_buku = pd.read_sql_query("SELECT * FROM qr_id_buku", conn)
        self.df_master_umum = pd.read_sql_query("SELECT * FROM data_buku_umum", conn)
        conn.close()
        
        total_judul_10 = total_judul_11 = total_judul_12 = total_judul_umum = 0
        total_eks_10 = total_eks_11 = total_eks_12 = total_eks_umum = total_eks_semua = 0
        total_terpinjam = total_tersedia = 0
        
        if not self.df_master_buku.empty:
            total_eks_semua += len(self.df_master_buku)
            set_judul_10 = set(); set_judul_11 = set(); set_judul_12 = set()
            for _, r in self.df_master_buku.iterrows():
                kode_satuan = str(r['label_buku']).strip()
                judul_real = self.dapatkan_judul_buku(kode_satuan).strip().upper()
                if     kode_satuan.startswith("10"): total_eks_10 += 1; set_judul_10.add(judul_real)
                elif   kode_satuan.startswith("11"): total_eks_11 += 1; set_judul_11.add(judul_real)
                elif   kode_satuan.startswith("12"): total_eks_12 += 1; set_judul_12.add(judul_real)
            total_judul_10 = len(set_judul_10); total_judul_11 = len(set_judul_11); total_judul_12 = len(set_judul_12)

        dict_agregat_umum = {}
        if not self.df_rekap_log.empty:
            df_isbn_log = self.df_rekap_log[self.df_rekap_log['kode_satuan_buku'].str.contains("-") & 
                                            (self.df_rekap_log['kode_satuan_buku'].str.startswith("978") | 
                                             self.df_rekap_log['kode_satuan_buku'].str.startswith("979"))]
            if not df_isbn_log.empty:
                ids_fiksi_unik = df_isbn_log['kode_satuan_buku'].unique()
                total_eks_umum = len(ids_fiksi_unik)
                total_eks_semua += total_eks_umum
                for id_f in ids_fiksi_unik:
                    j_title = self.dapatkan_judul_buku(id_f).strip().upper()
                    dict_agregat_umum[j_title] = dict_agregat_umum.get(j_title, 0) + 1
                total_judul_umum = len(dict_agregat_umum)

        if not self.df_rekap_log.empty:
            df_active_pinjam = self.df_rekap_log[self.df_rekap_log['status'].str.upper() == 'DIPINJAM']
            total_terpinjam = len(df_active_pinjam['kode_satuan_buku'].unique())
            
        total_tersedia = total_eks_semua - total_terpinjam
        if total_tersedia < 0: total_tersedia = 0

        # --- CARDS DASHBOARD ---
        frame_cards = tk.Frame(self.root, bg="#f4f6f9", pady=5)
        frame_cards.pack(fill=tk.X, padx=15, pady=5)
        
        c1 = tk.Frame(frame_cards, bg="#eef4ff", bd=1, relief=tk.SOLID, padx=8, pady=5); c1.grid(row=0, column=0, padx=3, sticky="nsew")
        tk.Label(c1, text="BUKU KELAS 10", font=("Arial", 9, "bold"), bg="#eef4ff", fg="#1a73e8").pack()
        tk.Label(c1, text=f"{total_judul_10} Judul\n{total_eks_10} Eks", font=("Arial", 11, "bold"), bg="#eef4ff", fg="#202124").pack(pady=3)
        
        c2 = tk.Frame(frame_cards, bg="#efffaf", bd=1, relief=tk.SOLID, padx=8, pady=5); c2.grid(row=0, column=1, padx=3, sticky="nsew")
        tk.Label(c2, text="BUKU KELAS 11", font=("Arial", 9, "bold"), bg="#efffaf", fg="#188038").pack()
        tk.Label(c2, text=f"{total_judul_11} Judul\n{total_eks_11} Eks", font=("Arial", 11, "bold"), bg="#efffaf", fg="#202124").pack(pady=3)

        c3 = tk.Frame(frame_cards, bg="#fff0f0", bd=1, relief=tk.SOLID, padx=8, pady=5); c3.grid(row=0, column=2, padx=3, sticky="nsew")
        tk.Label(c3, text="BUKU KELAS 12", font=("Arial", 9, "bold"), bg="#fff0f0", fg="#c5221f").pack()
        tk.Label(c3, text=f"{total_judul_12} Judul\n{total_eks_12} Eks", font=("Arial", 11, "bold"), bg="#fff0f0", fg="#202124").pack(pady=3)

        c4 = tk.Frame(frame_cards, bg="#fef3d6", bd=1, relief=tk.SOLID, padx=8, pady=5); c4.grid(row=0, column=3, padx=3, sticky="nsew")
        tk.Label(c4, text="NON-PAKET (FIKSI)", font=("Arial", 9, "bold"), bg="#fef3d6", fg="#b06000").pack()
        tk.Label(c4, text=f"{total_judul_umum} Judul\n{total_eks_umum} Eks", font=("Arial", 11, "bold"), bg="#fef3d6", fg="#202124").pack(pady=3)

        c5 = tk.Frame(frame_cards, bg="#e6fffa", bd=1, relief=tk.SOLID, padx=8, pady=5); c5.grid(row=0, column=4, padx=3, sticky="nsew")
        tk.Label(c5, text="LIVE STATUS RAK", font=("Arial", 9, "bold"), bg="#e6fffa", fg="#008577").pack()
        tk.Label(c5, text=f"Tersedia: {total_tersedia} E\nPinjam: {total_terpinjam} E", font=("Arial", 10, "bold"), bg="#e6fffa", fg="#202124").pack(pady=3)

        c6 = tk.Frame(frame_cards, bg="#343a40", bd=1, relief=tk.SOLID, padx=8, pady=5); c6.grid(row=0, column=5, padx=3, sticky="nsew")
        tk.Label(c6, text="TOTAL KESELURUH", font=("Arial", 9, "bold"), bg="#343a40", fg="white").pack()
        tk.Label(c6, text=f"{total_eks_semua}\nEksemplar", font=("Arial", 11, "bold"), bg="#343a40", fg="#00ff00").pack(pady=3)
        frame_cards.columnconfigure((0,1,2,3,4,5), weight=1)

        tab_lap = ttk.Notebook(self.root)
        tab_per_rombel = tk.Frame(tab_lap, bg="#f4f6f9")
        tab_stok_judul = tk.Frame(tab_lap, bg="#f4f6f9")
        tab_lap.add(tab_per_rombel, text="  [1] TANGGUNGAN PINJAM SISWA PER ROMBEL  ")
        tab_lap.add(tab_stok_judul, text="  [2] DAFTAR JUMLAH EKSEMPLAR PER JUDUL BUKU (ALL DATA)  ")
        tab_lap.pack(expand=1, fill="both", padx=15, pady=5)

        daftar_rombel = sorted(self.df_rekap_log['rombel'].dropna().unique()) if not self.df_rekap_log.empty else []
        frame_filter = tk.Frame(tab_per_rombel, bg="#f4f6f9", pady=5); frame_filter.pack(fill=tk.X)
        tk.Label(frame_filter, text="Pilih Rombel Kelas:", font=("Arial", 10, "bold"), bg="#f4f6f9").pack(side=tk.LEFT, padx=5)
        self.cb_rombel = ttk.Combobox(frame_filter, values=daftar_rombel, state="readonly", font=("Arial", 10), width=18); self.cb_rombel.pack(side=tk.LEFT, padx=5); self.cb_rombel.bind("<<ComboboxSelected>>", self.proses_tampilkan_rekap_rombel)
        
        self.tree_rombel = ttk.Treeview(tab_per_rombel, columns=("NISN", "Nama", "Jumlah"), show="headings", height=8)
        self.tree_rombel.heading("NISN", text="NISN"); self.tree_rombel.heading("Nama", text="Nama Lengkap Siswa"); self.tree_rombel.heading("Jumlah", text="Jumlah Buku Di Tangan")
        self.tree_rombel.column("NISN", width=120, anchor=tk.CENTER); self.tree_rombel.column("Nama", width=420, anchor=tk.W); self.tree_rombel.column("Jumlah", width=200, anchor=tk.CENTER); self.tree_rombel.pack(fill=tk.BOTH, expand=True, pady=5)

        tree_stok = ttk.Treeview(tab_stok_judul, columns=("No", "Jenis", "Judul", "TotalEks"), show="headings", height=9)
        tree_stok.heading("No", text="No"); tree_stok.heading("Jenis", text="Kategori"); tree_stok.heading("Judul", text="Judul Buku Paket / Fiksi Master"); tree_stok.heading("TotalEks", text="Stok Eksemplar Terdaftar")
        tree_stok.column("No", width=40, anchor=tk.CENTER); tree_stok.column("Jenis", width=100, anchor=tk.CENTER); tree_stok.column("Judul", width=450, anchor=tk.W); tree_stok.column("TotalEks", width=150, anchor=tk.CENTER); tree_stok.pack(fill=tk.BOTH, expand=True, pady=5)
        
        dict_final_katalog = {}
        if not self.df_master_buku.empty:
            for _, r in self.df_master_buku.iterrows():
                j_p = self.dapatkan_judul_buku(str(r['label_buku'])).strip().upper()
                dict_final_katalog[j_p] = dict_final_katalog.get(j_p, {'Kategori': 'BUKU PAKET', 'Stok': 0})
                dict_final_katalog[j_p]['Stok'] += 1
        for j_u, val_stok in dict_agregat_umum.items():
            dict_final_katalog[j_u] = {'Kategori': 'NON-PAKET', 'Stok': val_stok}
            
        idx_stok = 1
        for key_j, data_b in sorted(dict_final_katalog.items()):
            tree_stok.insert("", tk.END, values=(idx_stok, data_b['Kategori'], key_j, f"{data_b['Stok']} Eksemplar"))
            idx_stok += 1

        frame_aksi = tk.Frame(self.root, bg="#f4f6f9", pady=5)
        frame_aksi.pack(fill=tk.X, padx=15)
        self.btn_lihat_detil = tk.Button(frame_aksi, text="Tampilkan Detil Daftar Buku", font=("Arial", 11, "bold"), bg="#ff9900", fg="white", padx=15, pady=6, state=tk.DISABLED, command=self.pop_up_jendela_detil_buku)
        self.btn_lihat_detil.pack(side=tk.RIGHT, padx=5)
        tk.Button(frame_aksi, text="Ekspor Laporan Resmi (Excel)", font=("Arial", 11, "bold"), bg="#188038", fg="white", padx=15, pady=6, command=self.proses_ekspor_laporan_excel).pack(side=tk.RIGHT, padx=5)
        tk.Button(frame_aksi, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=6, command=self.tampilkan_menu_utama).pack(side=tk.LEFT)

    def bersihkan_nisn_ke_string(self, val):
        """Fungsi pembantu untuk menormalisasi NISN (menghapus desimal, spasi, dan leading zero agar matching 100% akurat)"""
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip().split('.')[0] # Hapus desimal .0 jika ada
        s = s.lstrip('0') # Hapus nol di depan untuk perbandingan murni
        return s

    def proses_tampilkan_rekap_rombel(self, event=None):
        rombel_pilihan = self.cb_rombel.get()
        for item in self.tree_rombel.get_children(): 
            self.tree_rombel.delete(item)
            
        if self.df_rekap_log.empty: return
        
        # 1. Filter dasar: Rombel & Status DIPINJAM
        df_filtered = self.df_rekap_log[(self.df_rekap_log['rombel'].astype(str).str.strip() == rombel_pilihan) & 
                                        (self.df_rekap_log['status'].astype(str).str.strip().str.upper() == 'DIPINJAM')].copy()
        
        if df_filtered.empty: 
            self.btn_lihat_detil.config(state=tk.DISABLED)
            messagebox.showinfo("Bersih", f"Kelas {rombel_pilihan} tidak memiliki tanggungan peminjaman.")
            return
        
        # Normalisasi NISN secara seragam
        df_filtered['nisn_clean'] = df_filtered['nisn'].apply(self.bersihkan_nisn_ke_string)
        
        # 2. Dapatkan Judul Buku Bersih
        df_filtered['Judul_Buku'] = df_filtered['kode_satuan_buku'].apply(lambda x: self.dapatkan_judul_buku(str(x)).strip().upper())
        
        # 3. FILTER LOGIS RUMPUN KEJURUAN (TEKNO vs BISMEN)
        def is_valid_buku_rombel(row):
            kode = str(row['kode_satuan_buku']).strip().upper()
            rombel = str(row['rombel']).strip().upper()
            
            # Filter Buku Fiksi ISBN (978/979)
            if kode.startswith("978") or kode.startswith("979"):
                return False

            # KLASIFIKASI RUMPUN TEKNO (TJKT / TKJ)
            if "TJKT" in rombel or "TKJ" in rombel:
                # Siswa TJKT HANYA BISA memegang buku TJKT/TEKNO/UMUM, BLOKIR BISMEN, MPLB, AKL, BD, MP
                if any(x in kode for x in ["BISMEN", "MPLB", "AKL", "BD", "MP"]): 
                    return False

            # KLASIFIKASI RUMPUN BISMEN (AKL / MPLB)
            elif "AKL" in rombel or "MPLB" in rombel:
                # Siswa BISMEN HANYA BISA memegang buku BISMEN/AKL/MPLB/UMUM, BLOKIR TEKNO, TJKT, TKJ
                if any(x in kode for x in ["TEKNO", "TJKT", "TKJ"]): 
                    return False
                
                # Spesifik internal Bisman (MPLB vs AKL)
                if "MPLB" in rombel and "AKL" in kode: return False
                if "AKL" in rombel and "MPLB" in kode: return False
                
            return True

        df_filtered = df_filtered[df_filtered.apply(is_valid_buku_rombel, axis=1)]

        # 4. Filter Judul Unik per Siswa (Mencegah double count untuk matapelajaran/judul yang sama)
        df_clean = df_filtered.drop_duplicates(subset=['nisn_clean', 'Judul_Buku']).copy()
        
        self.df_rombel_aktif_clean = df_clean
        
        # 5. Hitung Jumlah Buku Bersih per Siswa
        df_agregat = df_clean.groupby(['nisn_clean', 'nisn', 'nama_siswa'])['Judul_Buku'].count().reset_index(name='Jumlah')
        
        for _, r in df_agregat.iterrows(): 
            nisn_tampil = str(r['nisn']).split('.')[0].strip()
            self.tree_rombel.insert("", tk.END, values=(nisn_tampil, r['nama_siswa'], f"{r['Jumlah']} Buku"))
            
        self.btn_lihat_detil.config(state=tk.NORMAL)

    def pop_up_jendela_detil_buku(self):
        selected_item = self.tree_rombel.selection()
        if not selected_item:
            messagebox.showwarning("Pilih Siswa", "Silakan pilih salah satu siswa dari tabel terlebih dahulu.")
            return
            
        item_data = self.tree_rombel.item(selected_item[0])['values']
        raw_nisn_siswa = item_data[0]
        nisn_clean_target = self.bersihkan_nisn_ke_string(raw_nisn_siswa)
        nama_siswa = str(item_data[1])
        
        if self.df_rekap_log.empty:
            messagebox.showwarning("Kosong", "Data log transaksi tidak tersedia.")
            return

        # Ambil data transaksi aktif siswa terpilih
        df_log_siswa = self.df_rekap_log.copy()
        df_log_siswa['nisn_clean'] = df_log_siswa['nisn'].apply(self.bersihkan_nisn_ke_string)
        
        df_detil = df_log_siswa[(df_log_siswa['nisn_clean'] == nisn_clean_target) & 
                                (df_log_siswa['status'].astype(str).str.strip().str.upper() == 'DIPINJAM')].copy()
        
        if not df_detil.empty:
            # 1. Dapatkan Judul Buku Bersih
            df_detil['Judul_Buku'] = df_detil['kode_satuan_buku'].apply(lambda x: self.dapatkan_judul_buku(str(x)).strip().upper())
            
            # 2. FILTER RUMPUN SANGAT KETAT UNTUK JENDELA DETIL
            rombel_siswa = str(df_detil['rombel'].iloc[0]).upper()
            
            # Filter Buku Fiksi ISBN (978/979) dari daftar paket
            df_detil = df_detil[~df_detil['kode_satuan_buku'].astype(str).str.upper().str.startswith(("978", "979"))]
            
            # A. SISWA RUMPUN TEKNO (TJKT / TKJ)
            if "TJKT" in rombel_siswa or "TKJ" in rombel_siswa:
                # Blokir total buku ber-kode BISMEN, MPLB, AKL, BD, MP
                df_detil = df_detil[~df_detil['kode_satuan_buku'].astype(str).str.upper().str.contains("BISMEN|MPLB|AKL|BD|MP", regex=True)]
                
            # B. SISWA RUMPUN BISMEN (MPLB / AKL)
            elif "MPLB" in rombel_siswa or "AKL" in rombel_siswa:
                # Blokir total buku ber-kode TEKNO, TJKT, TKJ
                df_detil = df_detil[~df_detil['kode_satuan_buku'].astype(str).str.upper().str.contains("TEKNO|TJKT|TKJ", regex=True)]
                
                # Spesifik Lintas Jurusan Bisman
                if "MPLB" in rombel_siswa:
                    df_detil = df_detil[~df_detil['kode_satuan_buku'].astype(str).str.upper().str.contains("AKL", regex=True)]
                elif "AKL" in rombel_siswa:
                    df_detil = df_detil[~df_detil['kode_satuan_buku'].astype(str).str.upper().str.contains("MPLB", regex=True)]

            # 3. Hilangkan duplikasi berdasarkan judul buku unik
            df_detil = df_detil.drop_duplicates(subset=['Judul_Buku'])

        top = tk.Toplevel(self.root)
        top.title(f"Detil Tanggungan Buku - {nama_siswa}")
        top.geometry("720x420")
        top.grab_set()
        
        tk.Label(top, text=f"DAFTAR BUKU DIPINJAM: {nama_siswa} (NISN: {raw_nisn_siswa})", font=("Arial", 11, "bold"), bg="#1a73e8", fg="white", pady=8).pack(fill=tk.X)
        
        tree_detil = ttk.Treeview(top, columns=("No", "Kode", "Judul", "TglPinjam"), show="headings")
        tree_detil.heading("No", text="No")
        tree_detil.heading("Kode", text="Kode Eksemplar")
        tree_detil.heading("Judul", text="Judul Buku / Mata Pelajaran")
        tree_detil.heading("TglPinjam", text="Tanggal Pinjam")
        
        tree_detil.column("No", width=40, anchor=tk.CENTER)
        tree_detil.column("Kode", width=150, anchor=tk.CENTER)
        tree_detil.column("Judul", width=360, anchor=tk.W)
        tree_detil.column("TglPinjam", width=120, anchor=tk.CENTER)
        tree_detil.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        if df_detil.empty:
            tree_detil.insert("", tk.END, values=("-", "-", "TIDAK ADA TANGGUNGAN BUKU AKTIF", "-"))
        else:
            no = 1
            for _, row in df_detil.iterrows():
                tgl = str(row.get('tanggal_pinjam', '-'))[:10]
                tree_detil.insert("", tk.END, values=(no, row['kode_satuan_buku'], row['Judul_Buku'], tgl))
                no += 1
            
        tk.Button(top, text="Tutup", font=("Arial", 10, "bold"), bg="#d93025", fg="white", command=top.destroy, padx=20, pady=5).pack(pady=10)

    # =====================================================================
    # UTILITAS PROSES BACKGROUND MASSAL LABEL KISSCUT A3+
    # =====================================================================
    def tampilkan_layar_proses_cetak_buku(self):
        self.buat_layar_monitoring("PROSES MASAL CETAK LABEL BUKU A3+", self.eksekusi_back_end_cetak_buku)

    def buat_layar_monitoring(self, judul_modul, target_fungsi):
        self.bersihkan_layar()
        banner = tk.Label(self.root, text=judul_modul, font=("Arial", 14, "bold"), bg="#202124", fg="white", pady=10)
        banner.pack(fill=tk.X)
        frame_progress = tk.Frame(self.root, bg="#f4f6f9", pady=15)
        frame_progress.pack(fill=tk.X, padx=20)
        self.lbl_status_progress = tk.Label(frame_progress, text="Mempersiapkan data dari SQLite...", font=("Arial", 11), bg="#f4f6f9")
        self.lbl_status_progress.pack(anchor=tk.W, pady=2)
        self.progress_bar = ttk.Progressbar(frame_progress, orient="horizontal", mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=5)
        frame_console = tk.LabelFrame(self.root, text=" Live Console Log ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=10)
        frame_console.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        self.txt_log = tk.Text(frame_console, font=("Courier New", 9), bg="black", fg="#00ff00", state=tk.DISABLED)
        self.txt_log.pack(fill=tk.BOTH, expand=True)
        self.btn_kembali_menu = tk.Button(self.root, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#5f6368", fg="white", padx=15, pady=8, state=tk.DISABLED, command=self.tampilkan_menu_utama)
        self.btn_kembali_menu.pack(pady=15)
        t = threading.Thread(target=target_fungsi); t.daemon = True; t.start()

    def eksekusi_back_end_cetak_buku(self):
        output_pdf = os.path.join(DIR_PERPUS, "Cetak_Label_Buku_A3Plus.pdf")
        output_mal = os.path.join(DIR_PERPUS, "Mal_Kisscut_Label_A3Plus.pdf")
        try:
            self.log_ke_gui(self.txt_log, "Membaca database SQLite qr_id_buku...")
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query("SELECT * FROM qr_id_buku", conn)
            conn.close()
            
            if df.empty:
                self.log_ke_gui(self.txt_log, "[ERROR] Tabel qr_id_buku di SQLite kosong.")
                return
                
            rows_data = [r for i, r in df.iterrows() if pd.notna(r['label_buku'])]
            total_data = len(rows_data)
            paper_width, paper_height = 320 * MM_TO_POINT, 480 * MM_TO_POINT; box_width = box_height = 25 * MM_TO_POINT; qr_render_size = 17.5 * MM_TO_POINT; gap_x, gap_y = 5 * MM_TO_POINT, 7 * MM_TO_POINT; margin_left, margin_top = 10 * MM_TO_POINT, 12 * MM_TO_POINT
            c_design = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height)); c_mal = canvas.Canvas(output_mal, pagesize=(paper_width, paper_height))
            
            col_idx, row_idx, total = 0, 0, 0
            for row in rows_data:
                qr_text = str(row['qr_code']) if pd.notna(row['qr_code']) else str(row['label_buku']); label_text = str(row['label_buku'])
                x = margin_left + (col_idx * (box_width + gap_x)); y = paper_height - margin_top - box_height - (row_idx * (box_height + gap_y))
                jurusan = str(row['jurusan']).strip().upper() if pd.notna(row['jurusan']) else ""
                
                bg_color = (0.88, 0.82, 0.98) if jurusan in ["TKJ", "TEKNO"] else (1.0, 0.8, 0.85) if jurusan in ["MP", "BISMEN"] else (0.75, 0.93, 0.75) if jurusan == "BD" else (1.0, 0.96, 0.7) if jurusan == "AKL" else (1.0, 1.0, 1.0)
                c_design.setFillColorRGB(*bg_color); c_design.setStrokeColorRGB(0.9, 0.9, 0.9); c_design.rect(x, y, box_width, box_height, stroke=1, fill=1)
                c_mal.setStrokeColorRGB(1.0, 0.0, 1.0); c_mal.setLineWidth(0.3); c_mal.rect(x, y, box_width, box_height, stroke=1, fill=0)
                qr_x = x + ((box_width - qr_render_size) / 2); qr_y = y + box_height - qr_render_size - (1.0 * MM_TO_POINT)
                try:
                    encoded = urllib.parse.quote(qr_text); req = urllib.request.Request(f"https://quickchart.io/qr?text={encoded}&size=150&margin=0", headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=5) as response: c_design.drawImage(ImageReader(BytesIO(response.read())), qr_x, qr_y, width=qr_render_size, height=qr_render_size)
                except: c_design.line(qr_x, qr_y, qr_x + qr_render_size, qr_y + qr_render_size)
                c_design.setFillColorRGB(0.0, 0.0, 0.0); c_design.setFont("Helvetica-Bold", 6.0); c_design.drawCentredString(x + (box_width / 2), y + (2.0 * MM_TO_POINT), label_text)
                
                total += 1; col_idx += 1; persentase = int((total / total_data) * 100); self.progress_bar['value'] = persentase
                if col_idx >= 10: col_idx = 0; row_idx += 1
                if row_idx >= 14: c_design.showPage(); c_mal.showPage(); col_idx, row_idx = 0, 0
            c_design.save(); c_mal.save(); messagebox.showinfo("Sukses", "Stiker label buku selesai diproduksi!"); os.startfile(DIR_PERPUS)
        except Exception as e: messagebox.showerror("Error", f"{e}")
        finally: self.btn_kembali_menu.config(state=tk.NORMAL, bg="#1a73e8")

    def proses_ekspor_laporan_excel(self):
        timestamp_file = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_output_laporan = os.path.join(DIR_PERPUS, f"Laporan_Resmi_Perpus_{timestamp_file}.xlsx")
        try:
            if self.df_rekap_log.empty: return
            df_mentah = self.df_rekap_log.copy()
            df_mentah['judul_buku'] = df_mentah['kode_satuan_buku'].apply(self.dapatkan_judul_buku)
            
            df_sheet2 = df_mentah[df_mentah['status'].str.upper() == 'DIPINJAM'].groupby(['nisn', 'nama_siswa', 'rombel']).size().reset_index(name='Total Buku Belum Kembali')
            with pd.ExcelWriter(file_output_laporan) as writer:
                df_mentah.to_excel(writer, sheet_name='Log_Sirkulasi_Lengkap', index=False)
                df_sheet2.to_excel(writer, sheet_name='Tanggungan_Buku_Siswa', index=False)
            messagebox.showinfo("Ekspor Berhasil", "Laporan resmi perpustakaan diekspor dari SQLite!"); os.startfile(DIR_PERPUS)
        except Exception as e: messagebox.showerror("Error", f"{e}")

    def tampilkan_layar_proses_cetak_kartu(self):
        self.tampilkan_layar_setup_cetak_kartu()

if __name__ == "__main__":
    root = tk.Tk()
    app = AplikasiPerpusTerintegrasi(root)
    root.mainloop()