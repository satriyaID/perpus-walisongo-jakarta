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
from tkinter import messagebox, ttk, simpledialog, filedialog

# Import ReportLab & Pillow untuk Cetak PDF
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from PIL import Image, ImageDraw, ImageFont

# =====================================================================
# PATH CONFIGURATION (DYNAMIC, RELATIVE & SECURE "DB_PERPUS" FOLDER)
# =====================================================================
if getattr(sys, 'frozen', False):
    FOLDER_KERJA = os.path.dirname(sys.executable)
else:
    FOLDER_KERJA = os.path.dirname(os.path.abspath(__file__))

DIR_PERPUS = os.path.join(FOLDER_KERJA, "DB_PERPUS")
os.makedirs(DIR_PERPUS, exist_ok=True)

FOLDER_BACKUP = os.path.join(DIR_PERPUS, "Backup_Data")
os.makedirs(FOLDER_BACKUP, exist_ok=True)

DB_PATH = os.path.join(DIR_PERPUS, "perpustakaan.db")

TEMPLATE_KARTU = os.path.join(FOLDER_KERJA, "Kartu Perpustakaan SMK Walisongo.png")
FONT_CHAU = os.path.join(FOLDER_KERJA, "ChauPhilomene-Regular.ttf")

EXCEL_BUKU_LAMA = os.path.join(FOLDER_KERJA, "QR_ID_BUKU.xlsx")
EXCEL_SISWA_LAMA = os.path.join(FOLDER_KERJA, "Data_Siswa_Perpus.xlsx")
EXCEL_LOG_LAMA = os.path.join(FOLDER_KERJA, "Data_Peminjaman_Buku.xlsx")
EXCEL_UMUM_LAMA = os.path.join(FOLDER_KERJA, "Data_Buku_Umum.xlsx")

MM_TO_POINT = 2.83465

class AplikasiPerpusTerintegrasi:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Informasi Perpustakaan - SMK Walisongo Jakarta")
        self.root.geometry("920x880")
        self.root.configure(bg="#f4f6f9")
        
        self.target_buku = {"10": 10, "11": 15, "12": 4}
        
        self.siswa_aktif = None
        self.daftar_buku_dijepit = []
        self.daftar_kembali_dijepit = []
        self.antrean_reprint = []
        self.dict_check_siswa = {} 
        self.siswa_terfilter_cetak = []
        
        with open(os.path.join(DIR_PERPUS, "⚠️ JANGAN DIHAPUS.txt"), "w", encoding="utf-8") as f:
            f.write("PENTING:\nFolder ini berisi database utama SQLite Aplikasi Perpustakaan SMK Walisongo.\n"
                    "Menghapus atau mengubah isi folder ini akan merusak sistem.")
            
        self.inisialisasi_database()
        self.migrasi_excel_ke_sqlite()
        self.jalankan_auto_backup_silent()
        
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
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        
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
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS data_siswa_perpus (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nisn TEXT,
                nama TEXT,
                jurusan TEXT,
                rombel TEXT
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS data_buku_umum (
                isbn TEXT PRIMARY KEY,
                judul_buku TEXT
            )
        ''')
        
        conn.commit()
        conn.close()

    def migrasi_excel_ke_sqlite(self):
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
            try:
                cursor.execute(f"SELECT COUNT(*) FROM {nama_tabel}")
                jumlah_baris = cursor.fetchone()[0]
            except sqlite3.OperationalError:
                jumlah_baris = 0
            
            if jumlah_baris > 0:
                continue
            
            if os.path.exists(path_excel):
                file_name = os.path.basename(path_excel)
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
                    
                    df_lama.to_sql(nama_tabel, conn, if_exists="append", index=False)
                    waktu_sekarang = datetime.now().strftime("%Y%m%d_%H%M%S")
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

    def bersihkan_nisn_ke_string(self, val):
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip().split('.')[0]
        return s.lstrip('0')

    def ekstrak_kode_induk(self, kode_lengkap):
        """Membuang 3 digit eksemplar terakhir dari kode buku (Format Induk)"""
        kode = str(kode_lengkap).strip()
        if '-' in kode:
            bagian = kode.split('-')
            if len(bagian[-1]) >= 1 and bagian[-1].isdigit():
                return '-'.join(bagian[:-1])
        return kode

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
        tk.Button(frame_tombol, text="[2] CETAK STIKER LABEL BUKU MASSAL / IMPOR BARU\n(Layout A3+ Precision Kisscut)", font=("Arial", 11, "bold"), bg="#1a73e8", fg="white", width=45, height=2, command=self.tampilkan_layar_proses_cetak_buku).grid(row=1, column=0, pady=4)
        tk.Button(frame_tombol, text="[3] CETAK KARTU PERPUS DENGAN CEKLIS\n(Pilih Rombel & Pilih Siswa Secara Selektif)", font=("Arial", 11, "bold"), bg="#ff9900", fg="white", width=45, height=2, command=self.tampilkan_layar_setup_cetak_kartu).grid(row=2, column=0, pady=4)
        tk.Button(frame_tombol, text="[4] CETAK ULANG KARTU (REPRINT KOLEKTIF)\n(Cari via NISN - Maks 9 Siswa di A4 Landscape)", font=("Arial", 11, "bold"), bg="#6f42c1", fg="white", width=45, height=2, command=self.tampilkan_layar_cetak_satuan).grid(row=3, column=0, pady=4)
        tk.Button(frame_tombol, text="[5] MODUL REKAP & KONTROL PEMERATAAN BUKU\n(Matriks Ceklis Rombel, Rekap Buku, Ekspor Excel & PDF)", font=("Arial", 11, "bold"), bg="#007afc", fg="white", width=45, height=2, command=self.tampilkan_layar_rekap).grid(row=4, column=0, pady=4)
        
        tk.Button(
            frame_tombol, 
            text="[🛠️] KOREKSI / TAMBAH DATA MASTER PEMINJAM (SISWA/GURU)\n(Mutasi NISN/NIP & Pendaftaran Siswa Baru)", 
            font=("Arial", 11, "bold"), 
            bg="#f57c00", 
            fg="white", 
            width=45, 
            height=2, 
            command=self.buka_popup_koreksi_data_master
        ).grid(row=5, column=0, pady=8)

        frame_utilitas = tk.LabelFrame(self.root, text=" Pengaman & Pemulihan Data Relasional ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5)
        frame_utilitas.pack(fill=tk.X, padx=30, pady=5)
        tk.Button(frame_utilitas, text="Amankan Cadangan SQLite (Backup)", font=("Arial", 9, "bold"), bg="#5f6368", fg="white", command=self.pemicu_backup_manual).pack(side=tk.LEFT, padx=10, expand=True, fill=tk.X)
        tk.Button(frame_utilitas, text="Buka Folder Cadangan (Recovery)", font=("Arial", 9, "bold"), bg="#7c7c7c", fg="white", command=self.pemicu_recovery_data).pack(side=tk.RIGHT, padx=10, expand=True, fill=tk.X)
        
        tk.Button(self.root, text="Keluar Aplikasi", font=("Arial", 10), bg="#d93025", fg="white", width=15, pady=5, command=self.root.quit).pack(pady=10)

    # =====================================================================
    # FORM POP-UP KOREKSI & TAMBAH DATA MASTER
    # =====================================================================
    def buka_popup_koreksi_data_master(self):
        top = tk.Toplevel(self.root)
        top.title("Koreksi & Input Data Master Peminjam (Siswa/Guru)")
        top.geometry("560x520")
        top.resizable(False, False)
        top.grab_set()
        
        mode_input = {"is_baru": False}
        
        lbl_header = tk.Label(
            top, 
            text="FORM KOREKSI / TAMBAH DATA MASTER PEMINJAM", 
            font=("Arial", 11, "bold"), 
            bg="#f57c00", 
            fg="white", 
            pady=12
        )
        lbl_header.pack(fill=tk.X)
        
        frame_form = tk.Frame(top, padx=25, pady=15)
        frame_form.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame_form, text="Kategori Peminjam:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=6)
        var_kategori = tk.StringVar(value="SISWA")
        
        def toggle_kategori():
            if var_kategori.get() == "GURU":
                lbl_identitas.config(text="Cari NISN / Nama / NIP (Lama):")
                lbl_identitas_baru.config(text="NIP / Kode Guru (Baru):")
                ent_rombel.delete(0, tk.END)
                ent_rombel.insert(0, "GURU")
                ent_rombel.config(state="disabled")
            else:
                lbl_identitas.config(text="Cari NISN / Nama (Lama):")
                lbl_identitas_baru.config(text="NISN Siswa (Baru):")
                ent_rombel.config(state="normal")
                ent_rombel.delete(0, tk.END)
                
        frame_kat = tk.Frame(frame_form)
        frame_kat.grid(row=0, column=1, sticky="w")
        rb_siswa = tk.Radiobutton(frame_kat, text="Siswa", variable=var_kategori, value="SISWA", command=toggle_kategori)
        rb_guru = tk.Radiobutton(frame_kat, text="Guru / Staf", variable=var_kategori, value="GURU", command=toggle_kategori)
        rb_siswa.pack(side=tk.LEFT, padx=(0, 10))
        rb_guru.pack(side=tk.LEFT)

        lbl_identitas = tk.Label(frame_form, text="Cari NISN / Nama (Lama):", font=("Arial", 10))
        lbl_identitas.grid(row=1, column=0, sticky="w", pady=6)
        ent_id_lama = tk.Entry(frame_form, font=("Arial", 10), width=25)
        ent_id_lama.grid(row=1, column=1, sticky="w", pady=6)
        
        def cari_data_existing():
            kata_kunci = ent_id_lama.get().strip()
            if not kata_kunci:
                messagebox.showwarning("Peringatan", "Silakan ketik NISN atau Nama yang ingin dicari.", parent=top)
                return
                
            id_clean = self.bersihkan_nisn_ke_string(kata_kunci)
            mode_input["is_baru"] = False
            
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                
                if kata_kunci.isdigit() and id_clean != "":
                    cursor.execute("SELECT nisn, nama, rombel FROM data_siswa_perpus WHERE nisn = ? OR nisn LIKE ?", 
                                   (kata_kunci, f"%{id_clean}%"))
                else:
                    if kata_kunci in ["0", "", "NULL", "NONE"]:
                        cursor.execute("SELECT nisn, nama, rombel FROM data_siswa_perpus WHERE nisn = '0' OR nisn IS NULL OR nisn = ''")
                    else:
                        cursor.execute("SELECT nisn, nama, rombel FROM data_siswa_perpus WHERE nama LIKE ?", 
                                       (f"%{kata_kunci}%",))
                                       
                rows = cursor.fetchall()
                conn.close()
                
                if not rows:
                    tanya_tambah = messagebox.askyesno(
                        "Data Tidak Ditemukan", 
                        f"Data '{kata_kunci}' TIDAK DITERDETEKSI di master database.\n\n"
                        f"Apakah Anda ingin menambahkan ini sebagai SISWA / PEMINJAM BARU?", 
                        parent=top
                    )
                    
                    if tanya_tambah:
                        mode_input["is_baru"] = True
                        lbl_header.config(text="FORM INPUT PEMINJAM BARU KE DATABASE", bg="#34a853")
                        btn_simpan.config(text="➕ Simpan Data Peminjam Baru", bg="#34a853")
                        
                        ent_id_baru.delete(0, tk.END)
                        if kata_kunci.isdigit():
                            ent_id_baru.insert(0, kata_kunci)
                        else:
                            ent_nama_baru.delete(0, tk.END)
                            ent_nama_baru.insert(0, kata_kunci.upper())
                            
                        ent_nama_baru.focus()
                        lbl_warning.config(
                            text="ℹ️ MODE SISWA BARU: Data akan didaftarkan ke Master Database Peminjam.",
                            fg="#188038"
                        )
                    return
                
                lbl_header.config(text="FORM KOREKSI DATA MASTER & SINKRONISASI", bg="#f57c00")
                btn_simpan.config(text="💾 Jalankan Koreksi Data", bg="#1a73e8")
                lbl_warning.config(
                    text="🔒 SISTEM PROTEKSI AKTIF:\nPerubahan data akan memperbarui tabel master & menyinkronkan seluruh\nriwayat log transaksi sirkulasi aktif tanpa menghapus data apa pun.",
                    fg="#d93025"
                )

                if len(rows) > 1:
                    popup_pilih = tk.Toplevel(top)
                    popup_pilih.title("Pilih Data Yang Sesuai")
                    popup_pilih.geometry("500x250")
                    popup_pilih.grab_set()
                    
                    tk.Label(popup_pilih, text="Ditemukan beberapa data yang cocok, silakan pilih:", font=("Arial", 10, "bold"), pady=5).pack()
                    lb = tk.Listbox(popup_pilih, font=("Arial", 10), width=65, height=8)
                    lb.pack(padx=10, pady=5)
                    
                    for r in rows:
                        lb.insert(tk.END, f"NISN: {str(r[0]).split('.')[0]} | Nama: {r[1]} | Rombel: {r[2]}")
                        
                    def konfirmasi_pilihan():
                        idx = lb.curselection()
                        if not idx:
                            messagebox.showwarning("Pilih", "Silakan klik salah satu data terlebih dahulu.", parent=popup_pilih)
                            return
                        row_terpilih = rows[idx[0]]
                        
                        ent_id_lama.delete(0, tk.END)
                        ent_id_lama.insert(0, str(row_terpilih[0]).split('.')[0].strip())
                        ent_id_baru.delete(0, tk.END)
                        ent_id_baru.insert(0, str(row_terpilih[0]).split('.')[0].strip())
                        ent_nama_baru.delete(0, tk.END)
                        ent_nama_baru.insert(0, str(row_terpilih[1]).upper())
                        
                        ent_rombel.config(state="normal")
                        ent_rombel.delete(0, tk.END)
                        ent_rombel.insert(0, str(row_terpilih[2]).upper())
                        if var_kategori.get() == "GURU": ent_rombel.config(state="disabled")
                        
                        popup_pilih.destroy()
                        
                    tk.Button(popup_pilih, text="✓ Pilihlah Data Ini", bg="#34a853", fg="white", font=("Arial", 9, "bold"), command=konfirmasi_pilihan).pack(pady=5)
                    
                else:
                    row = rows[0]
                    ent_id_lama.delete(0, tk.END)
                    ent_id_lama.insert(0, str(row[0]).split('.')[0].strip())
                    ent_id_baru.delete(0, tk.END)
                    ent_id_baru.insert(0, str(row[0]).split('.')[0].strip())
                    ent_nama_baru.delete(0, tk.END)
                    ent_nama_baru.insert(0, str(row[1]).upper())
                    
                    ent_rombel.config(state="normal")
                    ent_rombel.delete(0, tk.END)
                    ent_rombel.insert(0, str(row[2]).upper())
                    if var_kategori.get() == "GURU": ent_rombel.config(state="disabled")
                    
                    messagebox.showinfo("Data Ditemukan", f"Data Peminjam: '{row[1]}' berhasil dimuat.", parent=top)
                    
            except Exception as e:
                messagebox.showerror("Error Database", f"Gagal membaca data: {e}", parent=top)

        btn_cari = tk.Button(frame_form, text="🔍 Cari Data", bg="#5c6bc0", fg="white", font=("Arial", 9, "bold"), command=cari_data_existing)
        btn_cari.grid(row=1, column=2, padx=5)

        ttk.Separator(frame_form, orient='horizontal').grid(row=2, column=0, columnspan=3, sticky="ew", pady=12)

        lbl_identitas_baru = tk.Label(frame_form, text="NISN Siswa (Baru):", font=("Arial", 10, "bold"))
        lbl_identitas_baru.grid(row=3, column=0, sticky="w", pady=6)
        ent_id_baru = tk.Entry(frame_form, font=("Arial", 10), width=35)
        ent_id_baru.grid(row=3, column=1, columnspan=2, sticky="w", pady=6)

        tk.Label(frame_form, text="Nama Lengkap (Baru):", font=("Arial", 10, "bold")).grid(row=4, column=0, sticky="w", pady=6)
        ent_nama_baru = tk.Entry(frame_form, font=("Arial", 10), width=35)
        ent_nama_baru.grid(row=4, column=1, columnspan=2, sticky="w", pady=6)

        tk.Label(frame_form, text="Rombel / Kelas (Baru):", font=("Arial", 10, "bold")).grid(row=5, column=0, sticky="w", pady=6)
        ent_rombel = tk.Entry(frame_form, font=("Arial", 10), width=35)
        ent_rombel.grid(row=5, column=1, columnspan=2, sticky="w", pady=6)

        lbl_warning = tk.Label(
            frame_form, 
            text="🔒 SISTEM PROTEKSI AKTIF:\nPerubahan data akan memperbarui tabel master & menyinkronkan seluruh\nriwayat log transaksi sirkulasi aktif tanpa menghapus data apa pun.", 
            font=("Arial", 8, "italic"), 
            fg="#d93025", 
            justify="left"
        )
        lbl_warning.grid(row=6, column=0, columnspan=3, sticky="w", pady=(20, 5))

        def eksekusi_simpan_koreksi():
            id_lama = ent_id_lama.get().strip()
            id_baru = ent_id_baru.get().strip()
            nama_baru = ent_nama_baru.get().strip().upper()
            rombel_baru = ent_rombel.get().strip().upper()

            if not id_baru or not nama_baru or not rombel_baru:
                messagebox.showwarning("Form Kosong", "NISN/NIP Baru, Nama Lengkap, dan Rombel wajib diisi!", parent=top)
                return

            if mode_input["is_baru"]:
                pesan_tambah = (
                    f"➕ KONFIRMASI INPUT PEMINJAM BARU\n\n"
                    f"• NISN / NIP  : {id_baru}\n"
                    f"• Nama Lengkap: {nama_baru}\n"
                    f"• Rombel      : {rombel_baru}\n\n"
                    f"Daftarkan peminjam baru ini ke master database?"
                )
                if messagebox.askyesno("Tambah Peminjam Baru", pesan_tambah, parent=top):
                    try:
                        conn = sqlite3.connect(DB_PATH)
                        cursor = conn.cursor()
                        jurusan = rombel_baru.split('-')[1] if '-' in rombel_baru else "UMUM"
                        cursor.execute("""
                            INSERT INTO data_siswa_perpus (nisn, nama, jurusan, rombel)
                            VALUES (?, ?, ?, ?)
                        """, (id_baru, nama_baru, jurusan, rombel_baru))
                        conn.commit()
                        conn.close()
                        self.jalankan_auto_backup_silent()
                        messagebox.showinfo("Sukses Mendaftar", f"Siswa/Peminjam '{nama_baru}' sukses didaftarkan!", parent=top)
                        top.destroy()
                    except Exception as e:
                        messagebox.showerror("Error", f"Gagal menambahkan data baru: {e}", parent=top)
                return

            if not id_lama:
                messagebox.showwarning("Form Kosong", "Identitas lama wajib diisi saat mode Koreksi!", parent=top)
                return

            pesan_konfirmasi = (
                f"⚠️ PERINGATAN KOREKSI DATA SINKRON!\n\n"
                f"• Identitas Acuan Lama : {id_lama}\n"
                f"• NISN / NIP Baru      : {id_baru}\n"
                f"• Nama Terkoreksi      : {nama_baru}\n"
                f"• Rombel Baru          : {rombel_baru}\n\n"
                f"Lanjutkan pembaruan dua tahap ke master & seluruh log sirkulasi?"
            )

            if messagebox.askyesno("Konfirmasi Mutasi Data", pesan_konfirmasi, parent=top):
                try:
                    conn = sqlite3.connect(DB_PATH)
                    cursor = conn.cursor()
                    cursor.execute("BEGIN TRANSACTION;")

                    id_clean_lama = self.bersihkan_nisn_ke_string(id_lama)

                    cursor.execute("""
                        SELECT nama, nisn FROM data_siswa_perpus 
                        WHERE nisn = ? OR (nisn LIKE ? AND ? != '') OR nama LIKE ?
                        LIMIT 1
                    """, (id_lama, f"%{id_clean_lama}%", id_clean_lama, f"%{id_lama}%"))
                    row_master_lama = cursor.fetchone()
                    
                    nama_lama_db = row_master_lama[0] if row_master_lama else id_lama
                    nisn_lama_db = row_master_lama[1] if row_master_lama else id_lama

                    cursor.execute("""
                        UPDATE data_siswa_perpus 
                        SET nisn = ?, nama = ?, rombel = ? 
                        WHERE nisn = ? OR (nisn LIKE ? AND ? != '') OR nama LIKE ?
                    """, (id_baru, nama_baru, rombel_baru, id_lama, f"%{id_clean_lama}%", id_clean_lama, f"%{id_lama}%"))

                    cursor.execute("""
                        UPDATE data_peminjaman_buku 
                        SET nisn = ?, nama_siswa = ?, rombel = ? 
                        WHERE nisn = ? 
                           OR (nisn LIKE ? AND ? != '')
                           OR nama_siswa = ? 
                           OR nama_siswa = ?
                           OR nama_siswa LIKE ?
                    """, (id_baru, nama_baru, rombel_baru, id_lama, f"%{id_clean_lama}%", id_clean_lama, nama_lama_db, nisn_lama_db, f"%{id_lama}%"))

                    baris_terdampak = cursor.rowcount
                    conn.commit()
                    conn.close()

                    self.jalankan_auto_backup_silent()
                    if hasattr(self, 'df_rekap_log'):
                        conn_ref = sqlite3.connect(DB_PATH)
                        self.df_rekap_log = pd.read_sql_query("SELECT * FROM data_peminjaman_buku", conn_ref)
                        conn_ref.close()

                    messagebox.showinfo(
                        "Sukses Terkoreksi", 
                        f"Data master dan {baris_terdampak} log transaksi sirkulasi terkait berhasil disinkronkan!", 
                        parent=top
                    )
                    top.destroy()

                except Exception as e:
                    conn.rollback()
                    conn.close()
                    messagebox.showerror("Error Transaksi", f"Mutasi gagal, data dikembalikan ke awal: {e}", parent=top)

        frame_btn = tk.Frame(top, pady=10, padx=25)
        frame_btn.pack(fill=tk.X, side=tk.BOTTOM)

        btn_simpan = tk.Button(
            frame_btn, 
            text="💾 Jalankan Koreksi Data", 
            font=("Arial", 10, "bold"), 
            bg="#1a73e8", 
            fg="white", 
            padx=15, 
            pady=6, 
            command=eksekusi_simpan_koreksi
        )
        btn_simpan.pack(side=tk.RIGHT, padx=5)

        btn_batal = tk.Button(
            frame_btn, 
            text="Batal", 
            font=("Arial", 10), 
            bg="#e0e0e0", 
            fg="black", 
            padx=15, 
            pady=6, 
            command=top.destroy
        )
        btn_batal.pack(side=tk.RIGHT, padx=5)

    # =====================================================================
    # HALAMAN 2: MODUL SIRKULASI TERINTEGRASI (DENGAN PROTEKSI SESI & ANTI-MULTI-SCAN)
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
            
        self.df_siswa['nisn'] = self.df_siswa['nisn'].astype(str).str.strip()
        
        banner = tk.Label(self.root, text="MODUL SIRKULASI PERPUSTAKAAN DUA ARAH (SQLITE ENGINE)", font=("Arial", 14, "bold"), bg="#34a853", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        tabControl = ttk.Notebook(self.root)
        tab_pinjam = tk.Frame(tabControl, bg="#f4f6f9"); tab_kembali = tk.Frame(tabControl, bg="#f4f6f9")
        tabControl.add(tab_pinjam, text="  [+] LAYANAN PEMINJAMAN BUKU PAS  "); tabControl.add(tab_kembali, text="  [-] LAYANAN PENGEMBALIAN BUKU PAS  ")
        tabControl.pack(expand=1, fill="both", padx=10, pady=5)
        
        frame_scan_siswa = tk.LabelFrame(tab_pinjam, text=" Langkah 1: Scan Kartu / Input NISN / NIP Peminjam ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_scan_siswa.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_scan_siswa, text="Scan NISN / NIP:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_nisn = tk.Entry(frame_scan_siswa, font=("Arial", 12), width=30); self.ent_nisn.grid(row=0, column=1, padx=10); self.ent_nisn.bind("<Return>", self.proses_scan_siswa); self.ent_nisn.focus()
        self.lbl_info_siswa = tk.Label(frame_scan_siswa, text="[Silakan scan kartu peminjam untuk memulai]", font=("Arial", 11, "italic"), fg="#555", bg="#e8eaed", width=75, height=4, anchor=tk.W, justify=tk.LEFT, padx=10, pady=5); self.lbl_info_siswa.grid(row=1, column=0, columnspan=2, pady=10)
        
        self.frame_scan_buku = tk.LabelFrame(tab_pinjam, text=" Langkah 2: Scan Kode Buku Paket / Barcode ISBN Buku Umum ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); self.frame_scan_buku.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Label(self.frame_scan_buku, text="Scan Barcode:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_buku = tk.Entry(self.frame_scan_buku, font=("Arial", 12), width=30, state=tk.DISABLED); self.ent_buku.grid(row=0, column=1, padx=10); self.ent_buku.bind("<Return>", self.proses_scan_buku)
        
        self.tree_buku = ttk.Treeview(self.frame_scan_buku, columns=("No", "Kode Buku", "Nama Buku"), show="headings", height=6)
        self.tree_buku.heading("No", text="No"); self.tree_buku.heading("Kode Buku", text="ID Satuan / Eksemplar"); self.tree_buku.heading("Nama Buku", text="Judul Buku Terdeteksi")
        self.tree_buku.column("No", width=50, anchor=tk.CENTER); self.tree_buku.column("Kode Buku", width=200, anchor=tk.W); self.tree_buku.column("Nama Buku", width=450, anchor=tk.W)
        self.tree_buku.grid(row=1, column=0, columnspan=2, pady=5, sticky="nsew")
        
        self.lbl_counter = tk.Label(self.frame_scan_buku, text="Jumlah Buku Di-scan: 0", font=("Arial", 11, "bold"), bg="#f4f6f9", fg="#1a73e8"); self.lbl_counter.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=2)
        self.frame_scan_buku.grid_rowconfigure(1, weight=1); self.frame_scan_buku.grid_columnconfigure(1, weight=1)
        
        frame_scan_kembali = tk.LabelFrame(tab_kembali, text=" Masukkan / Scan Label Buku Yang Dikembalikan ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=10); frame_scan_kembali.pack(fill=tk.X, padx=10, pady=5)
        tk.Label(frame_scan_kembali, text="Scan Barcode:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_buku_kembali = tk.Entry(frame_scan_kembali, font=("Arial", 12), width=30); self.ent_buku_kembali.grid(row=0, column=1, padx=10, sticky=tk.W); self.ent_buku_kembali.bind("<Return>", self.proses_scan_pengembalian_buku)
        frame_tabel_kembali = tk.LabelFrame(tab_kembali, text=" Manifest Validasi Buku Yang Berhasil Dikembalikan ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=5); frame_tabel_kembali.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_kembali = ttk.Treeview(frame_tabel_kembali, columns=("No", "Kode Buku", "Nama Siswa", "Rombel", "Status"), show="headings", height=8); self.tree_kembali.heading("No", text="No"); self.tree_kembali.heading("Kode Buku", text="Kode Seri Buku"); self.tree_kembali.heading("Nama Siswa", text="Nama Mantan Peminjam"); self.tree_kembali.heading("Rombel", text="Rombel"); self.tree_kembali.heading("Status", text="Status Akhir")
        self.tree_kembali.column("No", width=40, anchor=tk.CENTER); self.tree_kembali.column("Kode Buku", width=160, anchor=tk.W); self.tree_kembali.column("Nama Siswa", width=250, anchor=tk.W); self.tree_kembali.column("Rombel", width=90, anchor=tk.CENTER); self.tree_kembali.column("Status", width=90, anchor=tk.CENTER); self.tree_kembali.pack(fill=tk.BOTH, expand=True, pady=5)
        
        frame_aksi_kembali = tk.Frame(tab_kembali, bg="#f4f6f9", pady=5); frame_aksi_kembali.pack(fill=tk.X, padx=10)
        self.btn_simpan_kembali = tk.Button(frame_aksi_kembali, text="Proses & Bukukan Pengembalian", font=("Arial", 11, "bold"), bg="#1a73e8", fg="white", padx=15, pady=6, command=self.eksekusi_simpan_pengembalian, state=tk.DISABLED); self.btn_simpan_kembali.pack(side=tk.RIGHT)

        self.lbl_marquee.pack_forget()
        self.lbl_marquee.pack(fill=tk.X, side=tk.BOTTOM)
        
        frame_navigasi_total = tk.Frame(self.root, bg="#f4f6f9", pady=10)
        
        frame_nav_kiri = tk.Frame(frame_navigasi_total, bg="#f4f6f9")
        frame_nav_kiri.pack(side=tk.LEFT, padx=15)
        btn_back_utama = tk.Button(frame_nav_kiri, text="◀ Kembali ke Menu Utama (Dashboard)", font=("Arial", 11, "bold"), bg="#d93025", fg="white", padx=20, pady=8, command=self.tampilkan_menu_utama)
        btn_back_utama.pack(anchor=tk.W)

        frame_aksi_sirkulasi = tk.Frame(frame_navigasi_total, bg="#f4f6f9")
        frame_aksi_sirkulasi.pack(side=tk.RIGHT, padx=15)
        self.btn_simpan = tk.Button(frame_aksi_sirkulasi, text="Simpan & Lanjut Peminjam Lain ▶", font=("Arial", 11, "bold"), bg="#188038", fg="white", padx=20, pady=8, command=self.konfirmasi_dan_simpan, state=tk.DISABLED)
        self.btn_simpan.pack(anchor=tk.E)
        
        frame_navigasi_total.pack(fill=tk.X, side=tk.BOTTOM)

    def proses_scan_siswa(self, event=None):
        nisn_input = self.ent_nisn.get().strip()
        if not nisn_input: return
        
        # PROTEKSI KETAT: Mencegah scan siswa baru jika sesi siswa sebelumnya belum di-commit/disimpan
        if self.siswa_aktif is not None and len(self.daftar_buku_dijepit) > 0:
            self.ent_nisn.delete(0, tk.END)
            messagebox.showwarning(
                "Sesi Belum Disimpan!", 
                f"⚠️ PERINGATAN KETAT:\n\n"
                f"Anda sedang memproses peminjaman untuk '{self.siswa_aktif['nama']}' "
                f"dan masih ada {len(self.daftar_buku_dijepit)} buku dalam antrean yang BELUM DISIMPAN!\n\n"
                f"Silakan klik tombol 'Simpan & Lanjut Peminjam Lain' terlebih dahulu."
            )
            return

        id_clean = self.bersihkan_nisn_ke_string(nisn_input)
        hasil = self.df_siswa[self.df_siswa['nisn'].apply(self.bersihkan_nisn_ke_string) == id_clean]
        
        if not hasil.empty:
            row = hasil.iloc[0]
            self.siswa_aktif = {
                "nisn": str(row['nisn']).strip(), 
                "nama": str(row['nama']).upper(), 
                "jurusan": str(row['jurusan']).strip(), 
                "rombel": str(row['rombel']).strip(), 
                "kelas": str(row['rombel']).split('-')[0]
            }
            kelas_digit = "".join(filter(str.isdigit, self.siswa_aktif["kelas"]))
            self.target_seharusnya = self.target_buku.get(kelas_digit, 10)
            info_text = f"Nama    : {self.siswa_aktif['nama']}\nNISN/NIP: {self.siswa_aktif['nisn']}\nRombel  : {self.siswa_aktif['rombel']} | Jurusan: {self.siswa_aktif['jurusan']}\nTarget Wajib: {self.target_seharusnya} Buku Paket"
            self.lbl_info_siswa.config(text=info_text, fg="#155724", bg="#d4edda", font=("Arial", 11, "bold"))
            self.ent_buku.config(state=tk.NORMAL); self.ent_buku.focus(); self.ent_nisn.config(state=tk.DISABLED)
        else:
            nisn_guru = f"GURU-{id_clean}" if id_clean else "GURU-001"
            nama_guru = simpledialog.askstring("Identitas Guru / Staf", f"ID {nisn_input} tidak ditemukan di master siswa.\n\nMasukkan Nama Guru / Staf Peminjam:", parent=self.root)
            if nama_guru:
                self.siswa_aktif = {
                    "nisn": nisn_guru, "nama": nama_guru.upper(), "jurusan": "PENDIDIK", "rombel": "GURU", "kelas": "GURU"
                }
                info_text = f"Nama    : {self.siswa_aktif['nama']}\nID Guru : {self.siswa_aktif['nisn']}\nKategori: GURU / STAF PENDIDIK\nTarget Wajib: Bebas Peminjaman"
                self.lbl_info_siswa.config(text=info_text, fg="#0c5460", bg="#d1ecf1", font=("Arial", 11, "bold"))
                self.ent_buku.config(state=tk.NORMAL); self.ent_buku.focus(); self.ent_nisn.config(state=tk.DISABLED)
            else:
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

        # PROTEKSI ANTI-MULTI-SCAN / DUPLIKASI DALAM SESI LOKAL
        if any(buku_id == id_final_buku for buku_id, _, _ in self.daftar_buku_dijepit):
            messagebox.showwarning("Duplikasi Sesi", f"Buku dengan kode '{id_final_buku}' sudah di-scan dalam daftar antrean peminjaman siswa ini.")
            return

        # PROTEKSI DATABASE: Memastikan 1 kode satuan buku hanya bisa dipinjam 1 orang secara aktif
        try:
            conn = sqlite3.connect(DB_PATH)
            df_cek_db = pd.read_sql_query(
                "SELECT nama_siswa, rombel FROM data_peminjaman_buku WHERE kode_satuan_buku = ? AND UPPER(status) = 'DIPINJAM'", 
                conn, 
                params=(id_final_buku,)
            )
            conn.close()
            
            if not df_cek_db.empty:
                peminjam_lama = df_cek_db.iloc[0]['nama_siswa']
                rombel_lama = df_cek_db.iloc[0]['rombel']
                messagebox.showerror(
                    "Buku Sedang Dipinjam Orang Lain!", 
                    f"❌ TRANSAKSI DITOLAK OLEH SISTEM!\n\n"
                    f"Buku dengan kode satuan '{id_final_buku}' ({judul_terdeteksi})\n"
                    f"statusnya MASIH TERCATAT DIPINJAM oleh:\n\n"
                    f"• Nama  : {peminjam_lama}\n"
                    f"• Rombel: {rombel_lama}\n\n"
                    f"Harap lakukan pengembalian buku terlebih dahulu sebelum meminjamkannya ke siswa lain."
                )
                return
        except Exception as e:
            print(f"Error pengecekan database: {e}")

        self.daftar_buku_dijepit.append((id_final_buku, judul_terdeteksi, input_raw))
        no_urut = len(self.daftar_buku_dijepit)
        self.tree_buku.insert("", tk.END, values=(no_urut, id_final_buku, judul_terdeteksi))
        self.lbl_counter.config(text=f"Jumlah Buku Di-scan: {no_urut} Buku")
        self.btn_simpan.config(state=tk.NORMAL)

    def konfirmasi_dan_simpan(self):
        if messagebox.askyesno("Konfirmasi", f"Simpan seluruh peminjaman {self.siswa_aktif['nama']} ke SQLite?"):
            tgl_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                for id_buku, _, _ in self.daftar_buku_dijepit:
                    cursor.execute('''
                        INSERT INTO data_peminjaman_buku (tanggal_pinjam, nisn, nama_siswa, rombel, kode_satuan_buku, status)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (tgl_sekarang, str(self.siswa_aktif["nisn"]), str(self.siswa_aktif["nama"]), str(self.siswa_aktif["rombel"]), id_buku, "DIPINJAM"))
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
        
        if any(b['Kode Buku'] == kode_input for b in self.daftar_kembali_dijepit):
            messagebox.showwarning("Duplikasi", f"Buku dengan kode '{kode_input}' sudah ada dalam antrean pengembalian saat ini.")
            return
            
        conn = sqlite3.connect(DB_PATH)
        
        if len(kode_input) == 13 and (kode_input.startswith("978") or kode_input.startswith("979")):
            df_log = pd.read_sql_query(
                "SELECT rowid, * FROM data_peminjaman_buku WHERE kode_satuan_buku LIKE ? AND UPPER(status) LIKE '%PINJAM%'", 
                conn, 
                params=(f"{kode_input}%",)
            )
        else:
            df_log = pd.read_sql_query(
                "SELECT rowid, * FROM data_peminjaman_buku WHERE kode_satuan_buku = ? AND UPPER(status) LIKE '%PINJAM%'", 
                conn, 
                params=(kode_input,)
            )
            
        conn.close()
        
        if not df_log.empty:
            row = df_log.iloc[-1]
            
            if 'id' in row and pd.notna(row['id']):
                id_db = str(row['id'])
            elif 'rowid' in row and pd.notna(row['rowid']):
                id_db = str(row['rowid'])
            else:
                id_db = str(row.name)

            kode_real_db = str(row['kode_satuan_buku'])
            nama_peminjam = str(row['nama_siswa']).upper()
            rombel_peminjam = str(row['rombel']).upper()
            
            judul_buku = self.dapatkan_judul_buku(kode_real_db)
            
            pesan_konfirmasi = (
                f"📌 DETEKSI TRANSAKSI PEMINJAMAN DITEMUKAN!\n\n"
                f"• Peminjam  : {nama_peminjam} ({rombel_peminjam})\n"
                f"• Judul Buku: {judul_buku}\n"
                f"• Kode Seri : {kode_real_db}\n\n"
                f"Masukkan buku ini ke dalam antrean pengembalian?"
            )
            
            if messagebox.askyesno("Konfirmasi Pengembalian Buku", pesan_konfirmasi):
                buku_kembali = {
                    "id_pk": id_db, 
                    "Kode Buku": kode_real_db, 
                    "Nama Siswa": nama_peminjam, 
                    "Rombel": rombel_peminjam, 
                    "Judul Buku": judul_buku,
                    "Status": "KEMBALI"
                }
                
                self.daftar_kembali_dijepit.append(buku_kembali)
                no_urut = len(self.daftar_kembali_dijepit)
                
                self.tree_kembali.insert("", tk.END, values=(no_urut, kode_real_db, nama_peminjam, rombel_peminjam, "VALIDASI ✓"))
                self.btn_simpan_kembali.config(state=tk.NORMAL, bg="#1a73e8", fg="white")
                
        else:
            messagebox.showwarning(
                "Buku Tidak Sedang Dipinjam", 
                f"Buku dengan kode '{kode_input}' tidak terdeteksi sedang dipinjam oleh siswa/guru mana pun."
            )

    def eksekusi_simpan_pengembalian(self):
        if not self.daftar_kembali_dijepit:
            messagebox.showwarning("Kosong", "Tidak ada antrean buku untuk dikembalikan.")
            return
            
        if messagebox.askyesno("Konfirmasi Pengembalian", f"Proses pengembalian resmi untuk {len(self.daftar_kembali_dijepit)} buku ini?"):
            try:
                conn = sqlite3.connect(DB_PATH)
                cursor = conn.cursor()
                tgl_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for b in self.daftar_kembali_dijepit:
                    cursor.execute(
                        "UPDATE data_peminjaman_buku SET status = 'KEMBALI', tanggal_kembali = ? WHERE rowid = ? OR id = ?", 
                        (tgl_sekarang, b['id_pk'], b['id_pk'])
                    )
                
                conn.commit()
                conn.close()
                
                self.jalankan_auto_backup_silent()
                messagebox.showinfo("Pengembalian Sukses", f"Berhasil mengembalikan {len(self.daftar_kembali_dijepit)} buku ke dalam database!")
                
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
    # MODUL 4: REPRINT KOLEKTIF DARI NISN (MAKS 9 SISWA DI A4)
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
    # MODUL REKAP & LIVE ANALYTICS (ON-DEMAND SEARCH & AGREGASI KELOMPOK)
    # =====================================================================
    def tampilkan_layar_rekap(self):
        self.bersihkan_layar()
        
        banner = tk.Label(self.root, text="MODUL REKAP SIRKULASI & KONTROL PEMERATAAN BUKU", font=("Arial", 14, "bold"), bg="#007afc", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("UPDATE data_peminjaman_buku SET nisn = TRIM(nisn), kode_satuan_buku = TRIM(kode_satuan_buku), status = TRIM(status)")
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
        self.df_master_siswa = pd.read_sql_query("SELECT * FROM data_siswa_perpus", conn)
        conn.close()

        tab_lap = ttk.Notebook(self.root)
        tab_matriks_rombel = tk.Frame(tab_lap, bg="#f4f6f9")
        tab_per_buku = tk.Frame(tab_lap, bg="#f4f6f9")
        
        tab_lap.add(tab_matriks_rombel, text="  [1] REKAP PER ROMBEL (CEKLIS KELENGKAPAN)  ")
        tab_lap.add(tab_per_buku, text="  [2] REKAP PER BUKU (AGREGASI DISTRIBUSI KELOMPOK)  ")
        tab_lap.pack(expand=1, fill="both", padx=15, pady=10)

        # -----------------------------------------------------------------
        # TAB 1: REKAP PER ROMBEL (MATRIKS CEKLIS BUKU PAKET)
        # -----------------------------------------------------------------
        daftar_rombel = sorted(self.df_master_siswa['rombel'].dropna().unique()) if not self.df_master_siswa.empty else []
        
        frame_fltr_rombel = tk.Frame(tab_matriks_rombel, bg="#f4f6f9", pady=5)
        frame_fltr_rombel.pack(fill=tk.X)
        
        tk.Label(frame_fltr_rombel, text="Pilih Rombel Target:", font=("Arial", 10, "bold"), bg="#f4f6f9").pack(side=tk.LEFT, padx=5)
        self.cb_rombel_matriks = ttk.Combobox(frame_fltr_rombel, values=daftar_rombel, state="readonly", font=("Arial", 10), width=18)
        self.cb_rombel_matriks.pack(side=tk.LEFT, padx=5)
        self.cb_rombel_matriks.bind("<<ComboboxSelected>>", self.muat_matriks_rekap_rombel)

        frame_btn_exp_r = tk.Frame(frame_fltr_rombel, bg="#f4f6f9")
        frame_btn_exp_r.pack(side=tk.RIGHT, padx=5)
        
        self.btn_exp_r_excel = tk.Button(frame_btn_exp_r, text="📊 Ekspor Excel", font=("Arial", 9, "bold"), bg="#188038", fg="white", state=tk.DISABLED, command=self.ekspor_matriks_rombel_excel)
        self.btn_exp_r_excel.pack(side=tk.LEFT, padx=3)
        self.btn_exp_r_pdf = tk.Button(frame_btn_exp_r, text="📄 Ekspor PDF", font=("Arial", 9, "bold"), bg="#d93025", fg="white", state=tk.DISABLED, command=self.ekspor_matriks_rombel_pdf)
        self.btn_exp_r_pdf.pack(side=tk.LEFT, padx=3)

        frame_tabel_m = tk.Frame(tab_matriks_rombel, bg="#f4f6f9")
        frame_tabel_m.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.tree_matriks = ttk.Treeview(frame_tabel_m, show="headings", height=12)
        scrollbar_m_y = ttk.Scrollbar(frame_tabel_m, orient="vertical", command=self.tree_matriks.yview)
        scrollbar_m_x = ttk.Scrollbar(frame_tabel_m, orient="horizontal", command=self.tree_matriks.xview)
        self.tree_matriks.configure(yscrollcommand=scrollbar_m_y.set, xscrollcommand=scrollbar_m_x.set)
        
        scrollbar_m_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_m_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_matriks.pack(fill=tk.BOTH, expand=True)

        # -----------------------------------------------------------------
        # TAB 2: REKAP PER BUKU (DENGAN RINCIAN KODE SATUAN & ON-DEMAND)
        # -----------------------------------------------------------------
        frame_fltr_buku = tk.Frame(tab_per_buku, bg="#f4f6f9", pady=5)
        frame_fltr_buku.pack(fill=tk.X)
        
        tk.Label(frame_fltr_buku, text="Cari Judul / Kode Induk:", font=("Arial", 10, "bold"), bg="#f4f6f9").pack(side=tk.LEFT, padx=5)
        
        self.ent_cari_buku_rekap = tk.Entry(frame_fltr_buku, font=("Arial", 10), width=30)
        self.ent_cari_buku_rekap.pack(side=tk.LEFT, padx=5)
        self.ent_cari_buku_rekap.bind("<Return>", self.muat_rekap_per_buku) # Hanya mencari saat Enter ditekan
        
        tk.Button(frame_fltr_buku, text="🔍 Cari Rekap", font=("Arial", 9, "bold"), bg="#5c6bc0", fg="white", command=self.muat_rekap_per_buku).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_fltr_buku, text="Tampilkan Semua", font=("Arial", 9), bg="#e8eaed", command=self.reset_rekap_per_buku).pack(side=tk.LEFT, padx=5)

        frame_btn_exp_b = tk.Frame(frame_fltr_buku, bg="#f4f6f9")
        frame_btn_exp_b.pack(side=tk.RIGHT, padx=5)
        
        tk.Button(frame_btn_exp_b, text="📊 Ekspor Excel", font=("Arial", 9, "bold"), bg="#188038", fg="white", command=self.ekspor_per_buku_excel).pack(side=tk.LEFT, padx=3)
        tk.Button(frame_btn_exp_b, text="📄 Ekspor PDF", font=("Arial", 9, "bold"), bg="#d93025", fg="white", command=self.ekspor_per_buku_pdf).pack(side=tk.LEFT, padx=3)
        
        frame_tabel_b = tk.Frame(tab_per_buku, bg="#f4f6f9")
        frame_tabel_b.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.tree_per_buku = ttk.Treeview(frame_tabel_b, columns=("Kode", "Judul", "Rombel", "Siswa", "Satuan"), show="headings", height=12)
        self.tree_per_buku.heading("Kode", text="Kelompok Kode Induk")
        self.tree_per_buku.heading("Judul", text="Judul Buku")
        self.tree_per_buku.heading("Rombel", text="Rombel")
        self.tree_per_buku.heading("Siswa", text="Peminjam")
        self.tree_per_buku.heading("Satuan", text="Daftar Kode Satuan Buku")
        
        self.tree_per_buku.column("Kode", width=130, anchor=tk.CENTER)
        self.tree_per_buku.column("Judul", width=250, anchor=tk.W)
        self.tree_per_buku.column("Rombel", width=80, anchor=tk.CENTER)
        self.tree_per_buku.column("Siswa", width=180, anchor=tk.W)
        self.tree_per_buku.column("Satuan", width=180, anchor=tk.W)
        
        sb_b_y = ttk.Scrollbar(frame_tabel_b, orient="vertical", command=self.tree_per_buku.yview)
        self.tree_per_buku.configure(yscrollcommand=sb_b_y.set)
        sb_b_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_per_buku.pack(fill=tk.BOTH, expand=True)

        self.muat_rekap_per_buku()

        frame_aksi = tk.Frame(self.root, bg="#f4f6f9", pady=5)
        frame_aksi.pack(fill=tk.X, padx=15)
        tk.Button(frame_aksi, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=6, command=self.tampilkan_menu_utama).pack(side=tk.LEFT)

    def muat_matriks_rekap_rombel(self, event=None):
        rombel_pilihan = self.cb_rombel_matriks.get()
        if not rombel_pilihan: return
        
        df_siswa_rombel = self.df_master_siswa[self.df_master_siswa['rombel'] == rombel_pilihan].copy()
        if df_siswa_rombel.empty:
            messagebox.showinfo("Kosong", f"Tidak ada data siswa terdaftar di rombel {rombel_pilihan}")
            self.btn_exp_r_excel.config(state=tk.DISABLED)
            self.btn_exp_r_pdf.config(state=tk.DISABLED)
            return

        df_log_active = self.df_rekap_log[
            (self.df_rekap_log['rombel'] == rombel_pilihan) & 
            (self.df_rekap_log['status'].str.upper() == 'DIPINJAM')
        ].copy()
        
        df_log_active['Judul_Buku'] = df_log_active['kode_satuan_buku'].apply(lambda x: self.dapatkan_judul_buku(str(x)).strip().upper())

        daftar_judul_unik = sorted(list(df_log_active['Judul_Buku'].unique())) if not df_log_active.empty else []
        
        kolom_header = ["NISN", "Nama Siswa"] + daftar_judul_unik
        self.tree_matriks.config(columns=kolom_header)
        
        for item in self.tree_matriks.get_children():
            self.tree_matriks.delete(item)
            
        self.tree_matriks.heading("NISN", text="NISN")
        self.tree_matriks.column("NISN", width=110, anchor=tk.CENTER)
        self.tree_matriks.heading("Nama Siswa", text="Nama Lengkap Siswa")
        self.tree_matriks.column("Nama Siswa", width=220, anchor=tk.W)
        
        for j_title in daftar_judul_unik:
            self.tree_matriks.heading(j_title, text=j_title)
            self.tree_matriks.column(j_title, width=150, anchor=tk.CENTER)
            
        self.data_matriks_export = []
        for _, r_siswa in df_siswa_rombel.iterrows():
            nisn = str(r_siswa['nisn']).split('.')[0].strip()
            nama = str(r_siswa['nama']).upper()
            nisn_clean = self.bersihkan_nisn_ke_string(nisn)
            
            row_vals = [nisn, nama]
            for j_title in daftar_judul_unik:
                has_book = False
                if not df_log_active.empty:
                    df_cek = df_log_active[
                        (df_log_active['nisn'].apply(self.bersihkan_nisn_ke_string) == nisn_clean) & 
                        (df_log_active['Judul_Buku'] == j_title)
                    ]
                    if not df_cek.empty:
                        has_book = True
                row_vals.append("✓" if has_book else "✗")
                
            self.tree_matriks.insert("", tk.END, values=row_vals)
            self.data_matriks_export.append(row_vals)
            
        self.headers_matriks_export = kolom_header
        self.btn_exp_r_excel.config(state=tk.NORMAL)
        self.btn_exp_r_pdf.config(state=tk.NORMAL)

    def reset_rekap_per_buku(self):
        if hasattr(self, 'ent_cari_buku_rekap'):
            self.ent_cari_buku_rekap.delete(0, tk.END)
        self.muat_rekap_per_buku()

    def muat_rekap_per_buku(self, event=None):
        keyword = self.ent_cari_buku_rekap.get().strip().upper() if hasattr(self, 'ent_cari_buku_rekap') else ""
        
        for item in self.tree_per_buku.get_children():
            self.tree_per_buku.delete(item)
            
        if self.df_rekap_log.empty: return
        
        df_log_active = self.df_rekap_log[self.df_rekap_log['status'].str.upper() == 'DIPINJAM'].copy()
        if df_log_active.empty: return
        
        df_log_active['Judul_Buku'] = df_log_active['kode_satuan_buku'].apply(lambda x: self.dapatkan_judul_buku(str(x)).strip().upper())
        df_log_active['Kode_Induk'] = df_log_active['kode_satuan_buku'].apply(self.ekstrak_kode_induk)
        
        if keyword:
            df_log_active = df_log_active[
                (df_log_active['Kode_Induk'].str.upper().str.contains(keyword)) | 
                (df_log_active['Judul_Buku'].str.contains(keyword)) |
                (df_log_active['nama_siswa'].str.upper().str.contains(keyword)) |
                (df_log_active['rombel'].str.upper().str.contains(keyword)) |
                (df_log_active['kode_satuan_buku'].str.upper().str.contains(keyword))
            ]
            
        df_agregat = df_log_active.groupby(['Kode_Induk', 'Judul_Buku', 'rombel', 'nama_siswa'])['kode_satuan_buku'].apply(lambda x: ", ".join(sorted(x.unique()))).reset_index(name='Daftar_Satuan')
        df_agregat = df_agregat.sort_values(by=['Judul_Buku', 'rombel'])
            
        self.data_per_buku_export = []
        for _, r in df_agregat.iterrows():
            row_data = [
                r['Kode_Induk'],
                r['Judul_Buku'],
                r['rombel'],
                r['nama_siswa'],
                r['Daftar_Satuan']
            ]
            self.tree_per_buku.insert("", tk.END, values=row_data)
            self.data_per_buku_export.append(row_data)

    # =====================================================================
    # FITUR EKSPOR EXCEL & PDF KELENGKAPAN ROMBEL DAN REKAP BUKU
    # =====================================================================
    def ekspor_matriks_rombel_excel(self):
        rombel = self.cb_rombel_matriks.get()
        if not hasattr(self, 'data_matriks_export') or not self.data_matriks_export:
            messagebox.showwarning("Kosong", "Tidak ada data matriks untuk diekspor.")
            return
            
        filename = os.path.join(DIR_PERPUS, f"Rekap_Matriks_Kelengkapan_{rombel}.xlsx")
        try:
            df = pd.DataFrame(self.data_matriks_export, columns=self.headers_matriks_export)
            df.to_excel(filename, index=False)
            messagebox.showinfo("Sukses", f"Berhasil mengekspor matriks {rombel} ke Excel:\n{filename}")
            os.startfile(DIR_PERPUS)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor Excel: {e}")

    def ekspor_matriks_rombel_pdf(self):
        rombel = self.cb_rombel_matriks.get()
        if not hasattr(self, 'data_matriks_export') or not self.data_matriks_export:
            messagebox.showwarning("Kosong", "Tidak ada data matriks untuk diekspor.")
            return
            
        filename = os.path.join(DIR_PERPUS, f"Rekap_Matriks_Kelengkapan_{rombel}.pdf")
        try:
            doc = SimpleDocTemplate(filename, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
            elements = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=14, leading=16, alignment=1)
            sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=12, alignment=1)
            cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, leading=8)

            elements.append(Paragraph(f"LAPORAN MATRIKS KELENGKAPAN BUKU PAKET", title_style))
            elements.append(Paragraph(f"SMK WALISONGO JAKARTA - ROMBEL {rombel}", sub_style))
            elements.append(Spacer(1, 10))

            headers = [Paragraph(h, cell_bold) for h in self.headers_matriks_export]
            data_table = [headers]

            for row in self.data_matriks_export:
                row_cells = []
                for idx, val in enumerate(row):
                    if idx >= 2:
                        color = "green" if val == "✓" else "red"
                        row_cells.append(Paragraph(f"<font color='{color}'><b>{val}</b></font>", cell_style))
                    else:
                        row_cells.append(Paragraph(str(val), cell_style))
                data_table.append(row_cells)

            t = Table(data_table, repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#e8eaed")),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
            ]))
            elements.append(t)

            doc.build(elements)
            messagebox.showinfo("Sukses", f"Berhasil mencetak PDF matriks {rombel}:\n{filename}")
            os.startfile(filename)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mencetak PDF: {e}")

    def ekspor_per_buku_excel(self):
        if not hasattr(self, 'data_per_buku_export') or not self.data_per_buku_export:
            messagebox.showwarning("Kosong", "Tidak ada data buku untuk diekspor.")
            return
            
        filename = os.path.join(DIR_PERPUS, f"Laporan_Distribusi_Peminjaman_Buku.xlsx")
        try:
            df = pd.DataFrame(self.data_per_buku_export, columns=["Kelompok Kode Induk", "Judul Buku", "Rombel", "Nama Siswa Peminjam", "Daftar Kode Satuan Buku"])
            df.to_excel(filename, index=False)
            messagebox.showinfo("Sukses", f"Berhasil mengekspor Laporan Distribusi Buku ke Excel:\n{filename}")
            os.startfile(DIR_PERPUS)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor Excel: {e}")

    def ekspor_per_buku_pdf(self):
        if not hasattr(self, 'data_per_buku_export') or not self.data_per_buku_export:
            messagebox.showwarning("Kosong", "Tidak ada data buku untuk diekspor.")
            return
            
        filename = os.path.join(DIR_PERPUS, f"Laporan_Distribusi_Peminjaman_Buku.pdf")
        try:
            doc = SimpleDocTemplate(filename, pagesize=landscape(A4), leftMargin=20, rightMargin=20, topMargin=20, bottomMargin=20)
            elements = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=13, leading=15, alignment=1)
            sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=11, alignment=1)
            cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, leading=10)

            elements.append(Paragraph("LAPORAN DISTRIBUSI PEMINJAMAN BUKU PERPUSTAKAAN", title_style))
            elements.append(Paragraph("SMK WALISONGO JAKARTA", sub_style))
            elements.append(Spacer(1, 10))

            headers = [
                Paragraph("Kelompok Kode Induk", cell_bold), Paragraph("Judul Buku", cell_bold),
                Paragraph("Rombel", cell_bold), Paragraph("Nama Siswa Peminjam", cell_bold), Paragraph("Daftar Kode Satuan Buku", cell_bold)
            ]
            data_table = [headers]

            for row in self.data_per_buku_export:
                data_table.append([
                    Paragraph(str(row[0]), cell_style), Paragraph(str(row[1]), cell_style),
                    Paragraph(str(row[2]), cell_style), Paragraph(str(row[3]), cell_style),
                    Paragraph(str(row[4]), cell_style)
                ])

            t = Table(data_table, colWidths=[110, 220, 70, 160, 210], repeatRows=1)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#e8eaed")),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
            ]))
            elements.append(t)

            doc.build(elements)
            messagebox.showinfo("Sukses", f"Berhasil mencetak PDF Laporan Distribusi Buku:\n{filename}")
            os.startfile(filename)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mencetak PDF: {e}")

    # =====================================================================
    # UTILITAS PROSES BACKGROUND MASSAL LABEL KISSCUT A3+ (DENGAN OPSI IMPOR)
    # =====================================================================
    def tampilkan_layar_proses_cetak_buku(self):
        self.bersihkan_layar()
        
        banner = tk.Label(self.root, text="MODUL CETAK STIKER LABEL BUKU MASSAL A3+", font=("Arial", 14, "bold"), bg="#202124", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        frame_opsi = tk.LabelFrame(self.root, text=" Pilih Metode Penarikan Data Buku ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=20, pady=20)
        frame_opsi.pack(expand=True, pady=10)
        
        tk.Button(
            frame_opsi, 
            text="📊 Cetak Seluruh Label Buku dari Database SQLite\n(Membaca semua data master yang terdaftar)", 
            font=("Arial", 11, "bold"), 
            bg="#1a73e8", 
            fg="white", 
            width=50, 
            height=3, 
            command=lambda: self.buat_layar_monitoring("PROSES MASAL CETAK LABEL BUKU A3+ (ALL DATA)", self.eksekusi_back_end_cetak_buku)
        ).pack(pady=10)
        
        tk.Button(
            frame_opsi, 
            text="🆕 Cetak Khusus Buku Baru via Impor File Excel (.xlsx)\n(Hanya mencetak data di dalam file tanpa duplikasi data lama)", 
            font=("Arial", 11, "bold"), 
            bg="#34a853", 
            fg="white", 
            width=50, 
            height=3, 
            command=self.pemicu_cetak_buku_baru_excel
        ).pack(pady=10)
        
        tk.Button(self.root, text="Kembali ke Menu Utama", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=8, command=self.tampilkan_menu_utama).pack(pady=20)

    def pemicu_cetak_buku_baru_excel(self):
        path_excel = filedialog.askopenfilename(
            title="Pilih File Excel Buku Baru",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not path_excel:
            return
            
        self.buat_layar_monitoring(
            "PROSES CETAK LABEL KISSCUT A3+ (IMPOR BARU)", 
            lambda: self.eksekusi_cetak_buku_dari_excel(path_excel)
        )

    def eksekusi_cetak_buku_dari_excel(self, path_excel):
        output_pdf = os.path.join(DIR_PERPUS, "Cetak_Label_Buku_Baru_Impor.pdf")
        output_mal = os.path.join(DIR_PERPUS, "Mal_Kisscut_Buku_Baru_Impor.pdf")
        
        try:
            self.log_ke_gui(self.txt_log, f"Membuka dan memvalidasi file: {os.path.basename(path_excel)}...")
            
            df = pd.read_excel(path_excel, header=None, dtype=str)
            df.columns = ["kolom_0", "qr_code", "kolom_2", "label_buku", "jurusan"][:len(df.columns)]
            
            rows_data = [r for i, r in df.iterrows() if pd.notna(r['label_buku'])]
            total_data = len(rows_data)
            
            if total_data == 0:
                self.log_ke_gui(self.txt_log, "[ERROR] File Excel kosong atau kolom 'label_buku' tidak teridentifikasi.")
                return
                
            self.log_ke_gui(self.txt_log, f"Ditemukan {total_data} baris data buku baru siap cetak.")
            
            paper_width, paper_height = 320 * MM_TO_POINT, 480 * MM_TO_POINT
            box_width = box_height = 25 * MM_TO_POINT
            qr_render_size = 17.5 * MM_TO_POINT
            gap_x, gap_y = 5 * MM_TO_POINT, 7 * MM_TO_POINT
            margin_left, margin_top = 10 * MM_TO_POINT, 12 * MM_TO_POINT
            
            c_design = canvas.Canvas(output_pdf, pagesize=(paper_width, paper_height))
            c_mal = canvas.Canvas(output_mal, pagesize=(paper_width, paper_height))
            
            col_idx, row_idx, total = 0, 0, 0
            for row in rows_data:
                qr_text = str(row['qr_code']) if pd.notna(row['qr_code']) else str(row['label_buku'])
                label_text = str(row['label_buku'])
                jurusan = str(row['jurusan']).strip().upper() if pd.notna(row['jurusan']) else ""
                
                x = margin_left + (col_idx * (box_width + gap_x))
                y = paper_height - margin_top - box_height - (row_idx * (box_height + gap_y))
                
                bg_color = (0.88, 0.82, 0.98) if jurusan in ["TKJ", "TEKNO", "TJKT"] else (1.0, 0.8, 0.85) if jurusan in ["MP", "BISMEN", "MPLB"] else (0.75, 0.93, 0.75) if jurusan == "BD" else (1.0, 0.96, 0.7) if jurusan == "AKL" else (1.0, 1.0, 1.0)
                
                c_design.setFillColorRGB(*bg_color)
                c_design.setStrokeColorRGB(0.9, 0.9, 0.9)
                c_design.rect(x, y, box_width, box_height, stroke=1, fill=1)
                
                c_mal.setStrokeColorRGB(1.0, 0.0, 1.0)
                c_mal.setLineWidth(0.3)
                c_mal.rect(x, y, box_width, box_height, stroke=1, fill=0)
                
                qr_x = x + ((box_width - qr_render_size) / 2)
                qr_y = y + box_height - qr_render_size - (1.0 * MM_TO_POINT)
                
                try:
                    encoded = urllib.parse.quote(qr_text)
                    req = urllib.request.Request(f"https://quickchart.io/qr?text={encoded}&size=150&margin=0", headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=5) as response:
                        c_design.drawImage(ImageReader(BytesIO(response.read())), qr_x, qr_y, width=qr_render_size, height=qr_render_size)
                except:
                    c_design.line(qr_x, qr_y, qr_x + qr_render_size, qr_y + qr_render_size)
                    
                c_design.setFillColorRGB(0.0, 0.0, 0.0)
                c_design.setFont("Helvetica-Bold", 6.0)
                c_design.drawCentredString(x + (box_width / 2), y + (2.0 * MM_TO_POINT), label_text)
                
                total += 1
                col_idx += 1
                persentase = int((total / total_data) * 100)
                self.progress_bar['value'] = persentase
                
                if col_idx >= 10:
                    col_idx = 0
                    row_idx += 1
                if row_idx >= 14:
                    c_design.showPage()
                    c_mal.showPage()
                    col_idx, row_idx = 0, 0
                    
            c_design.save()
            c_mal.save()
            
            self.log_ke_gui(self.txt_log, "✅ Pembuatan PDF stiker khusus buku baru selesai!")
            messagebox.showinfo("Sukses", "Stiker label khusus buku baru berhasil diproduksi!")
            os.startfile(output_pdf)
            
        except Exception as e:
            messagebox.showerror("Error Impor", f"Gagal memproses file cetak baru: {e}")
        finally:
            self.btn_kembali_menu.config(state=tk.NORMAL, bg="#1a73e8")

    def buat_layar_monitoring(self, judul_modul, target_fungsi):
        self.bersihkan_layar()
        banner = tk.Label(self.root, text=judul_modul, font=("Arial", 14, "bold"), bg="#202124", fg="white", pady=10)
        banner.pack(fill=tk.X)
        frame_progress = tk.Frame(self.root, bg="#f4f6f9", pady=15)
        frame_progress.pack(fill=tk.X, padx=20)
        self.lbl_status_progress = tk.Label(frame_progress, text="Mempersiapkan data...", font=("Arial", 11), bg="#f4f6f9")
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
                
                bg_color = (0.88, 0.82, 0.98) if jurusan in ["TKJ", "TEKNO", "TJKT"] else (1.0, 0.8, 0.85) if jurusan in ["MP", "BISMEN", "MPLB"] else (0.75, 0.93, 0.75) if jurusan == "BD" else (1.0, 0.96, 0.7) if jurusan == "AKL" else (1.0, 1.0, 1.0)
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