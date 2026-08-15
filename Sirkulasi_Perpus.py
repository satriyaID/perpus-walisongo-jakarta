import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, ttk

# Konfigurasi Folder Kerja Lokal (Drive G)
FOLDER_KERJA = r"G:\2025-2026\WS"
EXCEL_SISWA = os.path.join(FOLDER_KERJA, "Data_Siswa_Perpus.xlsx")
EXCEL_LOG_PINJAM = os.path.join(FOLDER_KERJA, "Data_Peminjaman_Buku.xlsx")

class AplikasiSirkulasiPerpus:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Sirkulasi Buku Paket - SMK Walisongo Jakarta")
        self.root.geometry("700x650")
        self.root.configure(bg="#f4f6f9")
        
        # Standar Jumlah Buku Paket Bawaan Sekolah per Tingkat
        self.target_buku = {"10": 10, "11": 15, "12": 4}
        
        # State Data Transaksi Aktif
        self.siswa_aktif = None
        self.daftar_buku_dijepit = []
        
        # Memuat Database Siswa
        self.load_database_siswa()
        
        # Membuat Komponen Antarmuka (GUI)
        self.create_widgets()
        
    def load_database_siswa(self):
        if os.path.exists(EXCEL_SISWA):
            self.df_siswa = pd.read_excel(EXCEL_SISWA, dtype=str)
            # Bersihkan spasi tak terlihat agar pencarian NISN akurat
            self.df_siswa['NISN'] = self.df_siswa['NISN'].str.strip()
        else:
            messagebox.showerror("Error", f"File database siswa tidak ditemukan di: {EXCEL_SISWA}")
            self.root.destroy()

    def create_widgets(self):
        # Title Banner Atas
        banner = tk.Label(self.root, text="LAYANAN SIRKULASI BUKU PAKET PERPUSTAKAAN", font=("Arial", 14, "bold"), bg="#1a73e8", fg="white", pady=10)
        banner.pack(fill=tk.X)
        
        # --- SEKSI 1: SCAN KARTU SISWA ---
        frame_scan_siswa = tk.LabelFrame(self.root, text=" Langkah 1: Scan Kartu Perpustakaan Siswa ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=10)
        frame_scan_siswa.pack(fill=tk.X, padx=15, pady=10)
        
        tk.Label(frame_scan_siswa, text="Input / Scan NISN:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_nisn = tk.Entry(frame_scan_siswa, font=("Arial", 12), width=30)
        self.ent_nisn.grid(row=0, column=1, padx=10)
        self.ent_nisn.bind("<Return>", self.proses_scan_siswa) # Trigger otomatis saat QR scanner mengirim enter
        self.ent_nisn.focus()
        
        # Info Panel Siswa (Akan berubah hijau jika siswa ditemukan)
        self.lbl_info_siswa = tk.Label(frame_scan_siswa, text="[Silakan scan kartu siswa untuk memulai]", font=("Arial", 11, "italic"), fg="#555", bg="#e8eaed", width=65, height=4, anchor=tk.W, justify=tk.LEFT, padx=10, pady=5)
        self.lbl_info_siswa.grid(row=1, column=0, columnspan=2, pady=10)

        # --- SEKSI 2: SCAN BUKU PAKET ---
        self.frame_scan_buku = tk.LabelFrame(self.root, text=" Langkah 2: Scan Label Buku Paket (Kode Satuan) ", font=("Arial", 10, "bold"), bg="#f4f6f9", padx=10, pady=10)
        self.frame_scan_buku.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        tk.Label(self.frame_scan_buku, text="Scan Label Buku:", font=("Arial", 11), bg="#f4f6f9").grid(row=0, column=0, sticky=tk.W)
        self.ent_buku = tk.Entry(self.frame_scan_buku, font=("Arial", 12), width=30, state=tk.DISABLED)
        self.ent_buku.grid(row=0, column=1, padx=10)
        self.ent_buku.bind("<Return>", self.proses_scan_buku)
        
        # Tabel Visual (Treeview) untuk menampung daftar buku yang sedang di-scan
        self.tree_buku = ttk.Treeview(self.frame_scan_buku, columns=("No", "Kode Buku"), show="headings", height=8)
        self.tree_buku.heading("No", text="No")
        self.tree_buku.heading("Kode Buku", text="Kode Seri / Label Buku")
        self.tree_buku.column("No", width=50, anchor=tk.CENTER)
        self.tree_buku.column("Kode Buku", width=450, anchor=tk.W)
        self.tree_buku.grid(row=1, column=0, columnspan=2, pady=10, sticky="nsew")
        
        # Status Counter Buku
        self.lbl_counter = tk.Label(self.frame_scan_buku, text="Jumlah Buku Di-scan: 0", font=("Arial", 11, "bold"), bg="#f4f6f9", fg="#1a73e8")
        self.lbl_counter.grid(row=2, column=0, sticky=tk.W)
        
        # --- SEKSI 3: TOMBOL AKSI & KONFIRMASI ---
        frame_aksi = tk.Frame(self.root, bg="#f4f6f9", pady=10)
        frame_aksi.pack(fill=tk.X, padx=15)
        
        # Perbaikan parameter dari px/py menjadi padx/pady
        self.btn_simpan = tk.Button(frame_aksi, text="Simpan & Lanjut Siswa Lain", font=("Arial", 11, "bold"), bg="#188038", fg="white", padx=15, pady=8, command=self.konfirmasi_dan_simpan, state=tk.DISABLED)
        self.btn_simpan.pack(side=tk.RIGHT, padx=5)
        
        btn_batal = tk.Button(frame_aksi, text="Reset / Batal", font=("Arial", 11), bg="#d93025", fg="white", padx=15, pady=8, command=self.reset_form)
        btn_batal.pack(side=tk.RIGHT, padx=5)

    def proses_scan_siswa(self, event=None):
        nisn_input = self.ent_nisn.get().strip()
        if not nisn_input:
            return
            
        # Cari data siswa berdasarkan NISN di master data
        hasil = self.df_siswa[self.df_siswa['NISN'] == nisn_input]
        
        if not hasil.empty:
            row = hasil.iloc[0]
            self.siswa_aktif = {
                "NISN": row['NISN'],
                "Nama": row['Nama'].upper(),
                "Jurusan": row['Jurusan'],
                "Rombel": row['Rombel'],
                "Kelas": str(row['Rombel']).split('-')[0]
            }
            
            # Ekstrak angka kelas untuk menentukan target buku wajib
            kelas_digit = "".join(filter(str.isdigit, self.siswa_aktif["Kelas"]))
            self.target_seharusnya = self.target_buku.get(kelas_digit, 10)
            
            # Update panel informasi siswa menjadi hijau tanda sukses
            info_text = f"Nama    : {self.siswa_aktif['Nama']}\nNISN    : {self.siswa_aktif['NISN']}\nRombel  : {self.siswa_aktif['Rombel']} | Jurusan: {self.siswa_aktif['Jurusan']}\nTarget Wajib: {self.target_seharusnya} Buku Paket"
            self.lbl_info_siswa.config(text=info_text, fg="#155724", bg="#d4edda", font=("Arial", 11, "bold"))
            
            # Kunci kolom siswa, lalu buka dan pindahkan fokus ke kolom buku
            self.ent_buku.config(state=tk.NORMAL)
            self.ent_buku.focus()
            self.ent_nisn.config(state=tk.DISABLED)
        else:
            messagebox.showwarning("Data Tidak Ditemukan", f"Siswa dengan NISN {nisn_input} tidak terdaftar di database.")
            self.ent_nisn.delete(0, tk.END)

    def proses_scan_buku(self, event=None):
        kode_buku = self.ent_buku.get().strip()
        if not kode_buku:
            return
            
        # Validasi mencegah double-scan buku yang sama pada siswa yang sama
        if kode_buku in self.daftar_buku_dijepit:
            messagebox.showwarning("Duplikasi", f"Buku {kode_buku} sudah ada di dalam daftar transaksi siswa ini.")
            self.ent_buku.delete(0, tk.END)
            return
            
        self.daftar_buku_dijepit.append(kode_buku)
        
        # Tambahkan data ke tabel visual layar
        no_urut = len(self.daftar_buku_dijepit)
        self.tree_buku.insert("", tk.END, values=(no_urut, kode_buku))
        
        # Perbarui informasi counter jumlah buku
        self.lbl_counter.config(text=f"Jumlah Buku Di-scan: {no_urut} dari {self.target_seharusnya} seharusnya")
        
        # Kosongkan kolom input stiker buku agar siap menerima scan berikutnya
        self.ent_buku.delete(0, tk.END)
        self.btn_simpan.config(state=tk.NORMAL)

    def konfirmasi_dan_simpan(self):
        jumlah_sekarang = len(self.daftar_buku_dijepit)
        
        # Cek kesesuaian jumlah stiker buku paket dengan aturan kurikulum
        if jumlah_sekarang != self.target_seharusnya:
            pesan_konfirmasi = f"Jumlah buku yang di-scan ({jumlah_sekarang} buku) TIDAK SAMA dengan standar kelas {self.siswa_aktif['Kelas']} ({self.target_seharusnya} buku).\n\nApakah Anda tetap ingin menyimpan transaksi sirkulasi ini?"
        else:
            pesan_konfirmasi = f"Jumlah buku sudah sesuai standar target ({jumlah_sekarang} buku).\n\nSimpan data peminjaman paket ini?"
            
        respon = messagebox.askyesno("Konfirmasi Jumlah Buku Paket", pesan_konfirmasi)
        
        if respon:
            self.simpan_ke_excel_log()
            messagebox.showinfo("Sukses", f"Data peminjaman {self.siswa_aktif['Nama']} berhasil dibukukan!")
            self.reset_form()

    def simpan_ke_excel_log(self):
        tgl_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        data_baru = []
        
        # Susun struktur data baris transaksi
        for buku in self.daftar_buku_dijepit:
            data_baru.append({
                "Tanggal Pinjam": tgl_sekarang,
                "NISN": self.siswa_aktif["NISN"],
                "Nama Siswa": self.siswa_aktif["Nama"],
                "Rombel": self.siswa_aktif["Rombel"],
                "Kode Satuan Buku": buku,
                "Status": "DIPINJAM"
            })
            
        df_baru = pd.DataFrame(data_baru)
        
        # Gabungkan data baru ke file log riwayat (jika sudah ada)
        if os.path.exists(EXCEL_LOG_PINJAM):
            df_lama = pd.read_excel(EXCEL_LOG_PINJAM, dtype=str)
            df_total = pd.concat([df_lama, df_baru], ignore_index=False)
        else:
            df_total = df_baru
            
        df_total.to_excel(EXCEL_LOG_PINJAM, index=False)

    def reset_form(self):
        # Kembalikan data state ke kondisi semula
        self.siswa_aktif = None
        self.daftar_buku_dijepit = []
        
        # Bersihkan & normalkan elemen antarmuka (GUI)
        self.ent_nisn.config(state=tk.NORMAL)
        self.ent_nisn.delete(0, tk.END)
        self.ent_nisn.focus()
        
        self.ent_buku.delete(0, tk.END)
        self.ent_buku.config(state=tk.DISABLED)
        
        self.lbl_info_siswa.config(text="[Silakan scan kartu siswa untuk memulai]", fg="#555", bg="#e8eaed", font=("Arial", 11, "italic"))
        self.lbl_counter.config(text="Jumlah Buku Di-scan: 0", fg="#1a73e8")
        
        # Kosongkan data tabel visual di layarmonitor
        for item in self.tree_buku.get_children():
            self.tree_buku.delete(item)
            
        self.btn_simpan.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = AplikasiSirkulasiPerpus(root)
    root.mainloop()