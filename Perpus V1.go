/**
 * SISTEM OTOMATISASI PERPUSTAKAAN BUNDLING PAKET
 * Dibuat khusus untuk struktur Google Sheets "DataPerpustakaanWS"
 */

// Konfigurasi Nama Sheet - Pastikan nama di bawah ini sama persis dengan tab di Sheets Anda
const SHEET_BUKU = "Buku";
const SHEET_SIRKULASI = "Sirkulasi";
const SHEET_MASTER_PAKET = "Master_Paket";

/**
 * 1. FUNGSI UTAMA: PEMINJAMAN 1 PAKET BUKU SECARA MASSAL
 * Fungsi ini akan dipanggil oleh AppSheet Automation ketika petugas memilih Paket Buku untuk Siswa.
 * * @param {string} idAnggota - ID/NIS Siswa dari AppSheet
 * @param {string} namaPaket - Nama Paket Buku yang dipilih (e.g., "PAKET-10")
 */
function pinjamPaketMassal(idAnggota, namaPaket) {
  if (!idAnggota || !namaPaket) {
    throw new Error("Gagal: ID Anggota atau Nama Paket tidak boleh kosong.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shMaster = ss.getSheetByName(SHEET_MASTER_PAKET);
  
  if (!shMaster) {
    throw new Error("Sheet '" + SHEET_MASTER_PAKET + "' tidak ditemukan.");
  }
  
  // Ambil semua data pemetaan paket
  const dataMaster = shMaster.getDataRange().getValues();
  
  // Filter mapel/buku apa saja yang terdaftar di dalam paket yang dipilih
  // Kolom 0: Nama Paket, Kolom 1: Kode Depan Mapel (e.g., "10-02")
  const daftarBukuPaket = dataMaster.filter(row => row[0].toString().trim() === namaPaket.trim());
  
  if (daftarBukuPaket.length === 0) {
    throw new Error("Gagal: Nama paket '" + namaPaket + "' tidak terdaftar di Master_Paket.");
  }
  
  let totalBukuBerhasil = 0;
  let bukuGagalList = [];
  
  // Looping untuk memproses setiap jenis buku dalam paket tersebut
  daftarBukuPaket.forEach(buku => {
    const kodeMapel = buku[1].toString().trim(); // Contoh: "10-02"
    
    // Cari eksemplar buku fisik yang saat ini berstatus 'Tersedia'
    const idBukuFisik = cariEksemplarTersedia(kodeMapel);
    
    if (idBukuFisik) {
      // 1. Catat transaksi ke sheet Sirkulasi
      buatLogSirkulasi(idAnggota, idBukuFisik, namaPaket);
      
      // 2. Ubah status buku fisik tersebut menjadi 'Dipinjam'
      ubahStatusBuku(idBukuFisik, "Dipinjam");
      
      totalBukuBerhasil++;
    } else {
      // Jika tidak ada satu pun eksemplar yang tersedia untuk mapel ini
      bukuGagalList.push(kodeMapel);
    }
  });
  
  // Jika ada buku yang habis, beri peringatan log ke petugas
  if (bukuGagalList.length > 0) {
    Logger.log("Peminjaman parsial berhasil. Namun kode mapel berikut habis stok: " + bukuGagalList.join(", "));
    throw new Error("Peringatan: Berhasil meminjam " + totalBukuBerhasil + " buku. Paket tidak lengkap karena stok eksemplar untuk mapel [" + bukuGagalList.join(", ") + "] sedang HABIS di perpustakaan.");
  }
  
  return "Sukses: " + totalBukuBerhasil + " buku dalam " + namaPaket + " berhasil didaftarkan ke Anggota " + idAnggota;
}

/**
 * 2. FUNGSI UTAMA: PENGEMBALIAN BUKU SATUAN (SCAN QR)
 * Fungsi ini dipanggil saat petugas men-scan QR Code pada buku fisik yang dikembalikan siswa.
 * * @param {string} idBukuDiscan - ID Buku unik fisik yang tertera di QR Code (e.g., "10-02-0001")
 */
function kembalikanBukuSatuan(idBukuDiscan) {
  if (!idBukuDiscan) {
    throw new Error("Gagal: ID Buku yang di-scan kosong.");
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSirkulasi = ss.getSheetByName(SHEET_SIRKULASI);
  const dataSirkulasi = shSirkulasi.getDataRange().getValues();
  
  let ditemukan = false;
  
  // Cari baris sirkulasi aktif di mana ID Buku ini berstatus 'Dipinjam'
  for (let i = 1; i < dataSirkulasi.length; i++) {
    // Kolom 2 (C): ID Buku, Kolom 5 (F): Status Transaksi
    if (dataSirkulasi[i][2].toString().trim() === idBukuDiscan.trim() && dataSirkulasi[i][5].toString().trim() === "Dipinjam") {
      
      const barisTarget = i + 1;
      const tanggalSekarang = new Date();
      
      // Update data sirkulasi di baris tersebut
      shSirkulasi.getRange(barisTarget, 5).setValue(tanggalSekarang); // Kolom E: Tanggal Kembali
      shSirkulasi.getRange(barisTarget, 6).setValue("Kembali");       // Kolom F: Status Transaksi
      
      // Kembalikan status buku di database master Buku menjadi 'Tersedia'
      ubahStatusBuku(idBukuDiscan, "Tersedia");
      
      ditemukan = true;
      break;
    }
  }
  
  if (!ditemukan) {
    // PENGAMAN/ANTI-KECURANGAN: Jika buku di-scan tapi statusnya di sirkulasi tidak sedang dipinjam
    throw new Error("SISTEM MENOLAK: Buku [" + idBukuDiscan + "] tidak terdeteksi sedang dipinjam oleh siswa ini (atau mungkin salah buku/milik siswa lain). Periksa fisik buku!");
  }
  
  return "Sukses: Buku " + idBukuDiscan + " telah dikembalikan ke rak.";
}

// =========================================================================
// FUNGSI UTILITAS / PEMBANTU (INTERNAL ENGINE - JANGAN DIUBAH)
// =========================================================================

/**
 * Mencari nomor eksemplar buku yang berstatus 'Tersedia' berdasarkan kode depan mapel
 */
function cariEksemplarTersedia(kodeMapel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBuku = ss.getSheetByName(SHEET_BUKU);
  const dataBuku = shBuku.getDataRange().getValues();
  
  // Looping dari baris kedua (lewati header)
  for (let i = 1; i < dataBuku.length; i++) {
    const idBuku = dataBuku[i][0].toString(); // Kolom A: ID Buku (e.g. 10-02-0001)
    const status = dataBuku[i][4].toString(); // Kolom E: Status Buku
    
    // Cek apakah ID Buku diawali dengan kode mapel (e.g. "10-02") DAN berstatus "Tersedia"
    if (idBuku.startsWith(kodeMapel) && status.trim() === "Tersedia") {
      return idBuku; // Kembalikan ID Eksemplar penuh yang ketemu pertama kali
    }
  }
  return null; // Jika stok eksemplar habis
}

/**
 * Mengubah status buku di Sheet master "Buku"
 */
function ubahStatusBuku(idBuku, statusBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shBuku = ss.getSheetByName(SHEET_BUKU);
  const dataBuku = shBuku.getDataRange().getValues();
  
  for (let i = 1; i < dataBuku.length; i++) {
    if (dataBuku[i][0].toString().trim() === idBuku.trim()) {
      shBuku.getRange(i + 1, 5).setValue(statusBaru); // Kolom E: Status
      return;
    }
  }
}

/**
 * Menulis baris log baru di sheet "Sirkulasi"
 */
function buatLogSirkulasi(idAnggota, idBuku, namaPaket) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSirkulasi = ss.getSheetByName(SHEET_SIRKULASI);
  
  const idTransaksi = "TRX-" + Utilities.getUuid().substring(0, 8).toUpperCase();
  const tanggalPinjam = new Date();
  
  // Format Baris: ID Transaksi, ID Anggota, ID Buku, Tanggal Pinjam, Tanggal Kembali (kosong), Status, Nama Paket
  shSirkulasi.appendRow([
    idTransaksi, 
    idAnggota, 
    idBuku, 
    tanggalPinjam, 
    "", 
    "Dipinjam", 
    namaPaket
  ]);
}