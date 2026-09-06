/**
 * SISTEM MANAJEMEN PELANGGAN WIFI PT. JALUR DATA INDONESIA (JDI)
 * Backend Google Apps Script (Code.gs)
 * Menangani database Google Sheets, API endpoints, sinkronisasi realtime, dan setup database.
 */

// Nama-nama sheet yang digunakan
const SHEETS = {
  PELANGGAN: 'Pelanggan',
  KAS_MASUK: 'Kas_Masuk',
  KAS_KELUAR: 'Kas_Keluar',
  RESERVASI: 'Reservasi_Area',
  PENGATURAN: 'Pengaturan',
  PAKET: 'Paket_Internet'
};

/**
 * Endpoint GET utama untuk menyajikan HTML Web App
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PT. JALUR DATA INDONESIA - Internet & WiFi Provider')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Inisialisasi Database Otomatis pada Spreadsheet aktif
 * Jalankan fungsi ini sekali dari editor Google Apps Script!
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Pelanggan
  let sheetPelanggan = ss.getSheetByName(SHEETS.PELANGGAN);
  if (!sheetPelanggan) {
    sheetPelanggan = ss.insertSheet(SHEETS.PELANGGAN);
    sheetPelanggan.appendRow([
      'ID_Pelanggan', 'Nama_Lengkap', 'Nomor_WA', 'Alamat', 'Paket', 'Harga', 
      'Tgl_Pemasangan', 'Tgl_Jatuh_Tempo', 'Status_Pelanggan', 'Status_Pembayaran', 'Foto_Pelanggan'
    ]);
    sheetPelanggan.getRange("A1:K1").setFontWeight("bold").setBackground("#e0f2fe");
  }

  // 2. Sheet Kas Masuk
  let sheetKasMasuk = ss.getSheetByName(SHEETS.KAS_MASUK);
  if (!sheetKasMasuk) {
    sheetKasMasuk = ss.insertSheet(SHEETS.KAS_MASUK);
    sheetKasMasuk.appendRow(['ID_Transaksi', 'Tanggal', 'ID_Pelanggan', 'Nama_Pelanggan', 'Keterangan', 'Metode_Bayar', 'Jumlah']);
    sheetKasMasuk.getRange("A1:G1").setFontWeight("bold").setBackground("#dcfce7");
  }

  // 3. Sheet Kas Keluar
  let sheetKasKeluar = ss.getSheetByName(SHEETS.KAS_KELUAR);
  if (!sheetKasKeluar) {
    sheetKasKeluar = ss.insertSheet(SHEETS.KAS_KELUAR);
    sheetKasKeluar.appendRow(['ID_Transaksi', 'Tanggal', 'Kategori', 'Keperluan', 'Jumlah', 'Penanggung_Jawab']);
    sheetKasKeluar.getRange("A1:F1").setFontWeight("bold").setBackground("#fee2e2");
  }

  // 4. Sheet Reservasi / Cek Area
  let sheetReservasi = ss.getSheetByName(SHEETS.RESERVASI);
  if (!sheetReservasi) {
    sheetReservasi = ss.insertSheet(SHEETS.RESERVASI);
    sheetReservasi.appendRow(['No_Reservasi', 'Tanggal', 'Nama_Lengkap', 'Nomor_WA', 'Desa', 'Foto_Lokasi', 'Status']);
    sheetReservasi.getRange("A1:G1").setFontWeight("bold").setBackground("#fef3c7");
  }

  // 5. Sheet Paket Internet
  let sheetPaket = ss.getSheetByName(SHEETS.PAKET);
  if (!sheetPaket) {
    sheetPaket = ss.insertSheet(SHEETS.PAKET);
    sheetPaket.appendRow(['ID_Paket', 'Nama_Paket', 'Kecepatan', 'Harga']);
    sheetPaket.appendRow(['PKT-1', '20 Mbps', '20 Mbps', 155000]);
    sheetPaket.appendRow(['PKT-2', '25 Mbps', '25 Mbps', 167000]);
    sheetPaket.appendRow(['PKT-3', '30 Mbps', '30 Mbps', 205000]);
    sheetPaket.appendRow(['PKT-4', '50 Mbps', '50 Mbps', 255000]);
    sheetPaket.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e7ff");
  }

  // 6. Sheet Pengaturan
  let sheetPengaturan = ss.getSheetByName(SHEETS.PENGATURAN);
  if (!sheetPengaturan) {
    sheetPengaturan = ss.insertSheet(SHEETS.PENGATURAN);
    sheetPengaturan.appendRow(['Kunci', 'Nilai']);
    sheetPengaturan.appendRow(['admin_password', 'admin123']);
    sheetPengaturan.appendRow(['logo_base64', '']);
    sheetPengaturan.appendRow(['qris_base64', '']);
    sheetPengaturan.appendRow(['nama_pt', 'PT. JALUR DATA INDONESIA']);
    sheetPengaturan.appendRow(['no_wa_admin', '081272737850']);
    sheetPengaturan.appendRow(['alamat', 'Jl. AMD RT 010 RW 03 No 35 Banyumas Kec. Banyumas Kab. Pringsewu']);
    sheetPengaturan.getRange("A1:B1").setFontWeight("bold").setBackground("#f1f5f9");
  }

  return "Setup Database Berhasil Dibuat!";
}

/**
 * Handler POST untuk API panggilan dari AJAX Frontend
 */
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.payload;

    let result = { success: false, message: 'Action tidak dikenal' };

    switch (action) {
      case 'getDataSemua':
        result = ambilSemuaData();
        break;
      case 'simpanReservasi':
        result = simpanReservasiArea(payload);
        break;
      case 'tambahPelanggan':
        result = tambahPelanggan(payload);
        break;
      case 'updatePelanggan':
        result = updatePelanggan(payload);
        break;
      case 'hapusPelanggan':
        result = hapusPelanggan(payload);
        break;
      case 'catatKasMasuk':
        result = catatKasMasuk(payload);
        break;
      case 'catatKasKeluar':
        result = catatKasKeluar(payload);
        break;
      case 'simpanPaket':
        result = simpanDaftarPaket(payload);
        break;
      case 'simpanPengaturan':
        result = simpanPengaturan(payload);
        break;
      case 'importCSVPelanggan':
        result = importCSV(payload);
        break;
      default:
        result = { success: false, message: 'Aksi tidak ditemukan' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Mengambil semua data untuk sinkronisasi Real-time
 */
function ambilSemuaData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const pelanggan = bacaSheetJson(ss.getSheetByName(SHEETS.PELANGGAN));
  const kasMasuk = bacaSheetJson(ss.getSheetByName(SHEETS.KAS_MASUK));
  const kasKeluar = bacaSheetJson(ss.getSheetByName(SHEETS.KAS_KELUAR));
  const reservasi = bacaSheetJson(ss.getSheetByName(SHEETS.RESERVASI));
  const paket = bacaSheetJson(ss.getSheetByName(SHEETS.PAKET));
  
  // Baca pengaturan
  const sheetPengaturan = ss.getSheetByName(SHEETS.PENGATURAN);
  let pengaturan = {};
  if (sheetPengaturan) {
    const data = sheetPengaturan.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        pengaturan[data[i][0]] = data[i][1];
      }
    }
  }

  return {
    success: true,
    data: {
      pelanggan,
      kasMasuk,
      kasKeluar,
      reservasi,
      paket,
      pengaturan,
      serverTime: new Date().toISOString()
    }
  };
}

/**
 * Helper mengubah range sheet menjadi array of object JSON
 */
function bacaSheetJson(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  const rows = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = {};
    for (let j = 0; j < headers.length; j++) {
      row[headers[j]] = data[i][j];
    }
    rows.push(row);
  }
  return rows;
}

/**
 * Simpan reservasi cek area pelanggan baru
 */
function simpanReservasiArea(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RESERVASI);
  
  // Buat No Reservasi Otomatis (JDI-YYYYMMDD-XXX)
  const tgl = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
  const random = Math.floor(100 + Math.random() * 900);
  const noReservasi = `JDI-${tgl}-${random}`;
  
  const waktu = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    noReservasi,
    waktu,
    payload.namaLengkap,
    payload.nomorWA,
    payload.desa,
    payload.fotoLokasi || '',
    'Menunggu Verifikasi'
  ]);

  return {
    success: true,
    noReservasi: noReservasi,
    message: 'Reservasi berhasil dikirim! Silahkan simpan nomor registrasi Anda.'
  };
}

/**
 * Tambah Pelanggan Baru
 */
function tambahPelanggan(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PELANGGAN);

  const idPelanggan = 'PLG-' + Math.floor(1000 + Math.random() * 9000);
  
  // Hitung jatuh tempo (1 bulan dari tanggal pasang otomatis)
  const tglPasang = new Date(payload.tglPemasangan);
  const tglJatuhTempo = new Date(tglPasang);
  tglJatuhTempo.setMonth(tglJatuhTempo.getMonth() + 1);
  const formatJatuhTempo = Utilities.formatDate(tglJatuhTempo, "Asia/Jakarta", "yyyy-MM-dd");

  sheet.appendRow([
    idPelanggan,
    payload.namaLengkap,
    payload.nomorWA,
    payload.alamat,
    payload.paket,
    payload.harga,
    payload.tglPemasangan,
    formatJatuhTempo,
    payload.statusPelanggan || 'Aktif',
    payload.statusPembayaran || 'Belum Lunas',
    payload.fotoPelanggan || ''
  ]);

  return { success: true, message: 'Pelanggan berhasil ditambahkan!' };
}

/**
 * Update Data Pelanggan
 */
function updatePelanggan(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PELANGGAN);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == payload.idPelanggan) {
      const rowIndex = i + 1;
      
      // Update jatuh tempo otomatis jika tanggal pasang berubah
      let formatJatuhTempo = payload.tglJatuhTempo;
      if (payload.tglPemasangan !== data[i][6]) {
        let tglPasang = new Date(payload.tglPemasangan);
        let tglJt = new Date(tglPasang);
        tglJt.setMonth(tglJt.getMonth() + 1);
        formatJatuhTempo = Utilities.formatDate(tglJt, "Asia/Jakarta", "yyyy-MM-dd");
      }

      sheet.getRange(rowIndex, 2).setValue(payload.namaLengkap);
      sheet.getRange(rowIndex, 3).setValue(payload.nomorWA);
      sheet.getRange(rowIndex, 4).setValue(payload.alamat);
      sheet.getRange(rowIndex, 5).setValue(payload.paket);
      sheet.getRange(rowIndex, 6).setValue(payload.harga);
      sheet.getRange(rowIndex, 7).setValue(payload.tglPemasangan);
      sheet.getRange(rowIndex, 8).setValue(formatJatuhTempo);
      sheet.getRange(rowIndex, 9).setValue(payload.statusPelanggan);
      sheet.getRange(rowIndex, 10).setValue(payload.statusPembayaran);
      if (payload.fotoPelanggan) {
        sheet.getRange(rowIndex, 11).setValue(payload.fotoPelanggan);
      }
      return { success: true, message: 'Data pelanggan berhasil diperbarui!' };
    }
  }

  return { success: false, message: 'ID Pelanggan tidak ditemukan!' };
}

/**
 * Hapus Pelanggan
 */
function hapusPelanggan(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PELANGGAN);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == payload.idPelanggan) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Pelanggan berhasil dihapus!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan' };
}

/**
 * Catat Kas Masuk (Pembayaran iuran / instalasi)
 */
function catatKasMasuk(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.KAS_MASUK);
  const idTrx = 'KM-' + Date.now().toString().slice(-6);

  sheet.appendRow([
    idTrx,
    payload.tanggal || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd"),
    payload.idPelanggan || '-',
    payload.namaPelanggan,
    payload.keterangan,
    payload.metodeBayar,
    Number(payload.jumlah)
  ]);

  // Update status pembayaran pelanggan jika ada ID
  if (payload.idPelanggan && payload.idPelanggan !== '-') {
    const sheetPelanggan = ss.getSheetByName(SHEETS.PELANGGAN);
    const pData = sheetPelanggan.getDataRange().getValues();
    for (let j = 1; j < pData.length; j++) {
      if (pData[j][0] == payload.idPelanggan) {
        sheetPelanggan.getRange(j + 1, 10).setValue('Lunas');
        break;
      }
    }
  }

  return { success: true, message: 'Kas Masuk berhasil dicatat!' };
}

/**
 * Catat Kas Keluar (Operasional, perbaikan kabel, sewa tiang, dll)
 */
function catatKasKeluar(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.KAS_KELUAR);
  const idTrx = 'KK-' + Date.now().toString().slice(-6);

  sheet.appendRow([
    idTrx,
    payload.tanggal || Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd"),
    payload.kategori,
    payload.keperluan,
    Number(payload.jumlah),
    payload.penanggungJawab
  ]);

  return { success: true, message: 'Kas Keluar berhasil dicatat!' };
}

/**
 * Simpan daftar paket internet
 */
function simpanDaftarPaket(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.PAKET);
  sheet.clearContents();
  sheet.appendRow(['ID_Paket', 'Nama_Paket', 'Kecepatan', 'Harga']);

  payload.forEach(item => {
    sheet.appendRow([item.id, item.nama, item.kecepatan, Number(item.harga)]);
  });

  return { success: true, message: 'Daftar paket berhasil disimpan!' };
}

/**
 * Simpan pengaturan admin (Logo, Password, QRIS)
 */
function simpanPengaturan(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PENGATURAN);
  
  const entries = Object.entries(payload);
  const data = sheet.getDataRange().getValues();

  entries.forEach(([key, val]) => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(val);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, val]);
    }
  });

  return { success: true, message: 'Pengaturan berhasil diperbarui!' };
}

/**
 * Import Data Pelanggan via CSV array
 */
function importCSV(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PELANGGAN);
  
  if (!rows || rows.length === 0) {
    return { success: false, message: 'Data CSV kosong' };
  }

  let count = 0;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] && r[1]) {
      sheet.appendRow(r);
      count++;
    }
  }

  return { success: true, message: `Berhasil mengimpor ${count} baris pelanggan.` };
}