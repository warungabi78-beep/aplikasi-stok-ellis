/**
 * ============================================================================
 * ELLIS ESTETIKA - MODULAR CONFIGURATION & HELPER (WITH NATIVE PDF GENERATOR)
 * ============================================================================
 */
const SPREADSHEET_ID = "1Kx1C4jzrDxz71gT_tedcPRBg6-JA5KqSBGPoT-kWDZY";

/**
 * Membuka Spreadsheet secara aman dengan verifikasi hak akses terpusat.
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    throw new Error("Gagal terhubung ke Google Spreadsheet. Pastikan ID Spreadsheet benar.");
  }
}

/**
 * Merender halaman utama HTML Web App dengan evaluasi skrip modul.
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('Ellis Estetika - Sistem Gudang & Klinik')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    return HtmlService.createHtmlOutput("<h2>Terjadi kesalahan kompilasi Server:</h2><p>" + e.toString() + "</p>");
  }
}

/**
 * Mengikutsertakan file HTML modular secara dinamis menggunakan Template Evaluation.
 */
function include(filename) {
  try {
    return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
  } catch (e) {
    return "<!-- Gagal memuat file modul: " + filename + " -->";
  }
}

/**
 * Mengonversi data sheet menjadi array objek secara dinamis.
 */
function readSheetAsObjects(sheetName) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      for (let j = 0; j < headers.length; j++) {
        let val = data[i][j];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
        }
        obj[headers[j]] = val;
        obj["_col" + j] = val;
      }
      result.push(obj);
    }
    return result;
  } catch (e) {
    Logger.log("Error readSheetAsObjects pada sheet '" + sheetName + "': " + e.toString());
    return [];
  }
}

/**
 * Memperoleh seluruh data aplikasi awal untuk mempercepat inisialisasi sisi klien.
 * PENTING: Memuat sheet "RETURN_BARANG" agar history retur tampil di layar HP.
 */
function getInitialAppData() {
  try {
    return {
      success: true,
      stok: readSheetAsObjects("STOK"),
      klinik: readSheetAsObjects("KLINIK"),
      supplier: readSheetAsObjects("SUPPLIER"),
      order: readSheetAsObjects("ORDER"),
      faktur: readSheetAsObjects("FAKTUR"),
      return_barang: readSheetAsObjects("RETURN_BARANG") // Integrasi data retur barang
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ============================================================================
 * FITUR REKAP PDF - MURNI GOOGLE APPS SCRIPT (OWNER ONLY)
 * ============================================================================
 */
function parseDateString(dateStr) {
  const parts = dateStr.split('-');
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

function parseAnyDate(val) {
  if (val instanceof Date) return val;
  const str = val.toString().trim();
  const parts = str.split('/');
  if (parts.length === 3) {
    return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? new Date(0) : d;
}

function formatDateDisplay(d) {
  return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
}

function formatRupiah(num) {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

/**
 * Membuat PDF Laporan Rekap Faktur Pembelian berdasarkan rentang tanggal
 */
function generatePdf(startDateStr, endDateStr) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("FAKTUR");
    if (!sheet) return { success: false, error: "Sheet FAKTUR tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: "Tidak ada data transaksi faktur." };
    
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    const tglCol = headers.indexOf("tanggal");
    const noCol = headers.indexOf("no_faktur");
    const supCol = headers.indexOf("supplier");
    const totCol = headers.indexOf("total_harga");
    const ketCol = headers.indexOf("keterangan");
    
    if (tglCol === -1 || noCol === -1 || supCol === -1 || totCol === -1) {
      return { success: false, error: "Struktur kolom pada sheet FAKTUR tidak sesuai." };
    }
    
    const startDate = parseDateString(startDateStr);
    const endDate = parseDateString(endDateStr);
    
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);
    
    let matchingRows = [];
    let grandTotal = 0;
    
    for (let i = 1; i < data.length; i++) {
      const rawDate = data[i][tglCol];
      const itemDate = parseAnyDate(rawDate);
      
      if (itemDate >= startDate && itemDate <= endDate) {
        const noFaktur = data[i][noCol];
        const supplier = data[i][supCol];
        const totalHarga = Number(data[i][totCol]) || 0;
        const keterangan = ketCol !== -1 ? data[i][ketCol].toString() : "-";
        
        matchingRows.push({
          tanggal: formatDateDisplay(itemDate),
          noFaktur: noFaktur,
          supplier: supplier,
          totalHarga: totalHarga,
          keterangan: keterangan
        });
        grandTotal += totalHarga;
      }
    }
    
    if (matchingRows.length === 0) {
      return { success: false, error: "Tidak ditemukan transaksi faktur pada rentang tanggal tersebut." };
    }
    
    matchingRows.sort((a, b) => {
      const d1 = a.tanggal.split('/');
      const d2 = b.tanggal.split('/');
      return new Date(d1[2], d1[1]-1, d1[0]) - new Date(d2[2], d2[1]-1, d2[0]);
    });
    
    let rowsHtml = "";
    matchingRows.forEach(r => {
      rowsHtml += `
        <tr>
          <td style="border: 1px solid #e2e8f0; padding: 10px; color: #4a5568;">${r.tanggal}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; color: #4a5568;">${r.noFaktur}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; color: #4a5568;">${r.supplier}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; text-align: right; color: #1a202c; font-weight: bold;">Rp ${formatRupiah(r.totalHarga)}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; color: #4a5568;">${r.keterangan}</td>
        </tr>
      `;
    });
    
    const htmlContent = `
      <html>
      <head>
        <style>
          body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; margin: 40px; color: #2d3748; }
          .header-box { text-align: center; border-bottom: 3px double #D9B56B; padding-bottom: 15px; margin-bottom: 25px; }
          h2 { color: #D9B56B; font-size: 24px; font-weight: bold; margin: 0 0 5px 0; text-transform: uppercase; }
          .subtitle { font-size: 13px; color: #718096; font-weight: bold; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th { background-color: #D9B56B; color: white; padding: 12px; text-align: left; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
          td { font-size: 11px; }
          .grand-total-box { margin-top: 30px; text-align: right; border-top: 2px solid #D9B56B; padding-top: 15px; }
          .grand-total-label { font-size: 13px; color: #4a5568; font-weight: bold; }
          .grand-total-val { font-size: 18px; color: #D9B56B; font-weight: 900; margin-left: 10px; }
        </style>
      </head>
      <body>
        <div class="header-box">
          <h2>REKAP FAKTUR PEMBELIAN</h2>
          <div class="subtitle">PERIODE: ${formatDateDisplay(startDate)} s/d ${formatDateDisplay(endDate)}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="border: 1px solid #D9B56B; width: 12%;">TANGGAL</th>
              <th style="border: 1px solid #D9B56B; width: 20%;">NO FAKTUR</th>
              <th style="border: 1px solid #D9B56B; width: 23%;">SUPPLIER</th>
              <th style="border: 1px solid #D9B56B; text-align: right; width: 20%;">TOTAL HARGA</th>
              <th style="border: 1px solid #D9B56B; width: 25%;">KETERANGAN</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <div class="grand-total-box">
          <span class="grand-total-label">GRAND TOTAL:</span>
          <span class="grand-total-val">Rp ${formatRupiah(grandTotal)}</span>
        </div>
      </body>
      </html>
    `;
    
    const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName(`REKAP_FAKTUR_${startDateStr}_TO_${endDateStr}.pdf`);
    
    const folderId = "1lIZUiZHPOVxh_aQzJBJbmumnD8XgG-0X";
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(pdfBlob);
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${file.getId()}`;
    const viewUrl = file.getUrl();
    
    return {
      success: true,
      url: downloadUrl,
      viewUrl: viewUrl,
      filename: file.getName()
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ============================================================================
 * GENERATE PDF STOK BARANG & AUTO ARCHIVE (ADMIN & OWNER ONLY)
 * ============================================================================
 */
function generateStokPdf() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("STOK");
    if (!sheet) return { success: false, error: "Sheet STOK tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: false, error: "Tidak ada data stok." };
    
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    const nameColIdx = headers.indexOf("nama_barang");
    const qtyColIdx = headers.indexOf("qty");
    const satColIdx = headers.indexOf("satuan");
    
    if (nameColIdx === -1 || qtyColIdx === -1) {
      return { success: false, error: "Struktur kolom pada sheet STOK tidak sesuai." };
    }
    
    let rowsHtml = "";
    let totalItems = 0;
    
    for (let i = 1; i < data.length; i++) {
      const name = data[i][nameColIdx];
      const qty = Number(data[i][qtyColIdx]) || 0;
      const satuan = satColIdx !== -1 ? data[i][satColIdx].toString() : "PCS";
      
      rowsHtml += `
        <tr>
          <td style="border: 1px solid #e2e8f0; padding: 10px; color: #4a5568; font-weight: bold; text-transform: uppercase;">${name}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; text-align: center; color: #1a202c; font-weight: bold; font-size: 12px;">${qty}</td>
          <td style="border: 1px solid #e2e8f0; padding: 10px; text-align: center; color: #718096; text-transform: uppercase;">${satuan}</td>
        </tr>
      `;
      totalItems++;
    }
    
    const today = new Date();
    const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    
    const htmlContent = `
      <html>
      <head>
        <style>
          body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; margin: 40px; color: #2d3748; }
          .header-box { text-align: center; border-bottom: 3px double #D9B56B; padding-bottom: 15px; margin-bottom: 25px; }
          h2 { color: #D9B56B; font-size: 24px; font-weight: bold; margin: 0 0 5px 0; text-transform: uppercase; }
          .subtitle { font-size: 13px; color: #718096; font-weight: bold; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th { background-color: #D9B56B; color: white; padding: 12px; text-align: left; font-size: 11px; font-weight: bold; text-transform: uppercase; letter-spacing: 0.5px; }
          td { font-size: 11px; }
          .grand-total-box { margin-top: 30px; text-align: right; border-top: 2px solid #D9B56B; padding-top: 15px; }
          .grand-total-label { font-size: 13px; color: #4a5568; font-weight: bold; }
        </style>
      </head>
      <body>
        <div class="header-box">
          <h2>LAPORAN STOK TERBARU</h2>
          <div class="subtitle">WAKTU CETAK: ${dateStr}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="border: 1px solid #D9B56B; width: 60%;">NAMA BARANG</th>
              <th style="border: 1px solid #D9B56B; text-align: center; width: 20%;">STOK</th>
              <th style="border: 1px solid #D9B56B; text-align: center; width: 20%;">SATUAN</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <div class="grand-total-box">
          <span class="grand-total-label">TOTAL JENIS BARANG: ${totalItems}</span>
        </div>
      </body>
      </html>
    `;
    
    const pdfBlob = HtmlService.createHtmlOutput(htmlContent)
      .getAs('application/pdf')
      .setName("Stok_terbaru.pdf");
    
    const shareFolderId = "1Es5CXoAzcgtP4W96jUMkvWKbRtf0Nted";
    const shareFolder = DriveApp.getFolderById(shareFolderId);
    
    const existingFiles = shareFolder.getFilesByName("Stok_terbaru.pdf");
    while (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setTrashed(true);
    }
    
    const newShareFile = shareFolder.createFile(pdfBlob);
    newShareFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const archiveFolderId = "1DEwMARwxIyA90bi-Cdkch95VZXEAb3wG";
    const archiveFolder = DriveApp.getFolderById(archiveFolderId);
    
    const timestamp = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd_MM_yyyy_HHmmss");
    const archiveFileName = `Stok_Arsip_${timestamp}.pdf`;
    
    const archiveBlob = pdfBlob.copyBlob().setName(archiveFileName);
    const archiveFile = archiveFolder.createFile(archiveBlob);
    archiveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${newShareFile.getId()}`;
    const viewUrl = newShareFile.getUrl();
    
    return {
      success: true,
      url: downloadUrl,
      viewUrl: viewUrl,
      filename: "Stok_terbaru.pdf",
      archiveFilename: archiveFileName
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * NEW: FUNGSI TES UNTUK MEMICU DIALOG OTORISASI GOOGLE DRIVE
 */
function testDrivePermission() {
  try {
    const folderId = "1lIZUiZHPOVxh_aQzJBJbmumnD8XgG-0X";
    const folder = DriveApp.getFolderById(folderId);
    Logger.log("Koneksi berhasil! Folder tujuan ditemukan: " + folder.getName());
    return { success: true, message: "Akses diizinkan ke folder: " + folder.getName() };
  } catch (e) {
    Logger.log("Otorisasi Gagal: " + e.toString());
    return { success: false, error: e.toString() };
  }
}
