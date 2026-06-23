/**
 * ============================================================================
 * ELLIS ESTETIKA - TRANSACTIONS MODULE (WITH ATOMIC RETURN BARANG SYSTEM)
 * ============================================================================
 */

/**
 * Menyimpan data barang masuk ke sheet ENTRY dan menambah stok di sheet STOK.
 */
function simpanBarangMasuk(record) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("ENTRY");
    if (!sheet) return { success: false, error: "Lembar kerja ENTRY tidak ditemukan." };
    
    const headers = sheet.getDataRange().getValues()[0].map(h => h.toString().trim().toLowerCase());
    const newRow = [];
    const nextId = "E" + Date.now();
    
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      let val = "";
      if (h === "id_entry") val = nextId;
      else if (h === "tanggal") val = record.tanggal;
      else if (h === "id_supplier") val = record.id_supplier;
      else if (h === "id_barang") val = record.id_barang;
      else if (h === "nama_barang") val = record.nama_barang;
      else if (h === "qty_masuk") val = Number(record.qty_masuk);
      else if (h === "no_faktur") val = record.no_faktur;
      else if (h === "expired") val = record.expired;
      else if (h === "note") val = record.note;
      newRow.push(val);
    }
    
    sheet.appendRow(newRow);
    adjustStockBalance(record.id_barang, Number(record.qty_masuk));
    
    return { success: true };
  } catch (e) {
    return { success: false, error: "Gagal menyimpan barang masuk: " + e.toString() };
  }
}

/**
 * Menyimpan pesanan keluar ke sheet ORDER dan mengurangi stok di sheet STOK.
 * TELAH DIOPTIMALKAN DENGAN LOCK SERVICE & VALIDASI SERVER REAL-TIME (ATOMIC TRANSACTION)
 */
function simpanOrder(record) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(15000);
  } catch (e) {
    return { 
      success: false, 
      error: "Sistem sedang padat melayani antrean order lain. Silakan tekan SIMPAN kembali dalam beberapa detik." 
    };
  }

  try {
    const ss = getSpreadsheet();
    const sheetStok = ss.getSheetByName("STOK");
    if (!sheetStok) return { success: false, error: "Lembar kerja STOK tidak ditemukan." };
    
    const stokData = sheetStok.getDataRange().getValues();
    if (stokData.length <= 1) return { success: false, error: "Daftar produk pada database stok kosong." };
    
    const stokHeaders = stokData[0].map(h => h.toString().trim().toLowerCase());
    const idColIdx = stokHeaders.indexOf("id_barang");
    const nameColIdx = stokHeaders.indexOf("nama_barang");
    const qtyColIdx = stokHeaders.indexOf("qty");
    const satColIdx = stokHeaders.indexOf("satuan");
    
    if (idColIdx === -1 || qtyColIdx === -1) {
      return { success: false, error: "Struktur kolom database STOK tidak sesuai." };
    }
    
    let itemRowIdx = -1;
    let currentQty = 0;
    let namaBarang = record.nama_barang || "";
    let satuan = record.satuan || "PCS";
    const targetId = record.id_barang.toString().trim().toUpperCase();
    
    for (let k = 1; k < stokData.length; k++) {
      if (stokData[k][idColIdx].toString().trim().toUpperCase() === targetId) {
        itemRowIdx = k + 1;
        currentQty = Number(stokData[k][qtyColIdx]) || 0;
        if (nameColIdx !== -1) namaBarang = stokData[k][nameColIdx].toString().trim();
        if (satColIdx !== -1) satuan = stokData[k][satColIdx].toString().trim();
        break;
      }
    }
    
    if (itemRowIdx === -1) {
      return { success: false, error: "ID Produk '" + record.id_barang + "' tidak ditemukan dalam katalog gudang." };
    }
    
    const qtyRequested = Number(record.qty_keluar);
    if (isNaN(qtyRequested) || qtyRequested <= 0) {
      return { success: false, error: "Kuantitas pesanan tidak valid." };
    }
    
    if (currentQty < qtyRequested) {
      return { 
        success: false, 
        error: "Stok tidak mencukupi. Sisa stok aktual saat ini: " + currentQty + " " + satuan 
      };
    }
    
    const nextQty = currentQty - qtyRequested;
    sheetStok.getRange(itemRowIdx, qtyColIdx + 1).setValue(nextQty);
    
    const sheetOrder = ss.getSheetByName("ORDER");
    if (!sheetOrder) return { success: false, error: "Lembar kerja ORDER tidak ditemukan." };
    
    const orderHeaders = sheetOrder.getDataRange().getValues()[0].map(h => h.toString().trim().toLowerCase());
    const newRow = [];
    const nextId = "O" + Date.now();
    
    for (let i = 0; i < orderHeaders.length; i++) {
      const h = orderHeaders[i];
      let val = "";
      if (h === "id_order") val = nextId;
      else if (h === "tanggal") val = record.tanggal;
      else if (h === "id_klinik") val = record.id_klinik;
      else if (h === "id_barang") val = record.id_barang;
      else if (h === "nama_barang") val = namaBarang;
      else if (h === "qty_keluar") val = qtyRequested;
      else if (h === "satuan") val = satuan;
      else if (h === "status") val = "BELUM";
      newRow.push(val);
    }
    
    sheetOrder.appendRow(newRow);
    SpreadsheetApp.flush();
    
    return { success: true };
  } catch (e) {
    return { success: false, error: "Gagal memproses transaksi internal: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * MENYIMPAN TRANSAKSI RETURN BARANG (DENGAN LOCKSERVICE & AUTO TAMBAH STOK DI GUDANG UTAMA)
 */
function simpanReturn(record) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(15000);
  } catch (e) {
    return { 
      success: false, 
      error: "Sistem sedang padat melayani transaksi lain. Silakan tekan PROSES RETUR kembali." 
    };
  }

  try {
    const ss = getSpreadsheet();
    const sheetStok = ss.getSheetByName("STOK");
    if (!sheetStok) return { success: false, error: "Lembar kerja STOK tidak ditemukan." };
    
    const stokData = sheetStok.getDataRange().getValues();
    const stokHeaders = stokData[0].map(h => h.toString().trim().toLowerCase());
    const idColIdx = stokHeaders.indexOf("id_barang");
    const qtyColIdx = stokHeaders.indexOf("qty");
    
    if (idColIdx === -1 || qtyColIdx === -1) {
      return { success: false, error: "Format kolom database STOK tidak sesuai." };
    }
    
    let itemRowIdx = -1;
    let currentQty = 0;
    const targetId = record.id_barang.toString().trim().toUpperCase();
    
    for (let k = 1; k < stokData.length; k++) {
      if (stokData[k][idColIdx].toString().trim().toUpperCase() === targetId) {
        itemRowIdx = k + 1;
        currentQty = Number(stokData[k][qtyColIdx]) || 0;
        break;
      }
    }
    
    if (itemRowIdx === -1) {
      return { success: false, error: "ID Produk '" + record.id_barang + "' tidak ditemukan dalam katalog gudang." };
    }
    
    const qtyReturn = Number(record.qty);
    if (isNaN(qtyReturn) || qtyReturn <= 0) {
      return { success: false, error: "Kuantitas retur tidak valid." };
    }
    
    // Tambahkan kembali kuantitas retur ke stok pusat
    const nextQty = currentQty + qtyReturn;
    sheetStok.getRange(itemRowIdx, qtyColIdx + 1).setValue(nextQty);
    
    // Catat ke sheet RETURN_BARANG
    const sheetReturn = ss.getSheetByName("RETURN_BARANG");
    if (!sheetReturn) return { success: false, error: "Lembar kerja RETURN_BARANG tidak ditemukan." };
    
    const returnHeaders = sheetReturn.getDataRange().getValues()[0].map(h => h.toString().trim().toLowerCase());
    const newRow = [];
    const nextId = "R" + Date.now();
    
    for (let i = 0; i < returnHeaders.length; i++) {
      const h = returnHeaders[i];
      let val = "";
      if (h === "id_return") val = nextId;
      else if (h === "tanggal") val = record.tanggal;
      else if (h === "klinik") val = record.klinik;
      else if (h === "nama_barang") val = record.nama_barang;
      else if (h === "qty") val = qtyReturn;
      else if (h === "satuan") val = record.satuan;
      else if (h === "expired") val = record.expired;
      else if (h === "alasan") val = record.alasan || "";
      newRow.push(val);
    }
    
    sheetReturn.appendRow(newRow);
    SpreadsheetApp.flush();
    
    return { success: true };
  } catch (e) {
    return { success: false, error: "Gagal menyimpan retur: " + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui status verifikasi pengiriman barang ("BELUM" <-> "SUDAH") di ORDER.
 */
function toggleOrderStatusInSheet(orderId, newStatus) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("ORDER");
    if (!sheet) return { success: false, error: "Lembar kerja ORDER tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    const idColIdx = headers.indexOf("id_order");
    const statusColIdx = headers.indexOf("status");
    
    if (idColIdx === -1 || statusColIdx === -1) {
      return { success: false, error: "Format kolom ID_ORDER atau STATUS tidak ditemukan." };
    }
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIdx].toString().trim() === orderId) {
        sheet.getRange(i + 1, statusColIdx + 1).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, error: "ID Transaksi tidak ditemukan di database." };
  } catch (e) {
    return { success: false, error: "Gagal mengubah status: " + e.toString() };
  }
}

/**
 * Penyesuai kuantitas barang otomatis di dalam sheet STOK.
 */
function adjustStockBalance(idBarang, qtyDelta) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName("STOK");
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  
  const idColIdx = headers.indexOf("id_barang");
  const qtyColIdx = headers.indexOf("qty");
  
  if (idColIdx === -1 || qtyColIdx === -1) return;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColIdx].toString().trim().toUpperCase() === idBarang.toUpperCase()) {
      const currentQty = Number(data[i][qtyColIdx]) || 0;
      const updatedQty = Math.max(0, currentQty + qtyDelta);
      sheet.getRange(i + 1, qtyColIdx + 1).setValue(updatedQty);
      break;
    }
  }
}

/**
 * Menyimpan faktur pembelian baru ke dalam sheet FAKTUR dengan tambahan KETERANGAN.
 */
function simpanFaktur(record) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("FAKTUR");
    if (!sheet) return { success: false, error: "Lembar kerja FAKTUR tidak ditemukan." };
    
    const headers = sheet.getDataRange().getValues()[0].map(h => h.toString().trim().toLowerCase());
    const newRow = [];
    
    for (let i = 0; i < headers.length; i++) {
      const h = headers[i];
      let val = "";
      if (h === "tanggal") val = record.tanggal;
      else if (h === "no_faktur") val = record.no_faktur;
      else if (h === "supplier") val = record.supplier;
      else if (h === "total_harga") val = Number(record.total_harga);
      else if (h === "keterangan") val = record.keterangan;
      newRow.push(val);
    }
    
    sheet.appendRow(newRow);
    return { success: true };
  } catch (e) {
    return { success: false, error: "Gagal menyimpan faktur: " + e.toString() };
  }
}
