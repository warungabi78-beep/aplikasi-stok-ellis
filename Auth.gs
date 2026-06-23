/**
 * ============================================================================
 * ELLIS ESTETIKA - AUTHENTICATION MODULE (WITH CLINIC MAPPING SUPPORT)
 * ============================================================================
 */

/**
 * Memproses verifikasi masuk dan memetakan klinik terkait ke dalam sesi asinkron.
 * Mendukung pencarian kolom ID_KLINIK secara opsional pada sheet LOGIN.
 */
function processLogin(username, password) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("LOGIN");
    if (!sheet) return { success: false, error: "Lembar kerja LOGIN tidak ditemukan." };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    
    const userCol = headers.indexOf("username");
    const passCol = headers.indexOf("password");
    const roleCol = headers.indexOf("role");
    const clinicCol = headers.indexOf("id_klinik"); // Kolom baru opsional untuk mapping klinik
    
    if (userCol === -1 || passCol === -1 || roleCol === -1) {
      return { success: false, error: "Struktur kolom pada sheet LOGIN salah." };
    }
    
    for (let i = 1; i < data.length; i++) {
      const u = data[i][userCol].toString().trim();
      const p = data[i][passCol].toString().trim();
      
      if (u.toLowerCase() === username.toLowerCase() && p === password) {
        const role = data[i][roleCol].toString().trim().toUpperCase();
        let id_klinik = "";
        
        // Ambil data ID Klinik jika kolomnya tersedia
        if (clinicCol !== -1) {
          id_klinik = data[i][clinicCol].toString().trim();
        }
        
        return {
          success: true,
          username: u,
          role: role,
          id_klinik: id_klinik
        };
      }
    }
    return { success: false, error: "Nama pengguna atau Password/PIN salah." };
  } catch (e) {
    return { success: false, error: "Gagal memproses otentikasi: " + e.toString() };
  }
}
