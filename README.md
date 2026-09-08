/**
 * SISTEM DAFTAR HADIR BIMBEL - BACKEND GOOGLE APPS SCRIPT
 * File: Kode.gs
 * Description: Backend script for managing spreadsheet database, CRUD operations, authentication, and attendance records.
 */

function doGet(e) {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('Index');
  } catch (err) {
    try {
      template = HtmlService.createTemplateFromFile('index');
    } catch (err2) {
      template = HtmlService.createTemplateFromFile('Index.html');
    }
  }
  return template
    .evaluate()
    .setTitle('Sistem Absensi Bimbel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Automatic Database Setup Script
 * Run this function once from Google Apps Script editor to initialize sheets and default data.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet structure definitions
  const sheets = {
    'PENGATURAN': [
      ['Key', 'Value'],
      ['nama_sekolah', 'Bimbel Prestasi Utama'],
      ['logo_url', 'https://cdn-icons-png.flaticon.com/512/2991/2991106.png'],
      ['admin_username', 'admin'],
      ['admin_password', 'admin123'],
      ['materi_list', 'Matematika, Bahasa Inggris, IPA, IPS, Fisika, Kimia']
    ],
    'GURU': [
      ['ID', 'NIP', 'Nama', 'Email', 'Password', 'Telepon', 'Status'],
      ['G001', '19850101', 'Ahmad Subagja, S.Pd', 'ahmad@bimbel.com', 'guru123', '081234567890', 'Aktif'],
      ['G002', '19900202', 'Siti Rahma, M.Pd', 'siti@bimbel.com', 'guru123', '082345678901', 'Aktif']
    ],
    'BIMBEL': [
      ['ID', 'Nama Bimbel', 'Materi', 'Keterangan'],
      ['B001', 'Reguler SD', 'Matematika, IPA', 'Bimbel harian SD'],
      ['B002', 'Intensif SMP', 'Matematika, Bahasa Inggris, Fisika', 'Persiapan ujian SMP'],
      ['B003', 'Focus SMA', 'Fisika, Kimia, Matematika', 'Persiapan masuk PTN']
    ],
    'KELAS': [
      ['ID', 'Nama Kelas', 'Tingkat'],
      ['K001', 'Kelas 6 SD - A', 'SD'],
      ['K002', 'Kelas 9 SMP - B', 'SMP'],
      ['K003', 'Kelas 12 SMA - IPA 1', 'SMA']
    ],
    'SISWA': [
      ['ID', 'NIS', 'Nama Siswa', 'Kelas ID', 'Bimbel ID'],
      ['S001', '1001', 'Budi Santoso', 'Kelas 6 SD - A', 'Reguler SD'],
      ['S002', '1002', 'Ani Wijaya', 'Kelas 6 SD - A', 'Reguler SD'],
      ['S003', '1003', 'Citra Dewi', 'Kelas 9 SMP - B', 'Intensif SMP'],
      ['S004', '1004', 'Deni Pratama', 'Kelas 12 SMA - IPA 1', 'Focus SMA']
    ],
    'ABSENSI': [
      ['ID', 'Tanggal', 'Waktu', 'Guru ID', 'Nama Guru', 'Siswa ID', 'Nama Siswa', 'Kelas', 'Bimbel', 'Status', 'Catatan']
    ]
  };

  for (const sheetName in sheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const data = sheets[sheetName];
      sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      sheet.getRange(1, 1, 1, data[0].length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
  }

  // Remove default sheet if present
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Lembar1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }

  return { status: 'success', message: 'Database setup berhasil dijalankan!' };
}

/**
 * Login Handler for Admin and Teacher
 */
function apiLogin(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check Admin Login
    const settingSheet = ss.getSheetByName('PENGATURAN');
    const settingData = settingSheet.getDataRange().getValues();
    let adminUser = 'admin';
    let adminPass = 'admin123';
    
    for (let i = 1; i < settingData.length; i++) {
      if (settingData[i][0] === 'admin_username') adminUser = settingData[i][1];
      if (settingData[i][0] === 'admin_password') adminPass = settingData[i][1];
    }
    
    if (username === adminUser && password === adminPass) {
      return { status: 'success', role: 'admin', user: { name: 'Administrator System', email: adminUser } };
    }
    
    // Check Teacher Login
    const guruSheet = ss.getSheetByName('GURU');
    const guruData = guruSheet.getDataRange().getValues();
    for (let i = 1; i < guruData.length; i++) {
      const email = guruData[i][3];
      const pass = guruData[i][4];
      const nama = guruData[i][2];
      const id = guruData[i][0];
      const status = guruData[i][6];
      
      if ((username === email || username === guruData[i][1]) && password === pass) {
        if (status !== 'Aktif') {
          return { status: 'error', message: 'Akun guru tidak aktif! Hubungi admin.' };
        }
        return { status: 'success', role: 'guru', user: { id: id, name: nama, email: email } };
      }
    }
    
    return { status: 'error', message: 'Username/Email atau Password salah!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal login: ' + err.toString() };
  }
}

/**
 * Fetch all initial application state (Guru, Bimbel, Kelas, Siswa, Pengaturan, Dashboard Stats)
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Fetch Settings
    const settings = {};
    const settingSheet = ss.getSheetByName('PENGATURAN');
    if (settingSheet) {
      const settingValues = settingSheet.getDataRange().getValues();
      for (let i = 1; i < settingValues.length; i++) {
        settings[settingValues[i][0]] = settingValues[i][1];
      }
    }

    // Helper to get array of objects
    function sheetToObjects(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) return [];
      const headers = values[0];
      const result = [];
      for (let i = 1; i < values.length; i++) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
          obj[headers[j]] = values[i][j];
        }
        result.push(obj);
      }
      return result;
    }

    const listGuru = sheetToObjects('GURU');
    const listBimbel = sheetToObjects('BIMBEL');
    const listKelas = sheetToObjects('KELAS');
    const listSiswa = sheetToObjects('SISWA');
    const listAbsensi = sheetToObjects('ABSENSI');

    // Calculate today stats
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const todayAbsen = listAbsensi.filter(a => {
      let t = a['Tanggal'];
      if (t instanceof Date) {
        t = Utilities.formatDate(t, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return t === todayStr;
    });

    const stats = {
      totalGuru: listGuru.length,
      totalBimbel: listBimbel.length,
      totalKelas: listKelas.length,
      totalSiswa: listSiswa.length,
      absenHariIni: todayAbsen.length,
      hadirToday: todayAbsen.filter(a => a['Status'] === 'H').length,
      izinToday: todayAbsen.filter(a => a['Status'] === 'I').length,
      sakitToday: todayAbsen.filter(a => a['Status'] === 'S').length,
      alphaToday: todayAbsen.filter(a => a['Status'] === 'A').length
    };

    return {
      status: 'success',
      data: {
        settings,
        guru: listGuru,
        bimbel: listBimbel,
        kelas: listKelas,
        siswa: listSiswa,
        absensi: listAbsensi,
        stats
      }
    };
  } catch (err) {
    return { status: 'error', message: 'Error loading data: ' + err.toString() };
  }
}

/**
 * Generic CRUD Handler for Master Data
 */
function apiSaveMaster(sheetName, action, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName(sheetName);
    }
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    
    if (action === 'CREATE') {
      // Auto Generate ID if not provided
      if (!payload['ID']) {
        payload['ID'] = sheetName.substring(0, 1) + String(values.length).padStart(3, '0');
      }
      const newRow = headers.map(h => payload[h] !== undefined ? payload[h] : '');
      sheet.appendRow(newRow);
      return { status: 'success', message: 'Data berhasil ditambahkan!', id: payload['ID'] };
    } 
    
    if (action === 'UPDATE') {
      const idCol = 0; // First column is ID
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idCol]) === String(payload['ID'])) {
          const rowNum = i + 1;
          headers.forEach((h, colIdx) => {
            if (payload[h] !== undefined) {
              sheet.getRange(rowNum, colIdx + 1).setValue(payload[h]);
            }
          });
          return { status: 'success', message: 'Data berhasil diperbarui!' };
        }
      }
      return { status: 'error', message: 'ID tidak ditemukan!' };
    }
    
    if (action === 'DELETE') {
      const idCol = 0;
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][idCol]) === String(payload['ID'])) {
          sheet.deleteRow(i + 1);
          return { status: 'success', message: 'Data berhasil dihapus!' };
        }
      }
      return { status: 'error', message: 'ID tidak ditemukan!' };
    }
    
    return { status: 'error', message: 'Aksi tidak valid' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

/**
 * Import batch list of Siswa or Bimbel from Excel payload
 */
function apiBatchImport(sheetName, itemsList) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'error', message: 'Sheet tidak ditemukan!' };
    
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    let count = 0;

    itemsList.forEach((item, index) => {
      if (!item['ID']) {
        item['ID'] = sheetName.substring(0, 1) + String(values.length + count).padStart(3, '0');
      }
      const row = headers.map(h => item[h] !== undefined ? item[h] : '');
      sheet.appendRow(row);
      count++;
    });

    return { status: 'success', message: `Berhasil mengimpor ${count} data!` };
  } catch (err) {
    return { status: 'error', message: 'Gagal impor: ' + err.toString() };
  }
}

/**
 * Save presence transaction (Absensi Batch)
 */
function apiSimpanAbsenBatch(guruId, namaGuru, tanggal, bimbel, items) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('ABSENSI');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('ABSENSI');
    }
    
    const nowTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "HH:mm:ss");
    
    items.forEach(item => {
      const idAbsen = 'ABS-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        idAbsen,
        tanggal,
        nowTime,
        guruId,
        namaGuru,
        item.siswaId,
        item.namaSiswa,
        item.kelas,
        bimbel,
        item.status,
        item.catatan || ''
      ]);
    });

    return { status: 'success', message: `Absensi untuk ${items.length} siswa berhasil disimpan!` };
  } catch (err) {
    return { status: 'error', message: 'Gagal menyimpan absensi: ' + err.toString() };
  }
}

/**
 * Save System Settings
 */
function apiSaveSettings(settingsObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('PENGATURAN');
    if (!sheet) {
      setupDatabase();
      sheet = ss.getSheetByName('PENGATURAN');
    }
    
    sheet.clear();
    sheet.appendRow(['Key', 'Value']);
    
    for (const key in settingsObj) {
      sheet.appendRow([key, settingsObj[key]]);
    }
    
    return { status: 'success', message: 'Pengaturan berhasil disimpan!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
