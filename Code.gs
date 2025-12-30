/**
 * ================================================================
 * SISTEM SURAT MENYURAT - BACKEND (Google Apps Script)
 * ================================================================
 * 
 * Deskripsi:
 * Aplikasi web untuk mengelola surat menyurat dengan fitur:
 * - Login system dengan role-based access
 * - Auto-generate nomor agenda
 * - Input surat dari Srikandi atau manual
 * - Dashboard & laporan
 * 
 * Author: [Nama Anda]
 * Created: 26 Desember 2025
 * Version: 1.0
 * 
 * ================================================================
 */


// ================================================================
// KONFIGURASI GLOBAL
// ================================================================

/**
 * Konfigurasi aplikasi
 * Ubah nilai-nilai di sini sesuai kebutuhan
 */
const CONFIG = {
  // Kode klasifikasi default
  DEFAULT_KODE_BELAKANG: '405.29.05',
  
  // Sheet names
  SHEET_USERS: 'USERS',
  SHEET_DATA_SURAT: 'DATA_SURAT',
  SHEET_COUNTER: 'COUNTER',
  SHEET_SESSIONS: 'SESSIONS',
  SHEET_AUDIT_LOG: 'AUDIT_LOG',
  SHEET_MASTER_KODE: 'MASTER_KODE',
  
  // Session timeout (dalam menit)
  SESSION_TIMEOUT: 30,
  
  // App metadata
  APP_NAME: 'Sistem Surat Menyurat',
  APP_VERSION: '1.0'
};


// ================================================================
// UTILITY FUNCTIONS - Helper functions untuk berbagai keperluan
// ================================================================

/**
 * Mendapatkan spreadsheet aktif
 * @return {Spreadsheet} Google Spreadsheet object
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Mendapatkan sheet berdasarkan nama
 * @param {string} sheetName - Nama sheet
 * @return {Sheet} Google Sheet object
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  // Jika sheet tidak ada, buat baru
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    initializeSheet(sheetName, sheet);
  }
  
  return sheet;
}

/**
 * Inisialisasi header sheet baru
 * @param {string} sheetName - Nama sheet
 * @param {Sheet} sheet - Sheet object
 */
function initializeSheet(sheetName, sheet) {
  let headers = [];
  
  switch(sheetName) {
    case CONFIG.SHEET_USERS:
      headers = ['User ID', 'Username', 'Password Hash', 'Nama Lengkap', 'Bidang', 'Role', 'Status', 'Created Date', 'Last Login'];
      break;
      
    case CONFIG.SHEET_DATA_SURAT:
      headers = ['ID', 'Timestamp', 'No Agenda', 'Nomor Surat Lengkap', 'Sumber', 'Kode Depan', 'Kode Belakang', 'Tanggal Surat', 'Bidang', 'Perihal', 'Jenis Surat', 'Tujuan', 'Keterangan', 'Created By', 'Modified By', 'Modified Time'];
      break;
      
    case CONFIG.SHEET_COUNTER:
      headers = ['Tahun', 'Nomor Terakhir', 'Last Updated'];
      // Set nilai awal
      sheet.getRange(2, 1, 1, 3).setValues([[new Date().getFullYear(), 0, new Date()]]);
      break;
      
    case CONFIG.SHEET_SESSIONS:
      headers = ['Session ID', 'User ID', 'Username', 'Login Time', 'Expire Time', 'Last Activity'];
      break;
      
    case CONFIG.SHEET_AUDIT_LOG:
      headers = ['Timestamp', 'User ID', 'Username', 'Action', 'Detail', 'IP Address'];
      break;
      
    case CONFIG.SHEET_MASTER_KODE:
      headers = ['Kode Depan', 'Kode Belakang', 'Deskripsi', 'Bidang'];
      // Set contoh data
      sheet.getRange(2, 1, 3, 4).setValues([
        ['100.1.2', '405.29.05', 'Kepegawaian', 'Kepegawaian'],
        ['200.5.1', '405.29.05', 'Keuangan', 'Keuangan'],
        ['300.2.3', '405.29.05', 'Umum', 'Umum']
      ]);
      break;
  }
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

/**
 * Generate unique ID
 * @param {string} prefix - Prefix untuk ID
 * @return {string} Unique ID
 */
function generateId(prefix = 'ID') {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return `${prefix}${timestamp}${random}`;
}

/**
 * Hash password (simple hash - untuk production gunakan library crypto yang lebih kuat)
 * @param {string} password - Password plain text
 * @return {string} Hashed password
 */
function hashPassword(password) {
  // Untuk production, gunakan library crypto yang lebih kuat
  // Ini hanya contoh sederhana
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password));
}

/**
 * Verifikasi password
 * @param {string} password - Password plain text
 * @param {string} hash - Hashed password
 * @return {boolean} True jika cocok
 */
function verifyPassword(password, hash) {
  return hashPassword(password) === hash;
}

/**
 * Log aktivitas ke audit log
 * @param {string} userId - User ID
 * @param {string} username - Username
 * @param {string} action - Aksi yang dilakukan
 * @param {string} detail - Detail aksi
 */
function logActivity(userId, username, action, detail) {
  const sheet = getSheet(CONFIG.SHEET_AUDIT_LOG);
  sheet.appendRow([
    new Date(),
    userId,
    username,
    action,
    detail,
    Session.getActiveUser().getEmail() || 'Unknown'
  ]);
}


// ================================================================
// WEB APP DEPLOYMENT - Entry point untuk web app
// ================================================================

/**
 * Entry point untuk GET request
 * Menampilkan halaman login atau dashboard
 */
function doGet(e) {
  const sessionId = e.parameter.session;
  
  // Cek apakah user sudah login
  if (sessionId) {
    const session = getSessionById(sessionId);
    if (session && isSessionValid(session)) {
      // User sudah login, tampilkan dashboard
      return showDashboard(session);
    }
  }
  
  // Belum login, tampilkan halaman login
  return showLogin();
}

/**
 * Entry point untuk POST request
 * Handle form submissions
 */
function doPost(e) {
  try {
    const action = e.parameter.action;
    
    if (!action) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'No action specified'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    switch(action) {
      case 'login':
        return handleLogin(e);
      case 'logout':
        return handleLogout(e);
      default:
        return ContentService.createTextOutput(JSON.stringify({
          success: false,
          message: 'Invalid action: ' + action
        })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    Logger.log('doPost ERROR: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Server error: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ================================================================
// AUTHENTICATION - Fungsi untuk login, logout, session management
// ================================================================

/**
 * Tampilkan halaman login
 */
function showLogin() {
  const template = HtmlService.createTemplateFromFile('Login');
  template.appName = CONFIG.APP_NAME;
  return template.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Handle proses login
 * @param {Object} e - Event object dari form
 */
function handleLogin(e) {
  try {
    const username = e.parameter.username;
    const password = e.parameter.password;
    
    Logger.log('Login attempt for user: ' + username);
    
    // Validasi input
    if (!username || !password) {
      Logger.log('Login failed: Missing username or password');
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Username dan password harus diisi'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Cari user di database
    const sheet = getSheet(CONFIG.SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    Logger.log('Searching user in database...');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [userId, dbUsername, passwordHash, namaLengkap, bidang, role, status] = row;
      
      if (dbUsername === username) {
        Logger.log('User found: ' + username + ', status: ' + status);
        
        if (status !== 'active') {
          Logger.log('Login failed: User inactive');
          return ContentService.createTextOutput(JSON.stringify({
            success: false,
            message: 'User tidak aktif. Hubungi administrator.'
          })).setMimeType(ContentService.MimeType.JSON);
        }
        
        // Verifikasi password
        if (verifyPassword(password, passwordHash)) {
          Logger.log('Password verified. Creating session...');
          
          // Login berhasil, buat session
          const sessionId = createSession(userId, username);
          Logger.log('Session created: ' + sessionId);
          
          // Update last login
          sheet.getRange(i + 1, 9).setValue(new Date());
          
          // Log aktivitas
          logActivity(userId, username, 'LOGIN', 'User login ke sistem');
          
          // Redirect ke dashboard
          const webAppUrl = ScriptApp.getService().getUrl();
          const redirectUrl = webAppUrl + '?session=' + sessionId;
          
          Logger.log('Login successful. Redirecting to: ' + redirectUrl);
          
          return ContentService.createTextOutput(JSON.stringify({
            success: true,
            message: 'Login berhasil',
            redirectUrl: redirectUrl
          })).setMimeType(ContentService.MimeType.JSON);
        } else {
          Logger.log('Login failed: Invalid password');
        }
      }
    }
    
    // Login gagal
    Logger.log('Login failed: User not found or invalid credentials');
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Username atau password salah'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('Login ERROR: ' + error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Terjadi kesalahan server: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Buat session baru
 * @param {string} userId - User ID
 * @param {string} username - Username
 * @return {string} Session ID
 */
function createSession(userId, username) {
  const sessionId = generateId('SESSION');
  const now = new Date();
  const expireTime = new Date(now.getTime() + (CONFIG.SESSION_TIMEOUT * 60 * 1000));
  
  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  sheet.appendRow([
    sessionId,
    userId,
    username,
    now,
    expireTime,
    now
  ]);
  
  return sessionId;
}

/**
 * Dapatkan session berdasarkan ID
 * @param {string} sessionId - Session ID
 * @return {Object|null} Session object atau null
 */
function getSessionById(sessionId) {
  const sheet = getSheet(CONFIG.SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === sessionId) {
      return {
        sessionId: row[0],
        userId: row[1],
        username: row[2],
        loginTime: row[3],
        expireTime: row[4],
        lastActivity: row[5],
        rowIndex: i + 1
      };
    }
  }
  
  return null;
}

/**
 * Cek apakah session masih valid
 * @param {Object} session - Session object
 * @return {boolean} True jika valid
 */
function isSessionValid(session) {
  const now = new Date();
  return session && new Date(session.expireTime) > now;
}

/**
 * Update last activity session
 * @param {string} sessionId - Session ID
 */
function updateSessionActivity(sessionId) {
  const session = getSessionById(sessionId);
  if (session) {
    const sheet = getSheet(CONFIG.SHEET_SESSIONS);
    const now = new Date();
    const expireTime = new Date(now.getTime() + (CONFIG.SESSION_TIMEOUT * 60 * 1000));
    
    sheet.getRange(session.rowIndex, 5).setValue(expireTime); // Update expire time
    sheet.getRange(session.rowIndex, 6).setValue(now); // Update last activity
  }
}

/**
 * Handle logout
 * @param {Object} e - Event object
 */
function handleLogout(e) {
  const sessionId = e.parameter.session;
  
  if (sessionId) {
    const session = getSessionById(sessionId);
    if (session) {
      // Hapus session
      const sheet = getSheet(CONFIG.SHEET_SESSIONS);
      sheet.deleteRow(session.rowIndex);
      
      // Log aktivitas
      logActivity(session.userId, session.username, 'LOGOUT', 'User logout dari sistem');
    }
  }
  
  // Redirect ke login
  const webAppUrl = ScriptApp.getService().getUrl();
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    redirectUrl: webAppUrl
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Dapatkan user info berdasarkan user ID
 * @param {string} userId - User ID
 * @return {Object|null} User object atau null
 */
function getUserById(userId) {
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0] === userId) {
      return {
        userId: row[0],
        username: row[1],
        namaLengkap: row[3],
        bidang: row[4],
        role: row[5],
        status: row[6]
      };
    }
  }
  
  return null;
}


// ================================================================
// DASHBOARD - Fungsi untuk menampilkan dashboard
// ================================================================

/**
 * Tampilkan dashboard
 * @param {Object} session - Session object
 */
function showDashboard(session) {
  // Update session activity
  updateSessionActivity(session.sessionId);
  
  // Get user info
  const user = getUserById(session.userId);
  
  if (!user) {
    return showLogin();
  }
  
  // Create template
  const template = HtmlService.createTemplateFromFile('Dashboard');
  template.appName = CONFIG.APP_NAME;
  template.user = user;
  template.sessionId = session.sessionId;
  
  // Get dashboard data
  template.dashboardData = getDashboardData(user);
  
  return template.evaluate()
    .setTitle(`${CONFIG.APP_NAME} - Dashboard`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Dapatkan data untuk dashboard
 * @param {Object} user - User object
 * @return {Object} Dashboard data
 */
function getDashboardData(user) {
  const sheet = getSheet(CONFIG.SHEET_DATA_SURAT);
  const data = sheet.getDataRange().getValues();
  
  // Filter data berdasarkan role
  let filteredData = [];
  if (user.role === 'admin') {
    // Admin bisa lihat semua data
    filteredData = data.slice(1);
  } else {
    // User biasa hanya lihat data bidangnya
    filteredData = data.slice(1).filter(row => row[8] === user.bidang);
  }
  
  // Hitung statistik
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  const today = now.toDateString();
  
  const totalSurat = filteredData.length;
  const suratBulanIni = filteredData.filter(row => {
    const date = new Date(row[7]);
    return date.getMonth() === currentMonth && date.getFullYear() === currentYear;
  }).length;
  
  const suratHariIni = filteredData.filter(row => {
    const date = new Date(row[7]);
    return date.toDateString() === today;
  }).length;
  
  // Nomor agenda terakhir
  const counterSheet = getSheet(CONFIG.SHEET_COUNTER);
  const counterData = counterSheet.getRange(2, 1, 1, 3).getValues()[0];
  const nomorTerakhir = counterData[1] || 0;
  const lastUpdate = counterData[2] || new Date();
  
  // Surat terbaru (5 terakhir)
  const recentSurat = filteredData
    .sort((a, b) => new Date(b[1]) - new Date(a[1]))
    .slice(0, 5)
    .map(row => ({
      tanggal: Utilities.formatDate(new Date(row[7]), Session.getScriptTimeZone(), 'dd/MM/yy'),
      nomorSurat: row[3],
      perihal: row[9],
      jenis: row[10]
    }));
  
  return {
    totalSurat: totalSurat,
    suratBulanIni: suratBulanIni,
    suratHariIni: suratHariIni,
    nomorTerakhir: String(nomorTerakhir).padStart(3, '0'),
    lastUpdate: Utilities.formatDate(new Date(lastUpdate), Session.getScriptTimeZone(), 'dd MMM yyyy, HH:mm'),
    recentSurat: recentSurat
  };
}


// ================================================================
// SETUP & INITIALIZATION - Fungsi untuk setup awal
// ================================================================

/**
 * Setup awal aplikasi
 * Jalankan fungsi ini sekali saat pertama kali deploy
 */
function setupApplication() {
  // Inisialisasi semua sheets
  getSheet(CONFIG.SHEET_USERS);
  getSheet(CONFIG.SHEET_DATA_SURAT);
  getSheet(CONFIG.SHEET_COUNTER);
  getSheet(CONFIG.SHEET_SESSIONS);
  getSheet(CONFIG.SHEET_AUDIT_LOG);
  getSheet(CONFIG.SHEET_MASTER_KODE);
  
  // Buat user admin default jika belum ada
  createDefaultAdmin();
  
  Logger.log('Setup selesai!');
  Logger.log('Username: admin');
  Logger.log('Password: admin123');
  Logger.log('Silakan ganti password setelah login pertama kali');
}

/**
 * Buat user admin default
 */
function createDefaultAdmin() {
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  // Cek apakah sudah ada admin
  const hasAdmin = data.slice(1).some(row => row[5] === 'admin');
  
  if (!hasAdmin) {
    const userId = generateId('USER');
    const username = 'admin';
    const password = 'admin123'; // Password default
    const passwordHash = hashPassword(password);
    const namaLengkap = 'Administrator';
    const bidang = 'Sekretariat';
    const role = 'admin';
    const status = 'active';
    const createdDate = new Date();
    
    sheet.appendRow([
      userId,
      username,
      passwordHash,
      namaLengkap,
      bidang,
      role,
      status,
      createdDate,
      ''
    ]);
    
    Logger.log('Admin user created successfully');
  }
}

/**
 * Include file HTML/CSS/JS ke dalam template
 * @param {string} filename - Nama file
 * @return {string} Konten file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Process login via google.script.run (alternative method)
 * @param {string} username - Username
 * @param {string} password - Password
 * @return {Object} Login result
 */
function processLogin(username, password) {
  try {
    Logger.log('processLogin called with username: ' + username);
    
    // Validasi input
    if (!username || !password) {
      return {
        success: false,
        message: 'Username dan password harus diisi'
      };
    }
    
    // Cari user di database
    const sheet = getSheet(CONFIG.SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const [userId, dbUsername, passwordHash, namaLengkap, bidang, role, status] = row;
      
      if (dbUsername === username) {
        if (status !== 'active') {
          return {
            success: false,
            message: 'User tidak aktif. Hubungi administrator.'
          };
        }
        
        // Verifikasi password
        if (verifyPassword(password, passwordHash)) {
          // Login berhasil, buat session
          const sessionId = createSession(userId, username);
          
          // Update last login
          sheet.getRange(i + 1, 9).setValue(new Date());
          
          // Log aktivitas
          logActivity(userId, username, 'LOGIN', 'User login ke sistem');
          
          // Redirect ke dashboard
          const webAppUrl = ScriptApp.getService().getUrl();
          const redirectUrl = webAppUrl + '?session=' + sessionId;
          
          Logger.log('Login successful via processLogin');
          
          return {
            success: true,
            message: 'Login berhasil',
            redirectUrl: redirectUrl
          };
        }
      }
    }
    
    // Login gagal
    return {
      success: false,
      message: 'Username atau password salah'
    };
    
  } catch (error) {
    Logger.log('processLogin ERROR: ' + error.toString());
    return {
      success: false,
      message: 'Terjadi kesalahan: ' + error.toString()
    };
  }
}


/**
 * Submit surat - Handle form input surat
 * @param {Object} formData - Data form dari frontend
 * @return {Object} Result
 */
function submitSurat(formData) {
  try {
    Logger.log('=== submitSurat START ===');
    Logger.log('Mode: ' + formData.mode);
    Logger.log('Bidang: ' + formData.bidang);
    Logger.log('Session ID: ' + formData.sessionId);
    
    // Validasi session
    const session = getSessionById(formData.sessionId);
    if (!session) {
      Logger.log('ERROR: Session not found');
      return {
        success: false,
        message: 'Session tidak ditemukan. Silakan login kembali.'
      };
    }
    
    if (!isSessionValid(session)) {
      Logger.log('ERROR: Session expired');
      return {
        success: false,
        message: 'Session expired. Silakan login kembali.'
      };
    }
    
    Logger.log('Session valid. User ID: ' + session.userId);
    
    const user = getUserById(session.userId);
    if (!user) {
      Logger.log('ERROR: User not found');
      return {
        success: false,
        message: 'User tidak ditemukan.'
      };
    }
    
    Logger.log('User found: ' + user.username);
    
    // Generate nomor atau gunakan nomor dari Srikandi
    let nomorSuratLengkap = '';
    let nomorAgenda = '';
    let kodeDepan = '';
    let kodeBelakang = '';
    const currentYear = new Date().getFullYear();
    
    if (formData.mode === 'srikandi') {
      // Mode Srikandi: gunakan nomor yang sudah ada
      nomorSuratLengkap = formData.nomorLengkap;
      
      if (!nomorSuratLengkap) {
        Logger.log('ERROR: Nomor lengkap kosong');
        return {
          success: false,
          message: 'Nomor surat lengkap harus diisi!'
        };
      }
      
      // Extract parts dari nomor lengkap
      const parts = nomorSuratLengkap.split('/');
      if (parts.length >= 3) {
        kodeDepan = parts[0] || '';
        nomorAgenda = parts[1] || '';
        kodeBelakang = parts[2] || CONFIG.DEFAULT_KODE_BELAKANG;
      } else {
        Logger.log('ERROR: Format nomor tidak valid');
        return {
          success: false,
          message: 'Format nomor surat tidak valid. Contoh: 100.1.2/046/405.29.05/2025'
        };
      }
      
      Logger.log('Mode Srikandi - Nomor: ' + nomorSuratLengkap);
      
    } else {
      // Mode Manual: auto-generate nomor agenda
      kodeDepan = formData.kodeDepan;
      kodeBelakang = formData.kodeBelakang || CONFIG.DEFAULT_KODE_BELAKANG;
      
      if (!kodeDepan) {
        Logger.log('ERROR: Kode depan kosong');
        return {
          success: false,
          message: 'Kode klasifikasi depan harus diisi!'
        };
      }
      
      Logger.log('Getting counter...');
      
      // Get nomor terakhir dari counter
      const counterSheet = getSheet(CONFIG.SHEET_COUNTER);
      const counterData = counterSheet.getRange(2, 1, 1, 3).getValues()[0];
      const tahunCounter = counterData[0];
      let nomorTerakhir = counterData[1] || 0;
      
      Logger.log('Counter - Tahun: ' + tahunCounter + ', Nomor: ' + nomorTerakhir);
      
      // Reset counter jika tahun berubah
      if (tahunCounter != currentYear) {
        Logger.log('Reset counter for new year: ' + currentYear);
        nomorTerakhir = 0;
        counterSheet.getRange(2, 1).setValue(currentYear);
      }
      
      // Increment nomor
      nomorTerakhir = parseInt(nomorTerakhir) + 1;
      nomorAgenda = String(nomorTerakhir).padStart(3, '0');
      
      Logger.log('New nomor agenda: ' + nomorAgenda);
      
      // Update counter
      counterSheet.getRange(2, 2).setValue(nomorTerakhir);
      counterSheet.getRange(2, 3).setValue(new Date());
      
      // Generate nomor lengkap
      nomorSuratLengkap = `${kodeDepan}/${nomorAgenda}/${kodeBelakang}/${currentYear}`;
      
      Logger.log('Mode Manual - Generated: ' + nomorSuratLengkap);
    }
    
    Logger.log('Checking for duplicates...');
    
    // Validasi nomor tidak duplikat
    const dataSheet = getSheet(CONFIG.SHEET_DATA_SURAT);
    const existingData = dataSheet.getDataRange().getValues();
    
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][3] === nomorSuratLengkap) {
        Logger.log('ERROR: Duplicate nomor found');
        return {
          success: false,
          message: 'Nomor surat ' + nomorSuratLengkap + ' sudah ada dalam database!'
        };
      }
    }
    
    Logger.log('No duplicates. Saving data...');
    
    // Simpan data
    const id = generateId('SURAT');
    const timestamp = new Date();
    const tanggalSurat = formData.tanggalSurat ? new Date(formData.tanggalSurat) : new Date();
    
    dataSheet.appendRow([
      id,
      timestamp,
      nomorAgenda,
      nomorSuratLengkap,
      formData.mode === 'srikandi' ? 'Srikandi' : 'Manual',
      kodeDepan,
      kodeBelakang,
      tanggalSurat,
      formData.bidang || user.bidang,
      formData.perihal || '',
      formData.jenisSurat || '',
      formData.tujuan || '',
      formData.keterangan || '',
      user.username,
      '',
      ''
    ]);
    
    Logger.log('Data saved successfully');
    
    // Log aktivitas
    logActivity(user.userId, user.username, 'INPUT_SURAT', 'Input surat: ' + nomorSuratLengkap);
    
    Logger.log('=== submitSurat SUCCESS ===');
    
    return {
      success: true,
      message: 'Surat berhasil disimpan! Nomor: ' + nomorSuratLengkap,
      nomorAgenda: parseInt(nomorAgenda) || 0,
      nomorSurat: nomorSuratLengkap
    };
    
  } catch (error) {
    Logger.log('=== submitSurat ERROR ===');
    Logger.log('Error message: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    
    return {
      success: false,
      message: 'Terjadi kesalahan: ' + error.message
    };
  }
}


/**
 * Get surat data for table - VERSI DEBUG
 * @param {string} sessionId - Session ID
 * @param {string} userRole - User role (admin/user)
 * @param {string} userBidang - User bidang
 * @return {Object} Result with data array
 */
function getSuratData(sessionId, userRole, userBidang) {
  try {
    Logger.log('=== getSuratData START ===');
    Logger.log('Parameters:');
    Logger.log('  sessionId: ' + (sessionId || 'EMPTY'));
    Logger.log('  userRole: ' + (userRole || 'EMPTY'));
    Logger.log('  userBidang: ' + (userBidang || 'EMPTY'));
    
    // Validate parameters
    if (!sessionId) {
      Logger.log('ERROR: sessionId is empty!');
      return {
        success: false,
        message: 'Session ID kosong'
      };
    }
    
    // Get session
    Logger.log('Getting session...');
    const session = getSessionById(sessionId);
    
    if (!session) {
      Logger.log('ERROR: Session not found for ID: ' + sessionId);
      return {
        success: false,
        message: 'Session tidak ditemukan'
      };
    }
    
    Logger.log('Session found: ' + session.username);
    
    // Check if valid
    if (!isSessionValid(session)) {
      Logger.log('ERROR: Session expired');
      return {
        success: false,
        message: 'Session expired'
      };
    }
    
    Logger.log('Session valid ✓');
    
    // Get sheet
    Logger.log('Getting sheet DATA_SURAT...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DATA_SURAT');
    
    if (!sheet) {
      Logger.log('ERROR: Sheet DATA_SURAT not found!');
      Logger.log('Available sheets:');
      ss.getSheets().forEach(function(s) {
        Logger.log('  - ' + s.getName());
      });
      
      return {
        success: false,
        message: 'Sheet DATA_SURAT tidak ditemukan'
      };
    }
    
    Logger.log('Sheet found ✓');
    
    // Get data
    const data = sheet.getDataRange().getValues();
    Logger.log('Total rows: ' + data.length);
    
    if (data.length <= 1) {
      Logger.log('No data rows (only header)');
      return {
        success: true,
        data: []
      };
    }
    
    // Parse data
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Filter by role
      if (userRole !== 'admin' && row[8] !== userBidang) {
        continue;
      }
      
      result.push({
        id: row[0] || '',
        timestamp: row[1] || '',
        nomorAgenda: row[2] || '',
        nomor: row[3] || '',
        sumber: row[4] || '',
        kodeDepan: row[5] || '',
        kodeBelakang: row[6] || '',
        tanggal: row[7] || '',
        bidang: row[8] || '',
        perihal: row[9] || '',
        jenis: row[10] || '',
        tujuan: row[11] || '',
        keterangan: row[12] || '',
        createdBy: row[13] || ''
      });
    }
    
    Logger.log('Returning ' + result.length + ' records');
    Logger.log('=== getSuratData SUCCESS ===');
    
    return {
      success: true,
      data: result
    };
    
  } catch (error) {
    Logger.log('=== getSuratData ERROR ===');
    Logger.log('Error name: ' + error.name);
    Logger.log('Error message: ' + error.message);
    Logger.log('Error stack: ' + error.stack);
    
    return {
      success: false,
      message: 'Error: ' + error.message
    };
  }
}

/**
 * Delete surat
 * @param {string} suratId - Surat ID to delete
 * @param {string} sessionId - Session ID
 * @return {Object} Result
 */
function deleteSurat(suratId, sessionId) {
  try {
    Logger.log('=== deleteSurat START ===');
    Logger.log('Surat ID: ' + suratId);
    
    // Validasi session
    const session = getSessionById(sessionId);
    if (!session || !isSessionValid(session)) {
      return {
        success: false,
        message: 'Session tidak valid'
      };
    }
    
    const user = getUserById(session.userId);
    
    // Find and delete surat
    const sheet = getSheet(CONFIG.SHEET_DATA_SURAT);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === suratId) {
        const nomorSurat = data[i][3];
        const bidang = data[i][8];
        
        // Check permission: admin bisa hapus semua, user hanya bisa hapus milik bidangnya
        if (user.role !== 'admin' && bidang !== user.bidang) {
          return {
            success: false,
            message: 'Anda tidak memiliki permission untuk menghapus surat ini'
          };
        }
        
        // Delete row
        sheet.deleteRow(i + 1);
        
        // Log aktivitas
        logActivity(user.userId, user.username, 'DELETE_SURAT', 'Hapus surat: ' + nomorSurat);
        
        Logger.log('Surat deleted: ' + nomorSurat);
        Logger.log('=== deleteSurat SUCCESS ===');
        
        return {
          success: true,
          message: 'Surat berhasil dihapus: ' + nomorSurat
        };
      }
    }
    
    return {
      success: false,
      message: 'Surat tidak ditemukan'
    };
    
  } catch (error) {
    Logger.log('=== deleteSurat ERROR ===');
    Logger.log('Error: ' + error.message);
    
    return {
      success: false,
      message: 'Terjadi kesalahan: ' + error.message
    };
  }
}


// ================================================================
// QUICK SETUP - VERSI SIMPLE UNTUK TESTING
// ================================================================

/**
 * Setup cepat - Jalankan ini jika setupApplication error
 * Function ini lebih simple dan mudah di-debug
 */
function quickSetup() {
  try {
    Logger.log('🚀 Memulai Quick Setup...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('✓ Spreadsheet berhasil diakses');
    
    // Setup USERS sheet
    let usersSheet = ss.getSheetByName('USERS');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('USERS');
      usersSheet.appendRow(['User ID', 'Username', 'Password Hash', 'Nama Lengkap', 'Bidang', 'Role', 'Status', 'Created Date', 'Last Login']);
      
      // Buat user admin
      const userId = 'USER' + new Date().getTime();
      const passwordHash = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'admin123'));
      usersSheet.appendRow([userId, 'admin', passwordHash, 'Administrator', 'Sekretariat', 'admin', 'active', new Date(), '']);
      
      Logger.log('✓ Sheet USERS dibuat dengan user admin');
    }
    
    // Setup DATA_SURAT sheet
    let dataSheet = ss.getSheetByName('DATA_SURAT');
    if (!dataSheet) {
      dataSheet = ss.insertSheet('DATA_SURAT');
      dataSheet.appendRow(['ID', 'Timestamp', 'No Agenda', 'Nomor Surat Lengkap', 'Sumber', 'Kode Depan', 'Kode Belakang', 'Tanggal Surat', 'Bidang', 'Perihal', 'Jenis Surat', 'Tujuan', 'Keterangan', 'Created By', 'Modified By', 'Modified Time']);
      Logger.log('✓ Sheet DATA_SURAT dibuat');
    }
    
    // Setup COUNTER sheet
    let counterSheet = ss.getSheetByName('COUNTER');
    if (!counterSheet) {
      counterSheet = ss.insertSheet('COUNTER');
      counterSheet.appendRow(['Tahun', 'Nomor Terakhir', 'Last Updated']);
      counterSheet.appendRow([new Date().getFullYear(), 0, new Date()]);
      Logger.log('✓ Sheet COUNTER dibuat');
    }
    
    // Setup SESSIONS sheet
    let sessionsSheet = ss.getSheetByName('SESSIONS');
    if (!sessionsSheet) {
      sessionsSheet = ss.insertSheet('SESSIONS');
      sessionsSheet.appendRow(['Session ID', 'User ID', 'Username', 'Login Time', 'Expire Time', 'Last Activity']);
      Logger.log('✓ Sheet SESSIONS dibuat');
    }
    
    // Setup AUDIT_LOG sheet
    let auditSheet = ss.getSheetByName('AUDIT_LOG');
    if (!auditSheet) {
      auditSheet = ss.insertSheet('AUDIT_LOG');
      auditSheet.appendRow(['Timestamp', 'User ID', 'Username', 'Action', 'Detail', 'IP Address']);
      Logger.log('✓ Sheet AUDIT_LOG dibuat');
    }
    
    // Setup MASTER_KODE sheet
    let masterSheet = ss.getSheetByName('MASTER_KODE');
    if (!masterSheet) {
      masterSheet = ss.insertSheet('MASTER_KODE');
      masterSheet.appendRow(['Kode Depan', 'Kode Belakang', 'Deskripsi', 'Bidang']);
      masterSheet.appendRow(['100.1.2', '405.29.05', 'Kepegawaian', 'Kepegawaian']);
      masterSheet.appendRow(['200.5.1', '405.29.05', 'Keuangan', 'Keuangan']);
      masterSheet.appendRow(['300.2.3', '405.29.05', 'Umum', 'Umum']);
      Logger.log('✓ Sheet MASTER_KODE dibuat');
    }
    
    Logger.log('');
    Logger.log('========================================');
    Logger.log('✅ QUICK SETUP SELESAI!');
    Logger.log('========================================');
    Logger.log('');
    Logger.log('📋 KREDENSIAL LOGIN:');
    Logger.log('   Username: admin');
    Logger.log('   Password: admin123');
    Logger.log('');
    Logger.log('📌 NEXT STEPS:');
    Logger.log('   1. Refresh spreadsheet (F5)');
    Logger.log('   2. Cek apakah 6 sheets sudah muncul');
    Logger.log('   3. Deploy sebagai Web App');
    Logger.log('   4. Test login');
    Logger.log('');
    
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    Logger.log('Detail: ' + error.stack);
  }
}


// ================================================================
// HELPER FUNCTION FOR DASHBOARD - Server-side data processing
// ================================================================

/**
 * Get surat data untuk ditampilkan di dashboard (server-side)
 * Dipanggil langsung dari template, bukan via google.script.run
 * @param {string} sessionId - Session ID
 * @param {string} userRole - User role
 * @param {string} userBidang - User bidang
 * @return {Array} Array of surat objects
 */
function getSuratDataForDashboard(sessionId, userRole, userBidang) {
  try {
    Logger.log('=== getSuratDataForDashboard START ===');
    
    // Get data surat
    const sheet = getSheet(CONFIG.SHEET_DATA_SURAT);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return [];
    }
    
    // Parse data
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Filter berdasarkan role
      if (userRole !== 'admin' && row[8] !== userBidang) {
        continue; // Skip data bidang lain jika bukan admin
      }
      
      result.push({
        id: row[0],
        timestamp: row[1],
        nomorAgenda: row[2],
        nomor: row[3],
        sumber: row[4],
        kodeDepan: row[5],
        kodeBelakang: row[6],
        tanggal: row[7],
        bidang: row[8],
        perihal: row[9],
        jenis: row[10],
        tujuan: row[11],
        keterangan: row[12],
        createdBy: row[13]
      });
    }
    
    Logger.log('Found ' + result.length + ' records for dashboard');
    return result;
    
  } catch (error) {
    Logger.log('=== getSuratDataForDashboard ERROR ===');
    Logger.log('Error: ' + error.message);
    return [];
  }
}