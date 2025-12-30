/**
 * ================================================================
 * USER MANAGEMENT HELPER FUNCTIONS
 * ================================================================
 * 
 * File ini berisi fungsi-fungsi helper untuk mengelola user:
 * - Menambah user baru
 * - Mengubah password
 * - Menonaktifkan user
 * - Menghapus user
 * 
 * Jalankan fungsi-fungsi ini dari Apps Script Editor
 * untuk mengelola user secara manual.
 * 
 * ================================================================
 */


/**
 * ================================================================
 * TAMBAH USER BARU
 * ================================================================
 * 
 * Cara pakai:
 * 1. Edit data di bawah sesuai kebutuhan
 * 2. Pilih function createNewUser dari dropdown
 * 3. Klik Run
 * 4. Cek sheet USERS untuk verifikasi
 */
function createNewUser() {
  // ===== EDIT DATA USER DI SINI =====
  const newUser = {
    username: 'budi.kh',           // Username untuk login (huruf kecil, tanpa spasi)
    password: 'password123',       // Password (minimal 6 karakter)
    namaLengkap: 'Budi Santoso',  // Nama lengkap
    bidang: 'Kepegawaian',        // Bidang: Kepegawaian, Keuangan, Umum, dll
    role: 'user'                   // Role: 'admin' atau 'user'
  };
  // ====================================
  
  // Validasi input
  if (!newUser.username || !newUser.password || !newUser.namaLengkap || !newUser.bidang) {
    Logger.log('❌ ERROR: Semua field harus diisi!');
    return;
  }
  
  if (newUser.password.length < 6) {
    Logger.log('❌ ERROR: Password minimal 6 karakter!');
    return;
  }
  
  // Cek apakah username sudah ada
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === newUser.username) {
      Logger.log('❌ ERROR: Username "' + newUser.username + '" sudah digunakan!');
      return;
    }
  }
  
  // Buat user baru
  const userId = generateId('USER');
  const passwordHash = hashPassword(newUser.password);
  const createdDate = new Date();
  
  sheet.appendRow([
    userId,
    newUser.username,
    passwordHash,
    newUser.namaLengkap,
    newUser.bidang,
    newUser.role,
    'active',
    createdDate,
    ''
  ]);
  
  // Log hasil
  Logger.log('✅ User berhasil dibuat!');
  Logger.log('');
  Logger.log('📋 Detail User:');
  Logger.log('   User ID      : ' + userId);
  Logger.log('   Username     : ' + newUser.username);
  Logger.log('   Password     : ' + newUser.password + ' (SIMPAN INI!)');
  Logger.log('   Nama Lengkap : ' + newUser.namaLengkap);
  Logger.log('   Bidang       : ' + newUser.bidang);
  Logger.log('   Role         : ' + newUser.role);
  Logger.log('   Status       : active');
  Logger.log('');
  Logger.log('💡 User dapat login dengan:');
  Logger.log('   Username: ' + newUser.username);
  Logger.log('   Password: ' + newUser.password);
  
  // Log aktivitas
  logActivity(userId, newUser.username, 'USER_CREATED', 'User baru dibuat oleh system');
}


/**
 * ================================================================
 * BATCH CREATE USERS - Buat banyak user sekaligus
 * ================================================================
 * 
 * Cara pakai:
 * 1. Edit array users di bawah
 * 2. Tambahkan data user sebanyak yang diperlukan
 * 3. Pilih function batchCreateUsers dari dropdown
 * 4. Klik Run
 */
function batchCreateUsers() {
  // ===== EDIT DATA USERS DI SINI =====
  const users = [
    {
      username: 'tapem',
      password: 'tapem123',
      namaLengkap: 'Refqi Elfida Mufchiroh',
      bidang: 'Tata Pemerintahan',
      role: 'user'
    },
    {
      username: 'Pemas',
      password: 'pemas123',
      namaLengkap: 'Budi Hermawan',
      bidang: 'Pemberdayaan Masyarakat',
      role: 'user'
    },
    {
      username: 'bendahara',
      password: 'bendahara123',
      namaLengkap: 'Budi',
      bidang: 'Kasubag Sungram',
      role: 'user'
    }
    // Tambahkan user lain di sini...
  ];
  // ====================================
  
  Logger.log('🚀 Memulai batch create users...');
  Logger.log('');
  
  let successCount = 0;
  let errorCount = 0;
  
  for (let i = 0; i < users.length; i++) {
    const user = users[i];
    
    try {
      // Validasi
      if (!user.username || !user.password || !user.namaLengkap || !user.bidang) {
        Logger.log('❌ User #' + (i+1) + ' gagal: Data tidak lengkap');
        errorCount++;
        continue;
      }
      
      // Cek duplikasi username
      const sheet = getSheet(CONFIG.SHEET_USERS);
      const data = sheet.getDataRange().getValues();
      let isDuplicate = false;
      
      for (let j = 1; j < data.length; j++) {
        if (data[j][1] === user.username) {
          isDuplicate = true;
          break;
        }
      }
      
      if (isDuplicate) {
        Logger.log('❌ User #' + (i+1) + ' gagal: Username "' + user.username + '" sudah ada');
        errorCount++;
        continue;
      }
      
      // Buat user
      const userId = generateId('USER');
      const passwordHash = hashPassword(user.password);
      const createdDate = new Date();
      
      sheet.appendRow([
        userId,
        user.username,
        passwordHash,
        user.namaLengkap,
        user.bidang,
        user.role,
        'active',
        createdDate,
        ''
      ]);
      
      Logger.log('✅ User #' + (i+1) + ' berhasil: ' + user.username + ' (' + user.namaLengkap + ')');
      successCount++;
      
      // Log aktivitas
      logActivity(userId, user.username, 'USER_CREATED', 'User baru dibuat via batch');
      
    } catch (error) {
      Logger.log('❌ User #' + (i+1) + ' error: ' + error.toString());
      errorCount++;
    }
  }
  
  Logger.log('');
  Logger.log('📊 HASIL BATCH CREATE:');
  Logger.log('   Total user  : ' + users.length);
  Logger.log('   Berhasil    : ' + successCount);
  Logger.log('   Gagal       : ' + errorCount);
}


/**
 * ================================================================
 * GANTI PASSWORD USER
 * ================================================================
 * 
 * Cara pakai:
 * 1. Edit username dan password baru di bawah
 * 2. Pilih function changeUserPassword dari dropdown
 * 3. Klik Run
 */
function changeUserPassword() {
  // ===== EDIT DATA DI SINI =====
  const username = 'admin';           // Username yang mau diganti password
  const newPassword = 'newpassword';  // Password baru
  // =============================
  
  if (!username || !newPassword) {
    Logger.log('❌ ERROR: Username dan password baru harus diisi!');
    return;
  }
  
  if (newPassword.length < 6) {
    Logger.log('❌ ERROR: Password minimal 6 karakter!');
    return;
  }
  
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const userId = data[i][0];
      const passwordHash = hashPassword(newPassword);
      
      // Update password
      sheet.getRange(i + 1, 3).setValue(passwordHash);
      
      Logger.log('✅ Password berhasil diganti!');
      Logger.log('');
      Logger.log('   Username     : ' + username);
      Logger.log('   Password Baru: ' + newPassword);
      Logger.log('');
      Logger.log('💡 User dapat login dengan password baru');
      
      // Log aktivitas
      logActivity(userId, username, 'PASSWORD_CHANGED', 'Password diubah oleh admin');
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log('❌ ERROR: Username "' + username + '" tidak ditemukan!');
  }
}


/**
 * ================================================================
 * NONAKTIFKAN USER
 * ================================================================
 * 
 * Cara pakai:
 * 1. Edit username di bawah
 * 2. Pilih function deactivateUser dari dropdown
 * 3. Klik Run
 */
function deactivateUser() {
  // ===== EDIT USERNAME DI SINI =====
  const username = 'budi.kh';  // Username yang mau dinonaktifkan
  // =================================
  
  if (!username) {
    Logger.log('❌ ERROR: Username harus diisi!');
    return;
  }
  
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const userId = data[i][0];
      const currentStatus = data[i][6];
      
      if (currentStatus === 'inactive') {
        Logger.log('⚠️ User "' + username + '" sudah dalam status inactive');
        return;
      }
      
      // Update status
      sheet.getRange(i + 1, 7).setValue('inactive');
      
      Logger.log('✅ User berhasil dinonaktifkan!');
      Logger.log('');
      Logger.log('   Username : ' + username);
      Logger.log('   Status   : inactive');
      Logger.log('');
      Logger.log('💡 User tidak dapat login sampai diaktifkan kembali');
      
      // Log aktivitas
      logActivity(userId, username, 'USER_DEACTIVATED', 'User dinonaktifkan oleh admin');
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log('❌ ERROR: Username "' + username + '" tidak ditemukan!');
  }
}


/**
 * ================================================================
 * AKTIFKAN USER
 * ================================================================
 * 
 * Cara pakai:
 * 1. Edit username di bawah
 * 2. Pilih function activateUser dari dropdown
 * 3. Klik Run
 */
function activateUser() {
  // ===== EDIT USERNAME DI SINI =====
  const username = 'budi.kh';  // Username yang mau diaktifkan
  // =================================
  
  if (!username) {
    Logger.log('❌ ERROR: Username harus diisi!');
    return;
  }
  
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const userId = data[i][0];
      const currentStatus = data[i][6];
      
      if (currentStatus === 'active') {
        Logger.log('⚠️ User "' + username + '" sudah dalam status active');
        return;
      }
      
      // Update status
      sheet.getRange(i + 1, 7).setValue('active');
      
      Logger.log('✅ User berhasil diaktifkan!');
      Logger.log('');
      Logger.log('   Username : ' + username);
      Logger.log('   Status   : active');
      Logger.log('');
      Logger.log('💡 User sekarang dapat login kembali');
      
      // Log aktivitas
      logActivity(userId, username, 'USER_ACTIVATED', 'User diaktifkan oleh admin');
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log('❌ ERROR: Username "' + username + '" tidak ditemukan!');
  }
}


/**
 * ================================================================
 * LIHAT SEMUA USER
 * ================================================================
 * 
 * Menampilkan daftar semua user di log
 */
function viewAllUsers() {
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  Logger.log('👥 DAFTAR SEMUA USER');
  Logger.log('='.repeat(80));
  Logger.log('');
  
  if (data.length <= 1) {
    Logger.log('Belum ada user');
    return;
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    Logger.log('User #' + i);
    Logger.log('   User ID      : ' + row[0]);
    Logger.log('   Username     : ' + row[1]);
    Logger.log('   Nama Lengkap : ' + row[3]);
    Logger.log('   Bidang       : ' + row[4]);
    Logger.log('   Role         : ' + row[5]);
    Logger.log('   Status       : ' + row[6]);
    Logger.log('   Created      : ' + row[7]);
    Logger.log('   Last Login   : ' + (row[8] || 'Belum pernah login'));
    Logger.log('');
  }
  
  Logger.log('='.repeat(80));
  Logger.log('Total: ' + (data.length - 1) + ' user');
}


/**
 * ================================================================
 * HAPUS USER (PERMANENT)
 * ================================================================
 * 
 * PERINGATAN: User akan dihapus permanent dari database!
 * 
 * Cara pakai:
 * 1. Edit username di bawah
 * 2. Pilih function deleteUser dari dropdown
 * 3. Klik Run
 */
function deleteUser() {
  // ===== EDIT USERNAME DI SINI =====
  const username = 'test.user';  // Username yang mau dihapus PERMANENT
  // =================================
  
  if (!username) {
    Logger.log('❌ ERROR: Username harus diisi!');
    return;
  }
  
  // Cegah hapus admin
  if (username === 'admin') {
    Logger.log('❌ ERROR: User admin tidak bisa dihapus!');
    return;
  }
  
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const userId = data[i][0];
      const namaLengkap = data[i][3];
      
      // Hapus row
      sheet.deleteRow(i + 1);
      
      Logger.log('✅ User berhasil dihapus!');
      Logger.log('');
      Logger.log('   Username     : ' + username);
      Logger.log('   Nama Lengkap : ' + namaLengkap);
      Logger.log('');
      Logger.log('⚠️ User telah dihapus permanent dari database');
      
      // Log aktivitas
      logActivity(userId, username, 'USER_DELETED', 'User dihapus oleh admin');
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log('❌ ERROR: Username "' + username + '" tidak ditemukan!');
  }
}


/**
 * ================================================================
 * RESET PASSWORD KE DEFAULT
 * ================================================================
 * 
 * Reset password user ke "password123"
 * 
 * Cara pakai:
 * 1. Edit username di bawah
 * 2. Pilih function resetPasswordToDefault dari dropdown
 * 3. Klik Run
 */
function resetPasswordToDefault() {
  // ===== EDIT USERNAME DI SINI =====
  const username = 'budi.kh';  // Username yang mau direset password
  // =================================
  
  const defaultPassword = 'password123';
  
  if (!username) {
    Logger.log('❌ ERROR: Username harus diisi!');
    return;
  }
  
  const sheet = getSheet(CONFIG.SHEET_USERS);
  const data = sheet.getDataRange().getValues();
  
  let found = false;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username) {
      const userId = data[i][0];
      const passwordHash = hashPassword(defaultPassword);
      
      // Update password
      sheet.getRange(i + 1, 3).setValue(passwordHash);
      
      Logger.log('✅ Password berhasil direset!');
      Logger.log('');
      Logger.log('   Username        : ' + username);
      Logger.log('   Password Default: ' + defaultPassword);
      Logger.log('');
      Logger.log('💡 Informasikan password ini kepada user');
      Logger.log('💡 User sebaiknya mengganti password setelah login');
      
      // Log aktivitas
      logActivity(userId, username, 'PASSWORD_RESET', 'Password direset ke default oleh admin');
      
      found = true;
      break;
    }
  }
  
  if (!found) {
    Logger.log('❌ ERROR: Username "' + username + '" tidak ditemukan!');
  }
}