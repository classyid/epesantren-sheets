// Konfigurasi
const API_KEY = "<apikey>"; // API Key dari Postman collection
const BASE_URL = "<link-Api-Epesantren>";

// ==========================================
// Setup Menu di Google Sheets
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistem Santri')
    .addItem('Sinkronisasi Data Santri', 'syncAllStudentData')
    .addItem('Cek Saldo Santri', 'searchStudentUI')
    .addItem('Proses Transaksi', 'processTransactionUI')
    .addItem('Generate Laporan', 'generateReportUI')
    .addToUi();
}

// ==========================================
// Fungsi Integrasi API
// ==========================================

/**
 * Mendapatkan semua data santri dari API
 */
function getAllStudentData() {
  const url = `${BASE_URL}/get_data`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ key: API_KEY })
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.status === 1) {
      return responseData.data;
    } else {
      Logger.log("Error getting data: " + responseData.message);
      return [];
    }
  } catch(e) {
    Logger.log("Error fetching data: " + e.toString());
    return [];
  }
}

/**
 * Mendapatkan data satu santri berdasarkan NIS
 */
function getStudentByNIS(nis) {
  const url = `${BASE_URL}/get_data_first`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ 
      key: API_KEY,
      nis: nis
    })
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.status === 1) {
      return responseData.data;
    } else {
      Logger.log("Error getting student data: " + responseData.message);
      return null;
    }
  } catch(e) {
    Logger.log("Error fetching student data: " + e.toString());
    return null;
  }
}

/**
 * Menggunakan saldo untuk transaksi
 */
function useSaving(nis, nominal) {
  const url = `${BASE_URL}/use_saving`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ 
      key: API_KEY,
      nis: nis,
      nominal: nominal
    })
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseData = JSON.parse(response.getContentText());
    
    return responseData;
  } catch(e) {
    Logger.log("Error processing transaction: " + e.toString());
    return { 
      status: 0, 
      message: "Terjadi kesalahan saat memproses transaksi: " + e.toString() 
    };
  }
}

// ==========================================
// Fungsi UI dan Interaksi Sheet
// ==========================================

/**
 * Sinkronisasi data santri dari API ke Sheet
 */
function syncAllStudentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek apakah sheet data santri sudah ada, jika belum, buat baru
  let sheet = ss.getSheetByName('Data Santri');
  if (!sheet) {
    sheet = ss.insertSheet('Data Santri');
    
    // Set header
    sheet.getRange('A1:F1').setValues([['NIS', 'Nama', 'Saldo', 'RFID', 'PIN', 'Status']]);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  } else {
    // Clear existing data but keep the header
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 6).clear();
    }
  }
  
  // Ambil data dari API
  const data = getAllStudentData();
  
  if (data.length === 0) {
    SpreadsheetApp.getUi().alert('Tidak ada data yang ditemukan atau terjadi kesalahan.');
    return;
  }
  
  // Format data untuk dimasukkan ke sheet
  const formattedData = data.map(student => [
    student.nis,
    student.nama,
    student.saldo,
    student.rfid,
    student.pin,
    student.status || 'N'
  ]);
  
  // Masukkan data ke sheet
  sheet.getRange(2, 1, formattedData.length, 6).setValues(formattedData);
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 6);
  
  // Sorting berdasarkan nama
  sheet.getRange(2, 1, formattedData.length, 6).sort({column: 2, ascending: true});
  
  SpreadsheetApp.getUi().alert(`Berhasil menyinkronkan ${data.length} data santri.`);
}

/**
 * UI untuk pencarian santri
 */
function searchStudentUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Cek Saldo Santri', 'Masukkan NIS santri:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    const nis = response.getResponseText().trim();
    if (nis) {
      const student = getStudentByNIS(nis);
      if (student) {
        ui.alert(
          'Informasi Santri',
          `NIS: ${student.nis}\nNama: ${student.nama}\nSaldo: Rp ${formatCurrency(student.saldo)}\nStatus: ${student.status || 'N'}`,
          ui.ButtonSet.OK
        );
      } else {
        ui.alert('Santri dengan NIS tersebut tidak ditemukan.');
      }
    } else {
      ui.alert('NIS tidak boleh kosong.');
    }
  }
}

/**
 * UI untuk memproses transaksi pengeluaran saldo
 */
function processTransactionUI() {
  const ui = SpreadsheetApp.getUi();
  
  // Input NIS
  const nisResponse = ui.prompt('Transaksi Saldo', 'Masukkan NIS santri:', ui.ButtonSet.OK_CANCEL);
  if (nisResponse.getSelectedButton() != ui.Button.OK) return;
  
  const nis = nisResponse.getResponseText().trim();
  if (!nis) {
    ui.alert('NIS tidak boleh kosong.');
    return;
  }
  
  // Cek data santri
  const student = getStudentByNIS(nis);
  if (!student) {
    ui.alert('Santri dengan NIS tersebut tidak ditemukan.');
    return;
  }
  
  // Tampilkan info santri
  const infoResponse = ui.alert(
    'Informasi Santri',
    `NIS: ${student.nis}\nNama: ${student.nama}\nSaldo: Rp ${formatCurrency(student.saldo)}\n\nLanjutkan transaksi?`,
    ui.ButtonSet.YES_NO
  );
  
  if (infoResponse != ui.Button.YES) return;
  
  // Input nominal transaksi
  const nominalResponse = ui.prompt(
    'Transaksi Saldo',
    `Masukkan nominal transaksi untuk ${student.nama}:`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (nominalResponse.getSelectedButton() != ui.Button.OK) return;
  
  const nominalStr = nominalResponse.getResponseText().trim().replace(/\D/g, '');
  if (!nominalStr) {
    ui.alert('Nominal tidak valid.');
    return;
  }
  
  const nominal = parseInt(nominalStr);
  if (isNaN(nominal) || nominal <= 0) {
    ui.alert('Nominal harus berupa angka positif.');
    return;
  }
  
  // Konfirmasi final
  const confirmResponse = ui.alert(
    'Konfirmasi Transaksi',
    `Anda akan melakukan transaksi:\nSantri: ${student.nama} (${student.nis})\nNominal: Rp ${formatCurrency(nominal)}\n\nLanjutkan?`,
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse != ui.Button.YES) return;
  
  // Proses transaksi
  const result = useSaving(nis, nominal);
  
  if (result.status === 1) {
    // Transaksi berhasil
    ui.alert('Transaksi Berhasil', result.message, ui.ButtonSet.OK);
    
    // Catat transaksi
    logTransaction(student.nis, student.nama, nominal, "SUCCESS");
    
    // Refresh data santri
    syncAllStudentData();
  } else {
    // Transaksi gagal
    ui.alert('Transaksi Gagal', result.message || 'Terjadi kesalahan saat memproses transaksi.', ui.ButtonSet.OK);
    
    // Catat transaksi gagal
    logTransaction(student.nis, student.nama, nominal, "FAILED: " + (result.message || "Unknown error"));
  }
}

/**
 * Mencatat transaksi ke sheet transaksi
 */
function logTransaction(nis, name, amount, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek apakah sheet transaksi sudah ada, jika belum, buat baru
  let sheet = ss.getSheetByName('Transaksi');
  if (!sheet) {
    sheet = ss.insertSheet('Transaksi');
    
    // Set header
    sheet.getRange('A1:E1').setValues([['Tanggal', 'NIS', 'Nama', 'Nominal', 'Status']]);
    sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  }
  
  // Ambil timestamp saat ini
  const timestamp = new Date();
  
  // Tambahkan data transaksi baru
  const newRow = [timestamp, nis, name, amount, status];
  sheet.appendRow(newRow);
  
  // Format tanggal dan nominal
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  sheet.getRange(lastRow, 4).setNumberFormat('#,##0');
  
  // Auto-resize columns
  sheet.autoResizeColumns(1, 5);
}

/**
 * UI untuk menghasilkan laporan
 */
function generateReportUI() {
  const ui = SpreadsheetApp.getUi();
  const reportTypes = ['Saldo Tertinggi', 'Saldo Terendah', 'Transaksi Harian', 'Transaksi Bulanan'];
  
  const response = ui.prompt(
    'Generate Laporan',
    'Pilih jenis laporan (1-4):\n1. Saldo Tertinggi\n2. Saldo Terendah\n3. Transaksi Harian\n4. Transaksi Bulanan',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) return;
  
  const choice = parseInt(response.getResponseText().trim());
  if (isNaN(choice) || choice < 1 || choice > 4) {
    ui.alert('Pilihan tidak valid. Silakan pilih angka 1-4.');
    return;
  }
  
  // Buat laporan sesuai pilihan
  switch (choice) {
    case 1: createTopBalanceReport(); break;
    case 2: createLowestBalanceReport(); break;
    case 3: createDailyTransactionReport(); break;
    case 4: createMonthlyTransactionReport(); break;
  }
}

/**
 * Membuat laporan saldo tertinggi
 */
function createTopBalanceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek sheet data santri
  const dataSheet = ss.getSheetByName('Data Santri');
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Data santri tidak ditemukan. Silakan sinkronisasi data terlebih dahulu.');
    return;
  }
  
  // Ambil data santri
  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Filter hanya header dan data (buang header)
  const header = data[0];
  const students = data.slice(1);
  
  // Sort berdasarkan saldo (descending)
  students.sort((a, b) => parseFloat(b[2]) - parseFloat(a[2]));
  
  // Ambil 20 santri dengan saldo tertinggi
  const topStudents = students.slice(0, 20);
  
  // Buat sheet laporan baru
  let reportSheet = ss.getSheetByName('Laporan Saldo Tertinggi');
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet('Laporan Saldo Tertinggi');
  
  // Set header dan judul
  reportSheet.getRange('A1').setValue('LAPORAN 20 SANTRI DENGAN SALDO TERTINGGI');
  reportSheet.getRange('A1:D1').merge();
  reportSheet.getRange('A1').setFontWeight('bold').setHorizontalAlignment('center');
  
  reportSheet.getRange('A3:C3').setValues([['NIS', 'Nama', 'Saldo']]);
  reportSheet.getRange('A3:C3').setFontWeight('bold').setBackground('#f3f3f3');
  
  // Isi data
  const reportData = topStudents.map(student => [student[0], student[1], student[2]]);
  reportSheet.getRange(4, 1, reportData.length, 3).setValues(reportData);
  
  // Format saldo
  reportSheet.getRange(4, 3, reportData.length, 1).setNumberFormat('#,##0');
  
  // Auto-resize dan styling
  reportSheet.autoResizeColumns(1, 3);
  reportSheet.getRange('A1:C1').setBackground('#e6e6e6');
  
  SpreadsheetApp.getUi().alert('Laporan berhasil dibuat.');
}

/**
 * Membuat laporan saldo terendah
 */
function createLowestBalanceReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek sheet data santri
  const dataSheet = ss.getSheetByName('Data Santri');
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('Data santri tidak ditemukan. Silakan sinkronisasi data terlebih dahulu.');
    return;
  }
  
  // Ambil data santri
  const dataRange = dataSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Filter hanya header dan data (buang header)
  const header = data[0];
  const students = data.slice(1);
  
  // Sort berdasarkan saldo (ascending)
  students.sort((a, b) => parseFloat(a[2]) - parseFloat(b[2]));
  
  // Ambil 20 santri dengan saldo terendah
  const lowestStudents = students.slice(0, 20);
  
  // Buat sheet laporan baru
  let reportSheet = ss.getSheetByName('Laporan Saldo Terendah');
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet('Laporan Saldo Terendah');
  
  // Set header dan judul
  reportSheet.getRange('A1').setValue('LAPORAN 20 SANTRI DENGAN SALDO TERENDAH');
  reportSheet.getRange('A1:D1').merge();
  reportSheet.getRange('A1').setFontWeight('bold').setHorizontalAlignment('center');
  
  reportSheet.getRange('A3:C3').setValues([['NIS', 'Nama', 'Saldo']]);
  reportSheet.getRange('A3:C3').setFontWeight('bold').setBackground('#f3f3f3');
  
  // Isi data
  const reportData = lowestStudents.map(student => [student[0], student[1], student[2]]);
  reportSheet.getRange(4, 1, reportData.length, 3).setValues(reportData);
  
  // Format saldo
  reportSheet.getRange(4, 3, reportData.length, 1).setNumberFormat('#,##0');
  
  // Auto-resize dan styling
  reportSheet.autoResizeColumns(1, 3);
  reportSheet.getRange('A1:C1').setBackground('#e6e6e6');
  
  SpreadsheetApp.getUi().alert('Laporan berhasil dibuat.');
}

/**
 * Membuat laporan transaksi harian
 */
function createDailyTransactionReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek sheet transaksi
  const transactionSheet = ss.getSheetByName('Transaksi');
  if (!transactionSheet) {
    SpreadsheetApp.getUi().alert('Data transaksi tidak ditemukan. Silakan lakukan transaksi terlebih dahulu.');
    return;
  }
  
  // Ambil data transaksi
  const dataRange = transactionSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Filter hanya header dan data (buang header)
  const header = data[0];
  const transactions = data.slice(1);
  
  // Input tanggal
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  
  const response = ui.prompt(
    'Laporan Harian',
    `Masukkan tanggal (format: ${formattedDate}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) return;
  
  const inputDate = response.getResponseText().trim();
  
  // Validasi dan parsing tanggal
  let targetDate;
  try {
    const parts = inputDate.split('/');
    if (parts.length !== 3) throw new Error('Format tanggal tidak valid');
    
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // Bulan di JavaScript dimulai dari 0
    const year = parseInt(parts[2], 10);
    
    targetDate = new Date(year, month, day);
    
    if (isNaN(targetDate.getTime())) throw new Error('Tanggal tidak valid');
  } catch (e) {
    ui.alert('Format tanggal tidak valid. Gunakan format DD/MM/YYYY.');
    return;
  }
  
  // Filter transaksi sesuai tanggal
  const filteredTransactions = transactions.filter(transaction => {
    const transactionDate = new Date(transaction[0]);
    return transactionDate.getDate() === targetDate.getDate() &&
           transactionDate.getMonth() === targetDate.getMonth() &&
           transactionDate.getFullYear() === targetDate.getFullYear();
  });
  
  if (filteredTransactions.length === 0) {
    ui.alert('Tidak ada transaksi pada tanggal tersebut.');
    return;
  }
  
  // Buat sheet laporan baru
  const reportSheetName = 'Laporan Transaksi ' + Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'dd-MM-yyyy');
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet(reportSheetName);
  
  // Set header dan judul
  reportSheet.getRange('A1').setValue('LAPORAN TRANSAKSI HARIAN');
  reportSheet.getRange('A1:E1').merge();
  reportSheet.getRange('A1').setFontWeight('bold').setHorizontalAlignment('center');
  
  reportSheet.getRange('A2').setValue('Tanggal: ' + Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'dd MMMM yyyy'));
  reportSheet.getRange('A2:E2').merge();
  reportSheet.getRange('A2').setHorizontalAlignment('center');
  
  reportSheet.getRange('A4:E4').setValues([['Waktu', 'NIS', 'Nama', 'Nominal', 'Status']]);
  reportSheet.getRange('A4:E4').setFontWeight('bold').setBackground('#f3f3f3');
  
  // Isi data
  const reportData = filteredTransactions.map(transaction => [
    Utilities.formatDate(new Date(transaction[0]), Session.getScriptTimeZone(), 'HH:mm:ss'),
    transaction[1],
    transaction[2],
    transaction[3],
    transaction[4]
  ]);
  reportSheet.getRange(5, 1, reportData.length, 5).setValues(reportData);
  
  // Format nominal
  reportSheet.getRange(5, 4, reportData.length, 1).setNumberFormat('#,##0');
  
  // Hitung total transaksi
  const totalAmount = filteredTransactions.reduce((sum, transaction) => sum + parseFloat(transaction[3]), 0);
  const rowTotal = reportData.length + 5;
  
  reportSheet.getRange(rowTotal, 1).setValue('TOTAL');
  reportSheet.getRange(rowTotal, 1, 1, 3).merge();
  reportSheet.getRange(rowTotal, 1).setFontWeight('bold').setHorizontalAlignment('right');
  reportSheet.getRange(rowTotal, 4).setValue(totalAmount).setNumberFormat('#,##0').setFontWeight('bold');
  
  // Auto-resize dan styling
  reportSheet.autoResizeColumns(1, 5);
  reportSheet.getRange('A1:E2').setBackground('#e6e6e6');
  
  ui.alert('Laporan berhasil dibuat.');
}

/**
 * Membuat laporan transaksi bulanan
 */
function createMonthlyTransactionReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cek sheet transaksi
  const transactionSheet = ss.getSheetByName('Transaksi');
  if (!transactionSheet) {
    SpreadsheetApp.getUi().alert('Data transaksi tidak ditemukan. Silakan lakukan transaksi terlebih dahulu.');
    return;
  }
  
  // Ambil data transaksi
  const dataRange = transactionSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Filter hanya header dan data (buang header)
  const header = data[0];
  const transactions = data.slice(1);
  
  // Input bulan dan tahun
  const ui = SpreadsheetApp.getUi();
  const today = new Date();
  const formattedMonth = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/yyyy');
  
  const response = ui.prompt(
    'Laporan Bulanan',
    `Masukkan bulan (format: ${formattedMonth}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() != ui.Button.OK) return;
  
  const inputMonth = response.getResponseText().trim();
  
  // Validasi dan parsing bulan/tahun
  let targetMonth, targetYear;
  try {
    const parts = inputMonth.split('/');
    if (parts.length !== 2) throw new Error('Format bulan tidak valid');
    
    targetMonth = parseInt(parts[0], 10) - 1; // Bulan di JavaScript dimulai dari 0
    targetYear = parseInt(parts[1], 10);
    
    if (isNaN(targetMonth) || isNaN(targetYear) || targetMonth < 0 || targetMonth > 11) {
      throw new Error('Bulan tidak valid');
    }
  } catch (e) {
    ui.alert('Format bulan tidak valid. Gunakan format MM/YYYY.');
    return;
  }
  
  // Filter transaksi sesuai bulan dan tahun
  const filteredTransactions = transactions.filter(transaction => {
    const transactionDate = new Date(transaction[0]);
    return transactionDate.getMonth() === targetMonth &&
           transactionDate.getFullYear() === targetYear;
  });
  
  if (filteredTransactions.length === 0) {
    ui.alert('Tidak ada transaksi pada bulan tersebut.');
    return;
  }
  
  // Buat sheet laporan baru
  const monthName = Utilities.formatDate(new Date(targetYear, targetMonth, 1), Session.getScriptTimeZone(), 'MMMM yyyy');
  const reportSheetName = 'Laporan Transaksi ' + monthName;
  let reportSheet = ss.getSheetByName(reportSheetName);
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  reportSheet = ss.insertSheet(reportSheetName);
  
  // Set header dan judul
  reportSheet.getRange('A1').setValue('LAPORAN TRANSAKSI BULANAN');
  reportSheet.getRange('A1:E1').merge();
  reportSheet.getRange('A1').setFontWeight('bold').setHorizontalAlignment('center');
  
  reportSheet.getRange('A2').setValue('Bulan: ' + monthName);
  reportSheet.getRange('A2:E2').merge();
  reportSheet.getRange('A2').setHorizontalAlignment('center');
  
  reportSheet.getRange('A4:E4').setValues([['Tanggal', 'NIS', 'Nama', 'Nominal', 'Status']]);
  reportSheet.getRange('A4:E4').setFontWeight('bold').setBackground('#f3f3f3');
  
  // Isi data
  const reportData = filteredTransactions.map(transaction => [
    Utilities.formatDate(new Date(transaction[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
    transaction[1],
    transaction[2],
    transaction[3],
    transaction[4]
  ]);
  reportSheet.getRange(5, 1, reportData.length, 5).setValues(reportData);
  
  // Format nominal
  reportSheet.getRange(5, 4, reportData.length, 1).setNumberFormat('#,##0');
  
  // Hitung total transaksi
  const totalAmount = filteredTransactions.reduce((sum, transaction) => sum + parseFloat(transaction[3]), 0);
  const rowTotal = reportData.length + 5;
  
  reportSheet.getRange(rowTotal, 1).setValue('TOTAL');
  reportSheet.getRange(rowTotal, 1, 1, 3).merge();
  reportSheet.getRange(rowTotal, 1).setFontWeight('bold').setHorizontalAlignment('right');
  reportSheet.getRange(rowTotal, 4).setValue(totalAmount).setNumberFormat('#,##0').setFontWeight('bold');
  
  // Auto-resize dan styling
  reportSheet.autoResizeColumns(1, 5);
  reportSheet.getRange('A1:E2').setBackground('#e6e6e6');
  
  ui.alert('Laporan berhasil dibuat.');
}

// ==========================================
// Fungsi Utilitas
// ==========================================

/**
 * Format angka menjadi format mata uang
 */
function formatCurrency(amount) {
  return parseInt(amount).toLocaleString('id-ID');
}
