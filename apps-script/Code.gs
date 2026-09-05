// ===================================================
// FUNGSI onOpen()
// Otomatis jalan setiap spreadsheet dibuka, nambahin menu custom.
// ===================================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Dashboard Tools')
      .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
      .addToUi();
}

// ===================================================
// FUNGSI refreshDashboard()
// Dijalankan waktu tombol menu "Refresh Dashboard" diklik.
// ===================================================
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName('Dashboard');
  SpreadsheetApp.flush();
  var now = new Date();
  dashboardSheet.getRange('N1').setValue('Terakhir di-refresh: ' + now.toLocaleString());
  ss.toast('Dashboard berhasil di-refresh!', 'Selesai', 3);
}

// ===================================================
// FUNGSI doGet()
// WAJIB ada supaya Web App bisa menampilkan halaman HTML.
// ===================================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Supermarket Sales Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ===================================================
// FUNGSI getDashboardData()
// TUGASNYA CUMA SATU: ambil data mentah dari sheet Raw_Data,
// lalu ubah jadi array of object yang rapi (bukan array angka biasa),
// supaya mudah dipakai & difilter di JavaScript sisi HTML nanti.
//
// Kita SENGAJA tidak menghitung SUM/AVERAGE apapun di sini.
// Semua penghitungan/analisis dikerjakan di Dashboard.html,
// supaya waktu user pilih filter, tidak perlu panggil fungsi ini lagi.
// ===================================================
function getDashboardData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Raw_Data');
  var data = sheet.getDataRange().getValues();

  var headers = data[0];
  var rows = data.slice(1);

  // Cari posisi (index) setiap kolom yang kita perlukan, berdasarkan nama headernya.
  // Kalau urutan kolom di sheet berubah nanti, kode ini tetap jalan karena
  // dia cari berdasarkan NAMA, bukan posisi tetap.
  var invoiceIdx   = headers.indexOf('Invoice ID');
  var branchIdx    = headers.indexOf('Branch');
  var cityIdx      = headers.indexOf('City');
  var custTypeIdx  = headers.indexOf('Customer type');
  var genderIdx    = headers.indexOf('Gender');
  var productIdx   = headers.indexOf('Product line');
  var totalIdx     = headers.indexOf('Total');
  var paymentIdx   = headers.indexOf('Payment');
  var ratingIdx    = headers.indexOf('Rating');
  var monthIdx     = headers.indexOf('Month');

  // Ubah setiap baris mentah (array angka/teks tanpa nama)
  // jadi object yang punya "label" jelas, contoh:
  // { branch: "A", product: "Health and beauty", total: 548.97, ... }
  var transactions = rows.map(function(row) {
    return {
      invoiceId:    row[invoiceIdx],
      branch:       row[branchIdx],
      city:         row[cityIdx],
      customerType: row[custTypeIdx],
      gender:       row[genderIdx],
      product:      row[productIdx],
      // Number(...) || 0  => pastikan selalu jadi angka, kalau kosong/rusak jadi 0
      total:        Number(row[totalIdx]) || 0,
      payment:      row[paymentIdx],
      rating:       Number(row[ratingIdx]) || 0,
      month:        Number(row[monthIdx]) || 0
    };
  });

  return transactions;
}
