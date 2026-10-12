
const SPREADSHEET_ID = "1rCp0bz-Fka5VhMLM9GD97v8BzsOKZTesFhaxIlTZtHc";

// ============================================================
// PRODUK
// ============================================================
const PRODUCTS = {
  "Nasi + Ayam": { harga: 10000 },
  "Ayam": { harga: 8000 },
  "Nasi": { harga: 3000 },
  "Saus / Sambal": { harga: 3000 }
};

// ============================================================
// SPREADSHEET HELPER
// ============================================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ============================================================
// NUMBER HELPER
// ============================================================
function toNumber(value) {
  const number = Number(value);
  return isFinite(number) ? number : 0;
}


// 1. HANDLER PERMINTAAN POST
function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents);
    var action = requestData.action;

    if (action === "recordAttendance") { return recordAttendance(requestData); }
    if (action === "sale") { return recordSale(requestData); }
    if (action === "updateSale") { return updateSaleRecord(requestData); } // FITUR EDIT
    if (action === "deleteSale") { return deleteSaleRecord(requestData); } // FITUR DELETE
    if (action === "saveHPP") { return saveHPP(requestData); }
    if (action === "updateMasterHPP") { return updateMasterHPP(requestData); }
    if (action === "archiveYearlyData") { return archiveYearlyData(); }
    if (action === "saveSalary") { return saveSalaryRecord(requestData); }
    if (action === "logError") { return recordErrorLog(requestData); }
    if (action === "saveAdvance") { return saveAdvanceRecord(requestData); }

    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: "Action tidak dikenal: " + action 
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "error", 
      message: err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 2. HANDLER PERMINTAAN GET
function doGet(e) {
  var action = e.parameter.action;
  if (action === "verifyPin") { return verifyPin(e.parameter.pin); }
  if (action === "getData") { return getData(); }
  return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Action GET tidak valid" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function verifyPin(pin) {
  var cleanPin = String(pin || "").trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Staff");

  if (sheet) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var rowPin = String(data[i][0] || "").trim();
      var rowName = String(data[i][1] || "").trim();
      var rowRole = String(data[i][2] || "").trim().toLowerCase();

      if (rowPin === cleanPin && cleanPin !== "") {
        return ContentService.createTextOutput(JSON.stringify({
          status: "success",
          valid: true,
          role: rowRole || "kasir",
          employeeName: rowName || "Kasir Shift"
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", valid: false })).setMimeType(ContentService.MimeType.JSON);
}

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var salesSheet = ss.getSheetByName("Penjualan");
  var salesData = salesSheet ? getSheetDataSafe(salesSheet) : [];

  var masterHppSheet = ss.getSheetByName("Master HPP");
  var masterHppData = masterHppSheet ? getSheetDataSafe(masterHppSheet) : [];

  var hppSheet = ss.getSheetByName("HPP");
  var hppData = hppSheet ? getSheetDataSafe(hppSheet) : [];

  var absensiSheet = ss.getSheetByName("Absensi");
  var absensiData = absensiSheet ? getSheetDataSafe(absensiSheet) : [];

  var gajiSheet = ss.getSheetByName("Gaji");
  var gajiData = gajiSheet ? getSheetDataSafe(gajiSheet) : [];

  var errorSheet = ss.getSheetByName("ErrorLogs");
  var errorData = errorSheet ? getSheetDataSafe(errorSheet) : [];

  var kasbonSheet = ss.getSheetByName("Kasbon");
  var kasbonData = kasbonSheet ? getSheetDataSafe(kasbonSheet) : [];

  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: salesData,
    masterHpp: masterHppData,
    hppData: hppData,
    absensiData: absensiData,
    gajiData: gajiData,
    errorData: errorData,
    kasbonData: kasbonData
  })).setMimeType(ContentService.MimeType.JSON);
}

function getSheetDataSafe(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      var headerName = String(headers[j]).trim();
      var val = data[i][j];

      if (val instanceof Date) {
        val = Utilities.formatDate(val, timeZone, "yyyy-MM-dd HH:mm:ss");
      }
      obj[headerName] = val;
    }
    obj["_rowIndex"] = i + 1; // Baris asli di Sheet untuk keperluan Edit/Delete
    result.push(obj);
  }
  return result;
}

function recordAttendance(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Absensi");

  if (!sheet) {
    sheet = ss.insertSheet("Absensi");
    sheet.appendRow(["ID Absen", "Tanggal", "Nama Karyawan", "Role", "Jam Masuk", "Jam Keluar"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var dateStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  var timeStr = Utilities.formatDate(now, timeZone, "HH:mm:ss");

  var name = data.employeeName || (data.role === "admin" ? "Owner / Admin" : "Kasir Shift");
  var role = data.role || "kasir";

  var values = sheet.getDataRange().getValues();
  var existingRowIndex = -1;
  var existingAttId = "";

  for (var i = 1; i < values.length; i++) {
    var rowDate = values[i][1];
    var rowName = values[i][2];

    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, timeZone, "yyyy-MM-dd");
    }

    if (String(rowDate).trim() === dateStr && String(rowName).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      existingRowIndex = i + 1;
      existingAttId = values[i][0];
      break;
    }
  }

  if (data.type === "clockIn") {
    if (existingRowIndex > 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "success", attendanceId: existingAttId })).setMimeType(ContentService.MimeType.JSON);
    } else {
      var newAttendanceId = "ABS-" + now.getTime();
      sheet.appendRow([newAttendanceId, dateStr, name, role, timeStr, "-"]);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", attendanceId: newAttendanceId })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  if (data.type === "clockOut") {
    if (existingRowIndex > 0) {
      sheet.getRange(existingRowIndex, 6).setValue(timeStr);
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "error" })).setMimeType(ContentService.MimeType.JSON);
}

function recordSale(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Penjualan");

  if (!sheet) {
    sheet = ss.insertSheet("Penjualan");
    sheet.appendRow(["ID Transaksi", "Tanggal & Waktu", "Nama Produk", "Jumlah", "Metode Pembayaran", "Harga Satuan", "Total Harga"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var timeStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss");
  var transactionId = "TRX-" + now.getTime();

  var priceMap = {
    "Nasi + Ayam": 10000,
    "Ayam": 8000,
    "Nasi": 3000,
    "Saus / Sambal": 3000
  };

  var items = data.items || [];
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var nama = item.namaProduk;
    var qty = Number(item.jumlah) || 1;
    var harga = priceMap[nama] || 0;
    var total = harga * qty;

    sheet.appendRow([transactionId, "'" + timeStr, nama, qty, data.metodePembayaran || "Tunai", harga, total]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", transactionId: transactionId })).setMimeType(ContentService.MimeType.JSON);
}

// 6B. EDIT TRANSAKSI PENJUALAN
function updateSaleRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Penjualan");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);

  var rowIndex = Number(data.rowIndex);
  if (!rowIndex || rowIndex <= 1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Baris tidak valid" })).setMimeType(ContentService.MimeType.JSON);

  var qty = Number(data.jumlah) || 1;
  var harga = Number(data.hargaSatuan) || 0;
  var total = qty * harga;

  sheet.getRange(rowIndex, 3).setValue(data.namaProduk);
  sheet.getRange(rowIndex, 4).setValue(qty);
  sheet.getRange(rowIndex, 5).setValue(data.metodePembayaran);
  sheet.getRange(rowIndex, 6).setValue(harga);
  sheet.getRange(rowIndex, 7).setValue(total);

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Transaksi berhasil diperbarui!" })).setMimeType(ContentService.MimeType.JSON);
}

// 6C. HAPUS TRANSAKSI PENJUALAN
function deleteSaleRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Penjualan");
  if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Sheet tidak ditemukan" })).setMimeType(ContentService.MimeType.JSON);

  var rowIndex = Number(data.rowIndex);
  if (!rowIndex || rowIndex <= 1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Baris tidak valid" })).setMimeType(ContentService.MimeType.JSON);

  sheet.deleteRow(rowIndex);
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Transaksi berhasil dihapus!" })).setMimeType(ContentService.MimeType.JSON);
}

function saveHPP(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("HPP");

  if (!sheet) {
    sheet = ss.insertSheet("HPP");
    sheet.appendRow(["Tanggal & Waktu", "Nama Bahan Baku", "Kuantitas Standar Harian", "Biaya Standar Harian (Rp)"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var timeStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss");
  var materials = data.materials || [];

  for (var i = 0; i < materials.length; i++) {
    var m = materials[i];
    sheet.appendRow([timeStr, m.nama, m.qty, m.biaya]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data HPP berhasil dicatat" })).setMimeType(ContentService.MimeType.JSON);
}

function updateMasterHPP(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Master HPP");

  if (!sheet) {
    sheet = ss.insertSheet("Master HPP");
  } else {
    sheet.clear();
  }

  sheet.appendRow(["Nama Bahan Baku", "Kuantitas Standar Harian", "Harga Satuan", "Biaya Standar Harian (Rp)"]);

  var list = data.masterMaterials || [];
  for (var i = 0; i < list.length; i++) {
    var item = list[i];
    var qty = Number(item.qty) || 0;
    var harga = Number(item.hargaSatuan) || 0;
    var biaya = qty * harga;
    sheet.appendRow([item.nama, qty, harga, biaya]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Master HPP berhasil diperbarui" })).setMimeType(ContentService.MimeType.JSON);
}

function archiveYearlyData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var year = new Date().getFullYear();
  var archiveSuffix = "_" + year;

  var sheetsToArchive = ["Penjualan", "HPP"];
  for (var i = 0; i < sheetsToArchive.length; i++) {
    var sName = sheetsToArchive[i];
    var s = ss.getSheetByName(sName);
    if (s && s.getLastRow() > 1) {
      var archiveSheetName = sName + archiveSuffix;
      var archiveSheet = ss.getSheetByName(archiveSheetName) || ss.insertSheet(archiveSheetName);
      
      var data = s.getDataRange().getValues();
      archiveSheet.clear();
      for (var r = 0; r < data.length; r++) {
        archiveSheet.appendRow(data[r]);
      }

      var headers = data[0];
      s.clear();
      s.appendRow(headers);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data tahun " + year + " berhasil diarsipkan!" })).setMimeType(ContentService.MimeType.JSON);
}

function saveSalaryRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Gaji");

  if (!sheet) {
    sheet = ss.insertSheet("Gaji");
    sheet.appendRow(["ID Gaji", "Tanggal Bayar", "Nama Karyawan", "Periode", "Hari Kerja", "Kasbon/Potongan", "Gaji Bersih", "Catatan"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var timeStr = Utilities.formatDate(now, timeZone, "dd/MM/yyyy HH:mm:ss");

  var netSalaryNum = Number(data.netSalary) || 0;
  var deductionNum = Number(data.deduction) || 0;
  var workDaysNum = Number(data.workDays) || 0;
  var salId = "SAL-" + now.getTime();
  var empName = String(data.employeeName || "-");
  var periodStr = String(data.period || "-");
  var notesStr = String(data.notes || "Gaji Kasir");

  var headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  
  if (!headers || headers.length === 0 || !headers[0]) {
    sheet.appendRow(["ID Gaji", "Tanggal Bayar", "Nama Karyawan", "Periode", "Hari Kerja", "Kasbon/Potongan", "Gaji Bersih", "Catatan"]);
    headers = ["ID Gaji", "Tanggal Bayar", "Nama Karyawan", "Periode", "Hari Kerja", "Kasbon/Potongan", "Gaji Bersih", "Catatan"];
  }

  var newRow = new Array(headers.length);

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (h.includes("id")) newRow[i] = salId;
    else if (h.includes("bayar") || h.includes("waktu") || h.includes("tgl") || h.includes("tanggal")) newRow[i] = timeStr;
    else if (h.includes("nama") || h.includes("karyawan")) newRow[i] = empName;
    else if (h.includes("periode")) newRow[i] = periodStr;
    else if (h.includes("hari") || h.includes("kerja")) newRow[i] = workDaysNum;
    else if (h.includes("potong") || h.includes("kasbon")) newRow[i] = deductionNum;
    else if (h.includes("bersih") || h.includes("gaji") || h.includes("total")) newRow[i] = netSalaryNum;
    else if (h.includes("catat") || h.includes("note")) newRow[i] = notesStr;
    else newRow[i] = "-";
  }

  sheet.appendRow(newRow);

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data gaji berhasil disimpan!" })).setMimeType(ContentService.MimeType.JSON);
}

function recordErrorLog(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("ErrorLogs");
  if (!sheet) {
    sheet = ss.insertSheet("ErrorLogs");
    sheet.appendRow(["ID Error", "Waktu", "User / Role", "Pesan Error", "Detail / Stack Trace"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var timeStr = Utilities.formatDate(now, timeZone, "dd/MM/yyyy HH:mm:ss");

  sheet.appendRow(["ERR-" + now.getTime(), timeStr, data.user || "System", data.message || "Unknown Error", data.detail || "-"]);
  return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
}

function saveAdvanceRecord(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Kasbon");

  if (!sheet) {
    sheet = ss.insertSheet("Kasbon");
    sheet.appendRow(["ID Kasbon", "Tanggal", "Nama Karyawan", "Nominal (Rp)", "Keterangan", "Status"]);
  }

  var now = new Date();
  var timeZone = ss.getSpreadsheetTimeZone();
  var timeStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    "KAS-" + now.getTime(),
    "'" + timeStr,
    data.employeeName || "-",
    Number(data.amount) || 0,
    data.notes || "Kasbon Karyawan",
    "Belum Lunas"
  ]);

  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Kasbon berhasil dicatat!" })).setMimeType(ContentService.MimeType.JSON);
}
