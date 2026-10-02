/**
 * GOOGLE APPS SCRIPT FOR MAKMAL KOMPUTER SKSA
 * Menyimpan Data Tempahan (Tab: Tempahan), Pengguna (Tab: Pengguna) & Senarai Kelas (Tab: Kelas)
 * dengan Sokongan UserID Automatik, Pengurusan Kelas Fleksibel, Robust In-Place Row Status Update,
 * dan Pemprosesan Pukal (BATCH_ADD) Berprestasi Tinggi dengan LockService.
 */

// [PILIHAN] Masukkan ID Google Sheet jika menggunakan skrip bebas (Standalone Script) di script.google.com
// Contoh ID dari URL Sheet: https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
// Jika skrip ini dibuka terus dari Google Sheet (Extensions > Apps Script), biarkan kosong "".
var SPREADSHEET_ID = "";

/**
 * FUNGSI UJIAN: Klik butang 'Run' / 'Jalankan' pada fungsi ini untuk menguji sambungan
 */
function testRun() {
  var ss = getSpreadsheet();
  if (!ss) {
    Logger.log("Ralat: Spreadsheet tidak dapat dikesan. Sila pastikan skrip dibuka dari Extensions > Apps Script pada Google Sheet, atau masukkan SPREADSHEET_ID.");
    return "Ralat: Spreadsheet tidak dikesan.";
  }
  Logger.log("Berjaya bersambung ke Spreadsheet: " + ss.getName());
  var sheet = getOrCreateSheet(ss, "Tempahan");
  Logger.log("Tab 'Tempahan' sedia: " + sheet.getName());
  var userSheet = getOrCreateSheet(ss, "Pengguna");
  Logger.log("Tab 'Pengguna' sedia: " + userSheet.getName());
  var classSheet = getOrCreateSheet(ss, "Kelas");
  Logger.log("Tab 'Kelas' sedia: " + classSheet.getName());
  return "Berjaya! Pangkalan data sedia.";
}

function getSpreadsheet() {
  var ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}

  if (!ss && typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID.trim());
    } catch (e2) {
      Logger.log("Ralat membuka Spreadsheet melalui SPREADSHEET_ID: " + e2.toString());
    }
  }

  if (!ss) {
    try {
      var active = SpreadsheetApp.getActiveSheet();
      if (active && active.getParent()) {
        ss = active.getParent();
      }
    } catch (e3) {}
  }

  return ss;
}

function getOrCreateSheet(ss, targetName) {
  if (!ss || typeof ss.getSheets !== 'function') {
    ss = getSpreadsheet();
  }
  if (!ss) {
    throw new Error("Pangkalan data Google Sheet tidak ditemui! Sila pastikan: (1) Skrip ini dibuka dari Google Sheet (Extensions > Apps Script), atau (2) Masukkan SPREADSHEET_ID pada baris atas skrip jika menggunakan skrip standalone.");
  }
  if (!targetName) {
    targetName = "Tempahan";
  }

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName().trim().toLowerCase();
    if (name === targetName.toLowerCase()) {
      return sheets[i];
    }
  }
  var newSheet = ss.insertSheet(targetName);
  return newSheet;
}

function formatDateStr(val) {
  if (!val) return "";
  if (val instanceof Date) {
    var yr = val.getFullYear();
    var mo = ("0" + (val.getMonth() + 1)).slice(-2);
    var dy = ("0" + val.getDate()).slice(-2);
    return yr + "-" + mo + "-" + dy;
  }
  var s = String(val).trim();
  if (s.indexOf("T") !== -1) {
    s = s.split("T")[0];
  }
  return s;
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : "";
  var ss = getSpreadsheet();

  // 1. DAPATKAN SENARAI PENGGUNA (GET_USERS)
  if (action === "GET_USERS") {
    var userSheet = getOrCreateSheet(ss, "Pengguna");
    var data = userSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return responseJSON({ status: "success", users: [] });
    }
    var users = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] || row[1]) {
        users.push({
          userId: String(row[0] || ""),
          email: String(row[1] || row[0] || ""),
          name: String(row[2] || ""),
          role: String(row[3] || "Guru"),
          phone: String(row[4] || ""),
          subject: String(row[5] || ""),
          loginTime: String(row[6] || ""),
          authProvider: String(row[7] || "")
        });
      }
    }
    return responseJSON({ status: "success", users: users });
  }

  // 2. DAPATKAN SENARAI KELAS (GET_CLASSES)
  if (action === "GET_CLASSES") {
    var classSheet = getOrCreateSheet(ss, "Kelas");
    if (classSheet.getLastRow() === 0) {
      classSheet.appendRow(["ID Kelas", "Nama Kelas", "Tahap", "Bilangan Murid / PC", "Catatan", "Tarikh Kemaskini"]);
      classSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e8f0fe");
      var initialClasses = [
        ["CLS-1UTARID", "1 UTARID", "Tahap 1", 35, "Tahap 1 (Waktu balik 12:30/1:00 PM, Rehat 10:00-10:30 AM)", new Date().toLocaleString('ms-MY')],
        ["CLS-2ZUHRAH", "2 ZUHRAH", "Tahap 1", 35, "Tahap 1 (Waktu balik 12:30/1:00 PM, Rehat 10:00-10:30 AM)", new Date().toLocaleString('ms-MY')],
        ["CLS-3MARIKH", "3 MARIKH", "Tahap 1", 35, "Tahap 1 (Waktu balik 12:30/1:00 PM, Rehat 10:00-10:30 AM)", new Date().toLocaleString('ms-MY')],
        ["CLS-4MUSYTARI", "4 MUSYTARI", "Tahap 2", 35, "Tahap 2 (Waktu balik 1:30 PM, Rehat 10:30-11:00 AM)", new Date().toLocaleString('ms-MY')],
        ["CLS-5ZUHAL", "5 ZUHAL", "Tahap 2", 35, "Tahap 2 (Waktu balik 1:30 PM, Rehat 10:30-11:00 AM)", new Date().toLocaleString('ms-MY')],
        ["CLS-6NEPTUN", "6 NEPTUN", "Tahap 2", 35, "Tahap 2 (Waktu balik 1:30 PM, Rehat 10:30-11:00 AM)", new Date().toLocaleString('ms-MY')]
      ];
      classSheet.getRange(2, 1, initialClasses.length, 6).setValues(initialClasses);
    }

    var data = classSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return responseJSON({ status: "success", classes: [] });
    }
    var classes = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0] || row[1]) {
        classes.push({
          id: String(row[0] || ""),
          name: String(row[1] || ""),
          level: String(row[2] || "Tahap 1"),
          pcs: Number(row[3]) || 35,
          notes: String(row[4] || ""),
          updatedAt: String(row[5] || "")
        });
      }
    }
    return responseJSON({ status: "success", classes: classes });
  }

  // 3. DAPATKAN SENARAI TEMPAHAN (Default GET)
  var bookingSheet = getOrCreateSheet(ss, "Tempahan");
  var data = bookingSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return responseJSON({ status: "success", data: [] });
  }

  var bookings = [];
  var hasUserIdCol = (data[0][1] && String(data[0][1]).toLowerCase().includes("user"));

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!row[0]) continue;
    
    var bId = String(row[0] || "").trim();
    var bUserId = hasUserIdCol ? String(row[1] || "") : "";
    var bDate = formatDateStr(hasUserIdCol ? row[2] : row[1]);
    var bSlot = String(hasUserIdCol ? row[3] : row[2] || "");
    var bApplicant = String(hasUserIdCol ? row[4] : row[3] || "");
    var bRole = String(hasUserIdCol ? row[5] : row[4] || "Guru");
    var bSubject = String(hasUserIdCol ? row[6] : row[5] || "");
    var bPcCount = hasUserIdCol ? row[7] : row[6];
    var bPurpose = hasUserIdCol ? row[8] : row[7];
    var bStatus = String(hasUserIdCol ? row[11] : (row[10] || row[8] || "Diluluskan"));
    var bCreatedAt = hasUserIdCol ? row[12] : (row[11] || row[9] || new Date().toISOString());

    bookings.push({
      id: bId,
      userId: bUserId,
      date: bDate,
      slot: bSlot,
      applicant: bApplicant,
      role: bRole,
      subject: bSubject,
      pcCount: bPcCount,
      purpose: bPurpose,
      status: bStatus,
      createdAt: bCreatedAt
    });
  }

  return responseJSON({ status: "success", data: bookings });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000); // Tunggu giliran sehingga 30 saat untuk elak konflik 'race condition'
  } catch (lockErr) {}

  try {
    var contents = e.postData.contents;
    var data = JSON.parse(contents);
    var ss = getSpreadsheet();

    // 1. TAMBAH / KEMASKINI REKOD PENGGUNA (Tab: Pengguna)
    if (data.action === "RECORD_USER_ACCOUNT") {
      var userSheet = getOrCreateSheet(ss, "Pengguna");
      
      if (userSheet.getLastRow() === 0) {
        userSheet.appendRow([
          "UserID", "Emel ID DELIMa", "Nama Penuh", "Peranan / Jawatan",
          "No. Telefon", "Mata Pelajaran", "Log Masuk Terakhir", "Penyedia Log Masuk"
        ]);
        userSheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e8f0fe");
      }

      var email = (data.email || "").toLowerCase().trim();
      var name = data.name || "";
      var role = data.role || "Guru";
      var phone = data.phone || "";
      var subject = data.subject || "";
      var loginTime = data.loginTime || new Date().toLocaleString('ms-MY');
      var authProvider = data.authProvider || "DELIMa / Google SSO";

      var userId = data.userId || "";
      if (!userId && email) {
        var numMatch = email.match(/\d+/);
        if (numMatch) {
          userId = "USR-" + numMatch[0];
        } else {
          var cleanEmail = email.split('@')[0].replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
          userId = "USR-" + cleanEmail;
        }
      }

      var allValues = userSheet.getDataRange().getValues();
      var rowIndex = -1;

      for (var i = 1; i < allValues.length; i++) {
        var rowCell0 = String(allValues[i][0] || "").toLowerCase().trim();
        var rowCell1 = String(allValues[i][1] || "").toLowerCase().trim();
        if ((email && (rowCell0 === email || rowCell1 === email)) || (userId && rowCell0 === userId.toLowerCase())) {
          rowIndex = i + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        userSheet.getRange(rowIndex, 1).setValue(userId);
        userSheet.getRange(rowIndex, 2).setValue(email);
        userSheet.getRange(rowIndex, 3).setValue(name);
        userSheet.getRange(rowIndex, 4).setValue(role);
        if (phone) userSheet.getRange(rowIndex, 5).setValue(phone);
        if (subject) userSheet.getRange(rowIndex, 6).setValue(subject);
        userSheet.getRange(rowIndex, 7).setValue(loginTime);
        userSheet.getRange(rowIndex, 8).setValue(authProvider);
      } else {
        userSheet.appendRow([userId, email, name, role, phone, subject, loginTime, authProvider]);
      }

      return responseJSON({ status: "success", userId: userId, message: "Rekod pengguna diselaraskan." });
    }

    // 2. PENGURUSAN KELAS OLEH PENYELARAS ICT (Tab: Kelas)
    if (data.action === "SYNC_CLASSES") {
      var classSheet = getOrCreateSheet(ss, "Kelas");
      classSheet.clearContents();
      classSheet.appendRow(["ID Kelas", "Nama Kelas", "Tahap", "Bilangan Murid / PC", "Catatan", "Tarikh Kemaskini"]);
      classSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e8f0fe");

      var list = data.classes || [];
      if (list.length > 0) {
        var rowsToAdd = list.map(function(c, idx) {
          var cid = c.id || ("CLS-" + (c.name ? c.name.replace(/\s+/g, '').toUpperCase() : (idx + 1)));
          return [
            cid,
            c.name || "",
            c.level || "Tahap 1",
            c.pcs || 35,
            c.notes || "",
            c.updatedAt || new Date().toLocaleString('ms-MY')
          ];
        });
        classSheet.getRange(2, 1, rowsToAdd.length, 6).setValues(rowsToAdd);
      }
      return responseJSON({
        status: "success",
        message: "Senarai " + list.length + " kelas berjaya diselaraskan ke Google Sheet.",
        count: list.length
      });
    }

    if (data.action === "SAVE_CLASS" || data.action === "ADD_CLASS") {
      var classSheet = getOrCreateSheet(ss, "Kelas");
      if (classSheet.getLastRow() === 0) {
        classSheet.appendRow(["ID Kelas", "Nama Kelas", "Tahap", "Bilangan Murid / PC", "Catatan", "Tarikh Kemaskini"]);
        classSheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#e8f0fe");
      }

      var cls = data.classObj || data.class || {};
      var cId = String(cls.id || "").trim();
      var cName = String(cls.name || "").trim();
      var cLevel = cls.level || "Tahap 1";
      var cPcs = cls.pcs || 35;
      var cNotes = cls.notes || "";
      var cTime = new Date().toLocaleString('ms-MY');

      if (!cId && cName) {
        cId = "CLS-" + cName.replace(/\s+/g, '').toUpperCase();
      }

      var allVals = classSheet.getDataRange().getValues();
      var foundRow = -1;
      for (var r = 1; r < allVals.length; r++) {
        var rowId = String(allVals[r][0] || "").trim().toLowerCase();
        var rowName = String(allVals[r][1] || "").trim().toLowerCase();
        if ((cId && rowId === cId.toLowerCase()) || (cName && rowName === cName.toLowerCase())) {
          foundRow = r + 1;
          break;
        }
      }

      if (foundRow !== -1) {
        classSheet.getRange(foundRow, 1, 1, 6).setValues([[cId, cName, cLevel, cPcs, cNotes, cTime]]);
      } else {
        classSheet.appendRow([cId, cName, cLevel, cPcs, cNotes, cTime]);
      }

      return responseJSON({ status: "success", message: "Kelas " + cName + " berjaya disimpan ke Google Sheet." });
    }

    if (data.action === "DELETE_CLASS") {
      var classSheet = getOrCreateSheet(ss, "Kelas");
      var cId = String(data.id || "").trim().toLowerCase();
      var cName = String(data.name || "").trim().toLowerCase();

      var allVals = classSheet.getDataRange().getValues();
      var delRow = -1;
      for (var r = 1; r < allVals.length; r++) {
        var rowId = String(allVals[r][0] || "").trim().toLowerCase();
        var rowName = String(allVals[r][1] || "").trim().toLowerCase();
        if ((cId && rowId === cId) || (cName && rowName === cName)) {
          delRow = r + 1;
          break;
        }
      }

      if (delRow !== -1) {
        classSheet.deleteRow(delRow);
        return responseJSON({ status: "success", message: "Kelas berjaya dipadam dari Google Sheet." });
      }
      return responseJSON({ status: "warning", message: "Kelas tidak ditemui dalam Google Sheet." });
    }

    // 3. REKOD TEMPAHAN PUKAL (BATCH_ADD / SYNC_ALL) - KESELAMATAN 100% TIADA SLOT TERCICIR
    if (data.action === "BATCH_ADD" || data.action === "SYNC_ALL" || (Array.isArray(data.bookings) && data.bookings.length > 0)) {
      var bookings = data.bookings || [];
      var bookingSheet = getOrCreateSheet(ss, "Tempahan");

      if (bookingSheet.getLastRow() === 0) {
        bookingSheet.appendRow([
          "Kod Tempahan", "UserID", "Tarikh", "Slot Masa", "Pemohon", "Peranan",
          "Kelas / Subjek", "Bilangan PC", "Tujuan", "Peralatan", "Nota", "Status", "Tarikh Dicipta"
        ]);
        bookingSheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#e8f0fe");
      }

      var existingData = bookingSheet.getDataRange().getValues();
      var hasUserIdCol = (existingData[0][1] && String(existingData[0][1]).toLowerCase().includes("user"));
      var statusCol = hasUserIdCol ? 12 : 11; // 1-based index

      // Petakan kod tempahan sedia ada untuk semakan sepantas kilat
      var existingMap = {};
      for (var j = 1; j < existingData.length; j++) {
        var idKey = String(existingData[j][0] || "").trim().toLowerCase();
        if (idKey) {
          existingMap[idKey] = j + 1; // 1-based row index
        }
      }

      var newRowsToAppend = [];
      var updatedCount = 0;

      for (var k = 0; k < bookings.length; k++) {
        var b = bookings[k];
        if (!b) continue;
        var bId = String(b.id || "").trim();
        var bStatus = b.status || "Diluluskan";
        var dateFormatted = formatDateStr(b.date);

        if (bId && existingMap[bId.toLowerCase()]) {
          // Jika sudah ada dalam helaian: kemaskini status sahaja
          var rowIdx = existingMap[bId.toLowerCase()];
          bookingSheet.getRange(rowIdx, statusCol).setValue(bStatus);
          updatedCount++;
        } else {
          // Jika belum ada: kumpul untuk dimasukkan serentak (bulk write)
          var bUserId = b.userId || "";
          if (!bUserId && b.applicant) {
            var numMatch = String(b.applicant).match(/\d+/);
            if (numMatch) bUserId = "USR-" + numMatch[0];
          }

          newRowsToAppend.push([
            bId,
            bUserId,
            "'" + dateFormatted, // Awalan ' memastikan teks tarikh kekal YYYY-MM-DD tanpa ralat zon masa
            b.slot || "",
            b.applicant || "",
            b.role || "Guru",
            b.subject || "",
            b.pcCount || 35,
            b.purpose || "",
            JSON.stringify(b.equipments || []),
            b.notes || "",
            bStatus,
            b.createdAt || new Date().toISOString()
          ]);
        }
      }

      // Masukkan semua baris baru serentak dalam 1 operasi atomik
      if (newRowsToAppend.length > 0) {
        var startRow = bookingSheet.getLastRow() + 1;
        bookingSheet.getRange(startRow, 1, newRowsToAppend.length, 13).setValues(newRowsToAppend);
      }

      return responseJSON({
        status: "success",
        addedCount: newRowsToAppend.length,
        updatedCount: updatedCount,
        totalReceived: bookings.length,
        message: "Penyelarasan pukal berjaya: " + newRowsToAppend.length + " ditambah, " + updatedCount + " dikemaskini."
      });
    }

    // 3. KEMASKINI ATAU TAMBAH TEMPAHAN TUNGGAL (Tab: Tempahan)
    if (data.action === "ADD" || data.booking || data.action === "UPDATE_STATUS" || data.action === "CANCEL") {
      var booking = data.booking || {};
      var targetId = String(data.id || booking.id || "").trim();
      var newStatus = data.status || booking.status || "";
      var isCancellationOrStatusUpdate = (data.action === "UPDATE_STATUS" || data.action === "CANCEL" || newStatus === "Dibatalkan");

      var bookingSheet = getOrCreateSheet(ss, "Tempahan");

      if (bookingSheet.getLastRow() === 0) {
        bookingSheet.appendRow([
          "Kod Tempahan", "UserID", "Tarikh", "Slot Masa", "Pemohon", "Peranan",
          "Kelas / Subjek", "Bilangan PC", "Tujuan", "Peralatan", "Nota", "Status", "Tarikh Dicipta"
        ]);
        bookingSheet.getRange(1, 1, 1, 13).setFontWeight("bold").setBackground("#e8f0fe");
      }

      var rows = bookingSheet.getDataRange().getValues();
      var hasUserIdCol = (rows[0][1] && String(rows[0][1]).toLowerCase().includes("user"));
      var dateColIdx = hasUserIdCol ? 2 : 1;
      var slotColIdx = hasUserIdCol ? 3 : 2;
      var statusCol = hasUserIdCol ? 12 : 11; // Lajur Status (1-based index)

      var existingRowIndex = -1;

      // 1. Cari mengikut Kod Tempahan (ID)
      if (targetId) {
        var cleanTargetId = targetId.replace(/^'/, '').trim().toLowerCase();
        for (var j = 1; j < rows.length; j++) {
          var rowId = String(rows[j][0]).replace(/^'/, '').trim().toLowerCase();
          if (rowId === cleanTargetId) {
            existingRowIndex = j + 1; // 1-based index
            break;
          }
        }
      }

      // 2. Jika tidak dijumpai mengikut ID, padankan mengikut Tarikh & Slot Masa
      if (existingRowIndex === -1 && booking.date && booking.slot) {
        var targetDateNorm = formatDateStr(booking.date);
        var targetSlotNorm = String(booking.slot).replace(/\s+/g, ' ').trim().toLowerCase();
        for (var k = 1; k < rows.length; k++) {
          var rDateNorm = formatDateStr(rows[k][dateColIdx]);
          var rSlotNorm = String(rows[k][slotColIdx]).replace(/\s+/g, ' ').trim().toLowerCase();
          if (rDateNorm === targetDateNorm && rSlotNorm === targetSlotNorm) {
            existingRowIndex = k + 1;
            break;
          }
        }
      }

      // JIKA BARIS SEDIA ADA DIJUMPAI: Kemaskini status pada baris asal yang sedia ada
      if (existingRowIndex !== -1) {
        if (newStatus) {
          bookingSheet.getRange(existingRowIndex, statusCol).setValue(newStatus);
        }
        return responseJSON({
          status: "success",
          message: "Status tempahan " + targetId + " telah berjaya dikemaskini kepada '" + newStatus + "' pada baris sedia ada (baris " + existingRowIndex + ")."
        });
      }

      // PENTING: JIKA TINDAKAN IALAH PEMBATALAN ATAU KEMASKINI STATUS, JANGAN SEKALI-KALI CIPTA BARIS BARU!
      if (isCancellationOrStatusUpdate) {
        return responseJSON({
          status: "warning",
          message: "Rekod tempahan " + targetId + " tidak ditemui untuk dikemaskini. Tiada baris baru dicipta kerana tindakan ialah pembatalan."
        });
      }

      // JIKA BARIS BELUM ADA DAN BUKAN PEMBATALAN: Tambah baris baru
      var bookingUserId = booking.userId || "";
      if (!bookingUserId && booking.applicant) {
        var numMatch = String(booking.applicant).match(/\d+/);
        if (numMatch) bookingUserId = "USR-" + numMatch[0];
      }

      var dateFormatted = formatDateStr(booking.date);

      bookingSheet.appendRow([
        targetId || booking.id || "",
        bookingUserId,
        "'" + dateFormatted,
        booking.slot || "",
        booking.applicant || "",
        booking.role || "Guru",
        booking.subject || "",
        booking.pcCount || 35,
        booking.purpose || "",
        JSON.stringify(booking.equipments || []),
        booking.notes || "",
        booking.status || newStatus || "Menunggu Kelulusan",
        booking.createdAt || new Date().toISOString()
      ]);

      return responseJSON({ status: "success", message: "Tempahan baru ditambah." });
    }

    return responseJSON({ status: "error", message: "Tindakan tidak sah." });

  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (lErr) {}
    }
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function testRecordUser() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getOrCreateSheet(ss, "Pengguna");
  sheet.appendRow(["USR-83920192", "g-83920192@moe-dl.edu.my", "Cikgu Ahmad Razali", "Guru", "0123456789", "Sains", new Date().toLocaleString('ms-MY'), "DELIMa SSO"]);
  Logger.log("Ujian Rekod Pengguna Berjaya!");
}
