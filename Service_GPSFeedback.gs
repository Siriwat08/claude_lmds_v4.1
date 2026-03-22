/**
 * VERSION: 4.1
 * 📍 Service: GPS Feedback & Queue Management (NEW v4.1)
 * หน้าที่: จัดการ GPS_Queue รับพิกัดจากคนขับ อนุมัติเข้า Database
 */

// ==========================================
// 1. CREATE GPS_QUEUE SHEET
// ==========================================

function createGPSQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  if (ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE)) {
    ui.alert("ℹ️ ชีต GPS_Queue มีอยู่แล้วครับ");
    return;
  }

  var sheet = ss.insertSheet(SCG_CONFIG.SHEET_GPS_QUEUE);

  var headers = ["Timestamp","ShipToName","UUID_DB","LatLng_Driver","LatLng_DB","Diff_Meters","Reason","Approve","Reject"];

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4f46e5");
  headerRange.setFontColor("white");

  sheet.getRange(2, 8, 500, 1).insertCheckboxes();
  sheet.getRange(2, 9, 500, 1).insertCheckboxes();

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 160);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 80);
  sheet.setColumnWidth(9, 80);

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();

  ui.alert("✅ สร้างชีต GPS_Queue สำเร็จแล้วครับ");
}

// ==========================================
// 2. APPLY APPROVED FEEDBACK
// ==========================================

function applyApprovedFeedback() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();

  var queueSheet  = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);
  var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!queueSheet || !masterSheet) {
    ui.alert("❌ ไม่พบชีต GPS_Queue หรือ Database");
    return;
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert("⚠️ ระบบคิวทำงาน", "มีผู้ใช้งานอื่นกำลังใช้งานอยู่ กรุณารอสักครู่", ui.ButtonSet.OK);
    return;
  }

  // [v4.1] ตรวจสอบ Schema ก่อนทำงาน
  try { preCheck_Approve(); } catch(e) {
    ui.alert("❌ Schema Error", e.message, ui.ButtonSet.OK);
    return;
  }

  try {
    var lastQueueRow = getRealLastRow_(queueSheet, 1);
    if (lastQueueRow < 2) {
      ui.alert("ℹ️ ไม่มีรายการใน GPS_Queue");
      return;
    }

    var queueData = queueSheet.getRange(2, 1, lastQueueRow - 1, 9).getValues();

    var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
    var dbData   = masterSheet.getRange(2, 1, lastRowM - 1, 22).getValues();

    var uuidMap = {};
    dbData.forEach(function(r, i) {
      if (r[CONFIG.C_IDX.UUID]) uuidMap[r[CONFIG.C_IDX.UUID]] = i;
    });

    var approvedCount = 0;
    var skippedCount  = 0;
    var ts = new Date();

    queueData.forEach(function(row, i) {
      var isApproved = row[7];
      var isRejected = row[8];
      var reason     = row[6];

      if (!isApproved || isRejected) { skippedCount++; return; }
      if (reason === "APPROVED" || reason === "REJECTED") return;

      var uuid        = row[2];
      var latLngDriver = row[3];

      if (!uuid || !latLngDriver) { skippedCount++; return; }

      var parts = latLngDriver.toString().split(",");
      if (parts.length !== 2) { skippedCount++; return; }

      var newLat = parseFloat(parts[0].trim());
      var newLng = parseFloat(parts[1].trim());
      if (isNaN(newLat) || isNaN(newLng)) { skippedCount++; return; }

      if (!uuidMap.hasOwnProperty(uuid)) { skippedCount++; return; }

      var dbRowNum = uuidMap[uuid] + 2;

      masterSheet.getRange(dbRowNum, CONFIG.COL_LAT).setValue(newLat);
      masterSheet.getRange(dbRowNum, CONFIG.COL_LNG).setValue(newLng);
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_SOURCE).setValue("Driver_GPS");
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_CONFIDENCE).setValue(95);
      masterSheet.getRange(dbRowNum, CONFIG.COL_COORD_LAST_UPDATED).setValue(ts);
      masterSheet.getRange(dbRowNum, CONFIG.COL_UPDATED).setValue(ts);

      queueSheet.getRange(i + 2, 7).setValue("APPROVED");
      approvedCount++;
    });

    if (typeof clearSearchCache === 'function') clearSearchCache();
    SpreadsheetApp.flush();

    var msg = "✅ อนุมัติเรียบร้อย!\n\n" +
      "📍 อัปเดตพิกัดใน Database: " + approvedCount + " ราย\n" +
      "⏭️ ข้ามไป: " + skippedCount + " ราย";

    if (approvedCount === 0) msg += "\n\nไม่มีรายการที่ติ๊ก Approve\nกรุณาติ๊ก Col H ก่อนรันครับ";
    ui.alert(msg);

  } catch (e) {
    console.error("applyApprovedFeedback Error: " + e.message);
    ui.alert("❌ เกิดข้อผิดพลาด: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 3. SHOW GPS QUEUE STATS
// ==========================================

function showGPSQueueStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var queueSheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  if (!queueSheet) {
    ui.alert("❌ ไม่พบชีต GPS_Queue\nกรุณารัน 'สร้างชีต GPS_Queue ใหม่' ก่อนครับ");
    return;
  }

  var lastRow = getRealLastRow_(queueSheet, 1);
  if (lastRow < 2) {
    ui.alert("ℹ️ GPS_Queue ว่างเปล่า ยังไม่มีรายการครับ");
    return;
  }

  var data = queueSheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var stats = { total: 0, pending: 0, approved: 0, rejected: 0, gpsDiff: 0, noGPS: 0 };

  data.forEach(function(row) {
    var reason  = row[6];
    var approve = row[7];
    var reject  = row[8];
    if (!row[0]) return;
    stats.total++;
    if (reason === "APPROVED")      stats.approved++;
    else if (reason === "REJECTED") stats.rejected++;
    else if (approve)               stats.approved++;
    else if (reject)                stats.rejected++;
    else                            stats.pending++;
    if (reason === "GPS_DIFF")      stats.gpsDiff++;
    else if (reason === "DB_NO_GPS") stats.noGPS++;
  });

  ui.alert(
    "📊 GPS Queue สถิติ\n━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📋 รายการทั้งหมด: " + stats.total    + " ราย\n" +
    "⏳ รอตรวจสอบ: "    + stats.pending  + " ราย\n" +
    "✅ อนุมัติแล้ว: "  + stats.approved + " ราย\n" +
    "❌ ปฏิเสธแล้ว: "   + stats.rejected + " ราย\n" +
    "━━━━━━━━━━━━━━━━━━━━━━━\n" +
    "📍 GPS ต่างกัน >50m: " + stats.gpsDiff + " ราย\n" +
    "🔍 DB ไม่มีพิกัด: "   + stats.noGPS   + " ราย"
  );
}

// ==========================================
// 4. UPGRADE GPS_QUEUE SHEET
// ==========================================

function upgradeGPSQueueSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

  if (!sheet) { ui.alert("❌ ไม่พบชีต GPS_Queue ครับ"); return; }

  var headers = ["Timestamp","ShipToName","UUID_DB","LatLng_Driver","LatLng_DB","Diff_Meters","Reason","Approve","Reject"];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#4f46e5");
  headerRange.setFontColor("white");

  var realLastRow = getRealLastRow_(sheet, 1);
  var maxRow = sheet.getMaxRows();
  if (maxRow > 1) sheet.getRange(2, 8, maxRow - 1, 2).clearContent();

  var checkboxEnd = Math.max(realLastRow, 1) + 1000;
  sheet.getRange(2, 8, checkboxEnd, 1).insertCheckboxes();
  sheet.getRange(2, 9, checkboxEnd, 1).insertCheckboxes();

  if (realLastRow > 1) {
    sheet.getRange(2, 6, realLastRow - 1, 1).setNumberFormat("#,##0");
    var reasonData = sheet.getRange(2, 7, realLastRow - 1, 1).getValues();
    reasonData.forEach(function(row, i) {
      var rowNum = i + 2;
      var bg = "#ffffff";
      if (row[0] === "GPS_DIFF")  bg = "#fff3cd";
      if (row[0] === "DB_NO_GPS") bg = "#f8d7da";
      if (row[0] === "NO_MATCH")  bg = "#d1ecf1";
      if (row[0] === "APPROVED")  bg = "#d4edda";
      if (row[0] === "REJECTED")  bg = "#e2e3e5";
      sheet.getRange(rowNum, 1, 1, 9).setBackground(bg);
    });
  }

  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
  ui.alert("✅ อัปเกรด GPS_Queue สำเร็จ!\n\nแถวข้อมูลจริง: " + (realLastRow - 1) + " รายการ");
}

// ==========================================
// 5. RESET SYNC STATUS (สำหรับทดสอบ)
// ==========================================

function resetSyncStatus() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);

  var lastRow = sheet.getLastRow();
  var syncCol = SCG_CONFIG.SRC_IDX_SYNC_STATUS;

  var result = ui.alert(
    "⚠️ Reset SYNC_STATUS?",
    "จะล้าง SYNCED ทั้งหมด " + (lastRow - 1) + " แถว เพื่อให้ระบบประมวลผลใหม่\n\nใช้สำหรับทดสอบเท่านั้น",
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) return;

  sheet.getRange(2, syncCol, lastRow - 1, 1).clearContent();
  SpreadsheetApp.flush();
  ui.alert("✅ Reset เรียบร้อยแล้วครับ");
}
