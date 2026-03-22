/**
 * VERSION: 4.1
 * 🧹 Service: System Maintenance & Alerts
 * [v4.1] ลบ sendLineNotify() และ sendTelegramNotify() ออก
 *        (ใช้จาก Service_Notify.gs เท่านั้น)
 */

function cleanupOldBackups() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { console.warn("[Maintenance] ข้ามการทำงาน"); return; }
  try {
    var ss          = SpreadsheetApp.getActiveSpreadsheet();
    var sheets      = ss.getSheets();
    var deletedCount = 0;
    var keepDays    = 30;
    var now         = new Date();
    var deletedNames = [];

    sheets.forEach(function(sheet) {
      var name = sheet.getName();
      if (name.startsWith("Backup_")) {
        var datePart = name.match(/(\d{4})(\d{2})(\d{2})/);
        if (datePart && datePart.length === 4) {
          var sheetDate = new Date(parseInt(datePart[1]), parseInt(datePart[2]) - 1, parseInt(datePart[3]));
          var diffDays  = Math.ceil(Math.abs(now - sheetDate) / (1000 * 60 * 60 * 24));
          if (diffDays > keepDays) {
            try { ss.deleteSheet(sheet); deletedCount++; deletedNames.push(name); }
            catch(e) { console.error("[Maintenance] Could not delete " + name); }
          }
        }
      }
    });

    if (deletedCount > 0) {
      var msg = `🧹 Maintenance Report:\nระบบได้ลบชีต Backup เก่ากว่า ${keepDays} วัน จำนวน ${deletedCount} ชีต`;
      sendLineNotify(msg);
      sendTelegramNotify(msg);
      SpreadsheetApp.getActiveSpreadsheet().toast(`ลบ Backup เก่าไป ${deletedCount} ชีต`, "Maintenance");
    }
  } catch (err) {
    console.error("[Maintenance] Error: " + err.message);
  } finally {
    lock.releaseLock();
  }
}

function checkSpreadsheetHealth() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var cellLimit = 10000000;
  var totalCells = 0;
  var sheetCount = 0;

  ss.getSheets().forEach(function(s) {
    totalCells += (s.getMaxRows() * s.getMaxColumns());
    sheetCount++;
  });

  var usagePercent = (totalCells / cellLimit) * 100;

  if (usagePercent > 80) {
    var warn = `⚠️ CRITICAL WARNING: ไฟล์ใกล้เต็มแล้ว!\nการใช้งาน ${usagePercent.toFixed(2)}% (${totalCells.toLocaleString()} Cells)`;
    sendLineNotify(warn, true);
    sendTelegramNotify(warn);
    SpreadsheetApp.getUi().alert("⚠️ SYSTEM ALERT", warn, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast(`System Health OK (${usagePercent.toFixed(1)}%)`, "Health Check", 5);
  }
}
