/**
 * VERSION: 4.1
 * 🏥 System Diagnostic Tool
 */

function RUN_SYSTEM_DIAGNOSTIC() {
  var ui   = SpreadsheetApp.getUi();
  var logs = [];

  function pass(msg) { logs.push("✅ " + msg); }
  function warn(msg) { logs.push("⚠️ " + msg); }
  function fail(msg) { logs.push("❌ " + msg); }

  try {
    if (typeof CONFIG !== 'undefined') pass("System Variables: มองเห็น CONFIG");
    else fail("System Variables: มองไม่เห็น CONFIG");

    if (typeof md5 === 'function') pass("Core Utils: มองเห็น md5()");
    else fail("Core Utils: มองไม่เห็น md5()");

    if (typeof normalizeText === 'function') pass("Core Utils: มองเห็น normalizeText()");
    else fail("Core Utils: มองไม่เห็น normalizeText()");

    if (typeof GET_ADDR_WITH_CACHE === 'function') {
      try {
        var testGeo = GET_ADDR_WITH_CACHE(13.746, 100.539);
        if (testGeo && testGeo !== "Error") pass("Google Maps API: ทำงานปกติ");
        else warn("Google Maps API: โหลดได้แต่ส่งค่าแปลกๆ");
      } catch (geoErr) { fail("Google Maps API: Error (" + geoErr.message + ")"); }
    } else {
      fail("Google Maps API: ไม่พบ GET_ADDR_WITH_CACHE");
    }

    try {
      if (CONFIG && CONFIG.GEMINI_API_KEY) pass("AI Engine: ตรวจพบ GEMINI_API_KEY พร้อมใช้งาน");
    } catch (e) { fail("AI Engine: ไม่พบ GEMINI_API_KEY (" + e.message + ")"); }

    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('LINE_NOTIFY_TOKEN')) pass("Notifications: ตรวจพบ LINE Notify Token");
    else warn("Notifications: ยังไม่ได้ตั้งค่า LINE Notify");

    if (props.getProperty('TG_BOT_TOKEN') && props.getProperty('TG_CHAT_ID')) pass("Notifications: ตรวจพบ Telegram Config");
    else warn("Notifications: ยังไม่ได้ตั้งค่า Telegram");

    ui.alert("🏥 รายงานผลการสแกนระบบ:\n\n" + logs.join("\n"));
  } catch (e) {
    ui.alert("🚨 ระบบตรวจพบ Error ร้ายแรง:\n" + e.message);
  }
}

function RUN_SHEET_DIAGNOSTIC() {
  var ui   = SpreadsheetApp.getUi();
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];

  function pass(msg) { logs.push("✅ " + msg); }
  function warn(msg) { logs.push("⚠️ " + msg); }
  function fail(msg) { logs.push("❌ " + msg); }

  try {
    var dbName   = CONFIG.SHEET_NAME;
    var dbSheet  = ss.getSheetByName(dbName);
    if (dbSheet) {
      var rows = getRealLastRow_(dbSheet, CONFIG.COL_NAME);
      if (rows >= 2) pass("Master DB: พบชีต '" + dbName + "' (มีข้อมูล " + rows + " แถว)");
      else warn("Master DB: พบชีต '" + dbName + "' แต่ข้อมูลว่างเปล่า");
    } else { fail("Master DB: ไม่พบชีต '" + dbName + "'"); }

    var srcSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    if (srcSheet) pass("Source Data: พบชีต '" + CONFIG.SOURCE_SHEET + "'");
    else warn("Source Data: ไม่พบชีต '" + CONFIG.SOURCE_SHEET + "'");

    var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
    if (mapSheet) {
      var mapCols = mapSheet.getLastColumn();
      if (mapCols >= 5) pass("Name Mapping: พบชีต และโครงสร้าง V4.0 ถูกต้อง");
      else warn("Name Mapping: พบชีตแต่มีแค่ " + mapCols + " คอลัมน์ (ควร Upgrade)");
    } else { fail("Name Mapping: ไม่พบชีต '" + CONFIG.MAPPING_SHEET + "'"); }

    if (ss.getSheetByName(SCG_CONFIG.SHEET_DATA))  pass("SCG Operation: พบชีต Data");
    else warn("SCG Operation: ไม่พบชีต Data");

    if (ss.getSheetByName(SCG_CONFIG.SHEET_INPUT)) pass("SCG Operation: พบชีต Input");
    else warn("SCG Operation: ไม่พบชีต Input");

    if (ss.getSheetByName(CONFIG.SHEET_POSTAL)) pass("Geo Database: พบชีต PostalRef");
    else warn("Geo Database: ไม่พบชีต PostalRef");

    if (ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE)) pass("GPS Queue: พบชีต GPS_Queue ✅");
    else warn("GPS Queue: ไม่พบชีต GPS_Queue (กรุณาสร้างด้วยเมนู)");

    ui.alert("🕵️‍♂️ รายงานผลการสแกนชีต:\n\n" + logs.join("\n"));
  } catch (e) {
    ui.alert("🚨 เกิด Error ระหว่างตรวจสอบชีต:\n" + e.message);
  }
}
