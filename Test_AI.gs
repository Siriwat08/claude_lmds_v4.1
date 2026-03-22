/**
 * VERSION: 4.1
 * 🧪 Test & Debug: AI Capabilities
 */

function forceRunAI_Now() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (typeof processAIIndexing_Batch !== 'function') throw new Error("ไม่พบฟังก์ชัน 'processAIIndexing_Batch'");
    ss.toast("🚀 กำลังเริ่มระบบ AI Indexing...", "Debug System", 10);
    processAIIndexing_Batch();
    ui.alert("✅ สั่งงานเรียบร้อย!\nตรวจสอบคอลัมน์ Normalized ว่ามี Tag '[AI]' หรือไม่");
  } catch (e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function debug_TestTier4SmartResolution() {
  var ui = SpreadsheetApp.getUi();
  try {
    if (typeof resolveUnknownNamesWithAI !== 'function') throw new Error("ไม่พบฟังก์ชัน 'resolveUnknownNamesWithAI'");
    var response = ui.alert("🧠 ยืนยันรันทดสอบ Tier 4", "ต้องการดึงรายชื่อที่ไม่มีพิกัดจากหน้า SCG Data\nไปให้ Gemini วิเคราะห์หรือไม่?", ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) resolveUnknownNamesWithAI();
  } catch (e) {
    ui.alert("❌ Error: " + e.message);
  }
}

function debugGeminiConnection() {
  var ui = SpreadsheetApp.getUi();
  var apiKey;
  try { apiKey = CONFIG.GEMINI_API_KEY; }
  catch (e) { ui.alert("❌ API Key Error", e.message, ui.ButtonSet.OK); return; }

  var testWord = "SCG (Bang Sue Branch)";
  try {
    var model   = (typeof CONFIG !== 'undefined' && CONFIG.AI_MODEL) ? CONFIG.AI_MODEL : "gemini-1.5-flash";
    var url     = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    var payload = { "contents": [{ "parts": [{ "text": `Hello Gemini, test connection. Say "Connection Success" and reply with Thai translation of ${testWord}` }] }] };
    var options = { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true };
    var res     = UrlFetchApp.fetch(url, options);

    if (res.getResponseCode() === 200) {
      var json = JSON.parse(res.getContentText());
      var text = (json.candidates && json.candidates[0].content) ? json.candidates[0].content.parts[0].text : "No Text";
      ui.alert("✅ API Ping Success!\n\nResponse:\n" + text);
    } else {
      ui.alert("❌ API Error: " + res.getContentText());
    }
  } catch (e) {
    ui.alert("❌ Connection Failed: " + e.message);
  }
}

function debug_ResetSelectedRowsAI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getActiveSheet();

  if (sheet.getName() !== CONFIG.SHEET_NAME) {
    ui.alert("⚠️ กรุณาไฮไลต์เลือก Cell ในชีต Database เท่านั้นครับ");
    return;
  }

  var range      = sheet.getActiveRange();
  var startRow   = range.getRow();
  var numRows    = range.getNumRows();
  var colIndex   = (typeof CONFIG !== 'undefined' && CONFIG.COL_NORMALIZED) ? CONFIG.COL_NORMALIZED : 6;
  var targetRange = sheet.getRange(startRow, colIndex, numRows, 1);
  var values     = targetRange.getValues();
  var resetCount = 0;

  for (var i = 0; i < values.length; i++) {
    var val = values[i][0] ? values[i][0].toString() : "";
    if (val.indexOf("[AI]") !== -1 || val.indexOf("[Agent_") !== -1) {
      var cleanedVal = val.replace(" [AI]", "").replace("[AI]", "").replace(/\[Agent_.*?\]/g, "").trim();
      values[i][0] = cleanedVal;
      resetCount++;
    }
  }

  if (resetCount > 0) {
    targetRange.setValues(values);
    ss.toast("🔄 Reset AI Status เรียบร้อย " + resetCount + " แถว", "Debug", 5);
  } else {
    ss.toast("ℹ️ ไม่พบรายการที่มี Tag AI ในส่วนที่เลือก", "Debug", 5);
  }
}
