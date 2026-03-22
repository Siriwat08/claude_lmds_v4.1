/**
 * VERSION: 4.1
 * 🔐 Security Setup Utility
 */

function setupEnvironment() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('🔐 Security Setup: Gemini API', 'กรุณากรอก Gemini API Key (ต้องขึ้นต้นด้วย AIza...):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var key = response.getResponseText().trim();
    if (key.length > 30 && key.startsWith("AIza")) {
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('✅ บันทึก API Key สำเร็จ!');
    } else {
      ui.alert('❌ API Key ไม่ถูกต้อง', 'Key ต้องขึ้นต้นด้วย "AIza"', ui.ButtonSet.OK);
    }
  }
}

function setupLineToken() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('🔔 Setup: LINE Notify', 'กรุณากรอก LINE Notify Token:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.OK) {
    var token = response.getResponseText().trim();
    if (token.length > 20) {
      PropertiesService.getScriptProperties().setProperty('LINE_NOTIFY_TOKEN', token);
      ui.alert('✅ บันทึก LINE Token สำเร็จ!');
    } else {
      ui.alert('❌ Token สั้นเกินไป');
    }
  }
}

function setupTelegramConfig() {
  var ui    = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var resBot = ui.prompt('✈️ Setup: Telegram', '1. กรุณากรอก Bot Token:', ui.ButtonSet.OK_CANCEL);
  if (resBot.getSelectedButton() !== ui.Button.OK) return;
  var botToken = resBot.getResponseText().trim();

  var resChat = ui.prompt('✈️ Setup: Telegram', '2. กรุณากรอก Chat ID:', ui.ButtonSet.OK_CANCEL);
  if (resChat.getSelectedButton() !== ui.Button.OK) return;
  var chatId = resChat.getResponseText().trim();

  if (botToken && chatId) {
    props.setProperty('TG_BOT_TOKEN', botToken);
    props.setProperty('TG_CHAT_ID', chatId);
    ui.alert('✅ บันทึก Telegram Config สำเร็จ!');
  } else {
    ui.alert('❌ ข้อมูลไม่ครบถ้วน');
  }
}

function resetEnvironment() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('⚠️ Danger Zone', 'คุณต้องการล้างรหัส API Key ของ Gemini ใช่หรือไม่?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    ui.alert('🗑️ ล้างการตั้งค่า Gemini API Key เรียบร้อยแล้ว');
  }
}
