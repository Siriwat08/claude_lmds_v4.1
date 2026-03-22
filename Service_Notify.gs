/**
 * VERSION: 4.1
 * 🔔 Service: Omni-Channel Notification Hub
 * เป็นไฟล์เดียวที่นิยาม sendLineNotify และ sendTelegramNotify
 */

function sendSystemNotify(message, isUrgent) {
  try { sendLineNotify_Internal_(message, isUrgent); }
  catch (e) { console.error("[Notify Hub] LINE Failed: " + e.message); }
  try { sendTelegramNotify_Internal_(message, isUrgent); }
  catch (e) { console.error("[Notify Hub] Telegram Failed: " + e.message); }
}

function sendLineNotify(message, isUrgent) { sendLineNotify_Internal_(message, isUrgent); }
function sendTelegramNotify(message, isUrgent) { sendTelegramNotify_Internal_(message, isUrgent); }

function sendLineNotify_Internal_(message, isUrgent) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_NOTIFY_TOKEN');
  if (!token) return;
  var prefix  = isUrgent ? "🚨 URGENT ALERT:\n" : "🤖 SYSTEM REPORT:\n";
  var fullMsg = prefix + message;
  try {
    var response = UrlFetchApp.fetch("https://notify-api.line.me/api/notify", {
      "method": "post", "headers": { "Authorization": "Bearer " + token },
      "payload": { "message": fullMsg }, "muteHttpExceptions": true
    });
    if (response.getResponseCode() !== 200) console.warn("[LINE API Error] " + response.getContentText());
  } catch (e) { console.warn("[LINE Exception] " + e.message); }
}

function sendTelegramNotify_Internal_(message, isUrgent) {
  var token  = PropertiesService.getScriptProperties().getProperty('TG_BOT_TOKEN');
  var chatId = PropertiesService.getScriptProperties().getProperty('TG_CHAT_ID');
  if (!token) token  = PropertiesService.getScriptProperties().getProperty('TELEGRAM_BOT_TOKEN');
  if (!chatId) chatId = PropertiesService.getScriptProperties().getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  var icon    = isUrgent ? "🚨" : "🤖";
  var title   = isUrgent ? "<b>SYSTEM ALERT</b>" : "<b>SYSTEM REPORT</b>";
  var htmlMsg = `${icon} ${title}\n\n${escapeHtml_(message)}`;

  try {
    var url     = "https://api.telegram.org/bot" + token + "/sendMessage";
    var payload = { "chat_id": chatId, "text": htmlMsg, "parse_mode": "HTML" };
    var response = UrlFetchApp.fetch(url, { "method": "post", "contentType": "application/json", "payload": JSON.stringify(payload), "muteHttpExceptions": true });
    if (response.getResponseCode() !== 200) console.warn("[Telegram API Error] " + response.getContentText());
  } catch (e) { console.warn("[Telegram Exception] " + e.message); }
}

function escapeHtml_(text) {
  if (!text) return "";
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function notifyAutoPilotStatus(scgStatus, aiCount, aiMappedCount) {
  var mappedMsg = aiMappedCount !== undefined ? `\n🎯 AI Tier-4 จับคู่สำเร็จ: ${aiMappedCount} ร้าน` : "";
  var msg = "------------------\n✅ AutoPilot V4.1 รอบล่าสุด:\n📦 ดึงงาน SCG: " + scgStatus + "\n🧠 AI Indexing: " + aiCount + " รายการ" + mappedMsg;
  sendSystemNotify(msg, false);
}
