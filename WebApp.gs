/**
 * VERSION: 4.1
 * 🌐 WebApp Controller
 */

function doGet(e) {
  try {
    console.info(`[WebApp] GET Request. Params: ${JSON.stringify(e.parameter)}`);
    var page     = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Index';
    var template = HtmlService.createTemplateFromFile(page);
    template.initialQuery = (e && e.parameter && e.parameter.q) ? e.parameter.q : "";
    template.appVersion   = new Date().getTime();
    template.isEnterprise = true;

    var output = template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
      .setTitle('🔍 Logistics Master Search (V4.1)')
      .setFaviconUrl('https://img.icons8.com/color/48/truck--v1.png');

    output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return output;
  } catch (err) {
    console.error(`[WebApp] GET Error: ${err.message}`);
    return HtmlService.createHtmlOutput(`<div style="font-family:sans-serif;padding:20px;text-align:center;background:#ffebee;"><h3 style="color:#d32f2f;">❌ System Error (V4.1)</h3><p>${err.message}</p></div>`);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData) throw new Error("No payload found in POST request.");
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;
    if (action === "triggerAIBatch") {
      if (typeof processAIIndexing_Batch === 'function') {
        processAIIndexing_Batch();
        return createJsonResponse_({ status: "success", message: "AI Batch Processing Triggered" });
      }
    }
    return createJsonResponse_({ status: "success", message: "Webhook received", data: payload });
  } catch (err) {
    return createJsonResponse_({ status: "error", message: err.message });
  }
}

function createJsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  try { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }
  catch (e) { return "<!-- Error: File '" + filename + "' not found. -->"; }
}

function getUserContext() {
  try {
    return { email: Session.getActiveUser().getEmail() || "anonymous", locale: Session.getActiveUserLocale() || "th" };
  } catch (e) {
    return { email: "unknown", locale: "th" };
  }
}
