/**
 * VERSION: 4.1
 * 📦 Service: SCG Operation (Enterprise Edition)
 * [v4.1] ลบ normalizeText() ซ้ำออก (ใช้จาก Utils_Common.gs แทน)
 * [v4.1] แก้ checklsEPOD → checkIsEPOD
 * [v4.1] แก้ Negative Row Count ใน applyMasterCoordinatesToDailyJob
 */

// ==========================================
// 1. MAIN OPERATION: FETCH DATA
// ==========================================

function fetchDataFromSCGJWD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    ui.alert("⚠️ ระบบคิวทำงาน", "มีผู้ใช้งานอื่นกำลังโหลดข้อมูล Shipment อยู่ กรุณารอสักครู่", ui.ButtonSet.OK);
    return;
  }

  try {
    const inputSheet = ss.getSheetByName(SCG_CONFIG.SHEET_INPUT);
    const dataSheet  = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
    if (!inputSheet || !dataSheet) throw new Error("CRITICAL: ไม่พบชีต Input หรือ Data");

    const cookie = inputSheet.getRange(SCG_CONFIG.COOKIE_CELL).getValue();
    if (!cookie) throw new Error("❌ กรุณาวาง Cookie ในช่อง " + SCG_CONFIG.COOKIE_CELL);

    const lastRow = inputSheet.getLastRow();
    if (lastRow < SCG_CONFIG.INPUT_START_ROW) throw new Error("ℹ️ ไม่พบเลข Shipment ในชีต Input");

    const shipmentNumbers = inputSheet
      .getRange(SCG_CONFIG.INPUT_START_ROW, 1, lastRow - SCG_CONFIG.INPUT_START_ROW + 1, 1)
      .getValues().flat().filter(String);

    if (shipmentNumbers.length === 0) throw new Error("ℹ️ รายการ Shipment ว่างเปล่า");

    const shipmentString = shipmentNumbers.join(',');
    inputSheet.getRange(SCG_CONFIG.SHIPMENT_STRING_CELL).setValue(shipmentString).setHorizontalAlignment("left");

    const payload = {
      DeliveryDateFrom: '', DeliveryDateTo: '', TenderDateFrom: '', TenderDateTo: '',
      CarrierCode: '', CustomerCode: '', OriginCodes: '', ShipmentNos: shipmentString
    };

    const options = { method: 'post', payload: payload, muteHttpExceptions: true, headers: { cookie: cookie } };

    ss.toast("กำลังเชื่อมต่อ SCG Server...", "System", 10);
    const responseText = fetchWithRetry_(SCG_CONFIG.API_URL, options, (CONFIG.API_MAX_RETRIES || 3));
    const json = JSON.parse(responseText);
    const shipments = json.data || [];

    if (shipments.length === 0) throw new Error("API Return Success แต่ไม่พบข้อมูล Shipment");

    ss.toast("กำลังแปลงข้อมูล " + shipments.length + " Shipments...", "Processing", 5);
    const allFlatData = [];
    let runningRow = 2;

    shipments.forEach(shipment => {
      const destSet = new Set();
      (shipment.DeliveryNotes || []).forEach(n => { if (n.ShipToName) destSet.add(n.ShipToName); });
      const destListStr = Array.from(destSet).join(", ");

      (shipment.DeliveryNotes || []).forEach(note => {
        (note.Items || []).forEach(item => {
          const dailyJobId = note.PurchaseOrder + "-" + runningRow;
          const row = [
            dailyJobId, note.PlanDelivery ? new Date(note.PlanDelivery) : null,
            String(note.PurchaseOrder), String(shipment.ShipmentNo),
            shipment.DriverName, shipment.TruckLicense, String(shipment.CarrierCode),
            shipment.CarrierName, String(note.SoldToCode), note.SoldToName,
            note.ShipToName, note.ShipToAddress,
            note.ShipToLatitude + ", " + note.ShipToLongitude,
            item.MaterialName, item.ItemQuantity, item.QuantityUnit, item.ItemWeight,
            String(note.DeliveryNo), destSet.size, destListStr,
            "รอสแกน", "ยังไม่ได้ส่ง", "", 0, 0, 0, "",
            "", shipment.ShipmentNo + "|" + note.ShipToName
          ];
          allFlatData.push(row);
          runningRow++;
        });
      });
    });

    const shopAgg = {};
    allFlatData.forEach(r => {
      const key = r[28];
      if (!shopAgg[key]) shopAgg[key] = { qty: 0, weight: 0, invoices: new Set(), epod: 0 };
      shopAgg[key].qty += Number(r[14]) || 0;
      shopAgg[key].weight += Number(r[16]) || 0;
      shopAgg[key].invoices.add(r[2]);
      // [v4.1] แก้ checklsEPOD → checkIsEPOD
      if (checkIsEPOD(r[9], r[2])) shopAgg[key].epod++;
    });

    allFlatData.forEach(r => {
      const agg = shopAgg[r[28]];
      const scanInv = agg.invoices.size - agg.epod;
      r[23] = agg.qty;
      r[24] = Number(agg.weight.toFixed(2));
      r[25] = scanInv;
      r[27] = `${r[9]} / รวม ${scanInv} บิล`;
    });

    const headers = [
      "ID_งานประจำวัน","PlanDelivery","InvoiceNo","ShipmentNo","DriverName",
      "TruckLicense","CarrierCode","CarrierName","SoldToCode","SoldToName",
      "ShipToName","ShipToAddress","LatLong_SCG","MaterialName","ItemQuantity",
      "QuantityUnit","ItemWeight","DeliveryNo","จำนวนปลายทาง_System","รายชื่อปลายทาง_System",
      "ScanStatus","DeliveryStatus","Email พนักงาน",
      "จำนวนสินค้ารวมของร้านนี้","น้ำหนักสินค้ารวมของร้านนี้","จำนวน_Invoice_ที่ต้องสแกน",
      "LatLong_Actual","ชื่อเจ้าของสินค้า_Invoice_ที่ต้องสแกน","ShopKey"
    ];

    dataSheet.clear();
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

    if (allFlatData.length > 0) {
      dataSheet.getRange(2, 1, allFlatData.length, headers.length).setValues(allFlatData);
      dataSheet.getRange(2, 2, allFlatData.length, 1).setNumberFormat("dd/mm/yyyy");
      dataSheet.getRange(2, 3, allFlatData.length, 1).setNumberFormat("@");
      dataSheet.getRange(2, 18, allFlatData.length, 1).setNumberFormat("@");
    }

    applyMasterCoordinatesToDailyJob();
    buildOwnerSummary();
    buildShipmentSummary();

    ui.alert(`✅ ดึงข้อมูลสำเร็จ!\n- จำนวนรายการ: ${allFlatData.length} แถว\n- จับคู่พิกัด: เรียบร้อย`);

  } catch (e) {
    console.error("[SCG API Error]: " + e.message);
    ui.alert("❌ เกิดข้อผิดพลาด: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. COORDINATE MATCHING
// ==========================================

function applyMasterCoordinatesToDailyJob() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet  = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  const dbSheet    = ss.getSheetByName(SCG_CONFIG.SHEET_MASTER_DB);
  const mapSheet   = ss.getSheetByName(SCG_CONFIG.SHEET_MAPPING);
  const empSheet   = ss.getSheetByName(SCG_CONFIG.SHEET_EMPLOYEE);

  if (!dataSheet || !dbSheet) return;
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) return;

  const masterCoords    = {};
  const masterUUIDCoords = {};

  if (dbSheet.getLastRow() > 1) {
    const maxCol = Math.max(CONFIG.COL_NAME, CONFIG.COL_LAT, CONFIG.COL_LNG, CONFIG.COL_UUID);
    const dbData = dbSheet.getRange(2, 1, dbSheet.getLastRow() - 1, maxCol).getValues();
    dbData.forEach(r => {
      const name = r[CONFIG.C_IDX.NAME];
      const lat  = r[CONFIG.C_IDX.LAT];
      const lng  = r[CONFIG.C_IDX.LNG];
      const uuid = r[CONFIG.C_IDX.UUID];
      if (name && lat && lng) {
        const coords = lat + ", " + lng;
        masterCoords[normalizeText(name)] = coords;
        if (uuid) masterUUIDCoords[uuid] = coords;
      }
    });
  }

  const aliasMap = {};
  if (mapSheet && mapSheet.getLastRow() > 1) {
    mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0] && r[1]) aliasMap[normalizeText(r[0])] = r[1];
    });
  }

  // [v4.1] แก้ Negative Row Count
  const empMap = {};
  if (empSheet) {
    var empLastRow = empSheet.getLastRow();
    if (empLastRow >= 2) {
      empSheet.getRange(2, 1, empLastRow - 1, 8).getValues().forEach(r => {
        if (r[1] && r[6]) empMap[normalizeText(r[1])] = r[6];
      });
    }
  }

  const values       = dataSheet.getRange(2, 1, lastRow - 1, 29).getValues();
  const latLongUpdates = [];
  const bgUpdates      = [];
  const emailUpdates   = [];

  values.forEach(r => {
    let newGeo = "";
    let bg     = null;
    let email  = r[22];

    if (r[10]) {
      let rawName   = normalizeText(r[10]);
      let targetUID = aliasMap[rawName];
      if (targetUID && masterUUIDCoords[targetUID]) {
        newGeo = masterUUIDCoords[targetUID]; bg = "#b6d7a8";
      } else if (masterCoords[rawName]) {
        newGeo = masterCoords[rawName]; bg = "#b6d7a8";
      } else {
        let branchMatch = tryMatchBranch_(rawName, masterCoords);
        if (branchMatch) { newGeo = branchMatch; bg = "#ffe599"; }
      }
    }

    latLongUpdates.push([newGeo]);
    bgUpdates.push([bg]);

    if (r[4]) {
      const cleanDriver = normalizeText(r[4]);
      if (empMap[cleanDriver]) email = empMap[cleanDriver];
    }
    emailUpdates.push([email]);
  });

  dataSheet.getRange(2, 27, latLongUpdates.length, 1).setValues(latLongUpdates);
  dataSheet.getRange(2, 27, bgUpdates.length, 1).setBackgrounds(bgUpdates);
  dataSheet.getRange(2, 23, emailUpdates.length, 1).setValues(emailUpdates);

  ss.toast("✅ อัปเดตพิกัดและข้อมูลพนักงานเรียบร้อย", "System");
}

// ==========================================
// 3. UTILITIES & HELPERS
// ==========================================

function fetchWithRetry_(url, options, maxRetries) {
  for (let i = 0; i < maxRetries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200) return response.getContentText();
      throw new Error("HTTP " + response.getResponseCode());
    } catch (e) {
      if (i === maxRetries - 1) throw e;
      Utilities.sleep(1000 * Math.pow(2, i));
    }
  }
}

function tryMatchBranch_(name, masterCoords) {
  const keywords = ["สาขา", "branch", "สำนักงาน", "store", "shop"];
  for (let k of keywords) {
    if (name.includes(k)) {
      let parts = name.split(k);
      if (parts.length > 0 && parts[0].length > 2) {
        let parentName = normalizeText(parts[0]);
        if (masterCoords[parentName]) return masterCoords[parentName];
      }
    }
  }
  return null;
}

/**
 * [v4.1] แก้ไข: ชื่อฟังก์ชันถูกต้องแล้ว (checkIsEPOD ไม่ใช่ checklsEPOD)
 */
function checkIsEPOD(ownerName, invoiceNo) {
  if (!ownerName || !invoiceNo) return false;
  const owner = String(ownerName).toUpperCase();
  const inv   = String(invoiceNo);

  const epodOwners = ["BETTERBE", "SCG EXPRESS", "เบทเตอร์แลนด์", "JWD TRANSPORT"];
  if (epodOwners.some(w => owner.includes(w.toUpperCase()))) return true;

  if (owner.includes("DENSO") || owner.includes("เด็นโซ่")) {
    if (inv.includes("_DOC")) return false;
    if (/^\d+(-.*)?$/.test(inv)) return true;
    return false;
  }

  return false;
}

// ==========================================
// 4. BUILD SUMMARY
// ==========================================

function buildOwnerSummary() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) return;

  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 29).getValues();
  const ownerMap = {};

  data.forEach(r => {
    const owner   = r[9];
    const invoice = String(r[2]);
    if (!owner) return;
    if (!ownerMap[owner]) ownerMap[owner] = { all: new Set(), epod: new Set() };
    if (!invoice) return;
    if (checkIsEPOD(owner, invoice)) { ownerMap[owner].epod.add(invoice); return; }
    ownerMap[owner].all.add(invoice);
  });

  const summarySheet = ss.getSheetByName("สรุป_เจ้าของสินค้า");
  if (!summarySheet) { SpreadsheetApp.getUi().alert("❌ ไม่พบชีต สรุป_เจ้าของสินค้า"); return; }

  const summaryLastRow = summarySheet.getLastRow();
  if (summaryLastRow > 1) summarySheet.getRange(2, 1, summaryLastRow - 1, 6).clearContent().setBackground(null);

  const rows = [];
  Object.keys(ownerMap).sort().forEach(owner => {
    const o = ownerMap[owner];
    rows.push(["", owner, "", o.all.size, o.epod.size, new Date()]);
  });

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 6).setValues(rows);
    summarySheet.getRange(2, 4, rows.length, 2).setNumberFormat("#,##0");
    summarySheet.getRange(2, 6, rows.length, 1).setNumberFormat("dd/mm/yyyy HH:mm");
  }
}

function buildShipmentSummary() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  if (!dataSheet || dataSheet.getLastRow() < 2) return;

  const data = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 29).getValues();
  const shipmentMap = {};

  data.forEach(r => {
    const shipmentNo = String(r[3]);
    const truck      = String(r[5]);
    const owner      = r[9];
    const invoice    = String(r[2]);
    if (!shipmentNo || !truck) return;

    const key = shipmentNo + "_" + truck;
    if (!shipmentMap[key]) shipmentMap[key] = { shipmentNo: shipmentNo, truck: truck, all: new Set(), epod: new Set() };
    if (!invoice) return;
    if (checkIsEPOD(owner, invoice)) { shipmentMap[key].epod.add(invoice); return; }
    shipmentMap[key].all.add(invoice);
  });

  const summarySheet = ss.getSheetByName("สรุป_Shipment");
  if (!summarySheet) { SpreadsheetApp.getUi().alert("❌ ไม่พบชีต สรุป_Shipment"); return; }

  const summaryLastRow = summarySheet.getLastRow();
  if (summaryLastRow > 1) summarySheet.getRange(2, 1, summaryLastRow - 1, 7).clearContent().setBackground(null);

  const rows = [];
  Object.keys(shipmentMap).sort().forEach(key => {
    const s = shipmentMap[key];
    rows.push([key, s.shipmentNo, s.truck, "", s.all.size, s.epod.size, new Date()]);
  });

  if (rows.length > 0) {
    summarySheet.getRange(2, 1, rows.length, 7).setValues(rows);
    summarySheet.getRange(2, 5, rows.length, 2).setNumberFormat("#,##0");
    summarySheet.getRange(2, 7, rows.length, 1).setNumberFormat("dd/mm/yyyy HH:mm");
  }
}

// ==========================================
// 5. CLEAR FUNCTIONS
// ==========================================

function clearDataSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCG_CONFIG.SHEET_DATA);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 1 && lastCol > 0) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent().setBackground(null);
}

function clearInputSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SCG_CONFIG.SHEET_INPUT);
  if (!sheet) return;
  sheet.getRange(SCG_CONFIG.COOKIE_CELL).clearContent();
  sheet.getRange(SCG_CONFIG.SHIPMENT_STRING_CELL).clearContent();
  const lastRow = sheet.getLastRow();
  if (lastRow >= SCG_CONFIG.INPUT_START_ROW) {
    sheet.getRange(SCG_CONFIG.INPUT_START_ROW, 1, lastRow - SCG_CONFIG.INPUT_START_ROW + 1, 1).clearContent();
  }
}

function clearSummarySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("สรุป_เจ้าของสินค้า");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent().setBackground(null);
}

function clearShipmentSummarySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("สรุป_Shipment");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent().setBackground(null);
}

function clearSummarySheet_UI() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert('⚠️ ยืนยัน', 'ต้องการล้างชีต สรุป_เจ้าของสินค้า ใช่ไหม?', ui.ButtonSet.YES_NO);
  if (result === ui.Button.YES) { clearSummarySheet(); SpreadsheetApp.getUi().alert('✅ ล้างข้อมูลเรียบร้อยแล้ว'); }
}

/**
 * [v4.1] ล้างทั้งหมด: Input + Data + สรุป_เจ้าของสินค้า + สรุป_Shipment
 */
function clearAllSCGSheets_UI() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    '🔥 ยืนยันการล้างข้อมูลทั้งหมด',
    'ต้องการล้างข้อมูลใน:\n- Input\n- Data\n- สรุป_เจ้าของสินค้า\n- สรุป_Shipment\nทั้งหมดหรือไม่?',
    ui.ButtonSet.YES_NO
  );

  if (result === ui.Button.YES) {
    clearInputSheet();
    clearDataSheet();
    clearSummarySheet();
    clearShipmentSummarySheet();
    ui.alert('✅ ล้างข้อมูลทั้งหมดเรียบร้อยแล้ว');
  }
}
