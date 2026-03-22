/**
 * VERSION: 4.1
 * 🛠️ System Upgrade Tool
 * [v4.1] อัปเดต upgradeDatabaseStructure ให้ใช้ GPS Tracking Columns (Col 18-20)
 * [v4.1] ลบ Haversine Fallback ออก (มีใน Utils_Common.gs แล้ว)
 */

function upgradeDatabaseStructure() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var ui  = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) { ui.alert("❌ Critical Error: ไม่พบชีต " + CONFIG.SHEET_NAME); return; }

  // [v4.1] Col 18-20 ถูกใช้สำหรับ GPS Tracking แล้ว
  var gpsHeaders = [
    "Coord_Source",        // Col 18 (R)
    "Coord_Confidence",    // Col 19 (S)
    "Coord_Last_Updated"   // Col 20 (T)
  ];

  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var missingHeaders = [];
  gpsHeaders.forEach(function(header) {
    if (currentHeaders.indexOf(header) === -1) missingHeaders.push(header);
  });

  if (missingHeaders.length === 0) {
    ui.alert("✅ Database Structure เป็นปัจจุบันแล้ว\n\nCol 18: Coord_Source ✅\nCol 19: Coord_Confidence ✅\nCol 20: Coord_Last_Updated ✅");
    return;
  }

  var response = ui.alert(
    "⚠️ พบคอลัมน์ขาดหาย",
    "GPS Tracking Columns ขาดหาย " + missingHeaders.length + " รายการ:\n" + missingHeaders.join(", ") + "\n\nต้องการเพิ่มทันทีหรือไม่?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var startCol = sheet.getLastColumn() + 1;
  var range    = sheet.getRange(1, startCol, 1, missingHeaders.length);
  range.setValues([missingHeaders]);
  range.setFontWeight("bold");
  range.setBackground("#d0f0c0");
  range.setBorder(true, true, true, true, true, true);
  sheet.autoResizeColumns(startCol, missingHeaders.length);

  ui.alert("✅ เพิ่มคอลัมน์ GPS Tracking สำเร็จ!");
}

function upgradeNameMappingStructure_V4() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  var ui    = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert("❌ Critical Error: ไม่พบชีต " + CONFIG.MAPPING_SHEET); return; }

  var targetHeaders = ["Variant_Name", "Master_UID", "Confidence_Score", "Mapped_By", "Timestamp"];
  var range = sheet.getRange(1, 1, 1, 5);
  range.setValues([targetHeaders]);
  range.setFontWeight("bold");
  range.setFontColor("white");
  range.setBackground("#7c3aed");
  range.setBorder(true, true, true, true, true, true);

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 150);
  sheet.setFrozenRows(1);

  ui.alert("✅ Schema Upgrade V4.0 สำเร็จ!\nอัปเกรดชีต NameMapping เป็น 5 คอลัมน์เรียบร้อย");
}

function findHiddenDuplicates() {
  console.time("HiddenDupesCheck");
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var idxLat  = CONFIG.C_IDX.LAT;
  var idxLng  = CONFIG.C_IDX.LNG;
  var idxName = CONFIG.C_IDX.NAME;

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var data       = sheet.getRange(2, 1, lastRow - 1, 15).getValues();
  var duplicates = [];
  var grid       = {};

  for (var i = 0; i < data.length; i++) {
    var lat = parseFloat(data[i][idxLat]);
    var lng = parseFloat(data[i][idxLng]);
    if (isNaN(lat) || isNaN(lng)) continue;
    var gridKey = Math.floor(lat * 100) + "_" + Math.floor(lng * 100);
    if (!grid[gridKey]) grid[gridKey] = [];
    grid[gridKey].push({ index: i, row: data[i] });
  }

  for (var key in grid) {
    var bucket = grid[key];
    if (bucket.length < 2) continue;
    for (var a = 0; a < bucket.length; a++) {
      for (var b = a + 1; b < bucket.length; b++) {
        var item1 = bucket[a];
        var item2 = bucket[b];
        var dist  = getHaversineDistanceKM(item1.row[idxLat], item1.row[idxLng], item2.row[idxLat], item2.row[idxLng]);
        if (dist <= 0.05) {
          var name1 = normalizeText(item1.row[idxName]);
          var name2 = normalizeText(item2.row[idxName]);
          if (name1 !== name2) {
            duplicates.push({ row1: item1.index + 2, name1: item1.row[idxName], row2: item2.index + 2, name2: item2.row[idxName], distance: (dist * 1000).toFixed(0) + " ม." });
          }
        }
      }
    }
  }

  console.timeEnd("HiddenDupesCheck");

  if (duplicates.length > 0) {
    var msg = "⚠️ พบพิกัดทับซ้อน " + duplicates.length + " คู่:\n\n";
    duplicates.slice(0, 15).forEach(function(d) {
      msg += `• แถว ${d.row1} vs ${d.row2}: ${d.name1} / ${d.name2} (ห่าง ${d.distance})\n`;
    });
    if (duplicates.length > 15) msg += `\n...และอีก ${duplicates.length - 15} คู่`;
    ui.alert(msg);
  } else {
    ui.alert("✅ ไม่พบข้อมูลซ้ำซ้อนในระยะ 50 เมตร");
  }
}

function verifyDatabaseStructure() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  console.log("คอลัมน์ทั้งหมด: " + lastCol);
  var expected = ["Coord_Source", "Coord_Confidence", "Coord_Last_Updated"];
  expected.forEach(function(h, i) {
    var actual = headers[17 + i];
    if (actual === h) console.log("✅ Col " + (18 + i) + ": " + h);
    else console.log("❌ Col " + (18 + i) + ": คาดว่า '" + h + "' แต่เจอ '" + actual + "'");
  });
}
