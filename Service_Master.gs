/**
 * VERSION: 4.1
 * 🧠 Service: Master Data Management
 * [v4.1] getRealLastRow_: แก้ปัญหา Checkbox overflow
 * [v4.1] syncNewDataToMaster: เพิ่ม GPS Feedback Loop + SYNC_STATUS Checkpoint
 * [v4.1] finalizeAndClean_MoveToMapping: แก้ lastRow ghost rows ด้วย deleteRows
 * [v4.1] processClustering_GridOptimized: ดักจับ NaN_NaN + แก้ COL_CONFIDENCE เป็น %
 * [v4.1] recalculateAllConfidence, recalculateAllQuality: ฟังก์ชันใหม่
 */

// ==========================================
// 1. HELPER
// ==========================================

function getRealLastRow_(sheet, columnIndex) {
  var data = sheet.getRange(1, columnIndex, sheet.getMaxRows(), 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== "" && data[i][0] !== null && typeof data[i][0] !== 'boolean') {
      return i + 1;
    }
  }
  return 1;
}

// ==========================================
// 2. SYNC NEW DATA + GPS FEEDBACK LOOP
// ==========================================

function syncNewDataToMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    ui.alert("⚠️ ระบบคิวทำงาน", "มีผู้ใช้งานอื่นกำลังอัปเดตฐานข้อมูลอยู่ กรุณาลองใหม่ในอีก 15 วินาทีครับ", ui.ButtonSet.OK);
    return;
  }

  // [v4.1] ตรวจสอบ Schema ก่อนทำงาน
  try { preCheck_Sync(); } catch(e) {
    ui.alert("❌ Schema Error", e.message, ui.ButtonSet.OK);
    return;
  }

  try {
    var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET);
    var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var queueSheet  = ss.getSheetByName(SCG_CONFIG.SHEET_GPS_QUEUE);

    if (!sourceSheet || !masterSheet) {
      ui.alert("❌ CRITICAL: ไม่พบ Sheet (Source หรือ Database)");
      return;
    }

    if (!queueSheet) {
      ui.alert("❌ CRITICAL: ไม่พบชีต GPS_Queue\nกรุณาสร้างชีตก่อนครับ");
      return;
    }

    // --- โหลด Database เข้า Memory ---
    var lastRowM = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
    var existingNames = {};
    var existingUUIDs = {};
    var dbData = [];

    if (lastRowM > 1) {
      var maxCol = Math.max(
        CONFIG.COL_NAME, CONFIG.COL_LAT, CONFIG.COL_LNG,
        CONFIG.COL_UUID, CONFIG.COL_COORD_SOURCE,
        CONFIG.COL_COORD_CONFIDENCE, CONFIG.COL_COORD_LAST_UPDATED
      );
      dbData = masterSheet.getRange(2, 1, lastRowM - 1, maxCol).getValues();
      dbData.forEach(function(r, i) {
        if (r[CONFIG.C_IDX.NAME]) existingNames[normalizeText(r[CONFIG.C_IDX.NAME])] = i;
        if (r[CONFIG.C_IDX.UUID]) existingUUIDs[r[CONFIG.C_IDX.UUID]] = i;
      });
    }

    // --- โหลด NameMapping เข้า Memory ---
    var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
    var aliasToUUID = {};
    if (mapSheet && mapSheet.getLastRow() > 1) {
      mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues()
        .forEach(function(r) {
          if (r[0] && r[1]) aliasToUUID[normalizeText(r[0])] = r[1];
        });
    }

    // --- อ่านข้อมูลจาก Source Sheet ---
    var lastRowS = sourceSheet.getLastRow();
    if (lastRowS < 2) {
      ui.alert("ℹ️ ไม่มีข้อมูลในชีตต้นทาง");
      return;
    }

    // [v4.1] ใช้ getLastColumn() แทน hardcode
    var lastColS = sourceSheet.getLastColumn();
    var sData = sourceSheet.getRange(2, 1, lastRowS - 1, lastColS).getValues();

    // --- ตัวแปรเก็บผลลัพธ์ ---
    var newEntries   = [];
    var queueEntries = [];
    var dbUpdates    = {};
    var currentBatch = new Set();
    var ts           = new Date();

    sData.forEach(function(row, rowIndex) {
      // [v4.1] ข้ามแถวที่ SYNCED แล้ว
      var syncStatus = row[SCG_CONFIG.SRC_IDX_SYNC_STATUS - 1];
      if (syncStatus === SCG_CONFIG.SYNC_STATUS_DONE) return;

      var name = row[SCG_CONFIG.SRC_IDX.NAME];
      var lat  = parseFloat(row[SCG_CONFIG.SRC_IDX.LAT]);
      var lng  = parseFloat(row[SCG_CONFIG.SRC_IDX.LNG]);

      if (!name || isNaN(lat) || isNaN(lng)) return;

      var cleanName = normalizeText(name);

      // --- หา match ใน Database ---
      var matchIdx = -1;

      if (existingNames.hasOwnProperty(cleanName)) {
        matchIdx = existingNames[cleanName];
      } else if (aliasToUUID.hasOwnProperty(cleanName)) {
        var uid = aliasToUUID[cleanName];
        if (existingUUIDs.hasOwnProperty(uid)) matchIdx = existingUUIDs[uid];
      }

      // กรณีที่ 1: ชื่อใหม่
      if (matchIdx === -1) {
        if (!currentBatch.has(cleanName)) {
          var newRow = new Array(22).fill("");
          newRow[CONFIG.C_IDX.NAME]               = name;
          newRow[CONFIG.C_IDX.LAT]                = lat;
          newRow[CONFIG.C_IDX.LNG]                = lng;
          newRow[CONFIG.C_IDX.VERIFIED]           = false;
          newRow[CONFIG.C_IDX.SYS_ADDR]           = row[SCG_CONFIG.SRC_IDX.SYS_ADDR] || "";
          newRow[CONFIG.C_IDX.UUID]               = generateUUID();
          newRow[CONFIG.C_IDX.CREATED]            = ts;
          newRow[CONFIG.C_IDX.UPDATED]            = ts;
          newRow[CONFIG.C_IDX.COORD_SOURCE]       = "SCG_System";
          newRow[CONFIG.C_IDX.COORD_CONFIDENCE]   = 50;
          newRow[CONFIG.C_IDX.COORD_LAST_UPDATED] = ts;
          newRow[CONFIG.C_IDX.RECORD_STATUS]      = "Active";

          newEntries.push(newRow);
          currentBatch.add(cleanName);
          existingNames[cleanName] = -999;
        }
        return;
      }

      // [v4.1] ถ้า matchIdx = -999 หมายถึงเพิ่งเพิ่มในรอบนี้
      if (matchIdx === -999) return;

      var dbRow = dbData[matchIdx];
      if (!dbRow) return;

      var dbLat  = parseFloat(dbRow[CONFIG.C_IDX.LAT]);
      var dbLng  = parseFloat(dbRow[CONFIG.C_IDX.LNG]);
      var dbUUID = dbRow[CONFIG.C_IDX.UUID];

      // กรณีที่ 2: Database ไม่มีพิกัด
      if (isNaN(dbLat) || isNaN(dbLng)) {
        queueEntries.push([ts, name, dbUUID, lat + ", " + lng, "ไม่มีพิกัดใน DB", "", "DB_NO_GPS", false, false]);
        return;
      }

      var diffKm     = getHaversineDistanceKM(lat, lng, dbLat, dbLng);
      var diffMeters = Math.round(diffKm * 1000);
      var threshold  = SCG_CONFIG.GPS_THRESHOLD_METERS / 1000;

      // กรณีที่ 3: diff ≤ 50m
      if (diffKm <= threshold) {
        if (!dbUpdates.hasOwnProperty(matchIdx)) dbUpdates[matchIdx] = ts;
        return;
      }

      // กรณีที่ 4: diff > 50m
      queueEntries.push([ts, name, dbUUID, lat + ", " + lng, dbLat + ", " + dbLng, diffMeters, "GPS_DIFF", false, false]);
    });

    // --- เขียนผลลัพธ์ ---
    var summary = [];

    if (newEntries.length > 0) {
      masterSheet.getRange(lastRowM + 1, 1, newEntries.length, 22).setValues(newEntries);
      summary.push("➕ เพิ่มลูกค้าใหม่: " + newEntries.length + " ราย");
    }

    var updateCount = Object.keys(dbUpdates).length;
    if (updateCount > 0) {
      Object.keys(dbUpdates).forEach(function(idx) {
        var rowNum = parseInt(idx) + 2;
        masterSheet.getRange(rowNum, CONFIG.COL_COORD_LAST_UPDATED).setValue(dbUpdates[idx]);
      });
      summary.push("🕐 อัปเดต timestamp: " + updateCount + " ราย");
    }

    if (queueEntries.length > 0) {
      var lastQueueRow = getRealLastRow_(queueSheet, 1);
      queueSheet.getRange(lastQueueRow + 1, 1, queueEntries.length, 9).setValues(queueEntries);
      summary.push("📋 ส่งเข้า GPS_Queue: " + queueEntries.length + " ราย");
    }

    // Mark SYNCED
    var syncColIndex = SCG_CONFIG.SRC_IDX_SYNC_STATUS;
    sData.forEach(function(row, i) {
      var name = row[SCG_CONFIG.SRC_IDX.NAME];
      var lat  = parseFloat(row[SCG_CONFIG.SRC_IDX.LAT]);
      var lng  = parseFloat(row[SCG_CONFIG.SRC_IDX.LNG]);
      var currentStatus = row[syncColIndex - 1];
      if (name && !isNaN(lat) && !isNaN(lng) && currentStatus !== SCG_CONFIG.SYNC_STATUS_DONE) {
        sourceSheet.getRange(i + 2, syncColIndex).setValue(SCG_CONFIG.SYNC_STATUS_DONE);
      }
    });

    SpreadsheetApp.flush();

    if (summary.length === 0) {
      ui.alert("👌 ไม่มีข้อมูลใหม่ที่ต้องประมวลผลครับ");
    } else {
      ui.alert("✅ Sync สำเร็จ!\n\n" + summary.join("\n"));
    }

  } catch (error) {
    console.error("Sync Error: " + error.message);
    ui.alert("❌ เกิดข้อผิดพลาด: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function cleanDistance_Helper(val) {
  if (!val) return "";
  if (typeof val === 'number') return val;
  return parseFloat(val.toString().replace(/,/g, '').replace('km', '').trim()) || "";
}

// ==========================================
// 3. GEO DATA & DEEP CLEAN
// ==========================================

function updateGeoData_SmartCache() { runDeepCleanBatch_100(); }
function autoGenerateMasterList_Smart() { processClustering_GridOptimized(); }

function runDeepCleanBatch_100() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var props    = PropertiesService.getScriptProperties();
  var startRow = parseInt(props.getProperty('DEEP_CLEAN_POINTER') || '2');

  if (startRow > lastRow) {
    ui.alert("🎉 ตรวจครบทุกแถวแล้ว (Pointer Reset)");
    props.deleteProperty('DEEP_CLEAN_POINTER');
    return;
  }

  var endRow  = Math.min(startRow + CONFIG.DEEP_CLEAN_LIMIT - 1, lastRow);
  var numRows = endRow - startRow + 1;

  var range  = sheet.getRange(startRow, 1, numRows, 22);
  var values = range.getValues();

  var origin       = CONFIG.DEPOT_LAT + "," + CONFIG.DEPOT_LNG;
  var updatedCount = 0;

  for (var i = 0; i < values.length; i++) {
    var row     = values[i];
    var lat     = row[CONFIG.C_IDX.LAT];
    var lng     = row[CONFIG.C_IDX.LNG];
    var changed = false;

    if (lat && lng && !row[CONFIG.C_IDX.GOOGLE_ADDR]) {
      try {
        var addr = GET_ADDR_WITH_CACHE(lat, lng);
        if (addr && addr !== "Error") { row[CONFIG.C_IDX.GOOGLE_ADDR] = addr; changed = true; }
      } catch (e) { console.warn("Geo Error: " + e.message); }
    }

    if (lat && lng && !row[CONFIG.C_IDX.DIST_KM]) {
      var km = CALCULATE_DISTANCE_KM(origin, lat + "," + lng);
      if (km) { row[CONFIG.C_IDX.DIST_KM] = km; changed = true; }
    }

    if (!row[CONFIG.C_IDX.UUID]) {
      row[CONFIG.C_IDX.UUID]    = generateUUID();
      row[CONFIG.C_IDX.CREATED] = row[CONFIG.C_IDX.CREATED] || new Date();
      changed = true;
    }

    var gAddr = row[CONFIG.C_IDX.GOOGLE_ADDR];
    if (gAddr && (!row[CONFIG.C_IDX.PROVINCE] || !row[CONFIG.C_IDX.DISTRICT])) {
      var parsed = parseAddressFromText(gAddr);
      if (parsed && parsed.province) {
        row[CONFIG.C_IDX.PROVINCE] = parsed.province;
        row[CONFIG.C_IDX.DISTRICT] = parsed.district;
        row[CONFIG.C_IDX.POSTCODE] = parsed.postcode;
        changed = true;
      }
    }

    // [v4.1] คำนวณ QUALITY Score
    var qualityScore = 0;
    var rowName = row[CONFIG.C_IDX.NAME];
    if (rowName && rowName.toString().length >= 3) qualityScore += 10;
    var rowLat = parseFloat(row[CONFIG.C_IDX.LAT]);
    var rowLng = parseFloat(row[CONFIG.C_IDX.LNG]);
    if (!isNaN(rowLat) && !isNaN(rowLng)) {
      qualityScore += 20;
      if (rowLat >= 6 && rowLat <= 21 && rowLng >= 97 && rowLng <= 106) qualityScore += 10;
    }
    if (row[CONFIG.C_IDX.GOOGLE_ADDR]) qualityScore += 15;
    if (row[CONFIG.C_IDX.PROVINCE] && row[CONFIG.C_IDX.DISTRICT]) qualityScore += 10;
    if (row[CONFIG.C_IDX.POSTCODE]) qualityScore += 5;
    if (row[CONFIG.C_IDX.UUID]) qualityScore += 10;
    if (row[CONFIG.C_IDX.VERIFIED] === true) qualityScore += 20;
    row[CONFIG.C_IDX.QUALITY] = Math.min(qualityScore, 100);

    if (!row[CONFIG.C_IDX.RECORD_STATUS]) {
      row[CONFIG.C_IDX.RECORD_STATUS] = "Active";
      changed = true;
    }

    if (changed) { row[CONFIG.C_IDX.UPDATED] = new Date(); updatedCount++; }
  }

  if (updatedCount > 0) range.setValues(values);
  props.setProperty('DEEP_CLEAN_POINTER', (endRow + 1).toString());
  ss.toast("✅ Processed rows " + startRow + "-" + endRow + " (Updated: " + updatedCount + ")", "Deep Clean");
}

function resetDeepCleanMemory() {
  PropertiesService.getScriptProperties().deleteProperty('DEEP_CLEAN_POINTER');
  SpreadsheetApp.getActiveSpreadsheet().toast("🔄 Memory Reset: ระบบถูกรีเซ็ต จะเริ่มตรวจสอบแถวที่ 2 ในรอบถัดไป", "System Ready");
}

// ==========================================
// 4. FINALIZE & CLEAN
// ==========================================

function finalizeAndClean_MoveToMapping() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    ui.alert("⚠️ ระบบคิวทำงาน", "มีผู้ใช้งานอื่นกำลังแก้ไขฐานข้อมูล กรุณารอสักครู่", ui.ButtonSet.OK);
    return;
  }

  try {
    var masterSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    var mapSheet    = ss.getSheetByName(CONFIG.MAPPING_SHEET);

    if (!masterSheet || !mapSheet) { ui.alert("❌ Error: Missing Sheets"); return; }

    // [v4.1] ใช้ getRealLastRow_ แทน getLastRow() เพราะ VERIFIED column มี Checkbox
    var lastRow = getRealLastRow_(masterSheet, CONFIG.COL_NAME);
    if (lastRow < 2) { ui.alert("ℹ️ Database is empty."); return; }

    var allData = masterSheet.getRange(2, 1, lastRow - 1, 17).getValues();
    var uuidMap = {};

    allData.forEach(function(row) {
      var uuid = row[CONFIG.C_IDX.UUID];
      if (uuid) {
        var n = normalizeText(row[CONFIG.C_IDX.NAME]);
        var s = normalizeText(row[CONFIG.C_IDX.SUGGESTED]);
        if (n) uuidMap[n] = uuid;
        if (s) uuidMap[s] = uuid;
      }
    });

    var backupName = "Backup_DB_" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd_HHmm");
    masterSheet.copyTo(ss).setName(backupName);

    var rowsToKeep       = [];
    var mappingToUpload  = [];
    var processedNames   = new Set();

    for (var i = 0; i < allData.length; i++) {
      var row           = allData[i];
      var rawName       = row[CONFIG.C_IDX.NAME];
      var suggestedName = row[CONFIG.C_IDX.SUGGESTED];
      var isVerified    = row[CONFIG.C_IDX.VERIFIED];
      var currentUUID   = row[CONFIG.C_IDX.UUID];

      if (isVerified === true) {
        rowsToKeep.push(row);
      } else if (suggestedName && suggestedName !== "") {
        if (rawName !== suggestedName && !processedNames.has(rawName)) {
          var targetUUID = uuidMap[normalizeText(suggestedName)] || currentUUID;
          var mapRow = new Array(5).fill("");
          mapRow[CONFIG.MAP_IDX.VARIANT]   = rawName;
          mapRow[CONFIG.MAP_IDX.UID]       = targetUUID;
          mapRow[CONFIG.MAP_IDX.CONFIDENCE] = 100;
          mapRow[CONFIG.MAP_IDX.MAPPED_BY] = "Human";
          mapRow[CONFIG.MAP_IDX.TIMESTAMP] = new Date();
          mappingToUpload.push(mapRow);
          processedNames.add(rawName);
        }
      }
    }

    if (mappingToUpload.length > 0) {
      var lastRowMap = mapSheet.getLastRow();
      mapSheet.getRange(lastRowMap + 1, 1, mappingToUpload.length, 5).setValues(mappingToUpload);
    }

    masterSheet.getRange(2, 1, lastRow - 1, 17).clearContent();

    if (rowsToKeep.length > 0) {
      masterSheet.getRange(2, 1, rowsToKeep.length, 17).setValues(rowsToKeep);

      // [v4.1] ลบ ghost rows ด้วย deleteRows
      var ghostStart = rowsToKeep.length + 2;
      var ghostCount = lastRow - 1 - rowsToKeep.length;
      if (ghostCount > 0) masterSheet.deleteRows(ghostStart, ghostCount);

      SpreadsheetApp.flush();
      ui.alert("✅ Finalize Complete:\n- New Mappings: " + mappingToUpload.length + "\n- Active Master Data: " + rowsToKeep.length);
    } else {
      masterSheet.getRange(2, 1, allData.length, 17).setValues(allData);
      ui.alert("⚠️ Warning: No Verified rows found. Data restored to original state.");
    }
  } catch (e) {
    console.error("Finalize Error: " + e.message);
    ui.alert("❌ CRITICAL WRITE ERROR: " + e.message + "\nPlease check Backup Sheet.");
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 5. UUID & MAPPING REPAIR
// ==========================================

function assignMissingUUIDs() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var ui   = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var range  = sheet.getRange(2, CONFIG.COL_UUID, lastRow - 1, 1);
  var values = range.getValues();
  var count  = 0;

  var newValues = values.map(function(r) {
    if (!r[0]) { count++; return [generateUUID()]; }
    return [r[0]];
  });

  if (count > 0) {
    range.setValues(newValues);
    ui.alert("✅ Generated " + count + " new UUIDs.");
  } else {
    ui.alert("ℹ️ All rows already have UUIDs.");
  }
}

function repairNameMapping_Full() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var ui      = SpreadsheetApp.getUi();
  var dbSheet  = ss.getSheetByName(CONFIG.SHEET_NAME);
  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);

  var dbData = dbSheet.getRange(2, 1, getRealLastRow_(dbSheet, CONFIG.COL_NAME) - 1, CONFIG.COL_UUID).getValues();
  var uuidMap = {};
  dbData.forEach(function(r) {
    if (r[CONFIG.C_IDX.UUID]) uuidMap[normalizeText(r[CONFIG.C_IDX.NAME])] = r[CONFIG.C_IDX.UUID];
  });

  var mapRange  = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 5);
  var mapValues = mapRange.getValues();
  var cleanList = [];
  var seen      = new Set();

  mapValues.forEach(function(r) {
    var oldN = r[CONFIG.MAP_IDX.VARIANT];
    var uid  = r[CONFIG.MAP_IDX.UID];
    var conf = r[CONFIG.MAP_IDX.CONFIDENCE] || 100;
    var by   = r[CONFIG.MAP_IDX.MAPPED_BY]  || "System_Repair";
    var ts   = r[CONFIG.MAP_IDX.TIMESTAMP]  || new Date();

    var normOld = normalizeText(oldN);
    if (!normOld) return;
    if (!uid) uid = uuidMap[normalizeText(r[1])] || generateUUID();

    if (!seen.has(normOld)) {
      seen.add(normOld);
      var mapRow = new Array(5).fill("");
      mapRow[CONFIG.MAP_IDX.VARIANT]    = oldN;
      mapRow[CONFIG.MAP_IDX.UID]        = uid;
      mapRow[CONFIG.MAP_IDX.CONFIDENCE] = conf;
      mapRow[CONFIG.MAP_IDX.MAPPED_BY]  = by;
      mapRow[CONFIG.MAP_IDX.TIMESTAMP]  = ts;
      cleanList.push(mapRow);
    }
  });

  if (cleanList.length > 0) {
    mapSheet.getRange(2, 1, mapSheet.getLastRow(), 5).clearContent();
    mapSheet.getRange(2, 1, cleanList.length, 5).setValues(cleanList);
    ui.alert("✅ Repair Complete. Total Mappings: " + cleanList.length);
  } else {
    ui.alert("ℹ️ No repair needed or mapping is empty.");
  }
}

// ==========================================
// 6. CLUSTERING
// ==========================================

function processClustering_GridOptimized() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var range  = sheet.getRange(2, 1, lastRow - 1, 15);
  var values = range.getValues();
  var clusters = [];
  var grid     = {};

  for (var i = 0; i < values.length; i++) {
    var r   = values[i];
    var lat = parseFloat(r[CONFIG.C_IDX.LAT]);
    var lng = parseFloat(r[CONFIG.C_IDX.LNG]);

    // [v4.1] ดักจับ NaN ก่อนสร้าง gridKey
    if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;

    var gridKey = Math.floor(lat * 10) + "_" + Math.floor(lng * 10);
    if (!grid[gridKey]) grid[gridKey] = [];
    grid[gridKey].push(i);

    if (r[CONFIG.C_IDX.VERIFIED] === true) {
      clusters.push({ lat: lat, lng: lng, name: r[CONFIG.C_IDX.SUGGESTED] || r[CONFIG.C_IDX.NAME],
        rowIndexes: [i], hasLock: true, gridKey: gridKey });
    }
  }

  for (var i = 0; i < values.length; i++) {
    if (values[i][CONFIG.C_IDX.VERIFIED] === true) continue;

    var lat = parseFloat(values[i][CONFIG.C_IDX.LAT]);
    var lng = parseFloat(values[i][CONFIG.C_IDX.LNG]);

    // [v4.1] ดักจับ NaN ทั้งสองค่า
    if (!lat || !lng || isNaN(lat) || isNaN(lng)) continue;

    var myGridKey = Math.floor(lat * 10) + "_" + Math.floor(lng * 10);
    var found = false;

    for (var c = 0; c < clusters.length; c++) {
      if (clusters[c].gridKey === myGridKey) {
        var dist = getHaversineDistanceKM(lat, lng, clusters[c].lat, clusters[c].lng);
        if (dist <= CONFIG.DISTANCE_THRESHOLD_KM) {
          clusters[c].rowIndexes.push(i); found = true; break;
        }
      }
    }

    if (!found) clusters.push({ lat: lat, lng: lng, rowIndexes: [i], hasLock: false, name: null, gridKey: myGridKey });
  }

  var updateCount = 0;
  clusters.forEach(function(g) {
    var candidateNames = [];
    g.rowIndexes.forEach(function(idx) {
      var rawName = values[idx][CONFIG.C_IDX.NAME];
      var existingSuggested = values[idx][CONFIG.C_IDX.SUGGESTED];
      candidateNames.push(rawName);
      if (existingSuggested && existingSuggested !== rawName) {
        candidateNames.push(existingSuggested, existingSuggested, existingSuggested);
      }
    });

    var winner = g.hasLock ? g.name : getBestName_Smart(candidateNames);

    // [v4.1] คำนวณ Confidence เป็น % จริงๆ
    var countScore    = Math.min(g.rowIndexes.length * 10, 40);
    var hasVerified   = g.rowIndexes.some(function(idx) { return values[idx][CONFIG.C_IDX.VERIFIED] === true; });
    var verifiedScore = hasVerified ? 40 : 0;
    var hasCoord      = !isNaN(parseFloat(values[g.rowIndexes[0]][CONFIG.C_IDX.LAT])) &&
                        !isNaN(parseFloat(values[g.rowIndexes[0]][CONFIG.C_IDX.LNG]));
    var coordScore    = hasCoord ? 20 : 0;
    var confidence    = Math.min(countScore + verifiedScore + coordScore, 100);

    g.rowIndexes.forEach(function(idx) {
      if (values[idx][CONFIG.C_IDX.VERIFIED] !== true) {
        var currentSuggested  = values[idx][CONFIG.C_IDX.SUGGESTED];
        var currentConfidence = values[idx][CONFIG.C_IDX.CONFIDENCE];
        if (currentSuggested !== winner || currentConfidence !== confidence) {
          values[idx][CONFIG.C_IDX.SUGGESTED]  = winner;
          values[idx][CONFIG.C_IDX.CONFIDENCE] = confidence;
          values[idx][CONFIG.C_IDX.NORMALIZED] = normalizeText(winner);
          updateCount++;
        }
      }
    });
  });

  if (updateCount > 0) {
    range.setValues(values);
    SpreadsheetApp.getActiveSpreadsheet().toast("✅ จัดกลุ่มสำเร็จ! (Updated: " + updateCount + " rows)", "Clustering V4.1");
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast("ℹ️ ข้อมูลจัดกลุ่มเรียบร้อยดีอยู่แล้ว ไม่มีการเปลี่ยนแปลง", "Clustering V4.1");
  }
}

// ==========================================
// 7. RECALCULATE CONFIDENCE & QUALITY
// ==========================================

function recalculateAllConfidence() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var maxCol = Math.max(17, CONFIG.COL_COORD_SOURCE);
  var data   = sheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
  var updatedCount = 0;

  data.forEach(function(row, i) {
    var name     = row[CONFIG.C_IDX.NAME];
    var lat      = parseFloat(row[CONFIG.C_IDX.LAT]);
    var lng      = parseFloat(row[CONFIG.C_IDX.LNG]);
    var verified = row[CONFIG.C_IDX.VERIFIED];
    if (!name) return;

    var verifiedScore = (verified === true) ? 40 : 0;
    var coordScore    = (!isNaN(lat) && !isNaN(lng)) ? 20 : 0;
    var addrScore     = row[CONFIG.C_IDX.GOOGLE_ADDR] ? 10 : 0;
    var geoScore      = (row[CONFIG.C_IDX.PROVINCE] && row[CONFIG.C_IDX.DISTRICT]) ? 10 : 0;
    var uuidScore     = row[CONFIG.C_IDX.UUID] ? 10 : 0;
    var sourceScore   = 0;
    if (row.length > CONFIG.C_IDX.COORD_SOURCE) {
      sourceScore = (row[CONFIG.C_IDX.COORD_SOURCE] === "Driver_GPS") ? 10 : 0;
    }

    var newConfidence = Math.min(verifiedScore + coordScore + addrScore + geoScore + uuidScore + sourceScore, 100);
    if (row[CONFIG.C_IDX.CONFIDENCE] !== newConfidence) {
      data[i][CONFIG.C_IDX.CONFIDENCE] = newConfidence;
      updatedCount++;
    }
  });

  if (updatedCount > 0) {
    sheet.getRange(2, 1, data.length, maxCol).setValues(data);
    SpreadsheetApp.flush();
  }

  ui.alert("✅ คำนวณ Confidence ใหม่เสร็จ!\nอัปเดต: " + updatedCount + " แถว");
}

function recalculateAllQuality() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var ui    = SpreadsheetApp.getUi();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
  if (lastRow < 2) return;

  var data = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
  var updatedCount = 0;

  data.forEach(function(row, i) {
    var name = row[CONFIG.C_IDX.NAME];
    if (!name) return;

    var qualityScore = 0;
    if (name.toString().length >= 3) qualityScore += 10;
    var lat = parseFloat(row[CONFIG.C_IDX.LAT]);
    var lng = parseFloat(row[CONFIG.C_IDX.LNG]);
    if (!isNaN(lat) && !isNaN(lng)) {
      qualityScore += 20;
      if (lat >= 6 && lat <= 21 && lng >= 97 && lng <= 106) qualityScore += 10;
    }
    if (row[CONFIG.C_IDX.GOOGLE_ADDR]) qualityScore += 15;
    if (row[CONFIG.C_IDX.PROVINCE] && row[CONFIG.C_IDX.DISTRICT]) qualityScore += 10;
    if (row[CONFIG.C_IDX.POSTCODE]) qualityScore += 5;
    if (row[CONFIG.C_IDX.UUID]) qualityScore += 10;
    if (row[CONFIG.C_IDX.VERIFIED] === true) qualityScore += 20;

    var newQuality = Math.min(qualityScore, 100);
    if (row[CONFIG.C_IDX.QUALITY] !== newQuality) {
      data[i][CONFIG.C_IDX.QUALITY] = newQuality;
      updatedCount++;
    }
  });

  if (updatedCount > 0) {
    sheet.getRange(2, 1, data.length, 17).setValues(data);
    SpreadsheetApp.flush();
  }

  var stats = { total: 0, high: 0, mid: 0, low: 0, zero: 0 };
  data.forEach(function(row) {
    var q = parseFloat(row[CONFIG.C_IDX.QUALITY]);
    if (isNaN(q) || !row[CONFIG.C_IDX.NAME]) return;
    stats.total++;
    if (q >= 80) stats.high++;
    else if (q >= 50) stats.mid++;
    else if (q > 0) stats.low++;
    else stats.zero++;
  });

  ui.alert(
    "✅ คำนวณ Quality Score เสร็จแล้ว!\n\nอัปเดต: " + updatedCount + " แถว\n\n" +
    "📊 สรุป:\n🟢 ≥80%: " + stats.high + " แถว\n🟡 50-79%: " + stats.mid + " แถว\n🔴 <50%: " + stats.low + " แถว"
  );
}
