/**
 * VERSION: 4.1
 * 🚛 Logistics Master Data System - Configuration V4.1 (Enterprise Edition)
 * ------------------------------------------------------------------
 * [v4.1] เพิ่ม GPS Feedback Config (SHEET_GPS_QUEUE, GPS_THRESHOLD_METERS)
 * [v4.1] ย้าย SRC_IDX จาก Service_Master.gs มาไว้ที่นี่ (Single Source of Truth)
 * [v4.1] เพิ่ม SYNC_STATUS Tracking สำหรับ Checkpoint
 * [v4.1] กำหนดใช้ Col 18-20 สำหรับ GPS Tracking
 * [v4.1] เพิ่ม Soft Delete Columns (COL_RECORD_STATUS, COL_MERGED_TO_UUID)
 */

var CONFIG = {
  // --- SHEET NAMES ---
  SHEET_NAME: "Database",
  MAPPING_SHEET: "NameMapping",
  SOURCE_SHEET: "SCGนครหลวงJWDภูมิภาค",
  SHEET_POSTAL: "PostalRef",

  // --- 🧠 AI CONFIGURATION (SECURED) ---
  get GEMINI_API_KEY() {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) throw new Error("CRITICAL ERROR: GEMINI_API_KEY is not set. Please run setupEnvironment() first.");
    return key;
  },
  USE_AI_AUTO_FIX: true,
  AI_MODEL: "gemini-1.5-flash",
  AI_BATCH_SIZE: 20,

  // --- 🔴 DEPOT LOCATION ---
  DEPOT_LAT: 14.164688,
  DEPOT_LNG: 100.625354,

  // --- SYSTEM THRESHOLDS & LIMITS ---
  DISTANCE_THRESHOLD_KM: 0.05,
  BATCH_LIMIT: 50,
  DEEP_CLEAN_LIMIT: 100,
  API_MAX_RETRIES: 3,
  API_TIMEOUT_MS: 30000,
  CACHE_EXPIRATION: 21600,

  // --- DATABASE COLUMNS INDEX (1-BASED) ---
  COL_NAME: 1,
  COL_LAT: 2,
  COL_LNG: 3,
  COL_SUGGESTED: 4,
  COL_CONFIDENCE: 5,
  COL_NORMALIZED: 6,
  COL_VERIFIED: 7,
  COL_SYS_ADDR: 8,
  COL_ADDR_GOOG: 9,
  COL_DIST_KM: 10,
  COL_UUID: 11,
  COL_PROVINCE: 12,
  COL_DISTRICT: 13,
  COL_POSTCODE: 14,
  COL_QUALITY: 15,
  COL_CREATED: 16,
  COL_UPDATED: 17,

  // --- [v4.1] GPS TRACKING COLUMNS ---
  COL_COORD_SOURCE:       18, // R: พิกัดมาจากไหน (SCG_System / Driver_GPS)
  COL_COORD_CONFIDENCE:   19, // S: ความน่าเชื่อถือ 0-100
  COL_COORD_LAST_UPDATED: 20, // T: อัปเดตพิกัดล่าสุดเมื่อไร

  // --- [v4.1] SOFT DELETE COLUMNS ---
  COL_RECORD_STATUS:  21, // U: Active / Inactive / Merged
  COL_MERGED_TO_UUID: 22, // V: UUID ที่ชี้ไปหลัง Merge

  // --- NAMEMAPPING COLUMNS INDEX (1-BASED) ---
  MAP_COL_VARIANT: 1,
  MAP_COL_UID: 2,
  MAP_COL_CONFIDENCE: 3,
  MAP_COL_MAPPED_BY: 4,
  MAP_COL_TIMESTAMP: 5,

  // --- DATABASE ARRAY INDEX MAPPING (0-BASED) ---
  get C_IDX() {
    return {
      NAME: this.COL_NAME - 1,
      LAT: this.COL_LAT - 1,
      LNG: this.COL_LNG - 1,
      SUGGESTED: this.COL_SUGGESTED - 1,
      CONFIDENCE: this.COL_CONFIDENCE - 1,
      NORMALIZED: this.COL_NORMALIZED - 1,
      VERIFIED: this.COL_VERIFIED - 1,
      SYS_ADDR: this.COL_SYS_ADDR - 1,
      GOOGLE_ADDR: this.COL_ADDR_GOOG - 1,
      DIST_KM: this.COL_DIST_KM - 1,
      UUID: this.COL_UUID - 1,
      PROVINCE: this.COL_PROVINCE - 1,
      DISTRICT: this.COL_DISTRICT - 1,
      POSTCODE: this.COL_POSTCODE - 1,
      QUALITY: this.COL_QUALITY - 1,
      CREATED: this.COL_CREATED - 1,
      UPDATED: this.COL_UPDATED - 1,
      COORD_SOURCE:       this.COL_COORD_SOURCE - 1,
      COORD_CONFIDENCE:   this.COL_COORD_CONFIDENCE - 1,
      COORD_LAST_UPDATED: this.COL_COORD_LAST_UPDATED - 1,
      RECORD_STATUS:      this.COL_RECORD_STATUS - 1,
      MERGED_TO_UUID:     this.COL_MERGED_TO_UUID - 1,
    };
  },

  // --- NAMEMAPPING ARRAY INDEX (0-BASED) ---
  get MAP_IDX() {
    return {
      VARIANT:   this.MAP_COL_VARIANT - 1,
      UID:       this.MAP_COL_UID - 1,
      CONFIDENCE: this.MAP_COL_CONFIDENCE - 1,
      MAPPED_BY: this.MAP_COL_MAPPED_BY - 1,
      TIMESTAMP: this.MAP_COL_TIMESTAMP - 1
    };
  }
};

// --- SCG SPECIFIC CONFIG ---
const SCG_CONFIG = {
  SHEET_DATA:     'Data',
  SHEET_INPUT:    'Input',
  SHEET_EMPLOYEE: 'ข้อมูลพนักงาน',
  API_URL: 'https://fsm.scgjwd.com/Monitor/SearchDelivery',
  INPUT_START_ROW: 4,
  COOKIE_CELL: 'B1',
  SHIPMENT_STRING_CELL: 'B3',
  SHEET_MASTER_DB: 'Database',
  SHEET_MAPPING:   'NameMapping',

  // --- [v4.1] GPS FEEDBACK ---
  SHEET_GPS_QUEUE:      'GPS_Queue',
  GPS_THRESHOLD_METERS: 50,

  // --- [v4.1] SCGนครหลวงJWDภูมิภาค COLUMN INDEX (0-based) ---
  SRC_IDX: {
    NAME:     12,  // Col M: ชื่อปลายทาง
    LAT:      14,  // Col O: LAT (GPS จริงจากคนขับ)
    LNG:      15,  // Col P: LONG (GPS จริงจากคนขับ)
    SYS_ADDR: 18,  // Col S: ที่อยู่ปลายทาง
    DIST:     23,  // Col X: ระยะทางจากคลัง_Km
    GOOG_ADDR:24   // Col Y: ชื่อที่อยู่จาก_LatLong
  },

  // --- [v4.1] SYNC TRACKING ---
  SRC_IDX_SYNC_STATUS: 37, // Col AK: สถานะการ Sync
  SYNC_STATUS_DONE: "SYNCED",

  JSON_MAP: {
    SHIPMENT_NO:   'shipmentNo',
    CUSTOMER_NAME: 'customerName',
    DELIVERY_DATE: 'deliveryDate'
  }
};

// --- System Health Check ---
CONFIG.validateSystemIntegrity = function() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var errors = [];

  var requiredSheets = [this.SHEET_NAME, this.MAPPING_SHEET, SCG_CONFIG.SHEET_INPUT, this.SHEET_POSTAL];
  requiredSheets.forEach(function(name) {
    if (!ss.getSheetByName(name)) errors.push("Missing Sheet: " + name);
  });

  try {
    var key = this.GEMINI_API_KEY;
    if (!key || key.length < 20) errors.push("Invalid Gemini API Key format");
  } catch (e) {
    errors.push("Gemini API Key is not set. Please run setupEnvironment() first.");
  }

  if (errors.length > 0) {
    var msg = "⚠️ SYSTEM INTEGRITY FAILED:\n" + errors.join("\n");
    console.error(msg);
    throw new Error(msg);
  } else {
    console.log("✅ System Integrity: OK");
    return true;
  }
};
