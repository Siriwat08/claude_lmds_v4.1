/**
 * VERSION: 4.1
 * 🔍 Service: Search Engine
 * [v4.1] แก้ UTF-8 Cache Bomb: วัดขนาดด้วย Bytes แทน .length
 */

function searchMasterData(keyword, page) {
  console.time("SearchLatency");
  try {
    var pageNum  = parseInt(page) || 1;
    var pageSize = 20;

    if (!keyword || keyword.toString().trim() === "") {
      return { items: [], total: 0, totalPages: 0, currentPage: 1 };
    }

    var rawKey       = keyword.toString().toLowerCase().trim();
    var searchTokens = rawKey.split(/\s+/).filter(function(k) { return k.length > 0; });
    if (searchTokens.length === 0) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var aliasMap = getCachedNameMapping_(ss);
    var sheet    = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var lastRow = getRealLastRow_(sheet, CONFIG.COL_NAME);
    if (lastRow < 2) return { items: [], total: 0, totalPages: 0, currentPage: 1 };

    var data    = sheet.getRange(2, 1, lastRow - 1, 17).getValues();
    var matches = [];

    for (var i = 0; i < data.length; i++) {
      var row  = data[i];
      var name = row[CONFIG.C_IDX.NAME];
      if (!name) continue;

      var address    = row[CONFIG.C_IDX.GOOGLE_ADDR] || row[CONFIG.C_IDX.SYS_ADDR] || "";
      var lat        = row[CONFIG.C_IDX.LAT];
      var lng        = row[CONFIG.C_IDX.LNG];
      var uuid       = row[CONFIG.C_IDX.UUID];
      var aiKeywords = row[CONFIG.C_IDX.NORMALIZED] ? row[CONFIG.C_IDX.NORMALIZED].toString().toLowerCase() : "";
      var normName   = normalizeText(name);
      var rawName    = name.toString().toLowerCase();
      var aliases    = uuid ? (aliasMap[uuid] || "") : "";
      var haystack   = (rawName + " " + normName + " " + aliases + " " + aiKeywords + " " + address.toString().toLowerCase());

      var isMatch = searchTokens.every(function(token) { return haystack.indexOf(token) > -1; });

      if (isMatch) {
        matches.push({
          name: name, address: address, lat: lat, lng: lng,
          mapLink: (lat && lng) ? "https://www.google.com/maps/dir/?api=1&destination=" + lat + "," + lng : "",
          uuid: uuid,
          score: aiKeywords.includes(rawKey) ? 10 : 1
        });
      }
    }

    matches.sort(function(a, b) { return b.score - a.score; });

    var totalItems = matches.length;
    var totalPages = Math.ceil(totalItems / pageSize);
    if (pageNum > totalPages && totalPages > 0) pageNum = 1;

    var startIndex = (pageNum - 1) * pageSize;
    var pagedItems = matches.slice(startIndex, startIndex + pageSize);

    return { items: pagedItems, total: totalItems, totalPages: totalPages, currentPage: pageNum };

  } catch (error) {
    console.error("[Search Error]: " + error.message);
    return { items: [], total: 0, totalPages: 0, currentPage: 1, error: error.message };
  } finally {
    console.timeEnd("SearchLatency");
  }
}

function getCachedNameMapping_(ss) {
  var cache     = CacheService.getScriptCache();
  var cachedMap = cache.get("NAME_MAPPING_JSON_V4");
  if (cachedMap) return JSON.parse(cachedMap);

  var mapSheet = ss.getSheetByName(CONFIG.MAPPING_SHEET);
  var aliasMap = {};

  if (mapSheet && mapSheet.getLastRow() > 1) {
    var mapData = mapSheet.getRange(2, 1, mapSheet.getLastRow() - 1, 2).getValues();
    mapData.forEach(function(row) {
      var variant = row[0];
      var uid     = row[1];
      if (variant && uid) {
        if (!aliasMap[uid]) aliasMap[uid] = "";
        var normVariant = normalizeText(variant);
        aliasMap[uid] += " " + normVariant + " " + variant.toString().toLowerCase();
      }
    });

    try {
      var jsonString = JSON.stringify(aliasMap);
      // [v4.1] วัดขนาด Bytes จริงแทน .length เพราะภาษาไทย 1 ตัว = 3 Bytes
      var byteSize = Utilities.newBlob(jsonString).getBytes().length;
      if (byteSize < 100000) {
        cache.put("NAME_MAPPING_JSON_V4", jsonString, 3600);
        console.log("[Cache] NameMapping cached (" + byteSize + " bytes)");
      } else {
        console.warn("[Cache] NameMapping too large (" + byteSize + " bytes), skipping cache.");
      }
    } catch (e) {
      console.warn("[Cache Error]: " + e.message);
    }
  }

  return aliasMap;
}

function clearSearchCache() {
  CacheService.getScriptCache().remove("NAME_MAPPING_JSON_V4");
  console.log("[Cache] Search Cache Cleared.");
}
