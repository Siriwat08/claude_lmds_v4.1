/**
 * VERSION: 4.1
 * 🌍 Service: Geo Address & Google Maps Formulas
 * [v4.1] เพิ่ม clearPostalCache() สำหรับล้าง cache จากเมนู
 */

const POSTAL_COL = { ZIP: 0, DISTRICT: 2, PROVINCE: 3 };
var _POSTAL_CACHE = null;

function parseAddressFromText(fullAddress) {
  var result = { province: "", district: "", postcode: "" };
  if (!fullAddress) return result;

  var addrStr  = fullAddress.toString().trim();
  var zipMatch = addrStr.match(/(\d{5})/);
  if (zipMatch && zipMatch[1]) result.postcode = zipMatch[1];

  var postalDB = getPostalDataCached();
  if (postalDB && result.postcode && postalDB.byZip[result.postcode]) {
    var infoList = postalDB.byZip[result.postcode];
    if (infoList.length > 0) {
      result.province = infoList[0].province;
      result.district = infoList[0].district;
      return result;
    }
  }

  var provMatch = addrStr.match(/(?:จ\.|จังหวัด)\s*([ก-๙a-zA-Z0-9]+)/i);
  if (provMatch && provMatch[1]) result.province = provMatch[1].trim();

  var distMatch = addrStr.match(/(?:อ\.|อำเภอ|เขต)\s*([ก-๙a-zA-Z0-9]+)/i);
  if (distMatch && distMatch[1]) result.district = distMatch[1].trim();

  if (!result.province && (addrStr.includes("กรุงเทพ") || addrStr.includes("Bangkok") || addrStr.includes("กทม"))) {
    result.province = "กรุงเทพมหานคร";
  }

  return result;
}

function getPostalDataCached() {
  if (_POSTAL_CACHE) return _POSTAL_CACHE;

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_POSTAL) ? CONFIG.SHEET_POSTAL : "PostalRef";
  var sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var db   = { byZip: {} };

  data.forEach(function(row) {
    if (row.length <= POSTAL_COL.PROVINCE) return;
    var pc = String(row[POSTAL_COL.ZIP]).trim();
    if (!pc) return;
    if (!db.byZip[pc]) db.byZip[pc] = [];
    db.byZip[pc].push({ postcode: pc, district: row[POSTAL_COL.DISTRICT], province: row[POSTAL_COL.PROVINCE] });
  });

  _POSTAL_CACHE = db;
  return db;
}

/**
 * [v4.1] ล้าง Postal Cache อย่างปลอดภัย
 */
function clearPostalCache() {
  _POSTAL_CACHE = null;
  console.log("[Cache] Postal Cache cleared.");
}

// Google Maps formula functions
const _mapsMd5 = (key = "") => {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key)
    .map((char) => (char + 256).toString(16).slice(-2)).join("");
};

const _mapsGetCache = (key) => {
  try { return CacheService.getDocumentCache().get(_mapsMd5(key)); } catch(e) { return null; }
};

const _mapsSetCache = (key, value) => {
  try {
    const expirationInSeconds = (typeof CONFIG !== 'undefined' && CONFIG.CACHE_EXPIRATION) ? CONFIG.CACHE_EXPIRATION : 21600;
    if (value && value.toString().length < 90000) {
      CacheService.getDocumentCache().put(_mapsMd5(key), value, expirationInSeconds);
    }
  } catch (e) {
    console.warn("[Geo Cache Warn]: Could not cache key " + key);
  }
};

const GOOGLEMAPS_DURATION = (origin, destination, mode = "driving") => {
  if (!origin || !destination) throw new Error("No address specified!");
  if (origin.map) return origin.map(o => GOOGLEMAPS_DURATION(o, destination, mode));
  const key   = ["duration", origin, destination, mode].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { routes: [data] = [] } = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination).setMode(mode).getDirections();
  if (!data) throw new Error("No route found!");
  const { legs: [{ duration: { text: time } } = {}] = [] } = data;
  _mapsSetCache(key, time);
  return time;
};

const GOOGLEMAPS_DISTANCE = (origin, destination, mode = "driving") => {
  if (!origin || !destination) throw new Error("No address specified!");
  if (origin.map) return origin.map(o => GOOGLEMAPS_DISTANCE(o, destination, mode));
  const key   = ["distance", origin, destination, mode].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { routes: [data] = [] } = Maps.newDirectionFinder().setOrigin(origin).setDestination(destination).setMode(mode).getDirections();
  if (!data) throw new Error("No route found!");
  const { legs: [{ distance: { text: distance } } = {}] = [] } = data;
  _mapsSetCache(key, distance);
  return distance;
};

const GOOGLEMAPS_LATLONG = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_LATLONG(a));
  const key   = ["latlong", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  const { geometry: { location: { lat, lng } } = {} } = data;
  const answer = `${lat}, ${lng}`;
  _mapsSetCache(key, answer);
  return answer;
};

const GOOGLEMAPS_ADDRESS = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_ADDRESS(a));
  const key   = ["address", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  const { formatted_address } = data;
  _mapsSetCache(key, formatted_address);
  return formatted_address;
};

const GOOGLEMAPS_REVERSEGEOCODE = (latitude, longitude) => {
  if (!latitude || !longitude) throw new Error("Lat/Lng not specified!");
  const key   = ["reverse", latitude, longitude].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { results: [data = {}] = [] } = Maps.newGeocoder().reverseGeocode(latitude, longitude);
  const { formatted_address } = data;
  if (!formatted_address) return "Address not found";
  _mapsSetCache(key, formatted_address);
  return formatted_address;
};

const GOOGLEMAPS_COUNTRY = (address) => {
  if (!address) throw new Error("No address specified!");
  if (address.map) return address.map(a => GOOGLEMAPS_COUNTRY(a));
  const key   = ["country", address].join(",");
  const value = _mapsGetCache(key);
  if (value !== null) return value;
  Utilities.sleep(150);
  const { results: [data = null] = [] } = Maps.newGeocoder().geocode(address);
  if (data === null) throw new Error("Address not found!");
  const [{ short_name, long_name } = {}] = data.address_components.filter(({ types: [level] }) => level === "country");
  if (!short_name) throw new Error("Country not found!");
  const answer = `${long_name} (${short_name})`;
  _mapsSetCache(key, answer);
  return answer;
};

function GET_ADDR_WITH_CACHE(lat, lng) {
  try { return GOOGLEMAPS_REVERSEGEOCODE(lat, lng); }
  catch (e) { console.error(`[GeoAddr API] Error (${lat}, ${lng}): ${e.message}`); return ""; }
}

function CALCULATE_DISTANCE_KM(origin, destination) {
  try {
    var distanceText = GOOGLEMAPS_DISTANCE(origin, destination, "driving");
    if (!distanceText) return "";
    var cleanStr = String(distanceText).replace(/,/g, "").replace(/[^0-9.]/g, "");
    var val = parseFloat(cleanStr);
    return isNaN(val) ? "" : val.toFixed(2);
  } catch (e) {
    console.error(`[GeoAddr API] Distance Error: ${e.message}`);
    return "";
  }
}
