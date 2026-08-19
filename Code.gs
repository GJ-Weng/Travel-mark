/**
 * ============================================================
 *  旅遊清單 Trip Tags — Google Apps Script 後端
 *  用途：作為 Google 試算表的 Web API，供前端 PWA 讀寫資料
 *
 *  安裝方式：
 *  1. 開一份新的 Google 試算表（表格內容不用管，腳本會自動建立分頁）
 *  2. 「擴充功能」→「Apps Script」，把這個檔案內容整個貼進去（取代預設內容）
 *  3. 上方「部署」→「新增部署作業」
 *       類型選「網頁應用程式」
 *       執行身分：我 (你的帳號)
 *       誰可以存取：任何人
 *  4. 部署後複製「網頁應用程式 URL」（結尾是 /exec），貼到前端 App 的
 *     「設定」→「雲端同步」欄位
 *  5. 之後若修改此腳本，記得「管理部署作業」→ 編輯 → 產生「新版本」
 *     才會生效（URL 不變）
 * ============================================================
 */

var SHEET_ITEMS = 'Items';
var SHEET_COUNTRIES = 'Countries';
var ITEM_HEADERS = ['id', 'country', 'category', 'name', 'address', 'note', 'done', 'favorite', 'createdAt', 'updatedAt', 'deleted'];

/* ---------------- Sheet helpers ---------------- */

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(name, headers) {
  var ss = getSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function itemsSheet_() { return getOrCreateSheet_(SHEET_ITEMS, ITEM_HEADERS); }
function countriesSheet_() { return getOrCreateSheet_(SHEET_COUNTRIES, ['name']); }

function truthy_(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }

/* ---------------- Read ---------------- */

function readItems_() {
  var sheet = itemsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, ITEM_HEADERS.length).getValues();
  var items = [];
  values.forEach(function (row) {
    var obj = {};
    ITEM_HEADERS.forEach(function (h, i) { obj[h] = row[i]; });
    if (!obj.id) return;
    if (truthy_(obj.deleted)) return;
    obj.done = truthy_(obj.done);
    obj.favorite = truthy_(obj.favorite);
    obj.createdAt = Number(obj.createdAt) || 0;
    obj.updatedAt = Number(obj.updatedAt) || 0;
    delete obj.deleted;
    items.push(obj);
  });
  return items;
}

function readCountries_() {
  var sheet = countriesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var out = [];
  values.forEach(function (r) { if (r[0]) out.push(String(r[0])); });
  return out;
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

/* ---------------- Write ---------------- */

function upsertItem_(payload) {
  payload = payload || {};
  var sheet = itemsSheet_();
  if (!payload.id) payload.id = Utilities.getUuid();
  var now = Date.now();
  var row = findRowById_(sheet, payload.id);
  var existing = null;
  if (row > -1) {
    var rowValues = sheet.getRange(row, 1, 1, ITEM_HEADERS.length).getValues()[0];
    existing = {};
    ITEM_HEADERS.forEach(function (h, i) { existing[h] = rowValues[i]; });
  }
  var merged = {
    id: payload.id,
    country: payload.country != null ? payload.country : (existing ? existing.country : ''),
    category: payload.category != null ? payload.category : (existing ? existing.category : 'other'),
    name: payload.name != null ? payload.name : (existing ? existing.name : ''),
    address: payload.address != null ? payload.address : (existing ? existing.address : ''),
    note: payload.note != null ? payload.note : (existing ? existing.note : ''),
    done: payload.done != null ? !!payload.done : (existing ? truthy_(existing.done) : false),
    favorite: payload.favorite != null ? !!payload.favorite : (existing ? truthy_(existing.favorite) : false),
    createdAt: existing ? (Number(existing.createdAt) || now) : (Number(payload.createdAt) || now),
    updatedAt: now,
    deleted: false
  };
  var rowArr = ITEM_HEADERS.map(function (h) { return merged[h]; });
  if (row > -1) {
    sheet.getRange(row, 1, 1, ITEM_HEADERS.length).setValues([rowArr]);
  } else {
    sheet.appendRow(rowArr);
  }
  delete merged.deleted;
  return merged;
}

function deleteItem_(id) {
  var sheet = itemsSheet_();
  var row = findRowById_(sheet, id);
  if (row > -1) {
    sheet.getRange(row, ITEM_HEADERS.indexOf('deleted') + 1, 1, 1).setValue(true);
    sheet.getRange(row, ITEM_HEADERS.indexOf('updatedAt') + 1, 1, 1).setValue(Date.now());
  }
  return { id: id, deleted: true };
}

function upsertCountry_(name) {
  name = String(name || '').trim();
  if (!name) return readCountries_();
  var sheet = countriesSheet_();
  var existing = readCountries_();
  if (existing.indexOf(name) === -1) {
    sheet.appendRow([name]);
  }
  return readCountries_();
}

function deleteCountry_(name) {
  var sheet = countriesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return readCountries_();
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(name)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  return readCountries_();
}

function buildState_() {
  return {
    items: readItems_(),
    countries: readCountries_(),
    serverTime: Date.now()
  };
}

/* ---------------- HTTP entry points ---------------- */

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (err) {}
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'ping') {
      return jsonOut_({ success: true, pong: true, serverTime: Date.now() });
    }
    var out = buildState_();
    out.success = true;
    return jsonOut_(out);
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (err) {}
  try {
    var body = {};
    try {
      body = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonOut_({ success: false, error: 'Invalid JSON body' });
    }
    var actions = Array.isArray(body.actions) ? body.actions : [];
    var results = [];
    actions.forEach(function (a) {
      try {
        if (a.type === 'upsert-item') {
          results.push({ type: a.type, clientId: a.clientId, ok: true, data: upsertItem_(a.payload || {}) });
        } else if (a.type === 'delete-item') {
          results.push({ type: a.type, clientId: a.clientId, ok: true, data: deleteItem_(a.id) });
        } else if (a.type === 'upsert-country') {
          results.push({ type: a.type, clientId: a.clientId, ok: true, data: upsertCountry_(a.name) });
        } else if (a.type === 'delete-country') {
          results.push({ type: a.type, clientId: a.clientId, ok: true, data: deleteCountry_(a.name) });
        } else if (a.type === 'replace-all') {
          // full overwrite: used by "上傳本機資料到雲端（覆蓋雲端）"
          replaceAll_(a.items || [], a.countries || []);
          results.push({ type: a.type, clientId: a.clientId, ok: true });
        } else {
          results.push({ type: a.type, clientId: a.clientId, ok: false, error: 'unknown action type' });
        }
      } catch (actionErr) {
        results.push({ type: a.type, clientId: a.clientId, ok: false, error: String(actionErr) });
      }
    });
    var out = buildState_();
    out.success = true;
    out.results = results;
    return jsonOut_(out);
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function replaceAll_(items, countries) {
  var ss = getSS_();
  var oldItems = ss.getSheetByName(SHEET_ITEMS);
  if (oldItems) ss.deleteSheet(oldItems);
  var oldCountries = ss.getSheetByName(SHEET_COUNTRIES);
  if (oldCountries) ss.deleteSheet(oldCountries);
  var itemSheet = getOrCreateSheet_(SHEET_ITEMS, ITEM_HEADERS);
  var now = Date.now();
  var rows = (items || []).map(function (it) {
    return [
      it.id || Utilities.getUuid(), it.country || '', it.category || 'other', it.name || '',
      it.address || '', it.note || '', !!it.done, !!it.favorite,
      Number(it.createdAt) || now, Number(it.updatedAt) || now, false
    ];
  });
  if (rows.length) itemSheet.getRange(2, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

  var countrySheet = getOrCreateSheet_(SHEET_COUNTRIES, ['name']);
  var uniqueCountries = [];
  (countries || []).forEach(function (c) { if (c && uniqueCountries.indexOf(c) === -1) uniqueCountries.push(c); });
  if (uniqueCountries.length) countrySheet.getRange(2, 1, uniqueCountries.length, 1).setValues(uniqueCountries.map(function (c) { return [c]; }));
}
