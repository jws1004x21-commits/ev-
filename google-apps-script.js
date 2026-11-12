/**
 * 충전소 점검앱 v12 — Google Apps Script
 *
 * 사용법:
 * 1. 구글 시트 → 확장 프로그램 → Apps Script
 * 2. 기존 코드 전부 지우고 이 파일 내용을 붙여넣기
 * 3. 저장 (Ctrl+S)
 * 4. 배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포
 * 5. 권한 요청이 나오면 고급 → (안전하지 않은 페이지로 이동) → 허용
 *    드라이브 저장 권한이 추가로 필요합니다.
 *
 * 저장 위치:
 * 내 드라이브 / 전기차충전소점검_대전팀 / 점검사진 / 지역 / 시군 / 충전소 / 거점1_장소이름
 */

var ROOT_FOLDER_NAME = '전기차충전소점검_대전팀';
var PHOTO_ROOT_NAME = '점검사진';

function setupDrive() {
  var root = DriveApp.getRootFolder();
  var folder = getOrCreateFolder_(root, ROOT_FOLDER_NAME);
  var photos = getOrCreateFolder_(folder, PHOTO_ROOT_NAME);
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;
  var p = e.parameter || {};
  var payload = { ok: true, version: 'v12', message: '충전소 점검 서버 정상' };
  if (action === 'progress') {
    payload.progress = loadProgress_();
    payload.locks = loadLocks_();
  }
  if (action === 'record') {
    payload.record = loadRecord_(p);
  }
  if (action === 'file') {
    payload.file = loadFileData_(p);
  }
  if (action === 'count') {
    payload.count = folderFileCount_(ensureZoneFolder_(p));
  }
  if (action === 'lock') {
    payload.lock = takeLock_(p);
  }
  if (action === 'unlock') {
    payload.unlocked = releaseLock_(p);
  }
  if (action === 'heartbeat') {
    payload.lock = refreshLock_(p);
  }
  if (action === 'geocode') {
    payload.geos = geocodeAddresses_(p.addresses || '');
  }
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function loadProgress_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < values.length; i++) {
    var region = String(values[i][3] || '');
    var station = String(values[i][5] || '');
    if (!station || region === '테스트' || region === '집테스트') continue;
    if (station.indexOf('테스트') >= 0) continue;
    var rawDate = values[i][0];
    var dateStr = '';
    if (rawDate instanceof Date) {
      dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      dateStr = String(rawDate || '').substring(0, 10);
    }
    out.push({
      region: region,
      city: String(values[i][4] || ''),
      station: station,
      zone: String(values[i][6] || ''),
      inspector: String(values[i][2] || ''),
      date: dateStr,
      chargers: Number(values[i][7]) || 0,
      ground: String(values[i][16] || ''),
      v220: String(values[i][14] || ''),
      v380: String(values[i][15] || ''),
      mccb: String(values[i][18] || ''),
      meterNo: String(values[i][19] || ''),
      chargerNo: String(values[i][20] || ''),
      memo: String(values[i][10] || '')
    });
  }
  return out;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'complete';

    if (action === 'ping') {
      return json_({ ok: true, version: 'v11' });
    }

    if (action === 'photo') {
      var saved = savePhoto_(data);
      return json_({ ok: true, url: saved.url, folderUrl: saved.folderUrl });
    }

    if (action === 'deletePhoto') {
      return json_({ ok: true, deleted: deletePhoto_(data) });
    }

    var folder = ensureZoneFolder_(data);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeader_(sheet);
    var photoCount = folderFileCount_(folder);
    var rowValues = [
      data.date || '',
      data.time || '',
      data.inspector || '',
      data.region || '',
      data.city || '',
      data.station || '',
      data.zone || '',
      data.chargers || '',
      data.maker || '',
      data.gps || '',
      data.memo || '',
      data.address || '',
      data.photoCount || photoCount,
      folder.getUrl(),
      data.v220 || '',
      data.v380 || '',
      data.ground || '',
      data.insul || '',
      data.mccb || '',
      data.meterNo || '',
      data.chargerNo || ''
    ];
    var existingRow = findRow_(sheet, data.station, data.zone);
    if (action === 'update' || existingRow) {
      if (existingRow) {
        if (!data.date) rowValues[0] = sheet.getRange(existingRow, 1).getDisplayValue();
        rowValues[12] = photoCount;
        sheet.getRange(existingRow, 1, 1, rowValues.length).setValues([rowValues]);
        return json_({ ok: true, updated: true, folderUrl: folder.getUrl(), photoCount: photoCount });
      }
    }
    if (isDuplicateComplete_(sheet, data)) {
      return json_({ ok: true, duplicate: true, folderUrl: folder.getUrl() });
    }
    rowValues[12] = photoCount;
    sheet.appendRow(rowValues);

    if (data.photos && data.photos.length) {
      for (var i = 0; i < data.photos.length; i++) {
        savePhoto_({
          region: data.region,
          city: data.city,
          station: data.station,
          zone: data.zone,
          label: data.photos[i].label || data.photos[i].id || '사진',
          filename: data.photos[i].name || ('photo_' + (i + 1) + '.jpg'),
          data: data.photos[i].data
        });
      }
    }

    return json_({ ok: true, folderUrl: folder.getUrl() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function findRow_(sheet, station, zone) {
  station = String(station || '');
  zone = String(zone || '');
  if (!station || !zone) return 0;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][5]) === station && String(values[i][6]) === zone) return i + 1;
  }
  return 0;
}

function isDuplicateComplete_(sheet, data) {
  return findRow_(sheet, data.station, data.zone) > 0;
}

function folderFileCount_(folder) {
  var n = 0;
  var files = folder.getFiles();
  while (files.hasNext()) {
    files.next();
    n++;
  }
  return n;
}

function loadRecord_(p) {
  var station = String(p.station || '');
  var zone = String(p.zone || '');
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = findRow_(sheet, station, zone);
  var rec = { found: false, station: station, zone: zone, photos: [] };
  if (row) {
    var lastCol = Math.max(sheet.getLastColumn(), 14);
    var values = sheet.getRange(row, 1, 1, lastCol).getDisplayValues()[0];
    rec.found = true;
    rec.date = values[0];
    rec.time = values[1];
    rec.inspector = values[2];
    rec.region = values[3];
    rec.city = values[4];
    rec.memo = values[10];
    rec.address = values[11];
    rec.photoCount = values[12];
    rec.folderUrl = values[13];
    rec.v220 = values[14] || '';
    rec.v380 = values[15] || '';
    rec.ground = values[16] || '';
    rec.insul = values[17] || '';
    rec.mccb = values[18] || '';
    rec.meterNo = values[19] || '';
    rec.chargerNo = values[20] || '';
  }
  rec.photos = listPhotos_({
    region: rec.region || p.region,
    city: rec.city || p.city,
    station: station,
    zone: zone,
    zoneFolder: p.zoneFolder || '',
    zoneLabel: p.zoneLabel || ''
  });
  return rec;
}

function loadFileData_(p) {
  var file = null;
  var id = String(p.id || '');
  if (id) {
    try { file = DriveApp.getFileById(id); } catch (err) { file = null; }
  }
  if (!file && p.filename) {
    var folder = ensureZoneFolder_(p);
    var found = folder.getFilesByName(sanitize_(p.filename));
    if (found.hasNext()) file = found.next();
  }
  if (!file) return { ok: false };
  var blob = file.getBlob();
  var bytes = blob.getBytes();
  if (bytes.length > 1500000) return { ok: false, error: 'too big', name: file.getName() };
  return {
    ok: true,
    name: file.getName(),
    id: file.getId(),
    mime: blob.getContentType() || 'image/jpeg',
    data: Utilities.base64Encode(bytes)
  };
}

function listPhotos_(data) {
  var folder = ensureZoneFolder_(data);
  var files = folder.getFiles();
  var out = [];
  while (files.hasNext()) {
    var file = files.next();
    out.push({ name: file.getName(), id: file.getId() });
  }
  out.sort(function (a, b) {
    return String(a.name).localeCompare(String(b.name), 'ko', { numeric: true, sensitivity: 'base' });
  });
  return out;
}

function ensureHeader_(sheet) {
  var headers = [
    '점검일', '점검시간', '점검자', '지역', '시군', '충전소명', '거점',
    '충전기수', '제조사', 'GPS', '메모', '주소', '사진수', '사진폴더',
    '전압220', '전압380', '접지저항', '절연저항',
    'MCCB용량', '계량기번호', '충전기번호'
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var row1 = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  if (String(row1[14] || '') !== '전압220') {
    sheet.getRange(1, 15, 1, 4).setValues([['전압220', '전압380', '접지저항', '절연저항']]);
    sheet.getRange(1, 15, 1, 4).setFontWeight('bold');
  }
  if (String(row1[18] || '') !== 'MCCB용량') {
    sheet.getRange(1, 19, 1, 3).setValues([['MCCB용량', '계량기번호', '충전기번호']]);
    sheet.getRange(1, 19, 1, 3).setFontWeight('bold');
  }
}

function deletePhoto_(data) {
  var folder = ensureZoneFolder_(data);
  var name = sanitize_(data.filename || '');
  if (!name) return 0;
  var files = folder.getFilesByName(name);
  var count = 0;
  while (files.hasNext()) {
    files.next().setTrashed(true);
    count++;
  }
  return count;
}

function savePhoto_(data) {
  if (!data.data) {
    throw new Error('사진 데이터가 없습니다.');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    var folder = ensureZoneFolderUnlocked_(data);
    var raw = String(data.data);
    var comma = raw.indexOf(',');
    if (comma >= 0) raw = raw.substring(comma + 1);
    var bytes = Utilities.base64Decode(raw);
    var filename = sanitize_(data.filename || ((data.label || '사진') + '.jpg'));
    if (!/\.jpe?g$/i.test(filename)) filename += '.jpg';
    var existing = folder.getFilesByName(filename);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
    var file = folder.createFile(Utilities.newBlob(bytes, MimeType.JPEG, filename));
    file.setName(filename);
    forceIntoFolder_(file, folder);
    return { url: file.getUrl(), folderUrl: folder.getUrl() };
  } finally {
    lock.releaseLock();
  }
}

function zoneFolderName_(data) {
  var explicit = String(data.zoneFolder || '').trim();
  if (explicit) return explicit;
  var zone = String(data.zone || '').trim();
  var matched = zone.match(/^거점\s*(\d+)\s*:\s*(.*)$/);
  if (matched) {
    var tail = String(matched[2] || '').trim();
    return tail ? ('거점' + matched[1] + '_' + tail) : ('거점' + matched[1]);
  }
  return zone || '거점';
}

function childFolder_(parent, name) {
  name = sanitize_(name);
  if (!name) return null;
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function moveFolderFiles_(src, dest) {
  if (!src || !dest || src.getId() === dest.getId()) return;
  var files = src.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (dest.getFilesByName(file.getName()).hasNext()) continue;
    file.moveTo(dest);
  }
}

function forceIntoFolder_(file, folder) {
  var parents = file.getParents();
  var inDest = false;
  var extra = false;
  while (parents.hasNext()) {
    var parent = parents.next();
    if (parent.getId() === folder.getId()) inDest = true;
    else extra = true;
  }
  if (!inDest || extra) file.moveTo(folder);
}

function ensureZoneFolder_(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    return ensureZoneFolderUnlocked_(data);
  } finally {
    lock.releaseLock();
  }
}

function ensureZoneFolderUnlocked_(data) {
  var root = getOrCreateFolder_(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  var photos = getOrCreateFolder_(root, PHOTO_ROOT_NAME);
  var region = getOrCreateFolder_(photos, data.region || '미지정');
  var city = getOrCreateFolder_(region, data.city || '미지정');
  var station = getOrCreateFolder_(city, data.station || '미지정');
  var wanted = sanitize_(zoneFolderName_(data)) || '거점';
  var dest = childFolder_(station, wanted);
  if (!dest) dest = station.createFolder(wanted);
  absorbLegacyFolder_(station, dest, wanted, data.zone);
  absorbLegacyFolder_(station, dest, wanted, data.zoneLabel);
  return dest;
}

function absorbLegacyFolder_(station, dest, wanted, rawName) {
  var legacy = sanitize_(String(rawName || ''));
  if (!legacy || legacy === wanted) return;
  var oldFolder = childFolder_(station, legacy);
  if (oldFolder) moveFolderFiles_(oldFolder, dest);
}

function getOrCreateFolder_(parent, name) {
  name = sanitize_(name) || '미지정';
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

function sanitize_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 80);
}

var LOCK_MS = 2 * 60 * 60 * 1000;

function lockKey_(p) {
  return 'lock:' + [p.city || '', p.station || '', p.zone || ''].join('|');
}

function takeLock_(p) {
  var props = PropertiesService.getScriptProperties();
  var key = lockKey_(p);
  var now = Date.now();
  var inspector = String(p.inspector || '').trim();
  if (!inspector || !p.station || !p.zone) {
    return { ok: false, reason: 'missing' };
  }
  var raw = props.getProperty(key);
  if (raw) {
    try {
      var lock = JSON.parse(raw);
      if (lock.expires > now && lock.inspector && lock.inspector !== inspector) {
        return { ok: false, lockedBy: lock.inspector, started: lock.started };
      }
    } catch (err) {}
  }
  props.setProperty(key, JSON.stringify({
    inspector: inspector,
    started: now,
    expires: now + LOCK_MS,
    city: p.city || '',
    station: p.station || '',
    zone: p.zone || ''
  }));
  return { ok: true, inspector: inspector };
}

function refreshLock_(p) {
  var props = PropertiesService.getScriptProperties();
  var key = lockKey_(p);
  var raw = props.getProperty(key);
  var inspector = String(p.inspector || '').trim();
  var now = Date.now();
  if (!raw) return takeLock_(p);
  try {
    var lock = JSON.parse(raw);
    if (lock.inspector === inspector) {
      lock.expires = now + LOCK_MS;
      props.setProperty(key, JSON.stringify(lock));
      return { ok: true, inspector: inspector };
    }
    if (lock.expires > now) {
      return { ok: false, lockedBy: lock.inspector, started: lock.started };
    }
  } catch (err) {}
  return takeLock_(p);
}

function releaseLock_(p) {
  var props = PropertiesService.getScriptProperties();
  var key = lockKey_(p);
  var raw = props.getProperty(key);
  if (!raw) return true;
  try {
    var lock = JSON.parse(raw);
    var inspector = String(p.inspector || '').trim();
    if (lock.inspector && inspector && lock.inspector !== inspector && lock.expires > Date.now()) {
      return false;
    }
  } catch (err) {}
  props.deleteProperty(key);
  return true;
}

function loadLocks_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var now = Date.now();
  var out = [];
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('lock:') !== 0) return;
    try {
      var lock = JSON.parse(all[k]);
      if (!lock.expires || lock.expires < now) {
        props.deleteProperty(k);
        return;
      }
      out.push({
        city: lock.city || '',
        station: lock.station || '',
        zone: lock.zone || '',
        inspector: lock.inspector || '',
        started: lock.started || 0
      });
    } catch (err) {}
  });
  return out;
}

function geocodeAddresses_(raw) {
  var out = {};
  var list = String(raw || '').split(/\n+/);
  var geocoder = Maps.newGeocoder().setLanguage('ko');
  for (var i = 0; i < list.length && i < 40; i++) {
    var address = String(list[i] || '').trim();
    if (!address || out[address]) continue;
    try {
      var res = geocoder.geocode(address);
      if (res && res.status === 'OK' && res.results && res.results[0] && res.results[0].geometry) {
        var loc = res.results[0].geometry.location;
        out[address] = { lat: loc.lat, lng: loc.lng };
      }
      Utilities.sleep(80);
    } catch (err) {}
  }
  return out;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
