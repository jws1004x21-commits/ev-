/**
 * 충전소 점검앱 v6 — Google Apps Script
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
 * 내 드라이브 / 전기차충전소점검_대전팀 / 점검사진 / 지역 / 시군 / 충전소 / 거점
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
  var payload = { ok: true, version: 'v6', message: '충전소 점검 서버 정상' };
  if (action === 'progress') {
    payload.progress = loadProgress_();
  }
  if (action === 'record') {
    payload.record = loadRecord_(e.parameter || {});
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
      chargers: Number(values[i][7]) || 0
    });
  }
  return out;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'complete';

    if (action === 'ping') {
      return json_({ ok: true, version: 'v6' });
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
      folder.getUrl()
    ];
    var existingRow = findRow_(sheet, data.station, data.zone);
    if (action === 'update' || existingRow) {
      if (existingRow) {
        if (!data.date) rowValues[0] = sheet.getRange(existingRow, 1).getDisplayValue();
        rowValues[12] = photoCount;
        sheet.getRange(existingRow, 1, 1, 14).setValues([rowValues]);
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
    var values = sheet.getRange(row, 1, 1, 14).getDisplayValues()[0];
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
  }
  rec.photos = listPhotos_({
    region: rec.region || p.region,
    city: rec.city || p.city,
    station: station,
    zone: zone
  });
  return rec;
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
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '점검일', '점검시간', '점검자', '지역', '시군', '충전소명', '거점',
      '충전기수', '제조사', 'GPS', '메모', '주소', '사진수', '사진폴더'
    ]);
    sheet.getRange(1, 1, 1, 14).setFontWeight('bold');
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
  var folder = ensureZoneFolder_(data);
  var raw = String(data.data);
  var comma = raw.indexOf(',');
  if (comma >= 0) raw = raw.substring(comma + 1);
  var bytes = Utilities.base64Decode(raw);
  var filename = sanitize_(data.filename || ((data.label || '사진') + '.jpg'));
  if (!/\.jpe?g$/i.test(filename)) filename += '.jpg';
  var file = folder.createFile(Utilities.newBlob(bytes, MimeType.JPEG, filename));
  return { url: file.getUrl(), folderUrl: folder.getUrl() };
}

function ensureZoneFolder_(data) {
  var root = getOrCreateFolder_(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  var photos = getOrCreateFolder_(root, PHOTO_ROOT_NAME);
  var region = getOrCreateFolder_(photos, data.region || '미지정');
  var city = getOrCreateFolder_(region, data.city || '미지정');
  var station = getOrCreateFolder_(city, data.station || '미지정');
  return getOrCreateFolder_(station, data.zone || '거점');
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
