/**
 * 충전소 점검앱 v3 — Google Apps Script
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
  Logger.log('드라이브 권한 OK');
  Logger.log(folder.getUrl());
}

function doGet(e) {
  var payload = { ok: true, version: 'v3', message: '충전소 점검 서버 정상' };
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(payload);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || 'complete';

    if (action === 'ping') {
      return json_({ ok: true, version: 'v3' });
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
    sheet.appendRow([
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
      data.photoCount || 0,
      folder.getUrl()
    ]);

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
