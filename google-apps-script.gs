var TOTAL_SEATS = 20;
var DEFAULT_SHEET_NAME = '지원금증정';

var ALLOWED_SHEET_NAMES = {
  '실부담10%': true,
  '실부담0원': true,
  '지원금증정': true
};

var HEADERS = [
  '타임스탬프',
  '랜딩 버전',
  '기업명',
  '담당자 성명',
  '직책',
  '연락처',
  '이메일',
  '예상 수강 인원',
  '희망 일정',
  '희망 교육 주제',
  '희망 전문가'
];

function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
  var sheetName = getSheetName_(params);
  var sheet = getTargetSheet_(sheetName);
  ensureHeader_(sheet);

  var submissions = Math.max(0, sheet.getLastRow() - 1);
  var remaining = Math.max(0, TOTAL_SEATS - submissions);

  return json_({
    result: 'success',
    sheet: sheetName,
    total: TOTAL_SEATS,
    remaining: remaining
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    var data = parseBody_(e);
    var sheetName = getSheetName_(data);
    var sheet = getTargetSheet_(sheetName);
    ensureHeader_(sheet);

    sheet.appendRow([
      new Date(),
      data.landingVariant || sheetName,
      data.companyName || '',
      data.managerName || '',
      data.managerPosition || '',
      data.phone || '',
      data.email || '',
      data.headcount || '',
      data.preferredDate || '',
      data.topic || '',
      data.expert || ''
    ]);

    return json_({
      result: 'success',
      sheet: sheetName
    });
  } catch (err) {
    return json_({
      result: 'error',
      message: err && err.toString ? err.toString() : String(err)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignored) {}
  }
}

function getSheetName_(data) {
  var requested = data && data.sheetTab ? String(data.sheetTab).trim() : DEFAULT_SHEET_NAME;
  return ALLOWED_SHEET_NAMES[requested] ? requested : DEFAULT_SHEET_NAME;
}

function getTargetSheet_(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  return sheet;
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }

  var firstCell = sheet.getRange(1, 1).getValue();
  if (!firstCell) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  try {
    return JSON.parse(e.postData.contents);
  } catch (err) {
    return {};
  }
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
