/**
 * DatePact - GAS 백엔드 (심플 버전: 모임 1개 = 배포 1개)
 * 스프레드시트에 바인딩해서 사용하세요.
 * (Google Sheets에서 확장 프로그램 > Apps Script 로 열면 자동 바인딩됩니다)
 *
 * 새로운 약속을 시작할 땐 이 스프레드시트 + 앱스크립트 프로젝트를 통째로 복사해서
 * 새로 배포하세요. 그러면 약속끼리 서로 안 섞여요.
 */

var PARTICIPANTS_SHEET = 'Participants';
var META_SHEET = 'Meta';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('DatePact')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PARTICIPANTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PARTICIPANTS_SHEET);
    sheet.appendRow([
      'Name', 'WantsToHang', 'Dates', 'Times', 'Places',
      'Activities', 'FoodCandidates', 'FoodResult', 'UpdatedAt'
    ]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getMetaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(META_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(META_SHEET);
    sheet.appendRow(['Status', 'FinalDate', 'FinalTime', 'FinalPlace', 'FinalFood', 'ClosedAt']);
    sheet.appendRow(['open', '', '', '', '', '']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/** 모든 참여자의 답변을 가져옵니다. (다같이 비교보기 탭에서 사용) */
function getAllParticipants() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];

  var data = sheet.getRange(2, 1, last - 1, 9).getValues();
  var result = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = row[0];
    if (!name) continue;

    result.push({
      name: name,
      wantsToHang: row[1],
      dates: safeParse_(row[2], []),
      times: safeParse_(row[3], {}),
      places: safeParse_(row[4], []),
      activities: safeParse_(row[5], []),
      foodCandidates: safeParse_(row[6], []),
      foodResult: row[7] || null,
      updatedAt: row[8]
    });
  }
  return result;
}

/**
 * 한 사람의 답변을 저장(추가 또는 덮어쓰기)합니다.
 * entry = { name, wantsToHang, dates, times, places, activities, foodCandidates, foodResult }
 */
function saveParticipant(entry) {
  if (!entry || !entry.name) {
    throw new Error('이름이 필요해요');
  }

  var sheet = getSheet_();
  var last = sheet.getLastRow();
  var rowIndex = -1;

  if (last >= 2) {
    var names = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (names[i][0] === entry.name) { rowIndex = i + 2; break; }
    }
  }

  var rowValues = [
    entry.name,
    entry.wantsToHang || '',
    JSON.stringify(entry.dates || []),
    JSON.stringify(entry.times || {}),
    JSON.stringify(entry.places || []),
    JSON.stringify(entry.activities || []),
    JSON.stringify(entry.foodCandidates || []),
    entry.foodResult || '',
    new Date().toISOString()
  ];

  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  }

  return { success: true };
}

/** 한 사람의 답변을 삭제합니다. */
function deleteParticipant(name) {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return { success: false, message: '참여자를 찾을 수 없어요' };

  var names = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (names[i][0] === name) {
      sheet.deleteRow(i + 2);
      return { success: true };
    }
  }
  return { success: false, message: '참여자를 찾을 수 없어요' };
}

/** 약속 상태(진행중/마감) + 확정 정보 */
function getMeta() {
  var sheet = getMetaSheet_();
  var row = sheet.getRange(2, 1, 1, 6).getValues()[0];
  return {
    status: row[0] || 'open',
    finalDate: row[1] || '',
    finalTime: row[2] || '',
    finalPlace: row[3] || '',
    finalFood: row[4] || '',
    closedAt: row[5] || ''
  };
}

/** 약속 마감 + 최종 확정 내용 저장 */
function closeMeetup(final) {
  var sheet = getMetaSheet_();
  sheet.getRange(2, 1, 1, 6).setValues([[
    'closed',
    (final && final.date) || '',
    (final && final.time) || '',
    (final && final.place) || '',
    (final && final.food) || '',
    new Date().toISOString()
  ]]);
  return { success: true };
}

/** 마감된 약속을 다시 열기 */
function reopenMeetup() {
  var sheet = getMetaSheet_();
  sheet.getRange(2, 1).setValue('open');
  return { success: true };
}

/**
 * 제출용 통합 엔드포인트: 마감 여부 확인 + 저장을 한 번의 호출로 처리해서 왕복을 줄입니다.
 */
function submitAnswer(entry) {
  var meta = getMeta();
  if (meta.status === 'closed') {
    return { closed: true, meta: meta };
  }
  saveParticipant(entry);
  return { closed: false, meta: meta };
}

/**
 * 비교 탭용 통합 엔드포인트: meta + 참여자 목록을 한 번에 반환합니다.
 */
function getCompareData() {
  return {
    meta: getMeta(),
    participants: getAllParticipants()
  };
}

function safeParse_(value, fallback) {
  if (!value) return fallback;
  try {
    return JSON.parse(value);
  } catch (e) {
    return fallback;
  }
}
