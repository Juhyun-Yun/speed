/**
 * ⚡ 속력의 숲 - Apps Script 웹앱 (HTML + 데이터 프리로드 + 캐시)
 *
 * ───────────────────────────────────────────────
 * 🚀 빠른 로딩 비법
 * ───────────────────────────────────────────────
 * - HTML 페이지 응답 안에 학생 명단 JSON을 미리 심음 → 별도 fetch 없이 즉시 표시
 * - CacheService로 시트 읽기 결과를 10분 캐시 → 두 번째부터 시트 안 읽고 응답
 * - getRange로 A, B 두 열만 읽음 → getDataRange보다 빠름
 *
 * ───────────────────────────────────────────────
 * 📋 시트 양식 (1행 헤더)
 * ───────────────────────────────────────────────
 *   |  A   |   B    |
 * 1 | 번호 |  이름  |
 * 2 |  1   | 김민수 |
 * 3 |  2   | 이지은 |
 * ...
 *
 * ───────────────────────────────────────────────
 * 🔄 명단을 즉시 갱신하려면
 * ───────────────────────────────────────────────
 * 시트 수정 후 캐시(최대 10분)가 만료될 때까지 기다리거나,
 * Apps Script 편집기에서 clearCache() 함수를 한 번 실행.
 */

function doGet(e) {
  // ① API 모드: JSON
  if (e && e.parameter && e.parameter.api === 'students') {
    return ContentService
      .createTextOutput(JSON.stringify(getStudentsCached()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ② 기본 모드: HTML 페이지 + 학생 명단 프리로드
  var template = HtmlService.createTemplateFromFile('index');
  // 단일 인용부호는 안전하게 이스케이프 (이름에 ' 가 있어도 깨지지 않게)
  template.preloadedJson = JSON.stringify(getStudentsCached())
    .replace(/'/g, '\\u0027');
  return template.evaluate()
    .setTitle('속력의 숲')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getStudentsCached() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('students_v2');
  if (cached) {
    return JSON.parse(cached);
  }
  var data = readStudentsFromSheet();
  cache.put('students_v2', JSON.stringify(data), 600);  // 10분 캐시
  return data;
}

function readStudentsFromSheet() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, count: 0, students: [] };

    // A, B 두 열만 빠르게 읽음
    var values = sheet.getRange(1, 1, lastRow, 2).getValues();
    var students = [];
    for (var i = 1; i < values.length; i++) {
      var num = values[i][0];
      var name = values[i][1];
      var hasNum = (num !== '' && num !== null && num !== undefined);
      var hasName = (name !== '' && name !== null && name !== undefined);
      if (!hasNum && !hasName) continue;
      students.push({
        number: hasNum ? String(num).trim() : '',
        name: hasName ? String(name).trim() : ''
      });
    }
    return { ok: true, count: students.length, students: students };
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

/**
 * 캐시 즉시 비우기 (시트 수정 후 바로 반영하고 싶을 때 실행)
 */
function clearCache() {
  CacheService.getScriptCache().remove('students_v2');
  Logger.log('✅ 캐시 비웠습니다. 다음 호출부터 시트를 새로 읽습니다.');
}

/**
 * 동작 확인용
 */
function testReadSheet() {
  var result = readStudentsFromSheet();
  Logger.log(JSON.stringify(result, null, 2));
  Logger.log('학생 수: ' + (result.students ? result.students.length : 0));
}
