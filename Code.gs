/**
 * ============================================================
 * © 2026 GEG화성 (깊이 e끌림). All rights reserved.
 *
 * 본 코드는 「저작권법」상 보호받는 저작물입니다.
 * - 복제권(제16조)·공중송신권(제18조)·배포권(제20조)은
 *   저작권자에게 있습니다.
 * - 정식 경로로 받은 이용자라도 코드의 무단 복제·재배포·
 *   재판매·리브랜딩은 허용되지 않습니다.
 * - 무단 이용 시 「저작권법」 제136조(5년 이하 징역 또는
 *   5천만 원 이하 벌금) 및 제125조(손해배상) 적용 대상이
 *   될 수 있습니다.
 * - 이용 문의: bacusiki777@gmail.com, for2102@jimj.kr
 * ============================================================
 */

// 빌드 서명
const _BUILD_SIG = 'GEGHS-DEEPE-2026';

// 출처 확인용 함수
function getBuildInfo() {
  return {
    sig: _BUILD_SIG,
    owner: 'GEG화성 (깊이 e끌림)',
    year: 2026
  };
}

/**
 * 달려라, 달려! — 누가 가장 빠를까? · Apps Script 백엔드
 *
 * 역할:
 *  1) 학생 명단(번호·이름·모둠) 탭을 읽어 앱에 전달
 *  2) 같은 모둠 학생끼리 실시간 경주가 되도록 경주 상태 동기화
 *  3) 경주 결과를 시트에 기록(학생별 최신 기록 + 시간순 기록 탭)
 *  4) 페이지 첫 로드 시 학생 명단을 HTML에 미리 주입(preloadedJson)
 *
 * 탭 구성:
 *  - '학생명단' : A열 번호 / B열 이름 / C열 모둠 / D~F열 최신 기록
 *  - '경주기록' : 경주 결과가 시간순으로 쌓이는 탭
 *  - '사용 설명' : 선생님용 안내 탭(항상 첫 번째)
 */

// ──────────────────────────────────────────────────────────────
// 상수
// ──────────────────────────────────────────────────────────────
const STUDENT_SHEET_NAME = '학생명단';   // 학생 명단 탭
const LOG_SHEET_NAME     = '경주기록';   // 경주 결과가 시간순으로 쌓이는 탭
const GUIDE_SHEET_NAME   = '사용 설명';   // 선생님 안내 탭

const APP_COLOR       = '#ff8a3d';   // 앱 대표색(주황)
const APP_COLOR_LIGHT = '#ffe0b3';   // 옅은 대표색(표 헤더 배경)
const APP_COLOR_DEEP  = '#d35a17';   // 진한 대표색(제목 강조)

const STUDENTS_CACHE_KEY = 'students_v2';
const STUDENTS_CACHE_SEC = 600;      // 명단 10분 캐시

// ──────────────────────────────────────────────────────────────
// 메뉴 (선생님 메뉴)
// ──────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('선생님 메뉴')
    .addItem('사용 설명 탭 만들기', 'setupGuideSheet')
    .addItem('학생 명단 탭 만들기(학생1~학생30)', 'setupStudentSheet')
    .addSeparator()
    .addItem('명단 새로고침(캐시 비우기)', 'clearCache')
    .addToUi();
}

// ──────────────────────────────────────────────────────────────
// HtmlService / API 라우팅
// ──────────────────────────────────────────────────────────────
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'race')    return handleRace_(p);
  if (p.api === 'record')  return handleRecord_(p);
  if (p.api === 'students') {
    return jsonOut_({ ok: true, students: getStudentsCached() });
  }

  // 그 외에는 HTML 페이지 서빙 (Apps Script HtmlService로 직접 열 때)
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.preloadedJson = JSON.stringify({ ok: true, students: getStudentsCached() })
    .replace(/'/g, '\\u0027');
  return tmpl.evaluate()
    .setTitle('달려라, 달려!')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// 일부 환경에서는 POST로 올 수 있으므로 동일 로직을 doPost에도 연결
function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.api === 'race')   return handleRace_(p);
  if (p.api === 'record') return handleRecord_(p);
  return jsonOut_({ ok: false, error: 'unknown api' });
}

// ──────────────────────────────────────────────────────────────
// 학생 명단 로드 (캐시)
// ──────────────────────────────────────────────────────────────
function getStudentsCached() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(STUDENTS_CACHE_KEY);
  if (cached) {
    try { return JSON.parse(cached); } catch (err) { /* 캐시 깨짐 → 재읽기 */ }
  }
  const data = loadStudents_();
  cache.put(STUDENTS_CACHE_KEY, JSON.stringify(data), STUDENTS_CACHE_SEC);
  return data;
}

function loadStudents_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(STUDENT_SHEET_NAME) || ss.getActiveSheet();
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 3).getValues();
  return rows
    .map(r => ({
      number: r[0] === '' || r[0] == null ? null : Number(r[0]),
      name: String(r[1] || '').trim(),
      group: r[2] === '' || r[2] == null ? '' : String(r[2]).trim()
    }))
    .filter(s => s.name && !isSampleName_(s.name));
}

// 안내용 임시 이름/예시 행은 자동으로 걸러낸다.
function isSampleName_(name) {
  return /^학생\s*\d+$/.test(String(name).trim());
}

/**
 * 명단 새로고침 — 시트 수정 후 바로 반영하고 싶을 때 실행
 */
function clearCache() {
  CacheService.getScriptCache().remove(STUDENTS_CACHE_KEY);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('명단을 새로 불러오도록 했어요.', '완료', 3);
  } catch (e) { /* 편집기에서 실행한 경우 무시 */ }
}

/**
 * 동작 확인용
 */
function testReadSheet() {
  const students = loadStudents_();
  Logger.log(JSON.stringify(students, null, 2));
  Logger.log('학생 수: ' + students.length);
}

// ──────────────────────────────────────────────────────────────
// 경주 결과 기록
//   파라미터:
//     name:  학생 이름
//     group: 모둠(없으면 개인)
//     game:  'sametime' | 'samedist'
//     value: 기록 값 (sametime=간 거리 m, samedist=걸린 시간 초)
//     rank:  등수(선택)
// ──────────────────────────────────────────────────────────────
function handleRecord_(p) {
  const name = String(p.name || '').trim();
  const game = String(p.game || '').trim();
  if (!name || (game !== 'sametime' && game !== 'samedist')) {
    return jsonOut_({ ok: false, error: 'invalid record' });
  }
  const group = String(p.group || '').trim();
  const value = Number(p.value);
  const rank  = p.rank == null || p.rank === '' ? '' : Number(p.rank);
  if (!Number.isFinite(value)) {
    return jsonOut_({ ok: false, error: 'invalid value' });
  }

  const lock = LockService.getScriptLock();
  try { lock.waitLock(3000); } catch (e) { /* 무시 */ }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gameLabel = (game === 'sametime') ? '같은 시간 동안' : '같은 거리';
    const valueLabel = (game === 'sametime')
      ? (Math.round(value) + 'm')
      : (value.toFixed(1) + '초');
    const stamp = formatStamp_(new Date());

    // 1) 시간순 기록 탭에 한 줄 추가
    const log = ensureLogSheet_(ss);
    log.appendRow([stamp, name, group, gameLabel, valueLabel, rank]);

    // 2) 학생 명단 탭의 그 학생 줄에 최신 기록 저장
    writeLatestToRoster_(ss, name, game, valueLabel, stamp);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 무시 */ }
  }
}

// 명단 탭의 학생 줄(이름 옆)에 최신 기록을 저장한다.
function writeLatestToRoster_(ss, name, game, valueLabel, stamp) {
  const sh = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sh) return;
  ensureRosterHeaders_(sh);
  const last = sh.getLastRow();
  if (last < 2) return;
  const names = sh.getRange(2, 2, last - 1, 1).getValues();  // B열
  let rowIndex = -1;
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === name) { rowIndex = i + 2; break; }
  }
  if (rowIndex < 0) return;  // 명단에 없는 이름이면 기록 탭에만 남긴다.
  // D열=같은 시간 최신, E열=같은 거리 최신, F열=갱신 시각
  const col = (game === 'sametime') ? 4 : 5;
  sh.getRange(rowIndex, col).setValue(valueLabel);
  sh.getRange(rowIndex, 6).setValue(stamp);
}

function ensureLogSheet_(ss) {
  let sh = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET_NAME);
    const header = [['시각', '이름', '모둠', '게임', '기록', '등수']];
    sh.getRange(1, 1, 1, 6).setValues(header)
      .setFontWeight('bold')
      .setBackground(APP_COLOR_LIGHT);
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 160);
    sh.setColumnWidth(2, 120);
    sh.setColumnWidth(4, 120);
    sh.setColumnWidth(5, 120);
  }
  return sh;
}

// 명단 탭에 최신 기록용 헤더(D~F)가 없으면 채운다.
function ensureRosterHeaders_(sh) {
  const head = sh.getRange(1, 1, 1, 6).getValues()[0];
  const wanted = ['번호', '이름', '모둠', '같은 시간 최신기록', '같은 거리 최신기록', '갱신 시각'];
  let changed = false;
  for (let c = 0; c < 6; c++) {
    if (String(head[c] || '').trim() === '') { head[c] = wanted[c]; changed = true; }
  }
  if (changed) {
    sh.getRange(1, 1, 1, 6).setValues([head])
      .setFontWeight('bold')
      .setBackground(APP_COLOR_LIGHT);
  }
}

// ──────────────────────────────────────────────────────────────
// 경주 룸 동기화 (같은 모둠 학생들의 실시간 경주)
//   파라미터:
//     action: 'sync' | 'start' | 'reset'
//     room:   '<group>:<gameType>'   예) '1:sametime'
//     name:   학생 이름
//     distance: (sync) 현재까지 달린 거리(m, 정수)
//     finishedAt: (sync, samedist 도착 시) 도착까지 걸린 시간(ms)
//
//   응답: { ok:true, state:{ status, startedAt, countdownEndsAt, players, serverTime } }
// ──────────────────────────────────────────────────────────────
const RACE_TTL_SEC = 600;
const STALE_MS = 4000;         // 4초 이상 sync 없으면 빠진 학생으로 처리
const GAME_DURATIONS = {
  sametime: 10000,             // 10초
  samedist: 30000              // 최장 30초
};

function handleRace_(p) {
  const action = (p.action || '').toLowerCase();
  const room = p.room || '';
  if (!room || !room.includes(':')) {
    return jsonOut_({ ok: false, error: 'invalid room' });
  }
  const cache = CacheService.getScriptCache();
  const cacheKey = 'race:' + room;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(2000); } catch (e) { /* 무시 */ }

  try {
    let state = readState_(cache, cacheKey);
    const now = Date.now();
    const gameType = room.split(':')[1];

    state = advanceStatus_(state, now, gameType);

    switch (action) {
      case 'sync': {
        const name = String(p.name || '').trim();
        if (!name) break;
        if (!state.players[name]) {
          state.players[name] = { distance: 0, finishedAt: null, lastSeenAt: now };
        }
        const d = Math.max(0, Number(p.distance || 0));
        if (Number.isFinite(d)) {
          state.players[name].distance = Math.max(state.players[name].distance, d);
        }
        const fa = p.finishedAt == null ? null : Number(p.finishedAt);
        if (fa && Number.isFinite(fa) && state.players[name].finishedAt == null) {
          state.players[name].finishedAt = fa;
        }
        state.players[name].lastSeenAt = now;
        break;
      }
      case 'start': {
        if (state.status === 'lobby' || state.status === 'done') {
          if (state.status === 'done') {
            state = makeFreshState_();
          }
          state.status = 'countdown';
          state.countdownEndsAt = now + 4000;   // 1초 준비 + 3초 카운트다운
          state.startedAt = state.countdownEndsAt;
        }
        break;
      }
      case 'reset': {
        state = makeFreshState_();
        break;
      }
    }

    state = advanceStatus_(state, now, gameType);

    const playersOut = {};
    Object.keys(state.players).forEach(nm => {
      const pl = state.players[nm];
      playersOut[nm] = {
        distance: pl.distance,
        finishedAt: pl.finishedAt,
        isFresh: (now - (pl.lastSeenAt || 0)) < STALE_MS
      };
    });

    cache.put(cacheKey, JSON.stringify(state), RACE_TTL_SEC);

    return jsonOut_({
      ok: true,
      state: {
        status: state.status,
        startedAt: state.startedAt,
        countdownEndsAt: state.countdownEndsAt,
        players: playersOut,
        serverTime: now
      }
    });
  } finally {
    try { lock.releaseLock(); } catch (e) { /* 무시 */ }
  }
}

function readState_(cache, key) {
  const raw = cache.get(key);
  if (!raw) return makeFreshState_();
  try { return JSON.parse(raw); }
  catch (e) { return makeFreshState_(); }
}

function makeFreshState_() {
  return {
    status: 'lobby',
    startedAt: null,
    countdownEndsAt: null,
    players: {}
  };
}

function advanceStatus_(state, now, gameType) {
  if (state.status === 'countdown' && state.countdownEndsAt && now >= state.countdownEndsAt) {
    state.status = 'running';
  }
  if (state.status === 'running') {
    const dur = GAME_DURATIONS[gameType] || 30000;
    if (state.startedAt && now - state.startedAt >= dur) {
      state.status = 'done';
    }
    if (gameType === 'samedist') {
      const names = Object.keys(state.players);
      const allFinished = names.length > 0 && names.every(n => state.players[n].finishedAt != null);
      if (allFinished) state.status = 'done';
    }
  }
  return state;
}

// ──────────────────────────────────────────────────────────────
// 학생 명단 탭 만들기 (학생1~학생30 미리 채움)
// ──────────────────────────────────────────────────────────────
function setupStudentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(STUDENT_SHEET_NAME);

  const header = ['번호', '이름', '모둠', '같은 시간 최신기록', '같은 거리 최신기록', '갱신 시각'];
  const rows = [header];
  for (let i = 1; i <= 30; i++) {
    rows.push([i, '학생' + i, '', '', '', '']);
  }
  sh.getRange(1, 1, rows.length, header.length).setValues(rows);
  sh.getRange(1, 1, 1, header.length)
    .setFontWeight('bold')
    .setBackground(APP_COLOR_LIGHT);
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 80);
  sh.setColumnWidth(4, 150);
  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 160);

  ensureLogSheet_(ss);
  clearCache();
  try {
    ss.toast('학생명단 탭을 만들었어요. 학생1~학생30 자리에 실제 이름을 넣어 주세요.', '완료', 5);
  } catch (e) { /* 무시 */ }
}

// ──────────────────────────────────────────────────────────────
// 사용 설명 탭 만들기
//  - 기존 안내 탭('사용 설명' 또는 옛 이름)은 지우고 새로 만든다.
//  - 섹션 번호는 순서대로 자동 부여, 항상 첫 번째 위치.
// ──────────────────────────────────────────────────────────────
function setupGuideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 임시 이름 탭을 먼저 만들어 두어(마지막 한 장 삭제 불가 문제 회피),
  // 기존 안내 탭을 안전하게 지운다.
  const tempName = '__안내_임시_' + Math.floor(new Date().getTime());
  const temp = ss.insertSheet(tempName);

  const oldNames = [GUIDE_SHEET_NAME, '사용법', '📋 사용법', '사용 안내', '안내'];
  ss.getSheets().forEach(s => {
    const nm = s.getName();
    if (nm !== tempName && oldNames.indexOf(nm) !== -1) {
      ss.deleteSheet(s);
    }
  });

  temp.setName(GUIDE_SHEET_NAME);
  ss.setActiveSheet(temp);
  ss.moveActiveSheet(1);
  const sh = temp;

  // ── 내용 구성: 섹션은 자동 번호 ──
  const rows = [];        // [[텍스트], ...]  (A열 한 줄)
  const titleRows = [];   // 큰 제목 행 인덱스
  const sectionRows = []; // 섹션 제목 행 인덱스
  let sectionNo = 0;

  const addTitle = (t) => { rows.push([t]); titleRows.push(rows.length); };
  const addSection = (t) => {
    sectionNo += 1;
    rows.push([sectionNo + '. ' + t]);
    sectionRows.push(rows.length);
  };
  const addLine = (t) => { rows.push([t]); };
  const addSteps = (arr) => { arr.forEach((t, i) => rows.push([(i + 1) + ') ' + t])); };
  const addBlank = () => { rows.push(['']); };

  addTitle('달려라, 달려! — 누가 가장 빠를까?  선생님 사용 설명');
  addBlank();
  addLine('데이터나 설정을 변경할 때는 앱 화면이 아니라 해당 시트 탭에서 직접 수정하세요. 탭 이름은 코드와 연결되어 있으므로 삭제하거나 변경하지 마세요.');
  addBlank();

  addSection('이 시트의 탭 구성');
  addLine('학생명단 : 우리 반 학생의 번호·이름·모둠을 적는 탭입니다. 학생들이 앱에서 이름을 고를 때 이 탭을 읽어 옵니다.');
  addLine('경주기록 : 학생들이 경주를 마치면 그 결과가 시간순으로 자동으로 쌓이는 탭입니다.');
  addLine('사용 설명 : 지금 보고 계신 이 안내 탭입니다.');
  addBlank();

  addSection('학생 명단 넣기');
  addLine('학생명단 탭에서 A열은 번호, B열은 이름, C열은 모둠입니다. 1행은 제목 행이고 2행부터 학생을 적습니다.');
  addLine('처음에는 학생1~학생30이 예시로 들어 있습니다. 그 자리에 실제 학생 번호와 이름을 덮어써 주세요.');
  addLine('예시로 남아 있는 학생1~학생30 이름은 앱 명단에 나타나지 않고 자동으로 걸러집니다.');
  addBlank();

  addSection('모둠(팀) 지정하는 방법');
  addLine('같이 경주할 학생들의 C열에 같은 값을 넣으면 같은 팀이 됩니다. 예: 네 명의 C열에 모두 1이라고 적으면 그 넷이 한 팀입니다.');
  addLine('C열을 비워 두면 그 학생은 혼자 컴퓨터 친구 쌩쌩이와 경주합니다.');
  addLine('모둠은 언제든 바꿔도 됩니다. 시트 값만 바꾸면 학생이 화면을 새로고침할 때 바로 반영됩니다.');
  addBlank();

  addSection('앱과 연결해서 우리 반 명단으로 쓰는 방법');
  addSteps([
    '상단 메뉴에서 확장 프로그램을 눌러 Apps Script를 엽니다.',
    '오른쪽 위 배포 버튼에서 새 배포를 고르고, 유형은 웹 앱, 액세스 권한은 모든 사용자로 배포합니다.',
    '배포가 끝나면 나오는 웹 앱 주소를 복사합니다. 주소는 exec 로 끝납니다.',
    '학생들이 쓰는 앱 첫 화면에서 시작하기 버튼 아래의 설정을 열어 그 주소를 붙여넣고 연결하기를 누릅니다.',
    '연결하면 그 아래에 학생용 주소가 만들어집니다. 복사 버튼으로 복사해 우리 반 학생들에게 나눠 주세요.'
  ]);
  addLine('학생은 그 주소로 들어오면 곧바로 우리 반 명단을 보게 되고, 설정을 볼 필요가 없습니다.');
  addBlank();

  addSection('경주 결과 확인하기');
  addLine('학생이 경주를 마치면 경주기록 탭에 시각·이름·모둠·게임·기록이 한 줄씩 쌓입니다.');
  addLine('학생명단 탭의 각 학생 줄에도 그 학생의 가장 최근 기록이 적힙니다.');
  addLine('설정에 연결하지 않은 체험 상태에서는 결과가 시트에 저장되지 않고 그 기기에만 잠시 남습니다.');
  addBlank();

  addSection('다른 선생님과 나누는 방법');
  addSteps([
    '파일 메뉴에서 사본 만들기로 이 시트를 복제합니다.',
    '복제한 시트를 다른 선생님께 드립니다.',
    '받으신 선생님이 위의 앱과 연결하기 방법대로 새로 배포하고, 자기 반 학생 주소를 만들어 쓰시면 됩니다.'
  ]);
  addBlank();

  addSection('자주 묻는 질문');
  addLine('명단을 바꿨는데 앱에 그대로예요 → 학생이 화면을 새로고침했는지 확인하세요. 또는 상단의 \'선생님 메뉴\'에서 명단 새로고침을 눌러 주세요.');
  addLine('모둠을 바꿀 때마다 다시 배포해야 하나요 → 아니요. 시트 값만 바꾸는 것은 바로 반영됩니다. 코드를 바꿀 때만 다시 배포하면 됩니다.');
  addLine('학생 이름이 명단에 안 보여요 → B열에 이름이 바르게 적혀 있는지, 학생1처럼 예시 형태가 아닌지 확인하세요.');

  // ── 시트에 한 번에 입력 ──
  const n = rows.length;
  sh.getRange(1, 1, n, 1).setValues(rows);

  // ── 서식 ──
  const all = sh.getRange(1, 1, n, 1);
  all.setWrap(true).setVerticalAlignment('top').setFontSize(11);
  sh.setColumnWidth(1, 900);

  titleRows.forEach(r => {
    sh.getRange(r, 1)
      .setFontSize(14).setFontWeight('bold').setFontColor(APP_COLOR_DEEP);
  });
  sectionRows.forEach(r => {
    sh.getRange(r, 1)
      .setFontWeight('bold').setBackground(APP_COLOR_LIGHT);
  });
  sh.getRange(1, 1, n, 1).setBorder(true, true, true, true, false, true);

  // 다른 탭이 활성화되어 있지 않게 안내 탭을 보여준다.
  ss.setActiveSheet(sh);
  try {
    ss.toast('사용 설명 탭을 새로 만들었어요.', '완료', 4);
  } catch (e) { /* 무시 */ }
}

// ──────────────────────────────────────────────────────────────
// 공통 유틸
// ──────────────────────────────────────────────────────────────
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatStamp_(date) {
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  return Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm:ss');
}
