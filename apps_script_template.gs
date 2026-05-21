/**
 * 속력의 숲 — Apps Script 백엔드 템플릿
 *
 * 역할:
 *  1) 학생 명단(이름·번호·모둠) 시트를 읽어 JSON으로 제공
 *  2) 페이지 첫 로드 시 학생 명단을 HTML에 미리 주입(preloadedJson)
 *  3) 모둠으로 묶인 학생들끼리 실시간 경주가 가능하도록
 *     CacheService를 이용한 race 룸 상태 동기화 엔드포인트 제공
 *
 * 시트 구조 (학생 명단 시트, 기본 이름 "학생명단"):
 *   A열: 번호 (정수)
 *   B열: 이름 (문자열)
 *   C열: 모둠 (문자/숫자, 비어 있으면 개인 모드 → 컴퓨터와 경주)
 *   ※ 1행은 제목 행. 2행부터 데이터.
 *
 * 배포:
 *   - "확장 프로그램 → Apps Script"에서 이 파일 전체를 복사해 붙여넣기
 *   - "배포 → 새 배포 → 유형: 웹 앱"
 *   - 실행 계정: 본인
 *   - 액세스 권한: 모든 사용자 (학생들이 로그인 없이 접근 가능)
 *   - 배포 후 "웹 앱 URL"을 학생들에게 공유
 *
 * 시트를 사본으로 복제한 다른 선생님은 같은 절차로 "새 배포"만 하면 됨.
 */

const STUDENT_SHEET_NAME = '학생명단';   // 학생 명단이 들어있는 시트 이름

// ──────────────────────────────────────────────────────────────
// HtmlService — 페이지를 서빙할 때 학생 명단을 미리 주입
// ──────────────────────────────────────────────────────────────
function doGet(e) {
  // race 액션은 별도 처리(JSON 응답)
  if (e && e.parameter && e.parameter.api === 'race') {
    return handleRace_(e.parameter);
  }
  // 학생 명단 JSON 요청
  if (e && e.parameter && e.parameter.api === 'students') {
    return jsonOut_({ ok: true, students: loadStudents_() });
  }

  // 그 외에는 HTML 페이지 서빙
  const tmpl = HtmlService.createTemplateFromFile('index');
  tmpl.preloadedJson = JSON.stringify({ ok: true, students: loadStudents_() });
  return tmpl.evaluate()
    .setTitle('속력의 숲')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// 일부 호스팅에서는 fetch가 POST를 쓸 수 있으므로 동일 로직을 doPost에도 연결
function doPost(e) {
  const params = (e && e.parameter) || {};
  if (params.api === 'race') return handleRace_(params);
  return jsonOut_({ ok: false, error: 'unknown api' });
}

// ──────────────────────────────────────────────────────────────
// 학생 명단 로드
// ──────────────────────────────────────────────────────────────
function loadStudents_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(STUDENT_SHEET_NAME);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, 3).getValues();
  return rows
    .map(r => ({
      number: r[0] === '' ? null : Number(r[0]),
      name: String(r[1] || '').trim(),
      group: r[2] === '' || r[2] == null ? '' : String(r[2]).trim()
    }))
    .filter(s => s.name);
}

// ──────────────────────────────────────────────────────────────
// race 룸 동기화
//   파라미터:
//     action: 'sync' | 'start' | 'reset'
//     room:   '<group>:<gameType>'   예) '1:sametime'
//     name:   학생 이름
//     distance: (sync일 때) 현재까지 달린 거리(m, 정수)
//     finishedAt: (sync일 때, samedist 도착 시) 도착까지 걸린 시간(ms)
//                 도착 전이면 0 또는 미전송
//
//   응답: { ok: true, state: {...} }
//     state.status: 'lobby' | 'countdown' | 'running' | 'done'
//     state.startedAt: 경주 시작(또는 시작 예정) ms
//     state.countdownEndsAt: 카운트다운 종료 ms
//     state.players: { name: { distance, finishedAt, isFresh } }
//     state.serverTime: 서버 현재 ms (클라이언트 시간 보정용)
//
//   캐시 키: race:<room>
//   TTL: 10분 (CacheService 최대 6시간)
// ──────────────────────────────────────────────────────────────
const RACE_TTL_SEC = 600;
const STALE_MS = 4000;         // 4초 이상 sync 없으면 stale(빠진 학생)
const GAME_DURATIONS = {
  sametime: 10000,             // 10초
  samedist: 30000              // 최장 30초 (그 안에 모두 도착 또는 강제 종료)
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
  // 동시 호출 충돌 방지(짧게)
  try { lock.waitLock(2000); } catch (e) { /* 무시 */ }

  try {
    let state = readState_(cache, cacheKey);
    const now = Date.now();
    const gameType = room.split(':')[1];   // 'sametime' | 'samedist'

    // 자동 전환 처리 (이전 상태가 stale했을 때 정리)
    state = advanceStatus_(state, now, gameType);

    switch (action) {
      case 'sync': {
        const name = String(p.name || '').trim();
        if (!name) break;
        if (!state.players[name]) {
          state.players[name] = { distance: 0, finishedAt: null, lastSeenAt: now };
        }
        // 학생이 보낸 진행도(클라 측에서 단조 증가만 보내야 함)
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
          // 'done'에서 누군가 '시작'을 다시 누르면 즉시 새 경주로 진행
          if (state.status === 'done') {
            state = makeFreshState_(state);
          }
          state.status = 'countdown';
          state.countdownEndsAt = now + 3000;
          state.startedAt = state.countdownEndsAt;   // 진짜 시작은 카운트다운 끝
        }
        break;
      }
      case 'reset': {
        state = makeFreshState_();
        break;
      }
    }

    // sync 이후 다시 한 번 자동 전환 적용
    state = advanceStatus_(state, now, gameType);

    // 응답엔 isFresh 추가
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
    // samedist: 모두 도착하면 종료
    if (gameType === 'samedist') {
      const names = Object.keys(state.players);
      const allFinished = names.length > 0 && names.every(n => state.players[n].finishedAt != null);
      if (allFinished) state.status = 'done';
    }
  }
  return state;
}

// ──────────────────────────────────────────────────────────────
// 공통: JSON 응답
// ──────────────────────────────────────────────────────────────
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
