/**
 * 옴의 법칙 무한 문제은행 — 결과 수집용 Google Apps Script
 *
 * 설치 방법
 * 1) 결과를 모을 구글 스프레드시트를 새로 만든다.
 * 2) 메뉴 [확장 프로그램] → [Apps Script] 를 열고, 이 코드를 전부 붙여넣고 저장한다.
 * 3) [배포] → [새 배포] → 유형: 웹 앱
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → 배포 후 나오는 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 복사한다.
 * 4) index.html 상단의 SCRIPT_URL = '' 따옴표 안에 그 주소를 붙여넣는다.
 *
 * 시트 구성 (자동 생성)
 *  - 요약 : 학생이 제출할 때마다 한 줄 (점수, 정답률, 직렬/병렬 정답률, 소요시간)
 *  - 상세 : 문항별 한 줄 (유형, 문제, 정답, 학생 답, O/X)
 *  - 학생별 : 학번별 누적 통계 (메뉴 [문제은행] → [학생별 통계 갱신])
 */

const SUMMARY = '요약';
const DETAIL = '상세';
const BY_STUDENT = '학생별';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const now = new Date();

    const s = sheet_(ss, SUMMARY, ['제출시각', '학번', '이름', '모드', '푼 문제', '정답 수', '정답률(%)', '직렬 정답률(%)', '병렬 정답률(%)', '소요시간(분)', '세션ID']);
    s.appendRow([now, String(d.sid), d.name, d.mode, d.total, d.correct, d.rate, d.seriesRate, d.parallelRate, d.minutes, d.session]);

    const t = sheet_(ss, DETAIL, ['제출시각', '학번', '이름', '세션ID', '순번', '구분', '유형', '문제', '정답', '학생 답', '결과']);
    const rows = (d.attempts || []).map(a => [now, String(d.sid), d.name, d.session, a.n, a.group, a.tag, a.q, a.answer, a.input, a.ok ? 'O' : 'X']);
    if (rows.length) t.getRange(t.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    return json_({ ok: true, saved: rows.length });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ ok: true, msg: '옴의 법칙 문제은행 결과 수집 서버가 작동 중입니다.' });
}

/** 스프레드시트를 열면 메뉴 추가 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('문제은행')
    .addItem('학생별 통계 갱신', 'buildStudentStats')
    .addToUi();
}

/** 상세 시트를 학번별로 묶어 누적 통계를 만든다 */
function buildStudentStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = ss.getSheetByName(DETAIL);
  if (!t || t.getLastRow() < 2) return;
  const data = t.getRange(2, 1, t.getLastRow() - 1, 11).getValues();
  const map = {};
  data.forEach(r => {
    const key = r[1];
    const m = map[key] = map[key] || { sid: r[1], name: r[2], n: 0, ok: 0, s: [0, 0], p: [0, 0], wrong: {} };
    const ok = r[10] === 'O';
    m.n++; if (ok) m.ok++;
    const g = r[5] === '직렬' ? m.s : m.p; g[0]++; if (ok) g[1]++;
    if (!ok) { const k = r[5] + '·' + r[6]; m.wrong[k] = (m.wrong[k] || 0) + 1; }
  });
  const pct = (a, b) => b ? Math.round(a / b * 100) : '';
  const out = Object.values(map).sort((a, b) => String(a.sid).localeCompare(String(b.sid))).map(m => {
    const weak = Object.entries(m.wrong).sort((a, b) => b[1] - a[1]).slice(0, 3).map(x => `${x[0]}(${x[1]})`).join(', ');
    return [m.sid, m.name, m.n, m.ok, pct(m.ok, m.n), pct(m.s[1], m.s[0]), pct(m.p[1], m.p[0]), weak];
  });
  let sh = ss.getSheetByName(BY_STUDENT);
  if (sh) sh.clear(); else sh = ss.insertSheet(BY_STUDENT);
  const head = ['학번', '이름', '누적 문제 수', '누적 정답', '정답률(%)', '직렬(%)', '병렬(%)', '많이 틀린 유형(횟수)'];
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (out.length) sh.getRange(2, 1, out.length, head.length).setValues(out);
  sh.setFrozenRows(1);
}

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
