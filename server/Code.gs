/**
 * 한나T 빡침모드 ON — 서버 (구글 시트 + Apps Script)
 *
 * 처음 한 번:
 *  1) (이미 만들어져 있음) 이 프로젝트는 "한나T 빡침모드 DB" 시트를 SHEET_ID로 연결해요
 *  2) 위쪽 함수 목록에서 setup 을 고르고 [실행] → 권한 허용
 *  3) [배포 → 새 배포 → 웹 앱] / 실행: 나 / 액세스 권한: 모든 사용자 → 배포 → 웹 앱 URL 복사
 *  4) 폰 앱의 [명렬표 → 서버 연결]에 URL과 앱 비밀번호(설정 탭에 적혀 있음)를 넣기
 *
 * 코드를 고친 뒤에는 [배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포] (URL은 그대로)
 */

const TZ = 'Asia/Seoul';
const SHEET_ID = '14lqKtwJApdlzrnhbLDOE9L4dLGj1s_T-J79G5lWnRNs';
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
const MAX_PER_WEEK = 4;
const CAL_NAME = '한나T 빡침모드'; // 예전에 쓰던 전용 캘린더 이름 (지우기용)
const SW_HEAD = ['id', '반', '번호', '이름', '체크일', '마감일', '제출', '제출일', '알림일정ID', '만든시각', '받아감', '받아간날'];
const LT_HEAD = ['id', '반', '번호', '이름', '지각일', '청소시작(월)', '캘린더ID', '만든시각'];
const RO_HEAD = ['반', '번호', '이름'];

/* ---------------- 처음 설정 ---------------- */
function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  ensureSheet_('명렬표', RO_HEAD);
  const sw = ensureSheet_('욕설', SW_HEAD);
  sw.getRange('A:L').setNumberFormat('@');
  sw.getRange(1, 1, 1, SW_HEAD.length).setValues([SW_HEAD]);
  ensureSheet_('지각', LT_HEAD).getRange('A:H').setNumberFormat('@');
  ensureSheet_('설정', ['항목', '값']);
  let key = String(getSetting_('앱비밀번호') || '');
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    setSetting_('앱비밀번호', key);
  }
  const msg = '설정 완료!\n앱 비밀번호: ' + key + '\n(시트의 "설정" 탭에도 적혀 있어요)\n스크립트 시간대: ' + Session.getScriptTimeZone();
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { /* 편집기에서 실행하면 로그에만 표시 */ }
}

/* ---------------- 웹 앱 입구 ---------------- */
function doGet() {
  return out_({ ok: true, app: '한나T 빡침모드 ON', msg: '서버가 켜져 있어요. 앱에서 연결하세요.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '요청을 읽지 못했어요.' }); }
  const key = String(getSetting_('앱비밀번호') || '');
  if (!key) return out_({ ok: false, error: '서버에서 setup을 먼저 실행해 주세요.' });
  if (String(req.key || '') !== key) return out_({ ok: false, error: '앱 비밀번호가 맞지 않아요.', code: 'bad_key' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (err) { return out_({ ok: false, error: '서버가 바빠요. 잠시 후 다시 눌러 주세요.' }); }
  try {
    const res = handle_(req) || {};
    res.ok = res.ok !== false;
    res.all = getAll_();
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: '서버 오류: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function handle_(req) {
  const p = req.payload || {};
  switch (req.action) {
    case 'all': return {};
    case 'addSwear': return addSwear_(p);
    case 'delSwear': return delSwear_(p);
    case 'setSubmitted': return setSubmitted_(p);
    case 'setReceived': return setReceived_(p);
    case 'addLate': return addLate_(p);
    case 'delLate': return delLate_(p);
    case 'setHomeroom': setSetting_('우리반', String(p.homeroom || '')); return {};
    case 'setRosters': return setRosters_(p.rosters || {}, true);
    case 'setRoster': { const r = {}; r[String(p.cls)] = p.students || []; return setRosters_(r, false); }
    default: return { ok: false, error: '알 수 없는 요청이에요.' };
  }
}

/* ---------------- 욕설 · 필사 ---------------- */
function addSwear_(p) {
  const cls = String(p.cls), no = Number(p.no), name = String(p.name || (no + '번')).slice(0, 30), date = cleanDate_(p.date);
  const t = table_('욕설');
  const dup = t.data.map(swObj_).find(r => r.cls === cls && r.no === no && r.date === date);
  if (dup) return { rec: dup };
  const due = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dueDate || '')) ? String(p.dueDate) : alertOf_(date);
  const rec = { id: Utilities.getUuid(), cls: cls, no: no, name: name, date: date, alertDate: due, submitted: false, submittedDate: '', calEventId: '', createdAt: Date.now(), received: !!p.received, receivedDate: p.received ? today_() : '' };
  t.sh.appendRow(swRow_(rec));
  return { rec: rec };
}

function delSwear_(p) {
  const t = table_('욕설');
  const i = t.data.findIndex(r => String(r[0]) === String(p.id));
  if (i < 0) return {};
  t.sh.deleteRow(i + 2);
  return {};
}

function setSubmitted_(p) {
  const t = table_('욕설');
  const i = t.data.findIndex(r => String(r[0]) === String(p.id));
  if (i < 0) return { ok: false, error: '기록을 찾지 못했어요.' };
  const rec = swObj_(t.data[i]);
  rec.submitted = !!p.submitted;
  rec.submittedDate = p.submitted ? today_() : '';
  t.sh.getRange(i + 2, 1, 1, SW_HEAD.length).setValues([swRow_(rec)]);
  return { rec: rec };
}

function setReceived_(p) {
  const t = table_('욕설');
  const i = t.data.findIndex(r => String(r[0]) === String(p.id));
  if (i < 0) return { ok: false, error: '기록을 찾지 못했어요.' };
  const rec = swObj_(t.data[i]);
  rec.received = !!p.received;
  rec.receivedDate = p.received ? today_() : '';
  t.sh.getRange(i + 2, 1, 1, SW_HEAD.length).setValues([swRow_(rec)]);
  return { rec: rec };
}

/* ---------------- 지각 · 청소 ---------------- */
function addLate_(p) {
  const homeroom = String(getSetting_('우리반') || '');
  if (!homeroom) return { ok: false, error: '우리 반이 정해져 있지 않아요.' };
  const no = Number(p.no), name = String(p.name || (no + '번')).slice(0, 30), date = cleanDate_(p.date);
  const t = table_('지각');
  const list = t.data.map(ltObj_);
  const dup = list.find(r => r.no === no && r.date === date);
  if (dup) return { rec: dup };
  let ws = nextMon_(date), moved = false;
  while (list.filter(r => r.weekStart === ws).length >= MAX_PER_WEEK) { ws = addDays_(ws, 7); moved = true; }
  const rec = { id: Utilities.getUuid(), cls: homeroom, no: no, name: name, date: date, weekStart: ws, calEventId: '', createdAt: Date.now() };
  t.sh.appendRow(ltRow_(rec));
  return { rec: rec, moved: moved };
}

function delLate_(p) {
  const t = table_('지각');
  const i = t.data.findIndex(r => String(r[0]) === String(p.id));
  if (i < 0) return {};
  t.sh.deleteRow(i + 2);
  return {};
}

/* ---------------- 명렬표 ---------------- */
function setRosters_(incoming, replaceAll) {
  const sh = ss_().getSheetByName('명렬표');
  const cur = replaceAll ? {} : getAll_().rosters;
  Object.keys(incoming).forEach(c => {
    const arr = (incoming[c] || []).map(s => ({ no: Number(s.no), name: String(s.name || '').trim().slice(0, 30) }))
      .filter(s => s.name && isFinite(s.no));
    cur[String(c)] = arr;
  });
  const rows = [];
  Object.keys(cur).sort((a, b) => Number(a) - Number(b)).forEach(c => {
    cur[c].slice().sort((a, b) => a.no - b.no).forEach(s => rows.push([String(c), s.no, s.name]));
  });
  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([RO_HEAD]);
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  return { count: rows.length };
}

/* ---------------- 전체 읽기 ---------------- */
function getAll_() {
  const rosters = {};
  table_('명렬표').data.forEach(r => {
    const c = String(r[0]).trim(), name = String(r[2]).trim();
    if (!c || !name) return;
    (rosters[c] = rosters[c] || []).push({ no: Number(r[1]), name: name });
  });
  Object.keys(rosters).forEach(c => rosters[c].sort((a, b) => a.no - b.no));
  return {
    homeroom: String(getSetting_('우리반') || ''),
    rosters: rosters,
    swear: table_('욕설').data.filter(r => r[0]).map(swObj_),
    late: table_('지각').data.filter(r => r[0]).map(ltObj_),
    today: today_()
  };
}

/* ---------------- 도우미 ---------------- */
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function ensureSheet_(name, head) {
  const ss = ss_();
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (String(sh.getRange(1, 1).getValue()) === '') sh.getRange(1, 1, 1, head.length).setValues([head]);
  sh.setFrozenRows(1);
  return sh;
}
function table_(name) {
  const sh = ss_().getSheetByName(name);
  const v = sh.getDataRange().getValues();
  return { sh: sh, data: v.slice(1) };
}
function getSetting_(k) {
  const v = table_('설정').data.find(r => String(r[0]).trim() === k);
  return v ? v[1] : '';
}
function setSetting_(k, val) {
  const t = table_('설정');
  const i = t.data.findIndex(r => String(r[0]).trim() === k);
  if (i < 0) t.sh.appendRow([k, String(val)]);
  else t.sh.getRange(i + 2, 2).setNumberFormat('@').setValue(String(val));
}
function str_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v == null ? '' : v);
}
function swObj_(r) {
  return { id: String(r[0]), cls: String(r[1]), no: Number(r[2]), name: String(r[3]), date: str_(r[4]), alertDate: str_(r[5]),
    submitted: r[6] === true || String(r[6]).toUpperCase() === 'Y' || String(r[6]).toUpperCase() === 'TRUE',
    submittedDate: str_(r[7]), calEventId: String(r[8] || ''), createdAt: Number(r[9]) || 0,
    received: String(r[10]).toUpperCase() === 'Y' || r[10] === true, receivedDate: str_(r[11]) };
}
function swRow_(o) {
  return [o.id, String(o.cls), String(o.no), o.name, o.date, o.alertDate, o.submitted ? 'Y' : '', o.submittedDate || '', o.calEventId || '', String(o.createdAt || ''), o.received ? 'Y' : '', o.receivedDate || ''];
}
function ltObj_(r) {
  return { id: String(r[0]), cls: String(r[1]), no: Number(r[2]), name: String(r[3]), date: str_(r[4]), weekStart: str_(r[5]),
    calEventId: String(r[6] || ''), createdAt: Number(r[7]) || 0 };
}
function ltRow_(o) {
  return [o.id, String(o.cls), String(o.no), o.name, o.date, o.weekStart, o.calEventId || '', String(o.createdAt || '')];
}
/* ---------------- 구글 캘린더 정리 (한 번만 실행) ---------------- */
// 이제 앱은 구글 캘린더를 쓰지 않아요. 예전에 넣은 일정을 모두 지우고, 시트의 일정ID 칸도 비워요.
function removeCalendar() {
  let n = 0;
  const id = String(getSetting_('캘린더ID') || '');
  const own = (id ? [CalendarApp.getCalendarById(id)] : []).concat(CalendarApp.getOwnedCalendarsByName(CAL_NAME)).filter(c => c);
  const seen = {};
  own.forEach(c => { if (seen[c.getId()]) return; seen[c.getId()] = 1; c.deleteCalendar(); n++; });
  const def = CalendarApp.getDefaultCalendar();
  [['욕설', 9, SW_HEAD.length], ['지각', 7, LT_HEAD.length]].forEach(([name, col]) => {
    const t = table_(name);
    t.data.forEach((r, i) => {
      const ev = String(r[col - 1] || '');
      if (!ev) return;
      try { const e = def.getEventById(ev); if (e) e.deleteEvent(); } catch (err) { /* 이미 없음 */ }
      t.sh.getRange(i + 2, col).setValue('');
    });
  });
  setSetting_('캘린더ID', '');
  Logger.log('지운 캘린더: ' + n + '개');
}
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function cleanDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s)) ? String(s) : today_(); }
function ymdParts_(s) { return String(s).split('-').map(Number); }
function addDays_(s, n) { const p = ymdParts_(s); return new Date(Date.UTC(p[0], p[1] - 1, p[2] + n, 12)).toISOString().slice(0, 10); }
function dow_(s) { const p = ymdParts_(s); return new Date(Date.UTC(p[0], p[1] - 1, p[2], 12)).getUTCDay(); }
function alertOf_(s) { let a = addDays_(s, 2); const w = dow_(a); if (w === 6) a = addDays_(a, 2); else if (w === 0) a = addDays_(a, 1); return a; }
function nextMon_(s) { const w = dow_(s); let n = (8 - w) % 7; if (!n) n = 7; return addDays_(s, n); }
