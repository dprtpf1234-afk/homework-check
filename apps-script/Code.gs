/**
 * 과제물 수합 시스템 — 구글 스프레드시트 API (Google Apps Script)
 *
 * 1) 구글 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣기
 * 2) 함수 선택 칸에서 setup 선택 후 ▶ 실행 (권한 승인 + 시트 생성)
 * 3) 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자) > 웹 앱 URL 복사
 * 4) 관리자 페이지 > 설정 > [구글 스프레드시트 연동] 에 URL 붙여넣기
 *
 * 관리자 비밀번호 초기값: 1234 (관리자 페이지에서 변경)
 * 비밀번호를 잊었다면: 프로젝트 설정 > 스크립트 속성에서 ADMIN_PIN 삭제 → 1234로 초기화
 */

const TZ = 'Asia/Seoul';
const DEFAULT_PIN = '1234';
const SHEET = {
  STUDENTS: { name: '학생명단', header: ['번호', '이름'] },
  TASKS: { name: '과제목록', header: ['과제명', '마감일', '등록일시'] },
  RECORDS: { name: '제출기록', header: ['제출일시', '번호', '이름', '과제명', '비고'] },
};

/* ───────── 요청 처리 ───────── */

function doGet() {
  return json_({ ok: true, data: { service: 'submit-tracker', version: 1 } });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: '잘못된 요청입니다.' });
  }
  try {
    if (PUBLIC[req.action]) return json_({ ok: true, data: PUBLIC[req.action](req) });
    if (ADMIN[req.action]) {
      checkPin_(req.pin);
      return json_({ ok: true, data: ADMIN[req.action](req) });
    }
    return json_({ ok: false, error: '알 수 없는 요청: ' + req.action });
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err), code: err.code || '' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 학생 화면에서 쓰는 요청 (비밀번호 불필요) */
const PUBLIC = {
  getStudentData: () => ({
    students: getStudents_(),
    tasks: getTasks_(),
    records: getRecords_().map(r => ({ no: r.no, name: r.name, task: r.task })),
  }),
  submit: p => submit_(String(p.no), String(p.name), [].concat(p.tasks || [])),
  verifyPin: p => {
    checkPin_(p.pin);
    return sheetInfo_();
  },
};

/** 관리자 페이지에서 쓰는 요청 (비밀번호 필요) */
const ADMIN = {
  getDashboard: () => Object.assign({
    students: getStudents_(),
    tasks: getTasks_(),
    records: getRecords_(),
    updated: fmt_(new Date(), 'HH:mm:ss'),
  }, sheetInfo_()),

  toggleRecord: p => withLock_(() => {
    const no = String(p.no), name = String(p.name), task = String(p.task);
    const sh = sheet_(SHEET.RECORDS);
    const idx = dataRows_(SHEET.RECORDS)
      .findIndex(r => String(r[1]) === no && String(r[2]) === name && String(r[3]) === task);
    if (idx >= 0) {
      sh.deleteRow(idx + 2);
      return { submitted: false };
    }
    sh.appendRow([new Date(), Number(no) || no, name, task, '교사 확인']);
    return { submitted: true };
  }),

  addTask: p => withLock_(() => {
    const name = String(p.name || '').trim();
    if (!name) throw new Error('과제명을 입력하세요.');
    if (getTasks_().some(t => t.name === name)) throw new Error('같은 이름의 과제가 이미 있습니다.');
    sheet_(SHEET.TASKS).appendRow([name, String(p.due || ''), new Date()]);
    return getTasks_();
  }),

  deleteTask: p => withLock_(() => {
    const name = String(p.name);
    const idx = dataRows_(SHEET.TASKS).findIndex(r => String(r[0]).trim() === name);
    if (idx < 0) throw new Error('과제를 찾을 수 없습니다.');
    sheet_(SHEET.TASKS).deleteRow(idx + 2); // 제출기록 시트의 기록은 그대로 보존
    return getTasks_();
  }),

  saveStudents: p => withLock_(() => {
    const list = (p.students || [])
      .map(s => [String(s.no).trim(), String(s.name).trim()])
      .filter(s => s[1]);
    const sh = sheet_(SHEET.STUDENTS);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    if (list.length) sh.getRange(2, 1, list.length, 2).setValues(list.map(s => [Number(s[0]) || s[0], s[1]]));
    return getStudents_();
  }),

  changePin: p => {
    const pin = String(p.newPin || '');
    if (!/^\d{4,8}$/.test(pin)) throw new Error('비밀번호는 숫자 4~8자리로 정하세요.');
    PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', pin);
    return true;
  },
};

/* ───────── 최초 설정 (편집기에서 한 번 실행) ───────── */

function setup() {
  Object.keys(SHEET).forEach(k => sheet_(SHEET[k]));
  const st = sheet_(SHEET.STUDENTS);
  if (st.getLastRow() < 2) {
    const names = '김민준 이서연 박도윤 최서준 정하은'.split(' ');
    st.getRange(2, 1, names.length, 2).setValues(names.map((n, i) => [i + 1, n]));
  }
  const tk = sheet_(SHEET.TASKS);
  if (tk.getLastRow() < 2) tk.appendRow(['현장체험학습 동의서', '', new Date()]);
}

/* ───────── 시트 읽기/쓰기 ───────── */

function sheet_(def) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, 1, def.header.length).setValues([def.header])
      .setFontWeight('bold').setBackground('#e8eef7');
    sh.setFrozenRows(1);
    if (def === SHEET.RECORDS) sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    if (def === SHEET.TASKS) sh.getRange('B:B').setNumberFormat('@'); // 마감일을 글자 그대로
  }
  return sh;
}

function dataRows_(def) {
  const values = sheet_(def).getDataRange().getValues();
  values.shift();
  return values;
}

function fmt_(d, pattern) {
  return Utilities.formatDate(d, TZ, pattern);
}

function getStudents_() {
  return dataRows_(SHEET.STUDENTS)
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({ no: String(r[0]), name: String(r[1]).trim() }));
}

function getTasks_() {
  return dataRows_(SHEET.TASKS)
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      name: String(r[0]).trim(),
      due: r[1] instanceof Date ? fmt_(r[1], 'yyyy-MM-dd') : String(r[1] || ''),
    }));
}

function getRecords_() {
  return dataRows_(SHEET.RECORDS)
    .filter(r => String(r[3]).trim() !== '')
    .map(r => ({
      time: r[0] instanceof Date ? fmt_(r[0], 'yyyy-MM-dd HH:mm:ss') : String(r[0]),
      no: String(r[1]),
      name: String(r[2]),
      task: String(r[3]),
      memo: String(r[4] || ''),
    }));
}

function sheetInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { sheetName: ss.getName(), sheetUrl: ss.getUrl() };
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 학생 제출 — 이미 제출한 과제는 중복 기록하지 않음 */
function submit_(no, name, tasks) {
  return withLock_(() => {
    if (!getStudents_().some(s => s.no === no && s.name === name))
      throw new Error('학생 명단이 바뀌었습니다. 화면을 새로고침해 주세요.');
    const taskNames = getTasks_().map(t => t.name);
    const records = getRecords_();
    const now = new Date();
    const added = [], dups = [], rows = [];
    tasks.forEach(t => {
      if (!taskNames.includes(t)) throw new Error(`'${t}' 과제가 목록에 없습니다. 새로고침해 주세요.`);
      if (records.some(r => r.no === no && r.name === name && r.task === t)) dups.push(t);
      else { added.push(t); rows.push([now, Number(no) || no, name, t, '']); }
    });
    if (rows.length) {
      const sh = sheet_(SHEET.RECORDS);
      sh.getRange(sh.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
    }
    return { added, dups, time: fmt_(now, 'yyyy-MM-dd HH:mm:ss') };
  });
}

/** 비밀번호 확인 — 10분 안에 10번 틀리면 잠시 잠금 */
function checkPin_(pin) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw codeErr_('비밀번호를 여러 번 틀렸습니다. 10분 뒤 다시 시도하세요.', 'PIN');
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN') || DEFAULT_PIN;
  if (String(pin) !== real) {
    cache.put('pinFails', String(fails + 1), 600);
    throw codeErr_('비밀번호가 올바르지 않습니다.', 'PIN');
  }
}

function codeErr_(msg, code) {
  const e = new Error(msg);
  e.code = code;
  return e;
}
