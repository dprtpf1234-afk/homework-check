/**
 * 구글 스프레드시트(Apps Script) API 연결 + 데모 모드
 *  - 연동 URL이 있으면 구글 시트에 기록
 *  - 없으면 이 브라우저(localStorage)에 임시 저장하는 데모 모드
 */
(function () {
  const CFG = window.APP_CONFIG || {};
  const URL_KEY = 'submitTracker.apiUrl';
  const DEMO_KEY = 'submitTracker.demo.v1';

  const store = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (e) {} },
  };
  const session = {
    get(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { v == null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch (e) {} },
  };

  const p2 = n => String(n).padStart(2, '0');
  const fmt = d => `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())} ${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`;
  const esc = s => String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const due = s => { const m = /^\d{4}-(\d{2})-(\d{2})/.exec(s || ''); return m ? `${+m[1]}/${+m[2]}` : (s || ''); };

  /* ───── 구글 시트 ───── */
  async function remote(url, action, params) {
    let res;
    try {
      res = await fetch(url, { method: 'POST', body: JSON.stringify(Object.assign({ action }, params)) });
    } catch (e) {
      throw new Error('구글 스프레드시트에 연결할 수 없습니다. 인터넷 연결과 웹 앱 URL을 확인하세요.');
    }
    let j;
    try { j = await res.json(); }
    catch (e) { throw new Error('구글 시트 응답을 읽을 수 없습니다. 웹 앱 URL과 배포 설정(액세스: 모든 사용자)을 확인하세요.'); }
    if (!j.ok) { const err = new Error(j.error || '알 수 없는 오류'); err.code = j.code; throw err; }
    return j.data;
  }

  /* ───── 데모 (브라우저 저장) ───── */
  let mem = null;
  function dsave(d) { mem = d; store.set(DEMO_KEY, JSON.stringify(d)); }
  function dload() {
    const s = store.get(DEMO_KEY);
    if (s) { try { return (mem = JSON.parse(s)); } catch (e) {} }
    if (mem) return mem;
    const names = '김민준 이서연 박도윤 최서준 정하은 강지호 조수아 윤예준 장지유 임시우 한채원 오건우 서다은 신유준 권하린 황준서 안소율 송민재 류지안 홍은우'.split(' ');
    const today = new Date();
    const day = n => { const d = new Date(today); d.setDate(d.getDate() + n); return fmt(d).slice(0, 10); };
    const d = {
      pin: '1234',
      students: names.map((n, i) => ({ no: String(i + 1), name: n })),
      tasks: [
        { name: '현장체험학습 동의서', due: day(-1) },
        { name: '방과후학교 신청서', due: day(2) },
        { name: '가정환경조사서', due: day(4) },
        { name: '독서감상문 과제', due: day(7) },
      ],
      records: [],
    };
    d.students.forEach((s, i) => d.tasks.forEach((t, j) => {
      if ((i * 7 + j * 3) % 5 < 4 - j)
        d.records.push({ time: fmt(new Date(Date.now() - (i * 5 + j * 97) * 60e3)), no: s.no, name: s.name, task: t.name, memo: '' });
    }));
    dsave(d);
    return d;
  }
  function pinErr() { const e = new Error('비밀번호가 올바르지 않습니다. (데모 비밀번호: 1234)'); e.code = 'PIN'; return e; }
  const has = (d, no, name, t) => d.records.some(r => r.no === no && r.name === name && r.task === t);

  const demo = {
    getStudentData() { const d = dload(); return { students: d.students, tasks: d.tasks, records: d.records }; },
    submit({ no, name, tasks }) {
      const d = dload(), time = fmt(new Date()), added = [], dups = [];
      tasks.forEach(t => { if (has(d, no, name, t)) dups.push(t); else { added.push(t); d.records.push({ time, no, name, task: t, memo: '' }); } });
      dsave(d); return { added, dups, time };
    },
    verifyPin({ pin }) { if (String(pin) !== dload().pin) throw pinErr(); return { sheetName: '', sheetUrl: '' }; },
  };
  const demoAdmin = {
    getDashboard() { const d = dload(); return { students: d.students, tasks: d.tasks, records: d.records, updated: fmt(new Date()).slice(11), sheetName: '', sheetUrl: '' }; },
    toggleRecord({ no, name, task }) {
      const d = dload(), i = d.records.findIndex(r => r.no === no && r.name === name && r.task === task);
      if (i >= 0) d.records.splice(i, 1); else d.records.push({ time: fmt(new Date()), no, name, task, memo: '교사 확인' });
      dsave(d); return { submitted: i < 0 };
    },
    addTask({ name, due }) {
      const d = dload(); name = String(name || '').trim();
      if (!name) throw new Error('과제명을 입력하세요.');
      if (d.tasks.some(t => t.name === name)) throw new Error('같은 이름의 과제가 이미 있습니다.');
      d.tasks.push({ name, due: due || '' }); dsave(d); return d.tasks;
    },
    deleteTask({ name }) { const d = dload(); d.tasks = d.tasks.filter(t => t.name !== name); dsave(d); return d.tasks; },
    saveStudents({ students }) { const d = dload(); d.students = students.filter(s => s.name); dsave(d); return d.students; },
    changePin({ newPin }) {
      if (!/^\d{4,8}$/.test(String(newPin))) throw new Error('비밀번호는 숫자 4~8자리로 정하세요.');
      const d = dload(); d.pin = String(newPin); dsave(d); return true;
    },
  };

  /* ───── 공개 API ───── */
  window.Api = {
    get url() {
      const v = store.get(URL_KEY);
      if (v === 'demo') return '';
      return v || CFG.apiUrl || '';
    },
    get isDemo() { return !this.url; },
    setUrl(u) { store.set(URL_KEY, u ? u.trim() : 'demo'); },

    async call(action, params = {}) {
      if (this.url) return remote(this.url, action, params);
      await new Promise(r => setTimeout(r, 150));
      if (demo[action]) return demo[action](params);
      if (demoAdmin[action]) { if (String(params.pin) !== dload().pin) throw pinErr(); return demoAdmin[action](params); }
      throw new Error('알 수 없는 요청: ' + action);
    },
    /** 저장하기 전에 URL과 비밀번호가 맞는지 확인 */
    test(url, pin) { return remote(url.trim(), 'verifyPin', { pin }); },
    resetDemo() { store.set(DEMO_KEY, null); mem = null; },
    DEMO_KEY,
  };
  window.U = { esc, fmt, due, session, CFG };
})();
