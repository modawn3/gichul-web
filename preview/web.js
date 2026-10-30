// 웹 버전 공통: 로그인, 서버 기록 불러오기·저장, 비공개 저장소의 문항 이미지 주소 받기.
// index.html(학생 화면)은 window.__APP, teacher.html(선생님 화면)은 window.__TEACHER를 정의하고, 로그인이 끝나면 여기서 실행한다.
// 과목: 수학(index.html) 또는 영어(english.html, data_en.js가 window.SUBJ='en'을 정함). 영어 문항 id는 'en_'으로 시작한다.
const SUBJ = window.SUBJ || 'math', isEn = SUBJ === 'en';
const mine = id => String(id).startsWith('en_') === isEn;
const sb = supabase.createClient(window.SB_URL, window.SB_KEY, { auth: { persistSession: true, autoRefreshToken: true } });
window.sb = sb;
let ME = null;

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const lsGet = k => { try { return JSON.parse(localStorage.getItem(k)) } catch (e) { return null } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)) } catch (e) { } };

// ---- 문항 이미지: 서명된 주소(3일)를 모아서 한 번에 받고, 브라우저에 기억해 같은 그림을 다시 받지 않게 한다 ----
const TTL = 3 * 86400, BLANK = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
const URLS = lsGet('gichul_img_urls') || {};
const want = new Set(), inflight = new Set();
let flushT = null;
function imgAttr(path) {
  const c = URLS[path];
  if (c && c[1] > Date.now() + 3600e3) return `src="${c[0]}" data-k="${path}"`;
  if (!inflight.has(path)) { want.add(path); clearTimeout(flushT); flushT = setTimeout(flushImgs, 0) }
  return `src="${BLANK}" data-k="${path}"`;
}
async function flushImgs() {
  const paths = [...want]; want.clear(); paths.forEach(p => inflight.add(p));
  for (let i = 0; i < paths.length; i += 200) {
    const chunk = paths.slice(i, i + 200);
    const { data, error } = await sb.storage.from('items').createSignedUrls(chunk, TTL);
    chunk.forEach(p => inflight.delete(p));
    if (error) { console.warn(error); continue }
    const exp = Date.now() + TTL * 1000;
    for (const d of data) {
      if (!d.signedUrl) continue;
      URLS[d.path] = [d.signedUrl, exp];
      document.querySelectorAll(`img[data-k="${d.path}"]`).forEach(im => im.src = d.signedUrl);
    }
  }
  for (const k in URLS) if (URLS[k][1] < Date.now()) delete URLS[k];
  lsSet('gichul_img_urls', URLS);
}
// 저장소 파일 이름은 영문만 받으므로 문항 id의 한글을 바꾼다: 2024_03월학력평가_01 -> 2024_03h_01
window.imgKey = id => id.replace('월학력평가', 'h').replace('월모의평가', 'm').replace('수능', 'csat').replace(/[^\x00-\x7f]/g, c => 'u' + c.charCodeAt(0).toString(16));
window.QIMG = id => imgAttr(`q/${imgKey(id)}.png`);
window.SIMG = id => Array.from({ length: (window.SCOUNT || {})[id] || 0 }, (_, k) => imgAttr(`s/${imgKey(id)}_${k + 1}.png`));

// ---- 기록: 풀이 한 번 = attempts 한 줄, 나머지(복습 일정·시험 결과·풀던 시험) = user_state 한 칸 ----
async function fetchAll(q) {
  const out = [];
  for (let from = 0; ; from += 1000) {
    const { data, error } = await q().range(from, from + 999);
    if (error) throw error;
    out.push(...data);
    if (data.length < 1000) return out;
  }
}
window.fetchAll = fetchAll;
const cidOf = (key, a) => key + '|' + a.t;
let SYNCED = new Set(), PKEY = '', SKEY = '', stateT = null, pushing = false, usedLocal = false;
let MC = null;
function isMc(id, lv) {   // 객관식 답은 숫자(1~5)로 되돌려야 화면이 지난 답을 제대로 표시한다
  if (!MC) MC = Object.fromEntries((window.ITEMS || []).map(i => [i.id, i.mc]));
  if (lv) { const v = ((window.VARS || {})[id] || [])[lv - 1]; return v ? !!v.mc : false }
  return !!MC[id];
}

async function loadDB() {
  PKEY = 'gichul_pending_' + (isEn ? 'en_' : '') + ME.id; SKEY = 'gichul_state_' + (isEn ? 'en_' : '') + ME.id;
  const [rows, st] = await Promise.all([
    fetchAll(() => sb.from('attempts').select('cid,item_id,variant,ok,answer,secs,mode,at').eq('user_id', ME.id).order('at').order('id')),
    sb.from('user_state').select('data,updated_at').eq('user_id', ME.id).maybeSingle(),
  ]);
  if (st.error) throw st.error;
  // user_state 한 칸에 수학 기록은 맨 위에, 영어 기록은 en 아래에 둔다
  const raw = st.data ? st.data.data || {} : {};
  let data = isEn ? (raw.en || {}) : (({ en, ...m }) => m)(raw);
  const local = lsGet(SKEY);   // 저장하기 전에 창을 닫았으면 이 기기의 것이 더 새것
  if (local && (!st.data || local.t > Date.parse(st.data.updated_at))) data = local.data;
  const DB = Object.assign({ rev: {}, results: [], sets: [] }, data, { att: {}, vatt: {} });
  const add = (r) => {
    const a = { t: Date.parse(r.at), ok: r.ok, a: r.answer == null ? null : isMc(r.item_id, r.variant) ? +r.answer : r.answer };
    if (r.secs != null) a.s = r.secs;
    if (r.mode && r.mode !== 'practice') a.exam = r.mode;
    if (r.variant) (DB.vatt[r.item_id + '#' + r.variant] = DB.vatt[r.item_id + '#' + r.variant] || []).push(a);
    else (DB.att[r.item_id] = DB.att[r.item_id] || []).push(a);
  };
  rows.forEach(r => { SYNCED.add(r.cid); if (mine(r.item_id)) add(r) });
  // 아직 서버에 못 보낸 기록(인터넷이 끊겼을 때)
  for (const r of lsGet(PKEY) || []) if (!SYNCED.has(r.cid) && mine(r.item_id)) add(r);
  usedLocal = !!(local && data === local.data);
  return DB;
}
function rowsToPush(DB) {
  const out = [];
  const each = (map, isV) => {
    for (const key in map) for (const a of map[key]) {
      const cid = cidOf(key, a);
      if (SYNCED.has(cid)) continue;
      const [id, lv] = isV ? key.split('#') : [key, 0];
      out.push({ cid, item_id: id, variant: +lv, ok: !!a.ok, answer: a.a == null || a.a === '' ? null : String(a.a), secs: a.s ?? null, mode: a.exam || 'practice', at: new Date(a.t).toISOString() });
    }
  };
  each(DB.att || {}, false); each(DB.vatt || {}, true);
  return out;
}
let lastDB = null;
async function pushNow() {
  if (pushing || !lastDB) return;
  pushing = true;
  try {
    const rows = rowsToPush(lastDB);
    lsSet(PKEY, rows);
    for (let i = 0; i < rows.length; i += 500) {
      const chunk = rows.slice(i, i + 500);
      const { error } = await sb.from('attempts').upsert(chunk, { onConflict: 'user_id,cid', ignoreDuplicates: true });
      if (error) throw error;
      chunk.forEach(r => SYNCED.add(r.cid));
    }
    lsSet(PKEY, rowsToPush(lastDB));
    setSync(true);
  } catch (e) { console.warn(e); setSync(false) }
  pushing = false;
}
function stateOf(DB) { const { att, vatt, en, ...rest } = DB; return rest }
async function pushState() {
  // 다른 과목 기록은 서버에 있는 그대로 두고 이 과목 부분만 바꾼다
  const { data: cur, error: e0 } = await sb.from('user_state').select('data').eq('user_id', ME.id).maybeSingle();
  if (e0) { console.warn(e0); setSync(false); return }
  const base = cur ? cur.data || {} : {}, mineSt = stateOf(lastDB);
  const data = isEn ? { ...base, en: mineSt } : { ...mineSt, ...(base.en ? { en: base.en } : {}) };
  const { error } = await sb.from('user_state').upsert({ user_id: ME.id, data, updated_at: new Date().toISOString() });
  if (error) { console.warn(error); setSync(false) } else lsSet(SKEY, null);
}
window.__save = function (DB) {
  lastDB = DB;
  pushNow();
  lsSet(SKEY, { t: Date.now(), data: stateOf(DB) });
  clearTimeout(stateT); stateT = setTimeout(pushState, 1500);
};
setInterval(() => { if (lastDB && rowsToPush(lastDB).length) pushNow() }, 30000);
window.addEventListener('online', () => { if (lastDB) { pushNow(); pushState() } });
function setSync(ok) { const el = document.getElementById('syncSt'); if (el) { el.textContent = ok ? '' : '저장 대기 중(인터넷 확인)'; el.title = ok ? '' : '기록은 이 기기에 남아 있고, 연결되면 자동으로 보냅니다.' } }

// ---- 로그인 화면 ----
const CSS = `
.gw-login{position:fixed;inset:0;background:var(--bg,#f6f7f9);display:flex;align-items:center;justify-content:center;z-index:50;padding:16px}
.gw-card{background:var(--panel,#fff);border:1px solid var(--line,#e2e5ea);border-radius:14px;padding:28px 26px;width:100%;max-width:360px;box-shadow:0 8px 30px rgba(0,0,0,.06)}
.gw-card h1{font-size:20px;margin:0 0 4px}.gw-card p{margin:0 0 18px;color:var(--muted,#6b7385);font-size:14px}
.gw-card label{display:block;font-size:13px;color:var(--muted,#6b7385);margin:10px 0 4px}
.gw-card input{width:100%;font:inherit;font-size:16px;padding:10px 12px;border:1px solid var(--line,#e2e5ea);border-radius:8px}
.gw-card button{width:100%;margin-top:18px;font:inherit;font-size:16px;padding:11px;border:0;border-radius:8px;background:var(--accent,#2f6fde);color:#fff;cursor:pointer}
.gw-card button:disabled{opacity:.6}.gw-err{color:var(--bad,#d64545);font-size:14px;margin-top:10px;min-height:20px}
.gw-who{display:flex;gap:8px;align-items:center;font-size:13px;white-space:nowrap}.gw-who a,.gw-who button{font:inherit;font-size:13px;border:1px solid var(--line,#e2e5ea);background:var(--panel,#fff);border-radius:7px;padding:5px 9px;cursor:pointer;color:inherit;text-decoration:none}
#syncSt{color:var(--warn,#b7791f)}
.gw-guide{display:block;text-align:center;margin-top:6px;font-size:14px;color:var(--accent,#2f6fde)}
header>span.small{display:none}#reset{display:none!important}
@media(max-width:820px){header{flex-wrap:wrap;gap:6px 10px;padding:8px 12px}header h1{font-size:16px}
 nav.tabs{order:3;width:100%;overflow-x:auto;flex-wrap:nowrap}nav.tabs button{white-space:nowrap;padding:6px 10px;font-size:14px}
 .gw-who{margin-left:auto;flex-wrap:wrap;justify-content:flex-end}aside{max-height:190px!important;overflow:auto!important;border-right:0;border-bottom:1px solid var(--line,#e2e5ea)}
 main{padding:12px!important}.grid{grid-template-columns:1fr!important}.overlay{padding:8px!important}}
.gw-loading{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;font-size:16px;color:#6b7385;background:var(--bg,#f6f7f9);z-index:49}`;
function addCss() { const s = document.createElement('style'); s.textContent = CSS; document.head.appendChild(s) }
function loading(on, msg) {
  let el = document.getElementById('gwLoad');
  if (!on) { if (el) el.remove(); return }
  if (!el) { el = document.createElement('div'); el.id = 'gwLoad'; el.className = 'gw-loading'; document.body.appendChild(el) }
  el.textContent = msg || '불러오는 중…';
}
function loginScreen() {
  return new Promise(resolve => {
    const setup = /[?&]setup\b/.test(location.search);
    const box = document.createElement('div'); box.className = 'gw-login';
    box.innerHTML = setup ? `<form class="gw-card"><h1>처음 설정</h1><p>관리자 계정을 하나 만듭니다. 관리자가 이미 있으면 만들어지지 않습니다.</p>
      <label>설정 코드</label><input name="scode" required autocomplete="off">
      <label>관리자 아이디 (영문 소문자·숫자)</label><input name="uid" required autocapitalize="off">
      <label>이름</label><input name="nm" required>
      <label>비밀번호 (6자 이상)</label><input name="pw" type="password" minlength="6" required>
      <button>관리자 계정 만들기</button><div class="gw-err"></div></form>`
      : `<form class="gw-card"><h1>수능 ${isEn ? '영어' : '수학'} 기출 학습</h1><p>선생님께 받은 아이디와 비밀번호로 들어오세요.</p>
      <label>아이디</label><input name="uid" required autocapitalize="off" autocomplete="username">
      <label>비밀번호</label><input name="pw" type="password" required autocomplete="current-password">
      <button>로그인</button><div class="gw-err"></div><a class="gw-guide" href="guide.html">처음이세요? 사용 설명서 보기</a></form>`;
    document.body.appendChild(box);
    const f = box.querySelector('form'), err = box.querySelector('.gw-err'), btn = f.querySelector('button');
    f.onsubmit = async e => {
      e.preventDefault(); err.textContent = ''; btn.disabled = true;
      const id = f.elements.uid.value.trim().toLowerCase(), pw = f.elements.pw.value;
      if (setup) {
        const { data, error } = await sb.functions.invoke('admin', { body: { action: 'bootstrap', code: f.elements.scode.value.trim(), user: { login_id: id, name: f.elements.nm.value.trim(), password: pw } } });
        if (error || !data.ok) { err.textContent = (data && data.error) || (await errText(error)); btn.disabled = false; return }
        history.replaceState(null, '', location.pathname);
      }
      const { error } = await sb.auth.signInWithPassword({ email: `${id}@${window.ID_DOMAIN}`, password: pw });
      btn.disabled = false;
      if (error) { err.textContent = /invalid/i.test(error.message) ? '아이디 또는 비밀번호가 맞지 않습니다.' : error.message; return }
      box.remove(); resolve();
    };
    f.querySelector('input').focus();
  });
}
async function errText(error) {
  try { const b = await error.context.json(); return b.error || error.message } catch (e) { return error ? error.message : '오류' }
}
window.errText = errText;
async function signOut() { await sb.auth.signOut(); location.reload() }
async function changePw() {
  const p = prompt('새 비밀번호 (6자 이상)'); if (!p) return;
  if (p.length < 6) { alert('6자 이상으로 정해 주세요.'); return }
  const { error } = await sb.auth.updateUser({ password: p });
  alert(error ? '바꾸지 못했습니다: ' + error.message : '비밀번호를 바꿨습니다.');
}
function whoBar(page) {
  const cls = ME.grade ? `${ME.grade}학년 ${ME.class ?? ''}반 ` : '';
  const role = ME.role === 'student' ? '' : ME.role === 'admin' ? ' (관리자)' : ' (선생님)';
  const el = document.createElement('div'); el.className = 'gw-who';
  el.innerHTML = `<span id="syncSt"></span><span>${esc(cls)}${esc(ME.name)}${role}</span>`
    + (page === 'teacher' ? '' : isEn ? '<a href="index.html">수학</a>' : '<a href="english.html">영어</a>')
    + (ME.role !== 'student' ? (page === 'teacher' ? '<a href="index.html">문제 풀기</a>' : '<a href="teacher.html">선생님 화면</a>') : '')
    + (page === 'teacher' ? '<a href="teacher_guide.html">사용 설명서</a>' : '<a href="guide.html">사용 설명서</a>')
    + '<button data-a="pw">비밀번호 바꾸기</button><button data-a="out">로그아웃</button>';
  el.querySelector('[data-a=pw]').onclick = changePw;
  el.querySelector('[data-a=out]').onclick = signOut;
  const hd = document.querySelector('header'); (hd || document.body).appendChild(el);
}
window.ME = () => ME;

async function boot() {
  addCss();
  let { data: { session } } = await sb.auth.getSession();
  if (!session) { await loginScreen(); ({ data: { session } } = await sb.auth.getSession()) }
  loading(true);
  const { data: p, error } = await sb.from('profiles').select('*').eq('id', session.user.id).maybeSingle();
  if (error || !p) { loading(true, '등록된 계정이 아닙니다. 선생님께 문의하세요.'); setTimeout(signOut, 2500); return }
  ME = p;
  try {
    if (window.__TEACHER) {
      if (ME.role === 'student') { location.href = 'index.html'; return }
      whoBar('teacher'); await window.__TEACHER(ME);
    } else {
      window.__DB = await loadDB();
      lastDB = window.__DB;
      whoBar('app'); window.__APP();
      if (rowsToPush(lastDB).length) pushNow();
      if (usedLocal) pushState();
    }
  } catch (e) { console.error(e); loading(true, '불러오지 못했습니다: ' + (e.message || e)); return }
  loading(false);
}
document.addEventListener('DOMContentLoaded', boot);
