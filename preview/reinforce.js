// 오답 보강 노트: 학교 마킹 리스트(xls)로 학생별 보강 문항·해설 인쇄본을 만든다.
// 파일은 이 브라우저 안에서만 읽고 어디에도 올리지 않는다. 이름 열은 읽지 않는다(반·번호만 쓴다).
// 규칙(2026-11-04 승진님): 국어·영어·수학만 / 과목당 10문항 내외 / 올린 시험의 오답을 모두 보고
//   쉬운 문항(난이도 하→중→상) 먼저, 같은 난이도면 두 번 이상 시험에서 틀린 영역 먼저 / 영역당 최대 3문항
//   / 비슷한 기출은 원래 문항보다 한 단계 넘게 어렵지 않게 / 변형 문항 없이 실제 기출만, 올린 시험 회차는 빼고
//   / 미적분·기하 선택자는 공통(1~22번)만 / 긴 그림은 글줄 사이 빈 줄에서만 잘라 다음 단으로.
(function () {
const R = window.Reinforce = {};
const SN = { ko: '국어', ma: '수학', en: '영어' }, SUBS = ['ko', 'ma', 'en'];
const DORD = { lo: 0, mid: 1, hi: 2 }, DN = { hi: '상', mid: '중', lo: '하' };
const W = { ma: { c: 4, code: 2.5 }, ko: { c: 3, code: 5 }, en: { c: 1.5, code: 5 } };
const SUBJMAP = { '화법과작문': ['ko', 'h'], '화작': ['ko', 'h'], '언어와매체': ['ko', 'm'], '언매': ['ko', 'm'], '확률과통계': ['ma', '확통'], '확통': ['ma', '확통'], '미적분': ['ma', '미적분'], '기하': ['ma', '기하'], '영어': ['en', ''] };
const CIR = '①②③④⑤';
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const NOT = { c: null, s: [], f: [], d: null, r: null, w: [] };

// ---- 과목 자료 불러오기(다른 화면이 쓰는 window 값은 되돌려 둔다) ----
const loadJs = src => new Promise((ok, no) => { const e = document.createElement('script'); e.src = src; e.onload = ok; e.onerror = () => no(new Error(src + ' 불러오기 실패')); document.head.appendChild(e) });
const KEEP = ['UNITS', 'ITEMS', 'TAGS', 'TAGDEF', 'SUBJ', 'PSG', 'SCOUNT'];
R.SD = null;
R.load = async function () {
  if (!window.CUTMODEL) try { await loadJs('cutmodel.js') } catch (e) { }   // 진단표의 수학 예상 등급용
  if (R.SD) return R.SD;
  const SD = {}, keep = KEEP.map(k => [k, window[k]]);
  try {
    for (const [s, sfx] of [['ma', ''], ['ko', '_ko'], ['en', '_en']]) {
      KEEP.forEach(k => { window[k] = undefined });
      await loadJs(`data${sfx}.js`); await loadJs(`tags${sfx}.js`); await loadJs(`scount${sfx}.js`);
      const S = { items: window.ITEMS, units: Object.fromEntries(window.UNITS.map(u => [u.code, u])), tags: window.TAGS || {}, td: window.TAGDEF || { c: {}, s: {}, f: {} }, psg: window.PSG || {}, scount: window.SCOUNT || {} };
      S.byid = Object.fromEntries(S.items.map(i => [i.id, i]));
      const DF = {};
      for (const i of S.items) { const t = S.tags[i.id] || NOT; (t.s || []).forEach(x => DF['s:' + x] = (DF['s:' + x] || 0) + 1); (t.f || []).forEach(x => DF['f:' + x] = (DF['f:' + x] || 0) + 1) }
      S.idf = k => Math.log((S.items.length + 1) / ((DF[k] || 0) + 1));
      SD[s] = S;
    }
  } finally { keep.forEach(([k, v]) => window[k] = v) }
  return R.SD = SD;
};
// 시험 목록(수학 자료 기준, 최근 것부터): e = '2026_10월학력평가'
R.exams = () => [...new Map(R.SD.ma.items.map(i => [i.e, i.label])).entries()].sort((a, b) => b[0].localeCompare(a[0])).map(([e, label]) => ({ e, label }));
R.short = e => { const m = /_(\d+)월(학력|모의)평가/.exec(e); return m ? `${+m[1]}월` : e.includes('수능') ? '수능' : e };
R.guessExam = name => {
  const ex = R.exams(), y = (/20(2\d)/.exec(name) || [])[0], m = /(\d{1,2})\s*월/.exec(name), csat = /수능/.test(name);
  const c = ex.filter(x => (!y || x.e.startsWith(y)) && (csat ? x.e.includes('수능') : m ? x.e.includes(`_${String(+m[1]).padStart(2, '0')}월`) : false));
  return (c[0] || {}).e || '';
};

// ---- 마킹 리스트 읽기: 원본(계열·학년·반·번호·이름·과목·1번~)과 줄인 형식(반·번호·과목/번호·1번~) 모두 ----
const norm = v => { v = String(v ?? '').trim(); if (v === '') return ''; const n = Number(v); return isFinite(n) ? String(Math.round(n)) : v };
R.parse = function (buf) {
  const wb = XLSX.read(buf, { type: 'array' });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' });
  let hc = null, a0 = -1, hr = -1;
  for (let r = 0; r < Math.min(rows.length, 8); r++) {
    const cells = rows[r].map(x => String(x).replace(/\s/g, ''));
    const ci = cells.indexOf('반'), ni = cells.indexOf('번호'), si = cells.findIndex(x => x.startsWith('과목'));
    if (ci >= 0 && ni >= 0 && si >= 0 && !hc) { hc = { ci, ni, si }; hr = r }
    const one = cells.indexOf('1번'); if (one >= 0 && a0 < 0) a0 = one;
  }
  if (!hc) throw new Error('제목 줄(반·번호·과목)을 찾지 못했습니다.');
  if (a0 < 0) a0 = hc.si + 1;
  const out = [], other = {};
  for (let r = hr + 1; r < rows.length; r++) {
    const row = rows[r], cls = norm(row[hc.ci]), no = norm(row[hc.ni]), sj = String(row[hc.si] || '').replace(/\s/g, '');
    if (!/^\d+$/.test(cls) || !/^\d+$/.test(no) || !sj) continue;
    if (!SUBJMAP[sj]) { other[sj] = (other[sj] || 0) + 1; continue }
    const [s, sel] = SUBJMAP[sj];
    out.push({ cls: +cls, no: +no, s, sel, subj: sj, ans: Array.from({ length: 45 }, (_, k) => norm(row[a0 + k])) });
  }
  return { rows: out, classes: [...new Set(out.map(x => x.cls))].sort((a, b) => a - b), other };
};

// ---- 채점 ----
function iid(s, e, n, sel) {
  const nn = String(n).padStart(2, '0');
  if (s === 'ma') return (n <= 22 || sel === '확통') ? `${e}_${nn}` : null;
  if (s === 'en') return n >= 18 ? `en_${e}_${nn}` : null;
  return n <= 34 ? `ko_${e}_${nn}` : `ko_${e}_${sel}${n}`;
}
function grade(e, rec) {
  const S = R.SD[rec.s], nq = rec.s === 'ma' ? 30 : 45, res = [];
  for (let n = 1; n <= nq; n++) {
    const id = iid(rec.s, e, n, rec.sel), it = id && S.byid[id];
    if (!it) continue;   // 정답이 자료에 없는 문항(영어 듣기, 미적분·기하 23~30번)은 채점하지 않는다
    const mine = rec.ans[n - 1], key = String(it.ans);
    res.push({ n, id, key, mine, ok: mine === key, it });
  }
  return res;
}
const tg = (s, id) => R.SD[s].tags[id] || NOT;
function sim(s, a, b) {
  const S = R.SD[s], ta = tg(s, a.id), tb = tg(s, b.id), P = [], w = W[s], U = S.units, td = S.td;
  if (ta.c && ta.c === tb.c) P.push([w.c, (td.c || {})[ta.c] || '']);
  const wa = ta.w || [], wb = tb.w || [];
  if (wa.length && wb.length) {
    if (wa[0] === wb[0]) P.push([3.5, String((td.w || {})[wa[0]] || '').split(' ').slice(1).join(' ') || String((td.w || {})[wa[0]] || '')]);
    else if (wa.find(v => wb.includes(v))) P.push([1.5, '']);
  }
  if (a.code === b.code) P.push([w.code, U[a.code].name]);
  else if ((a.tags || []).includes(b.code) || (b.tags || []).includes(a.code)) P.push([1.2, '']);
  else if (U[a.code].mid === U[b.code].mid) P.push([.6, '']);
  (ta.s || []).filter(x => (tb.s || []).includes(x)).forEach(x => P.push([.4 + S.idf('s:' + x) * .5, (td.s || {})[x] || '']));
  if (ta.d && tb.d) { const g = Math.abs(DORD[ta.d] - DORD[tb.d]); if (!g) P.push([1, '']); else if (g === 1) P.push([.3, '']) }
  if (a.pts === b.pts) P.push([.8, '']);
  (ta.f || []).filter(x => (tb.f || []).includes(x) && x !== 'mc' && x !== 'short').forEach(x => P.push([.3 + S.idf('f:' + x) * .2, '']));
  if (a.mc === b.mc) P.push([.2, '']);
  return [P.reduce((t, p) => t + p[0], 0), P.slice().sort((x, y) => y[0] - x[0]).map(p => p[1]).filter(Boolean).slice(0, 3)];
}
function pool(s, sel, taken) {
  return R.SD[s].items.filter(i => {
    if (taken.has(i.e.replace(/^(ko|en)_/, ''))) return false;
    if (s === 'ko' && i.sel && i.sel !== sel) return false;
    if (s === 'ma' && sel !== '확통' && i.n > 22) return false;
    return true;
  });
}
// recs: [{e, rec}] 시험 순서대로(한 학생·한 과목)
function plan(s, recs, taken, N) {
  const wrong = [], scores = {}; let sel = null, subj = null;
  for (const { e, rec } of recs) {
    if (!rec || !rec.ans.some(x => x)) continue;
    sel = rec.sel; subj = rec.subj;
    const g = grade(e, rec), lab = R.short(e);
    scores[e] = [g.filter(x => x.ok).length, g.length];
    for (const x of g) {
      if (x.ok) continue;
      const t = tg(s, x.id);
      wrong.push({ ...x, e, lab, r: t.r ?? .5, d: t.d || 'mid' });
    }
  }
  if (sel === null) return null;
  const area = {};
  wrong.forEach(w => (area[w.it.code] = area[w.it.code] || []).push(w));
  const info = {};
  for (const c in area) info[c] = { exams: [...new Set(area[c].map(w => w.e))].sort(), n: area[c].length };
  const rep = Object.keys(area).filter(c => info[c].exams.length >= 2)
    .sort((a, b) => info[b].exams.length - info[a].exams.length || info[b].n - info[a].n || (a < b ? -1 : 1));
  const allw = wrong.slice().sort((a, b) => DORD[a.d] - DORD[b.d] || info[b.it.code].exams.length - info[a.it.code].exams.length || info[b.it.code].n - info[a.it.code].n || a.r - b.r);
  const T = [], per = {}, seen = new Set();
  for (const cap of [3, 99]) for (const w of allw) {
    if (T.length >= N) break;
    if (seen.has(w.id) || (per[w.it.code] || 0) >= cap) continue;
    T.push(w); seen.add(w.id); per[w.it.code] = (per[w.it.code] || 0) + 1;
  }
  const P = pool(s, sel, taken), used = new Set(), usedpsg = new Set(), picks = [], PS = R.SD[s].psg;
  for (const z of T) {
    let best = null;
    for (const b of P) {
      if (used.has(b.id)) continue;
      const tb = tg(s, b.id), gap = tb.d ? DORD[tb.d] - DORD[z.d] : 0;
      if (gap > 1) continue;
      let [sc, why] = sim(s, z.it, b);
      if (gap === 1) sc -= 2.5;
      if (s === 'ko' && PS[b.id] && usedpsg.has(PS[b.id])) sc -= 3;
      if (!best || sc > best.sc) best = { sc, b, why };
    }
    if (best) { used.add(best.b.id); if (PS[best.b.id]) usedpsg.add(PS[best.b.id]); picks.push({ target: z, item: best.b, why: best.why, rep: rep.includes(z.it.code) }) }
  }
  return { sel, subj, scores, info, rep, picks };
}
function flags(st, files) {
  const out = [];
  for (const f of files) for (const s of SUBS) {
    const r = st.recs[f.e] && st.recs[f.e][s]; if (!r) continue;
    const a = r.ans.filter(x => /^[1-5]$/.test(x));
    if (!r.ans.some(x => x)) { out.push(`${R.short(f.e)} ${SN[s]} 답 없음`); continue }
    if (a.length >= 10) {
      const c = {}; a.forEach(x => c[x] = (c[x] || 0) + 1);
      const [k, v] = Object.entries(c).sort((x, y) => y[1] - x[1])[0];
      if (v / a.length >= .6) out.push(`${R.short(f.e)} ${SN[s]} 몰아 찍음?(${k}번 ${v}/${a.length})`);
    }
  }
  for (const f of files) if (!st.recs[f.e]) out.push(`${R.short(f.e)} 기록 없음`);
  return out;
}
// files: [{e, parsed}], 같은 시험이 두 번 있으면 뒤의 것을 쓴다
R.analyze = function (files, cls, opt) {
  files = files.filter(f => f.e).sort((a, b) => a.e.localeCompare(b.e));
  const taken = new Set(files.map(f => f.e)), st = {};
  for (const f of files) for (const r of f.parsed.rows) {
    if (r.cls !== cls) continue;
    const x = st[r.no] = st[r.no] || { cls, no: r.no, recs: {} };
    (x.recs[f.e] = x.recs[f.e] || {})[r.s] = r;
  }
  return Object.values(st).sort((a, b) => a.no - b.no).map(x => {
    x.plans = {};
    for (const s of opt.subs) x.plans[s] = plan(s, files.map(f => ({ e: f.e, rec: (x.recs[f.e] || {})[s] })), taken, opt.n);
    x.flags = flags(x, files);
    return x;
  });
};

// ---- 그림: 저장소 주소 받기 + 글줄 사이 빈 띠 찾기 ----
R.urlFor = null;   // async (paths[]) => {path: url}
const IMC = {};
function loadImg(url) {
  return new Promise(res => {
    const im = new Image(); im.crossOrigin = 'anonymous';
    im.onload = () => {
      let ink = null;
      try {
        const cv = document.createElement('canvas'); cv.width = im.naturalWidth; cv.height = im.naturalHeight;
        const cx = cv.getContext('2d'); cx.drawImage(im, 0, 0);
        const d = cx.getImageData(0, 0, cv.width, cv.height).data, w = cv.width; ink = new Uint32Array(cv.height);
        for (let y = 0; y < cv.height; y++) { let c = 0; for (let x = 0, o = y * w * 4; x < w; x++, o += 4) { if (d[o + 3] > 30 && (d[o] * 299 + d[o + 1] * 587 + d[o + 2] * 114) / 1000 < 170) c++ } ink[y] = c }
      } catch (e) { ink = null }
      res({ url, w: im.naturalWidth, h: im.naturalHeight, ink });
    };
    im.onerror = () => res(null);
    im.src = url;
  });
}
async function fetchImgs(paths, prog) {
  const need = [...new Set(paths)].filter(p => !(p in IMC));
  if (need.length) {
    const urls = await R.urlFor(need); let done = 0;
    const q = need.slice();
    const next = async () => { while (q.length) { const p = q.shift(); IMC[p] = urls[p] ? await loadImg(urls[p]) : null; done++; prog && prog(done, need.length) } };
    await Promise.all(Array.from({ length: 8 }, next));
  }
}
const qPath = id => `q/${imgKey(id)}.png`;
const sPaths = (s, id) => Array.from({ length: R.SD[s].scount[id] || 0 }, (_, k) => `s/${imgKey(id)}_${k + 1}.png`);
const LOG = [];
function findCut(im, miny, maxy, tag) {
  maxy = Math.min(maxy, im.h - 1); miny = Math.max(1, miny);
  if (im.ink) {
    const thr = Math.max(3, Math.floor(im.w * .006)), g = y => im.ink[y] <= thr;
    let y = maxy;
    while (y >= miny) {
      if (g(y)) { const e = y; while (y >= miny && g(y)) y--; if (e - y >= 3) return Math.floor((y + 1 + e) / 2) }
      y--;
    }
    let b = miny; for (let k = miny; k <= maxy; k++) if (im.ink[k] < im.ink[b]) b = k;
    LOG.push([tag, 'fallback', b]); return b;
  }
  LOG.push([tag, 'noink', maxy]); return maxy;
}

// ---- 2단 배치(단 너비 88mm, 단 높이 258mm) ----
const CW = 88, CAP = 258, GAP = 1.5;
class Flow {
  constructor(firstcap = CAP) { this.firstcap = firstcap; this.pages = []; this.newpage(); this.rem = firstcap }
  newpage() { this.pages.push([[], []]); this.ci = 0; this.rem = CAP }
  nextcol() { if (this.ci === 0) { this.ci = 1; this.rem = this.pages.length === 1 ? this.firstcap : CAP } else this.newpage() }
  put(h, mm) { this.pages[this.pages.length - 1][this.ci].push(h); this.rem -= mm }
  header(h, chars, keep = 30) { const mm = 4.6 * Math.max(1, Math.ceil(chars / 32)) + 5; if (this.rem < mm + keep) this.nextcol(); this.put(h, mm) }
  note(h) { this.header(h, 40, 20) }
  image(path, split, tag) {
    const im = IMC[path];
    if (!im) { this.note(`<div class="note">(그림을 불러오지 못했습니다: ${esc(path)})</div>`); return }
    const sc = CW / im.w; let y0 = 0;
    const piece = (a, b) => `<div class="cut" style="height:${((b - a) * sc).toFixed(2)}mm"><img src="${im.url}" style="margin-top:-${(a * sc).toFixed(2)}mm"></div>`;
    for (let guard = 0; guard < 200; guard++) {
      const hmm = (im.h - y0) * sc;
      if (hmm <= this.rem) { this.put(piece(y0, im.h), hmm + GAP); return }
      const whole = hmm <= CAP;
      if (((!split && whole) || this.rem < 40 || (whole && this.rem < hmm * .35)) && this.rem < CAP) { this.nextcol(); continue }
      const span = Math.floor(this.rem / sc), y = Math.max(y0 + 1, findCut(im, y0 + Math.floor(span * .45), y0 + span, tag));
      this.put(piece(y0, y), (y - y0) * sc + GAP); y0 = y; this.nextcol();
    }
  }
  html(foot) {
    const P = this.pages.filter(p => p[0].length || p[1].length);
    return P.map((p, k) => `<div class="pg"><div class="cols"><div class="col">${p[0].join('')}</div><div class="col c2">${p[1].join('')}</div></div><div class="foot">${esc(foot)} · ${k + 1}/${P.length}</div></div>`).join('');
  }
}
const A = it => { const a = String(it.ans); return it.mc && /^[1-5]$/.test(a) ? CIR[+a - 1] : a };

R.CSS = `@page{size:A4;margin:14mm 13mm}
#printArea{font-family:"NanumGothic","Malgun Gothic","Apple SD Gothic Neo",sans-serif;font-size:10.5pt;color:#222;background:#fff}
#printArea .stu{break-before:page}#printArea .stu:first-child{break-before:auto}
#printArea h1{font-size:17pt;margin:0 0 4px}#printArea h2{font-size:13.5pt;border-bottom:2px solid #333;padding-bottom:3px;margin:16px 0 8px}
#printArea .sub{color:#666;font-size:9.5pt}#printArea table{border-collapse:collapse;width:100%;margin:4px 0 10px;font-size:9.5pt}
#printArea td,#printArea th{border:1px solid #bbb;padding:3px 5px;text-align:left;vertical-align:top;white-space:normal;position:static;background:none;color:#222}
#printArea th{background:#f0f0f0}#printArea .rep{color:#b3261e;font-weight:bold}#printArea .ans td{text-align:center}#printArea .tagl{color:#555}
#printArea .note{font-size:9pt;color:#555;margin:2px 0;background:none;padding:0}
#printArea .pg{height:266mm;width:184mm;position:relative;break-before:page;overflow:hidden}
#printArea .cols{display:flex;justify-content:space-between;height:258mm}
#printArea .col{width:88mm;box-sizing:content-box}#printArea .c2{border-left:1px solid #ddd;padding-left:3mm}
#printArea .cut{width:88mm;overflow:hidden;margin-bottom:1.5mm}#printArea .cut img{width:88mm;display:block;max-width:none;margin-left:0;margin-right:0;margin-bottom:0}
#printArea .foot{position:absolute;left:0;right:0;bottom:0;text-align:center;font-size:8pt;color:#888}
#printArea .sech{font-size:13pt;font-weight:bold;border-bottom:2px solid #333;margin-bottom:6px}
#printArea .qh{background:#f4f6fa;border-left:4px solid #3b5bdb;padding:4px 8px;margin-bottom:6px;font-size:9.5pt}#printArea .qh b{font-size:11pt}
#printArea .sh{font-weight:bold;font-size:9.5pt;background:#f4f6fa;padding:2px 6px;margin:4px 0 3px}
@media screen{#printArea .pg,#printArea .first{box-shadow:0 1px 6px rgba(0,0,0,.18);margin:0 auto 14px;padding:14mm 13mm;box-sizing:content-box;background:#fff}#printArea .first{width:184mm}}
@media print{body>*:not(#printWrap){display:none!important}#printWrap{position:static!important;display:block!important;background:#fff!important;padding:0!important;overflow:visible!important}#printWrap .ptool{display:none!important}}`;

// 한 학생의 인쇄본 HTML(그림은 미리 받아 둔 상태여야 한다)
function studentHtml(x, files, title) {
  const who = `${x.cls}반 ${x.no}번`, H = [];
  const exs = files.map(f => f.e);
  H.push(`<div class="first"><h1>${title ? esc(title) + ' ' : ''}${who} 오답 보강 자료</h1><div class="sub">${esc(exs.map(e => R.short(e)).join(' · '))} 오답 분석 · 우선순위: ① 쉬운 문항(난이도 하→중→상) ② 같은 난이도면 여러 시험에서 반복해 틀린 영역</div>`);
  H.push(`<h2>한눈에 보기</h2><table><tr><th>과목</th>${exs.map(e => `<th>${R.short(e)}</th>`).join('')}<th>반복해서 틀린 영역 (틀린 시험 · 문항 수)</th></tr>`);
  for (const s of Object.keys(x.plans)) {
    const p = x.plans[s];
    if (!p) { H.push(`<tr><td>${SN[s]}</td><td colspan="${exs.length + 1}">응시 기록 없음</td></tr>`); continue }
    const U = R.SD[s].units;
    const reps = p.rep.slice(0, 8).map(c => `<span class="rep">${esc(U[c].name)}</span>(${p.info[c].exams.map(R.short).join('·')}, ${p.info[c].n})`).join(', ') || '없음 → 쉬운 문항부터';
    H.push(`<tr><td>${SN[s]}<br><span class="sub">${esc(p.subj)}</span></td>${exs.map(e => `<td>${p.scores[e] ? p.scores[e].join('/') : '-'}</td>`).join('')}<td>${reps}</td></tr>`);
  }
  H.push('</table><div class="note">점수 칸은 맞힌 문항/채점한 문항 수입니다. 영어는 읽기(18~45번)만, 수학 미적분·기하 선택자는 공통(1~22번)만 채점했습니다.</div>');
  for (const s of Object.keys(x.plans)) {
    const p = x.plans[s]; if (!p || !p.picks.length) continue;
    const U = R.SD[s].units;
    H.push(`<h2>${SN[s]} 보강 문항 목록</h2><table><tr><th>#</th><th>내가 틀린 문항</th><th>영역</th><th>보강 문항(출처)</th><th>비슷한 점</th></tr>`);
    p.picks.forEach((pk, k) => { const z = pk.target, it = pk.item; H.push(`<tr><td>${k + 1}</td><td>${z.lab} ${z.n}번 <span class="sub">(난이도 ${DN[z.d]})</span></td><td>${pk.rep ? '<span class="rep">반복</span> ' : ''}${esc(U[z.it.code].name)}</td><td>${esc(it.label)} ${it.n}번</td><td class="tagl">${esc(pk.why.join(' · '))}</td></tr>`) });
    H.push('</table>');
  }
  H.push('</div>');
  let F = new Flow();
  for (const s of Object.keys(x.plans)) {
    const p = x.plans[s]; if (!p || !p.picks.length) continue;
    const U = R.SD[s].units, PS = R.SD[s].psg, shown = new Set();
    if (F.pages[F.pages.length - 1][0].length || F.pages[F.pages.length - 1][1].length) F.newpage();
    F.header(`<div class="sech">${SN[s]} 보강 문항</div>`, 10);
    p.picks.forEach((pk, k) => {
      const z = pk.target, it = pk.item, ps = PS[it.id];
      const t = `${SN[s]} ${k + 1}. ${it.label} ${it.n}번 · ${U[it.code].name} — ${z.lab} ${z.n}번에서 틀린 유형`;
      F.header(`<div class="qh"><b>${SN[s]} ${k + 1}.</b> ${esc(it.label)} ${it.n}번 · ${esc(U[it.code].name)} — <span class="sub">${z.lab} ${z.n}번에서 틀린 유형</span></div>`, t.length, 40);
      if (ps && !shown.has(ps)) { F.image(qPath(ps), true, ps); shown.add(ps) }
      else if (ps) F.note(`<div class="note">※ 지문은 앞의 ${SN[s]} ${p.picks.findIndex(q => PS[q.item.id] === ps) + 1}번 문항과 같습니다.</div>`);
      F.image(qPath(it.id), false, it.id);
    });
  }
  H.push(F.html(who + ' 보강 문항'));
  const ans = ['<h2>정답</h2><table class="ans">'];
  for (const s of Object.keys(x.plans)) { const p = x.plans[s]; if (p && p.picks.length) ans.push(`<tr><th>${SN[s]}</th>${p.picks.map((pk, k) => `<td>${k + 1}. ${esc(A(pk.item))}</td>`).join('')}</tr>`) }
  ans.push('</table>');
  const rowsN = ans.length - 2;
  F = new Flow(CAP - (14 + rowsN * 8));
  F.header('<div class="sech">해설</div>', 4);
  for (const s of Object.keys(x.plans)) {
    const p = x.plans[s]; if (!p) continue;
    p.picks.forEach((pk, k) => {
      const t = `${SN[s]} ${k + 1}. (${pk.item.label} ${pk.item.n}번) 정답 ${A(pk.item)}`;
      F.header(`<div class="sh">${esc(t)}</div>`, t.length);
      const sp = sPaths(s, pk.item.id);
      if (!sp.length) F.note('<div class="note">(해설 그림 없음)</div>');
      sp.forEach(pth => F.image(pth, true, pth));
    });
  }
  // 정답표를 해설 첫 쪽 맨 위에 넣는다
  const sol = F.html(who + ' 정답·해설').replace('<div class="cols">', ans.join('') + '<div class="cols" style="height:' + (CAP - (14 + rowsN * 8)) + 'mm">');
  H.push(sol);
  return `<div class="stu">${H.join('')}</div>`;
}
// ---- 반 전체 통계(2026-11-17): 반 오답 분석지·학생 진단표가 같이 쓴다 ----
// 문항마다 반 응시·오답 수, 고른 답, 전국 오답률(태그 r, top이면 EBSi 실제 값). 학생마다 영역별 맞힌 수.
const pct = v => v == null || isNaN(v) ? '-' : Math.round(v * 100) + '%';
function mathPts(rec, g) {
  if (rec.s !== 'ma' || rec.sel !== '확통' || g.length < 30) return null;
  return g.reduce((t, x) => t + (x.ok ? +x.it.pts || 0 : 0), 0);
}
// 수학(확통) 예상 등급컷: 발표·예상 등급컷이 있으면 그것, 없으면 학생 화면과 같은 모델로 추정
const CUM = [.04, .11, .23, .40, .60, .77, .89, .96];
function cutsFor(e) {
  const M = window.CUTMODEL; if (!M) return null;
  if (M.cuts && M.cuts[e] && M.cuts[e].c.length === 8) return { c: M.cuts[e].c, src: M.cuts[e].src };
  const items = R.SD.ma.items.filter(i => i.e === e).sort((a, b) => a.n - b.n); if (items.length !== 30) return null;
  const TH = [], W = [], WE = []; let ws = 0, we = 0;
  for (let k = 0; k <= 200; k++) { const t = -5 + k * .05; TH.push(t); const w = Math.exp(-t * t / 2); W.push(w); ws += w; const z = (t - M.mu) / M.s, v = Math.exp(-z * z / 2); WE.push(v); we += v }
  const P = (c, b, t) => c + (1 - c) / (1 + Math.exp(-M.A * (t - b)));
  const wrongOf = it => { const o = M.odap[it.e] || {}; if (o[it.n] != null) return o[it.n]; const w15 = Math.min(...Object.values(o), 1) * .95; return Math.min(1 / (1 + Math.exp(-M.g[it.n - 1])), w15) };
  const bOf = (wrong, c) => { let lo = -8, hi = 8; for (let i = 0; i < 40; i++) { const m = (lo + hi) / 2; let p = 0; for (let j = 0; j < TH.length; j++) p += WE[j] / we * P(c, m, TH[j]); if (1 - p > wrong) hi = m; else lo = m } return (lo + hi) / 2 };
  let D = TH.map(() => { const a = new Float64Array(101); a[0] = 1; return a });
  for (const it of items) {
    const g = it.mc ? .2 : 0, b = bOf(wrongOf(it), g);
    D = D.map((row, j) => { const p = P(g, b, TH[j]), N = new Float64Array(101); for (let s = 0; s <= 100; s++) { if (!row[s]) continue; N[s] += row[s] * (1 - p); if (s + it.pts <= 100) N[s + it.pts] += row[s] * p } return N });
  }
  const dist = new Float64Array(101); D.forEach((row, j) => { for (let s = 0; s <= 100; s++) dist[s] += row[s] * W[j] / ws });
  const tail = new Float64Array(102); for (let s = 100; s >= 0; s--) tail[s] = tail[s + 1] + dist[s];
  return { c: CUM.map(q => { let s = 100; while (s > 0 && tail[s] < q) s--; return s }), src: '모델 추정' };
}
const gradeOf = (sc, c) => { for (let g = 0; g < 8; g++) if (sc >= c[g]) return g + 1; return 9 };
const CUTC = {};
R.grade = (e, sc) => { const c = e in CUTC ? CUTC[e] : (CUTC[e] = cutsFor(e)); return c ? { g: gradeOf(sc, c.c), src: c.src } : null };

// res: analyze 결과(한 반 전체)
R.classStats = function (res, files, subs) {
  files = files.filter(f => f.e).sort((a, b) => a.e.localeCompare(b.e));
  const CS = {};
  for (const s of subs) {
    const items = {}, exams = {}, area = {}, stu = {};
    for (const f of files) exams[f.e] = { n: 0, ok: 0, tot: 0, pts: [] };
    for (const x of res) {
      const me = stu[x.no] = { area: {}, ex: {} };
      for (const f of files) {
        const rec = (x.recs[f.e] || {})[s]; if (!rec || !rec.ans.some(v => v)) continue;
        const g = grade(f.e, rec), ok = g.filter(v => v.ok).length, pts = mathPts(rec, g);
        const E = exams[f.e]; E.n++; E.ok += ok; E.tot += g.length; if (pts != null) E.pts.push(pts);
        me.ex[f.e] = { ok, tot: g.length, pts, sel: rec.subj, wrong: g.filter(v => !v.ok) };
        for (const v of g) {
          const t = tg(s, v.id), I = items[v.id] = items[v.id] || { it: v.it, e: f.e, n: 0, wrong: 0, pick: {}, r: t.r, top: t.top, d: t.d || 'mid' };
          I.n++; if (!v.ok) { I.wrong++; if (/^[1-5]$/.test(v.mine)) I.pick[v.mine] = (I.pick[v.mine] || 0) + 1 }
          const c = v.it.code, A = area[c] = area[c] || { n: 0, ok: 0, exp: 0 }, M = me.area[c] = me.area[c] || { n: 0, ok: 0 };
          A.n++; M.n++; if (v.ok) { A.ok++; M.ok++ } A.exp += 1 - (t.r ?? .5);
        }
      }
    }
    CS[s] = { items, exams, area, stu };
  }
  return CS;
};
// 수업용 문항: 반 오답률 30% 이상이면서 전국보다 많이 틀린 순(전국 오답률 85% 이상 초고난도 제외), 영역당 2개, 비슷한 기출 1개씩
function lessonPicks(s, C, taken, N) {
  const L = Object.values(C.items).filter(I => I.n >= 5 && I.wrong / I.n >= .3 && (I.r ?? .5) < .85)
    .map(I => ({ ...I, cw: I.wrong / I.n, gap: I.wrong / I.n - (I.r ?? .5) }))
    .sort((a, b) => (b.gap + b.cw * .5) - (a.gap + a.cw * .5));
  const out = [], per = {}, used = new Set(), PS = R.SD[s].psg;
  for (const I of L) {
    if (out.length >= N) break;
    const c = I.it.code; if ((per[c] || 0) >= 2) continue; per[c] = (per[c] || 0) + 1;
    const sel = s === 'ma' ? '확통' : (I.it.sel || null);
    let best = null;
    for (const b of pool(s, sel, taken)) {
      if (used.has(b.id)) continue;
      const tb = tg(s, b.id), gap = tb.d ? DORD[tb.d] - DORD[I.d] : 0; if (gap > 1) continue;
      let [sc, why] = sim(s, I.it, b); if (gap === 1) sc -= 2.5;
      if (!best || sc > best.sc) best = { sc, b, why };
    }
    if (best) used.add(best.b.id);
    out.push({ I, sim: best && best.b, why: best ? best.why : [] });
  }
  return out;
}
const topPick = I => { const e = Object.entries(I.pick).sort((a, b) => b[1] - a[1])[0]; return e && e[1] >= 2 ? `${CIR[+e[0] - 1]} ${pct(e[1] / I.n)}` : '-' };
const bar = (v, ref) => `<span class="bar"><i style="width:${Math.round((v || 0) * 100)}%"></i>${ref != null ? `<u style="left:${Math.round(ref * 100)}%"></u>` : ''}</span>`;

// 반 오답 분석지(선생님용 요약 + 수업용 문항 묶음 + 정답·해설)
function classHtml(res, files, title, CS, N) {
  files = files.filter(f => f.e).sort((a, b) => a.e.localeCompare(b.e));
  const cls = res[0].cls, taken = new Set(files.map(f => f.e)), H = [], picks = {};
  H.push(`<div class="first"><h1>${title ? esc(title) + ' ' : ''}${cls}반 모의고사 오답 분석 <span class="sub">(선생님용)</span></h1><div class="sub">${esc(files.map(f => R.short(f.e)).join(' · '))} 마킹 리스트 기준 · ${res.length}명 · 전국 오답률은 EBSi 오답률(상위 15문항은 실제 값, 나머지는 추정)</div>`);
  for (const s of Object.keys(CS)) {
    const C = CS[s], U = R.SD[s].units; if (!Object.keys(C.items).length) continue;
    const P = picks[s] = lessonPicks(s, C, taken, N), inP = new Set(P.map(p => p.I.it.id));
    H.push(`<h2>${SN[s]}</h2><table><tr><th>시험</th><th>응시</th><th>반 평균 맞힌 문항</th>${s === 'ma' ? '<th>확통 응시자 평균 원점수</th>' : ''}</tr>`);
    for (const f of files) { const E = C.exams[f.e]; if (!E.n) { H.push(`<tr><td>${R.short(f.e)}</td><td colspan="3">기록 없음</td></tr>`); continue } H.push(`<tr><td>${R.short(f.e)}</td><td>${E.n}명</td><td>${(E.ok / E.n).toFixed(1)} / ${(E.tot / E.n).toFixed(0)}</td>${s === 'ma' ? `<td>${E.pts.length ? (E.pts.reduce((a, b) => a + b, 0) / E.pts.length).toFixed(1) + `점 (${E.pts.length}명)` : '-'}</td>` : ''}</tr>`) }
    H.push('</table>');
    const A = Object.entries(C.area).filter(([c, v]) => v.n >= 5).map(([c, v]) => ({ c, rate: v.ok / v.n, exp: v.exp / v.n, n: v.n })).sort((a, b) => (a.rate - a.exp) - (b.rate - b.exp)).slice(0, 6);
    H.push(`<table><tr><th>전국보다 약한 영역</th><th>푼 문항(연인원)</th><th>반 정답률</th><th>전국 예상</th><th>차이</th></tr>${A.map(a => `<tr><td>${esc(U[a.c].name)}</td><td>${a.n}</td><td>${bar(a.rate, a.exp)} ${pct(a.rate)}</td><td>${pct(a.exp)}</td><td class="${a.rate < a.exp ? 'rep' : ''}">${a.rate < a.exp ? '−' : '+'}${Math.abs(Math.round((a.rate - a.exp) * 100))}%p</td></tr>`).join('')}</table>`);
    const T = Object.values(C.items).filter(I => I.n >= 5).sort((a, b) => b.wrong / b.n - a.wrong / a.n).slice(0, 10);
    H.push(`<table><tr><th>많이 틀린 문항</th><th>영역</th><th>난이도</th><th>반 오답률</th><th>전국</th><th>많이 고른 오답</th></tr>${T.map(I => `<tr><td>${inP.has(I.it.id) ? '★ ' : ''}${R.short(I.e)} ${I.it.n}번</td><td>${esc(U[I.it.code].name)}</td><td>${DN[I.d]}</td><td>${pct(I.wrong / I.n)}</td><td>${pct(I.r)}${I.top ? '' : '<span class="sub">*</span>'}</td><td>${topPick(I)}</td></tr>`).join('')}</table>`);
    H.push(`<div class="note">★ = 뒤의 수업용 문항에 넣은 문항(반 오답률 30% 이상이면서 전국보다 많이 틀린 순, 전국 오답률 85% 이상 초고난도는 뺌, 영역당 2개). * 전국 오답률 추정값. 막대의 세로선 = 전국 예상 정답률.</div>`);
  }
  H.push('</div>');
  let F = new Flow(), cnt = 0;
  for (const s of Object.keys(picks)) {
    const P = picks[s], U = R.SD[s].units, PS = R.SD[s].psg, shown = new Set(); if (!P.length) continue;
    if (F.pages[F.pages.length - 1][0].length || F.pages[F.pages.length - 1][1].length) F.newpage();
    F.header(`<div class="sech">${SN[s]} 수업용 문항</div>`, 10);
    P.forEach((p, k) => {
      const it = p.I.it, t = `${SN[s]} ${k + 1}. ${R.short(p.I.e)} ${it.n}번 · ${U[it.code].name} — 반 오답률 ${pct(p.I.cw)} (전국 ${pct(p.I.r)})`;
      F.header(`<div class="qh"><b>${SN[s]} ${k + 1}.</b> ${esc(R.short(p.I.e))} ${it.n}번 · ${esc(U[it.code].name)} — <span class="sub">반 오답률 ${pct(p.I.cw)} (전국 ${pct(p.I.r)})</span></div>`, t.length, 40);
      const ps = PS[it.id]; if (ps && !shown.has(ps)) { F.image(qPath(ps), true, ps); shown.add(ps) }
      F.image(qPath(it.id), false, it.id); cnt++;
      if (p.sim) {
        const b = p.sim, t2 = `${SN[s]} ${k + 1}-1. 같은 유형 연습 · ${b.label} ${b.n}번 ${p.why.join(' · ')}`;
        F.header(`<div class="qh" style="border-left-color:#999"><b>${SN[s]} ${k + 1}-1.</b> 같은 유형 연습 · ${esc(b.label)} ${b.n}번 <span class="sub">${esc(p.why.join(' · '))}</span></div>`, t2.length, 40);
        const ps2 = PS[b.id]; if (ps2 && !shown.has(ps2)) { F.image(qPath(ps2), true, ps2); shown.add(ps2) }
        F.image(qPath(b.id), false, b.id);
      }
    });
  }
  if (!cnt) return `<div class="stu">${H.join('')}<p class="note">조건에 맞는 수업용 문항이 없습니다.</p></div>`;
  H.push(F.html(`${cls}반 수업용 문항`));
  const ans = ['<h2>정답</h2><table class="ans">'];
  for (const s of Object.keys(picks)) { const P = picks[s]; if (P.length) ans.push(`<tr><th>${SN[s]}</th>${P.map((p, k) => `<td>${k + 1}. ${esc(A(p.I.it))}${p.sim ? `<br>${k + 1}-1. ${esc(A(p.sim))}` : ''}</td>`).join('')}</tr>`) }
  ans.push('</table>');
  const rowsN = ans.length - 2, top = 14 + rowsN * 13;
  F = new Flow(CAP - top);
  F.header('<div class="sech">해설</div>', 4);
  for (const s of Object.keys(picks)) picks[s].forEach((p, k) => {
    for (const [it, lab] of [[p.I.it, `${k + 1}`], [p.sim, `${k + 1}-1`]]) {
      if (!it) continue;
      const t = `${SN[s]} ${lab}. (${it.label} ${it.n}번) 정답 ${A(it)}`;
      F.header(`<div class="sh">${esc(t)}</div>`, t.length);
      const sp = sPaths(s, it.id); if (!sp.length) F.note('<div class="note">(해설 그림 없음)</div>');
      sp.forEach(pth => F.image(pth, true, pth));
    }
  });
  H.push(F.html(`${cls}반 수업용 문항 정답·해설`).replace('<div class="cols">', ans.join('') + '<div class="cols" style="height:' + (CAP - top) + 'mm">'));
  return `<div class="stu">${H.join('')}</div>`;
}
R.classPaths = function (files, CS, N) {
  const taken = new Set(files.filter(f => f.e).map(f => f.e)), out = [];
  for (const s of Object.keys(CS)) for (const p of lessonPicks(s, CS[s], taken, N)) for (const it of [p.I.it, p.sim]) {
    if (!it) continue; const ps = R.SD[s].psg[it.id]; if (ps) out.push(qPath(ps)); out.push(qPath(it.id), ...sPaths(s, it.id));
  }
  return out;
};
// res: 반 전체 analyze 결과, N: 과목당 수업용 문항 수
R.renderClass = async function (res, files, title, N, prog) {
  LOG.length = 0;
  const subs = Object.keys((res[0] || {}).plans || {}), CS = R.classStats(res, files, subs);
  await fetchImgs(R.classPaths(files, CS, N), prog);
  return { html: classHtml(res, files, title, CS, N), log: LOG.slice() };
};

// 학생별 1쪽 진단표: 점수 흐름(반 평균과), 약한 영역(반 정답률과), 이번에 할 것, 선생님 메모 칸
function diagHtml(x, files, CS, title, withNote) {
  const who = `${x.cls}반 ${x.no}번`, H = [], todo = [];
  H.push(`<div class="first diag"><h1>${title ? esc(title) + ' ' : ''}${who} 모의고사 진단표</h1><div class="sub">${esc(files.map(f => R.short(f.e)).join(' · '))} 결과 · 막대 = 내 정답률, 빨간 세로선 = 반 정답률</div>`);
  for (const s of Object.keys(x.plans)) {
    const C = CS[s], me = C && C.stu[x.no], p = x.plans[s], U = R.SD[s].units;
    if (!me || !Object.keys(me.ex).length) { H.push(`<h2>${SN[s]}</h2><div class="note">응시 기록 없음</div>`); continue }
    const last = files.map(f => f.e).filter(e => me.ex[e]).pop(), sel = me.ex[last].sel;
    const cells = files.map(f => {
      const m = me.ex[f.e], E = C.exams[f.e]; if (!m) return `<td>결시</td>`;
      let g = ''; if (m.pts != null) { const gr = R.grade(f.e, m.pts); g = ` · <b>${m.pts}점</b>${gr ? ` (예상 ${gr.g}등급)` : ''}` }
      return `<td><b>${m.ok}</b>/${m.tot}문항${g}<br><span class="sub">반 평균 ${(E.ok / E.n).toFixed(1)}문항</span></td>`;
    }).join('');
    H.push(`<h2>${SN[s]} <span class="sub">${esc(sel || '')}</span></h2><table class="dg"><tr>${files.map(f => `<th>${R.short(f.e)}</th>`).join('')}</tr><tr>${cells}</tr></table>`);
    const ar = Object.entries(me.area).filter(([c, v]) => v.n >= 2).map(([c, v]) => ({ c, r: v.ok / v.n, n: v.n, cr: C.area[c].ok / C.area[c].n })).sort((a, b) => a.r - b.r || a.cr - b.cr);
    const weak = ar.filter(a => a.r < 1).slice(0, 3), strong = ar.filter(a => a.r >= .8 && a.n >= 3).sort((a, b) => b.r - a.r || b.n - a.n).slice(0, 3);
    H.push(`<table class="dg"><tr><th style="width:30%">약한 영역</th><th style="width:30%">내 정답률</th><th style="width:9%">반</th><th>틀린 문항</th></tr>${weak.map(a => `<tr><td>${(p && p.rep.includes(a.c)) ? '<span class="rep">반복</span> ' : ''}${esc(U[a.c].name)}</td><td>${bar(a.r, a.cr)} ${pct(a.r)} <span class="sub">(${a.n})</span></td><td>${pct(a.cr)}</td><td class="sub">${esc(files.filter(f => me.ex[f.e]).map(f => { const w = me.ex[f.e].wrong.filter(v => v.it.code === a.c).map(v => v.n); return w.length ? `${R.short(f.e)} ${w.join(',')}` : '' }).filter(Boolean).join(' / '))}</td></tr>`).join('') || '<tr><td colspan="4">두 문항 이상 푼 영역에서 틀린 곳이 없습니다.</td></tr>'}</table>`);
    if (strong.length) H.push(`<div class="note">잘하는 영역: ${strong.map(a => `${esc(U[a.c].name)} ${pct(a.r)}`).join(', ')}</div>`);
    // 이번에 할 것: 과목마다 하나(반복해서 틀린 영역 → 쉬운 문항 실수 → 가장 약한 영역 순)
    const lo = me.ex[last].wrong.filter(v => (tg(s, v.id).d || 'mid') === 'lo');
    const ref = c => { if (!withNote || !p) return ''; const k = p.picks.findIndex(q => q.target.it.code === c); return k >= 0 ? ` 보강 노트 ${SN[s]} ${k + 1}번부터 풀어 보세요.` : '' };
    if (p && p.rep.length) { const c = p.rep[0]; todo.push(`<b>${SN[s]}</b> '${esc(U[c].name)}' 영역을 ${p.info[c].exams.map(R.short).join('·')} 시험에서 계속 틀렸습니다.${ref(c)}`) }
    else if (lo.length >= 2) todo.push(`<b>${SN[s]}</b> ${R.short(last)} 시험의 쉬운 문항(난이도 하)에서 ${lo.length}개 틀렸습니다(${lo.map(v => v.n + '번').join(', ')}). 끝까지 읽고 검산하는 습관부터 잡으세요.`);
    else if (weak[0]) todo.push(`<b>${SN[s]}</b> '${esc(U[weak[0].c].name)}' 정답률이 ${pct(weak[0].r)}로 가장 낮습니다.${ref(weak[0].c)}`);
  }
  H.push(`<h2>이번에 할 것</h2><ol class="todo">${todo.slice(0, 3).map(t => `<li>${t}</li>`).join('') || '<li>틀린 문항을 다시 풀어 보세요.</li>'}</ol>`);
  if (x.flags.length) H.push(`<div class="note">확인할 점: ${x.flags.map(esc).join(', ')}</div>`);
  H.push('<div class="memo">선생님 메모</div></div>');
  return H.join('');
}
R.CSS += `
#printArea .bar{display:inline-block;position:relative;width:22mm;height:3mm;background:#e6e6e6;vertical-align:middle;margin-right:3px}#printArea .bar i{position:absolute;left:0;top:0;bottom:0;background:#555}#printArea .bar u{position:absolute;top:-1mm;bottom:-1mm;width:0;border-left:2px solid #b3261e}
#printArea .diag h2{margin:8px 0 3px;font-size:12pt}#printArea table.dg{margin:2px 0 3px;font-size:9pt}#printArea table.dg td,#printArea table.dg th{padding:1.5px 5px}
#printArea .todo{margin:4px 0 6px 0;padding-left:20px;line-height:1.55}#printArea .memo{border:1px solid #999;height:22mm;padding:2mm 3mm;font-size:9pt;color:#777;margin-top:6px}
#printArea .first{break-before:page}#printArea .stu>.first:first-child{break-before:auto}`;

R.paths = function (list) {
  const out = [];
  for (const x of list) for (const s of Object.keys(x.plans)) {
    const p = x.plans[s]; if (!p) continue;
    for (const pk of p.picks) { const ps = R.SD[s].psg[pk.item.id]; if (ps) out.push(qPath(ps)); out.push(qPath(pk.item.id), ...sPaths(s, pk.item.id)) }
  }
  return out;
};
// list: analyze 결과 중 인쇄할 학생들
// opt.all: 반 전체 analyze 결과(진단표의 반 평균용), opt.diag: 진단표를 맨 앞에, opt.only: 진단표만(상담용)
R.render = async function (list, files, title, prog, opt = {}) {
  LOG.length = 0;
  const all = opt.all || list, CS = (opt.diag || opt.only) ? R.classStats(all, files, Object.keys((all[0] || {}).plans || {})) : null;
  if (!opt.only) await fetchImgs(R.paths(list), prog);
  files = files.filter(f => f.e).sort((a, b) => a.e.localeCompare(b.e));
  const one = x => opt.only ? `<div class="stu">${diagHtml(x, files, CS, title, false)}</div>`
    : opt.diag ? studentHtml(x, files, title).replace('<div class="stu">', '<div class="stu">' + diagHtml(x, files, CS, title, true)) : studentHtml(x, files, title);
  return { html: list.map(one).join(''), log: LOG.slice() };
};
R.SN = SN; R.short = R.short; R.esc = esc;
})();
