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
R.paths = function (list) {
  const out = [];
  for (const x of list) for (const s of Object.keys(x.plans)) {
    const p = x.plans[s]; if (!p) continue;
    for (const pk of p.picks) { const ps = R.SD[s].psg[pk.item.id]; if (ps) out.push(qPath(ps)); out.push(qPath(pk.item.id), ...sPaths(s, pk.item.id)) }
  }
  return out;
};
// list: analyze 결과 중 인쇄할 학생들
R.render = async function (list, files, title, prog) {
  LOG.length = 0;
  await fetchImgs(R.paths(list), prog);
  files = files.filter(f => f.e).sort((a, b) => a.e.localeCompare(b.e));
  return { html: list.map(x => studentHtml(x, files, title)).join(''), log: LOG.slice() };
};
R.SN = SN; R.short = R.short; R.esc = esc;
})();
