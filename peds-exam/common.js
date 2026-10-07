/* 兒科題庫前後測：學員端與教師端共用的工具 */
(function () {
  'use strict';
  const CFG = window.PEDS_CONFIG || {};

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const LS = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} },
    del(k) { try { localStorage.removeItem(k); } catch (e) {} }
  };
  const SS = {
    get(k) { try { return sessionStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { sessionStorage.setItem(k, v); } catch (e) {} }
  };

  /* ---------- 題庫 ---------- */
  let banksCache = null;
  const bankCache = {};
  async function loadBanks() {
    if (banksCache) return banksCache;
    const r = await fetch(CFG.BANKS || 'data/banks.json', { cache: 'no-cache' });
    if (!r.ok) throw new Error('讀不到題庫清單（' + r.status + '）');
    banksCache = await r.json();
    return banksCache;
  }
  async function loadBank(id) {
    if (bankCache[id]) return bankCache[id];
    const list = await loadBanks();
    const meta = list.find(b => b.id === id);
    if (!meta) throw new Error('找不到題庫「' + id + '」');
    const r = await fetch(meta.file, { cache: 'no-cache' });
    if (!r.ok) throw new Error('讀不到題庫檔案（' + r.status + '）');
    const bank = await r.json();
    bank.letters = bank.optionLetters || 'ABCDE';
    bank.byNo = {};
    bank.questions.forEach(q => { bank.byNo[q.no] = q; });
    bankCache[id] = bank;
    return bank;
  }

  /* ---------- 亂數（可重現） ---------- */
  function hashStr(s) {   // FNV-1a 32-bit
    let h = 0x811c9dc5;
    for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 0x01000193); }
    return h >>> 0;
  }
  function rng(seed) {     // mulberry32
    let a = (typeof seed === 'number' ? seed : hashStr(String(seed))) >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function shuffle(arr, rand) {
    const a = arr.slice(); rand = rand || Math.random;
    for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(rand() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    return a;
  }

  /* ---------- 考卷：題目清單直接放在連結裡，前測與後測用同一串就一定是同一卷 ---------- */
  function paperCode(bankId, nums) {
    return hashStr(bankId + '|' + nums.join('.')).toString(36).toUpperCase().padStart(7, '0').slice(-7);
  }
  function parsePaper(search) {
    const p = new URLSearchParams(search);
    const nums = String(p.get('q') || '').split(/[.,\s]+/).map(Number).filter(n => Number.isInteger(n) && n > 0);
    const m = p.get('m') === 'post' ? 'post' : (p.get('m') === 'pre' ? 'pre' : 'practice');
    return {
      b: p.get('b') || '',
      nums,
      k: (p.get('k') || '').slice(0, 60),
      m,
      t: (p.get('t') || '').slice(0, 80),
      sq: p.get('sq') === '1',
      so: p.get('so') === '1',
      r: p.get('r') === '1'
    };
  }
  function paperQuery(paper) {
    const p = new URLSearchParams();
    p.set('b', paper.b);
    p.set('q', paper.nums.join('.'));
    if (paper.k) p.set('k', paper.k);
    p.set('m', paper.m);
    if (paper.t) p.set('t', paper.t);
    if (paper.sq) p.set('sq', '1');
    if (paper.so) p.set('so', '1');
    if (paper.r) p.set('r', '1');
    return '?' + p.toString();
  }
  const PHASE = { pre: '前測', post: '後測', practice: '練習' };

  /* 依科別抽題。mode: prop＝依題庫比例、equal＝每科平均、propmin＝依比例且每科至少 1 題 */
  function drawPaper(questions, n, mode, seed) {
    const rand = rng(seed == null ? Date.now() : seed);
    const byCat = {};
    questions.forEach(q => { (byCat[q.category] = byCat[q.category] || []).push(q); });
    const cats = Object.keys(byCat);
    n = Math.max(1, Math.min(n, questions.length));
    const quota = {};
    cats.forEach(c => { quota[c] = 0; });
    if (mode === 'equal') {
      let left = n, open = cats.slice();
      while (left > 0 && open.length) {
        open = shuffle(open, rand);
        for (const c of open) { if (!left) break; if (quota[c] < byCat[c].length) { quota[c]++; left--; } }
        open = open.filter(c => quota[c] < byCat[c].length);
      }
    } else {
      if (mode === 'propmin') cats.forEach(c => { if (n >= cats.length) quota[c] = 1; });
      const base = cats.reduce((s, c) => s + quota[c], 0);
      const left = n - base;
      const raw = cats.map(c => ({ c, x: left * byCat[c].length / questions.length }));
      raw.forEach(r => { quota[r.c] += Math.floor(r.x); });
      let rest = n - cats.reduce((s, c) => s + quota[c], 0);
      raw.sort((a, b) => (b.x - Math.floor(b.x)) - (a.x - Math.floor(a.x)) || rand() - 0.5);
      for (const r of raw) { if (rest <= 0) break; if (quota[r.c] < byCat[r.c].length) { quota[r.c]++; rest--; } }
      while (rest > 0) {   // 有科別題數不夠時，從還有剩的科別補
        const c = cats.find(x => quota[x] < byCat[x].length);
        if (!c) break; quota[c]++; rest--;
      }
      cats.forEach(c => { quota[c] = Math.min(quota[c], byCat[c].length); });
    }
    const picked = [];
    cats.forEach(c => { shuffle(byCat[c], rand).slice(0, quota[c]).forEach(q => picked.push(q.no)); });
    return picked.sort((a, b) => a - b);
  }

  /* ---------- 統計 ---------- */
  const mean = a => a.length ? a.reduce((s, x) => s + x, 0) / a.length : NaN;
  function sd(a) {
    if (a.length < 2) return NaN;
    const m = mean(a);
    return Math.sqrt(a.reduce((s, x) => s + (x - m) * (x - m), 0) / (a.length - 1));
  }
  function lgamma(x) {
    const c = [76.18009172947146, -86.50532032941677, 24.01409824083091, -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
    let y = x, tmp = x + 5.5;
    tmp -= (x + 0.5) * Math.log(tmp);
    let ser = 1.000000000190015;
    for (let j = 0; j < 6; j++) ser += c[j] / ++y;
    return -tmp + Math.log(2.5066282746310005 * ser / x);
  }
  function betacf(a, b, x) {
    const MAXIT = 200, EPS = 3e-14, FPMIN = 1e-300;
    let qab = a + b, qap = a + 1, qam = a - 1, c = 1, d = 1 - qab * x / qap;
    if (Math.abs(d) < FPMIN) d = FPMIN;
    d = 1 / d; let h = d;
    for (let m = 1; m <= MAXIT; m++) {
      const m2 = 2 * m;
      let aa = m * (b - m) * x / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; h *= d * c;
      aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
      c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
      d = 1 / d; const del = d * c; h *= del;
      if (Math.abs(del - 1) < EPS) break;
    }
    return h;
  }
  function ibeta(a, b, x) {   // 正規化不完全貝他函數 I_x(a,b)
    if (x <= 0) return 0; if (x >= 1) return 1;
    const bt = Math.exp(lgamma(a + b) - lgamma(a) - lgamma(b) + a * Math.log(x) + b * Math.log(1 - x));
    return x < (a + 1) / (a + b + 2) ? bt * betacf(a, b, x) / a : 1 - bt * betacf(b, a, 1 - x) / b;
  }
  function tP2(t, df) {       // 雙尾 p
    if (!isFinite(t)) return t === t ? 0 : NaN;
    return ibeta(df / 2, 0.5, df / (df + t * t));
  }
  function pairedT(pre, post) {
    const d = pre.map((x, i) => post[i] - x);
    const n = d.length, md = mean(d), s = sd(d);
    const out = { n, mPre: mean(pre), mPost: mean(post), mDiff: md, sd: s, t: NaN, df: n - 1, p: NaN, dz: NaN };
    if (n >= 2 && s > 0) { out.t = md / (s / Math.sqrt(n)); out.p = tP2(out.t, n - 1); out.dz = md / s; }
    else if (n >= 2 && s === 0) { out.dz = md === 0 ? 0 : Infinity; }
    return out;
  }

  /* ---------- 報表 ---------- */
  const normKey = s => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
  const personKey = r => normKey(r.pid) || ('名:' + normKey(r.name));

  /* rows：後端回傳的作答紀錄（同一梯次）。bank：題庫（用來補科別、題幹） */
  function buildReport(rows, bank) {
    const L = bank ? bank.letters : 'ABCDE';
    const sorted = rows.slice().sort((a, b) => String(a.submittedAt).localeCompare(String(b.submittedAt)));
    const first = { pre: {}, post: {} }, dupCount = { pre: 0, post: 0 };
    const papers = { pre: {}, post: {} };
    sorted.forEach(r => {
      if (r.phase !== 'pre' && r.phase !== 'post') return;
      const k = personKey(r);
      if (first[r.phase][k]) { dupCount[r.phase]++; return; }   // 同一人重考：以第一次為準
      first[r.phase][k] = r;
      papers[r.phase][r.paperCode] = (papers[r.phase][r.paperCode] || 0) + 1;
    });
    const pct = r => r.total ? r.score / r.total * 100 : 0;

    // 個人
    const keys = Array.from(new Set(Object.keys(first.pre).concat(Object.keys(first.post))));
    const people = keys.map(k => {
      const a = first.pre[k], b = first.post[k], any = a || b;
      return {
        key: k, name: any.name, pid: any.pid, hospital: any.hospital, level: any.level,
        pre: a ? pct(a) : null, post: b ? pct(b) : null,
        preRaw: a ? a.score + '/' + a.total : '', postRaw: b ? b.score + '/' + b.total : '',
        diff: a && b ? pct(b) - pct(a) : null
      };
    }).sort((x, y) => String(x.hospital).localeCompare(String(y.hospital)) || String(x.name).localeCompare(String(y.name)));
    const paired = people.filter(p => p.pre != null && p.post != null);
    const pairStat = pairedT(paired.map(p => p.pre), paired.map(p => p.post));

    // 逐題與科別
    const item = {}, cat = {};
    const touch = (no, q) => item[no] || (item[no] = {
      no, q, category: q ? q.category : '?',
      pre: { n: 0, ok: 0, blank: 0, c: {} }, post: { n: 0, ok: 0, blank: 0, c: {} }
    });
    ['pre', 'post'].forEach(ph => {
      Object.values(first[ph]).forEach(r => {
        (r.answers || []).forEach(a => {
          const q = bank ? bank.byNo[a.no] : null;
          const it = touch(a.no, q);
          if (!q && a.cat) it.category = a.cat;
          const s = it[ph];
          s.n++;
          if (!a.picked) s.blank++; else s.c[a.picked] = (s.c[a.picked] || 0) + 1;
          const ans = q ? q.answer : a.ans;
          const ok = !!a.picked && a.picked === ans;
          if (ok) s.ok++;
          const cname = it.category;
          const cs = cat[cname] || (cat[cname] = { name: cname, pre: { n: 0, ok: 0 }, post: { n: 0, ok: 0 }, items: new Set() });
          cs[ph].n++; if (ok) cs[ph].ok++; cs.items.add(a.no);
        });
      });
    });
    const items = Object.values(item).map(it => {
      const ans = it.q ? it.q.answer : '';
      const trap = ph => {
        let t = null, tn = 0;
        L.split('').forEach(x => { if (x !== ans && (it[ph].c[x] || 0) > tn) { t = x; tn = it[ph].c[x]; } });
        return { letter: t, n: tn };
      };
      return Object.assign(it, {
        answer: ans,
        preRate: it.pre.n ? it.pre.ok / it.pre.n * 100 : null,
        postRate: it.post.n ? it.post.ok / it.post.n * 100 : null,
        preTrap: trap('pre'), postTrap: trap('post')
      });
    }).sort((a, b) => a.no - b.no);

    const overall = ph => {
      const n = Object.values(cat).reduce((s, c) => s + c[ph].n, 0);
      const ok = Object.values(cat).reduce((s, c) => s + c[ph].ok, 0);
      return n ? ok / n * 100 : null;
    };
    const ovPre = overall('pre'), ovPost = overall('post');
    const ovDiff = ovPre != null && ovPost != null ? ovPost - ovPre : null;
    const cats = Object.values(cat).map(c => {
      const pre = c.pre.n ? c.pre.ok / c.pre.n * 100 : null;
      const post = c.post.n ? c.post.ok / c.post.n * 100 : null;
      const diff = pre != null && post != null ? post - pre : null;
      return {
        name: c.name, items: c.items.size, preN: c.pre.n, postN: c.post.n, pre, post, diff,
        net: diff != null && ovDiff != null ? diff - ovDiff : null,
        // 單一科別答對率的 95% 信賴區間半寬（常態近似），提醒「每科題數少時排名多半是雜訊」
        preCI: c.pre.n ? 196 * Math.sqrt(Math.max(pre / 100 * (1 - pre / 100), 0.25 / c.pre.n) / c.pre.n) : null
      };
    }).sort((a, b) => (a.pre == null ? 999 : a.pre) - (b.pre == null ? 999 : b.pre));

    const scores = ph => Object.values(first[ph]).map(pct);
    return {
      people, paired, pairStat, items, cats,
      overall: { pre: ovPre, post: ovPost, diff: ovDiff },
      count: { pre: Object.keys(first.pre).length, post: Object.keys(first.post).length },
      scores: { pre: scores('pre'), post: scores('post') },
      dupCount, papers,
      onlyPre: people.filter(p => p.pre != null && p.post == null),
      onlyPost: people.filter(p => p.pre == null && p.post != null)
    };
  }

  function toCSV(rows) {
    return '﻿' + rows.map(r => r.map(v => {
      const s = v == null ? '' : String(v);
      return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    }).join(',')).join('\r\n');
  }
  function download(name, text, type) {
    const blob = new Blob([text], { type: type || 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  const fmt = (x, d) => x == null || isNaN(x) ? '—' : (+x).toFixed(d == null ? 0 : d);

  async function getJSON(url, ms) {
    const ac = new AbortController(); const t = setTimeout(() => ac.abort(), ms || 15000);
    try { const r = await fetch(url, { signal: ac.signal }); return await r.json(); } finally { clearTimeout(t); }
  }

  window.Peds = {
    CFG, esc, LS, SS, loadBanks, loadBank, hashStr, rng, shuffle,
    paperCode, parsePaper, paperQuery, PHASE, drawPaper,
    mean, sd, tP2, pairedT, buildReport, personKey, toCSV, download, fmt, getJSON
  };
})();
