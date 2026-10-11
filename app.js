'use strict';
/* RECIE dashboard — reads recie_project_data.xlsx and renders every section by sector and governorate. */

// Data file names (Netlify paths are case-sensitive, so both spellings are tried)
const FILES = ['recie_project_data.xlsx', 'RECIE_Project_Data.xlsx'];
const CACHE_KEY = 'recie-data-v3';
const SHEETS = ['Config', 'Sectors', 'Governorates', 'Indicators', 'Indicator_Data', 'Activities', 'Activity_Data', 'Allocations', 'Updates'];
const XLSX_FALLBACK = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
const $ = id => document.getElementById(id);
const qs = new URLSearchParams(location.search);
const store = {
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* storage disabled or full */ } }
};

const T = {
  ar: {
    skip: 'انتقل إلى المحتوى',
    nav: ['نظرة عامة', 'لوحة المتابعة', 'مناطق الاستجابة', 'المؤشرات', 'الأنشطة', 'آخر المستجدات'],
    live: 'مباشر', cached: 'بيانات محفوظة',
    refresh: 'تحديث البيانات', print: 'طباعة أو حفظ PDF', theme: 'تبديل المظهر', top: 'العودة للأعلى',
    over: 'استجابة إنسانية متعددة القطاعات',
    desc: 'استجابة متكاملة لدعم الأطفال والأسر المتأثرة بالأزمات عبر التعليم والحماية والمياه والإصحاح والمشاركة المجتمعية.',
    donor: 'المانح', partner: 'الشريك المنفذ', duration: 'مدة المشروع', start: 'البداية', end: 'النهاية', updated: 'آخر تحديث',
    elapsed: 'المدة المنقضية', upcoming: 'قادم', ongoing: 'جارٍ', completed: 'مكتمل',
    daysLeft: n => `متبقٍ ${n} يوم`, daysTo: n => `يبدأ بعد ${n} يوم`,
    impactT: 'الأثر بالأرقام', impactS: 'تتحدث تلقائياً من ملف البيانات',
    impact: ['الأفراد المستهدفون', 'الأفراد الذين تم الوصول إليهم', 'إجمالي الميزانية', 'المحافظات', 'القطاعات', 'الأنشطة'],
    filterT: 'استكشف البيانات حسب القطاع والمحافظة',
    filterD: 'اختر قطاعاً أو محافظة، أو انقر على أي رسم أو بطاقة أو خلية لتصفية التقرير كله.',
    sector: 'القطاع', gov: 'المحافظة', reset: 'إعادة الضبط', context: 'عرض بيانات', filtering: 'التصفية الحالية', clear: 'إلغاء التصفية',
    peopleTarget: 'هدف الأفراد', peopleActual: 'الأفراد المحققون', activitiesTarget: 'هدف الأنشطة', activitiesActual: 'الأنشطة المحققة',
    ofTarget: 'من الهدف', remaining: 'متبقٍ',
    overallT: 'نسبة الإنجاز الكلية', overallS: 'الأفراد الذين تم الوصول إليهم مقارنة بالهدف',
    bySectorT: 'الوصول حسب القطاع', legendT: 'الهدف', legendA: 'المحقق',
    budgetT: 'توزيع الميزانية حسب القطاع', budgetS: 'بالدولار الأمريكي', total: 'الإجمالي',
    profBudget: 'ميزانية المحافظة', profInd: 'مؤشرات بلغت 75% أو أكثر', profAct: 'أنشطة بدأ تنفيذها', profNote: 'ملاحظة',
    areasT: 'مناطق الاستجابة', areasD: 'قارن الإنجاز بين المحافظات والقطاعات، وانقر على أي محافظة أو خلية لعرض بياناتها في كامل التقرير.',
    heatT: 'نسبة الوصول حسب المحافظة والقطاع', heatFull: '100% فأكثر', heatNone: 'لا هدف', heatOver: 'تجاوز الهدف',
    reached: 'تم الوصول', of: 'من', activitiesShort: 'نشاط', budgetShort: 'الميزانية',
    indT: 'المؤشرات', indD: 'تقدم كل مؤشر مقارنة بهدفه. افتح «حسب المحافظة» لرؤية مساهمة كل محافظة.',
    byGov: 'حسب المحافظة', target: 'الهدف', actual: 'الإنجاز',
    tags: { done: 'تحقق الهدف', ok: 'على المسار', mid: 'قيد التقدم', low: 'يحتاج تسريع', none: 'لم يبدأ' },
    actT: 'الأنشطة والميزانيات', actD: 'الأنشطة بأهدافها وإنجازها وميزانياتها. اختر محافظة لرؤية أنشطتها، أو افتح أي نشاط لرؤية توزيعه على المحافظات.',
    search: 'ابحث عن نشاط...', sortBudget: 'الميزانية', sortTarget: 'الهدف', sortName: 'الاسم', exportCsv: 'تصدير CSV', expand: 'عرض المحافظات',
    thAct: 'النشاط', thSector: 'القطاع', thTarget: 'الهدف', thActual: 'المنجز', thBudget: 'الميزانية', thStatus: 'الحالة', share: 'من الميزانية',
    status: { Planned: 'مخطط', Ongoing: 'قيد التنفيذ', Completed: 'مكتمل' },
    upT: 'آخر المستجدات والأخبار', upD: 'أحدث الأخبار والتحديثات المنشورة حول المشروع.',
    more: 'قراءة المزيد', none: 'لا توجد بيانات مطابقة.', noNews: 'لا توجد مستجدات منشورة.',
    noticeCache: 'تعذّر تحميل ملف البيانات؛ يتم عرض آخر نسخة محفوظة.', noticeFail: 'تعذّر تحميل ملف البيانات. يرجى التحقق من الاتصال ثم الضغط على زر التحديث.',
    warnT: 'تنبيه في ملف البيانات', warnMissing: 'الورقة غير موجودة', warnUnknown: 'معرّف غير معروف', warnMore: n => `و${n} تنبيهات أخرى`,
    footer: 'تنفذ منظمة صناع النهضة المشروع بالشراكة مع منظمة الأمم المتحدة للطفولة - يونيسف.'
  },
  en: {
    skip: 'Skip to content',
    nav: ['Overview', 'Dashboard', 'Response Areas', 'Indicators', 'Activities', 'Latest Updates'],
    live: 'Live', cached: 'Cached data',
    refresh: 'Refresh data', print: 'Print or save as PDF', theme: 'Toggle theme', top: 'Back to top',
    over: 'Multi-sector Humanitarian Response',
    desc: 'An integrated response supporting crisis-affected children and families through education, protection, WASH and community engagement.',
    donor: 'Donor', partner: 'Implementing Partner', duration: 'Project Duration', start: 'Start', end: 'End', updated: 'Last updated',
    elapsed: 'Time elapsed', upcoming: 'Upcoming', ongoing: 'Ongoing', completed: 'Completed',
    daysLeft: n => `${n} days left`, daysTo: n => `Starts in ${n} days`,
    impactT: 'Impact at a glance', impactS: 'Updated automatically from the data file',
    impact: ['People targeted', 'People reached', 'Total budget', 'Governorates', 'Sectors', 'Activities'],
    filterT: 'Explore data by sector and governorate',
    filterD: 'Pick a sector or governorate, or click any chart, card or cell to filter the whole report.',
    sector: 'Sector', gov: 'Governorate', reset: 'Reset', context: 'Showing', filtering: 'Current filter', clear: 'Clear filter',
    peopleTarget: 'People target', peopleActual: 'People reached', activitiesTarget: 'Activity target', activitiesActual: 'Activities achieved',
    ofTarget: 'of target', remaining: 'remaining',
    overallT: 'Overall achievement', overallS: 'People reached compared with target',
    bySectorT: 'Reach by sector', legendT: 'Target', legendA: 'Reached',
    budgetT: 'Budget by sector', budgetS: 'In US dollars', total: 'Total',
    profBudget: 'Governorate budget', profInd: 'Indicators at 75% or more', profAct: 'Activities started', profNote: 'Note',
    areasT: 'Response Areas', areasD: 'Compare achievement across governorates and sectors. Click a governorate or a cell to focus the whole report on it.',
    heatT: 'Reach by governorate and sector', heatFull: '100%+', heatNone: 'No target', heatOver: 'Target exceeded',
    reached: 'Reached', of: 'of', activitiesShort: 'activities', budgetShort: 'Budget',
    indT: 'Indicators', indD: 'Progress of each indicator against its target. Open “By governorate” to see each governorate’s share.',
    byGov: 'By governorate', target: 'Target', actual: 'Actual',
    tags: { done: 'Target met', ok: 'On track', mid: 'In progress', low: 'Needs acceleration', none: 'Not started' },
    actT: 'Activities and Budgets', actD: 'Activities with targets, achievement and budgets. Pick a governorate to see its activities, or open any activity to see its split by governorate.',
    search: 'Search activities...', sortBudget: 'Budget', sortTarget: 'Target', sortName: 'Name', exportCsv: 'Export CSV', expand: 'Show governorates',
    thAct: 'Activity', thSector: 'Sector', thTarget: 'Target', thActual: 'Achieved', thBudget: 'Budget', thStatus: 'Status', share: 'of budget',
    status: { Planned: 'Planned', Ongoing: 'Ongoing', Completed: 'Completed' },
    upT: 'Latest News and Updates', upD: 'The latest published news and updates about the project.',
    more: 'Read more', none: 'No matching data.', noNews: 'No updates are published.',
    noticeCache: 'The data file could not be loaded; showing the latest saved copy.', noticeFail: 'The data file could not be loaded. Check your connection and press refresh.',
    warnT: 'Data file warning', warnMissing: 'Sheet not found', warnUnknown: 'Unknown ID', warnMore: n => `and ${n} more`,
    footer: 'Nahda Makers Organization implements the project in partnership with the United Nations Children’s Fund - UNICEF.'
  }
};

let lang = qs.get('lang') || store.get('lang') || 'ar';
if (!T[lang]) lang = 'ar';
let D = null, X = null, sector = qs.get('sector') || 'ALL', gov = qs.get('gov') || 'ALL';
let fromCache = false, actSort = 'budget', actQuery = '', lastActRows = [];
const prevValues = new Map();
const reduced = matchMedia('(prefers-reduced-motion: reduce)').matches;

/* ---------- helpers ---------- */
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const loc = () => lang === 'ar' ? 'ar-YE-u-nu-latn' : 'en-US'; // Western digits in both languages, matching the dates in the data file
const nfCache = {};
const nf = (key, opts) => (nfCache[loc() + key] ||= new Intl.NumberFormat(loc(), opts));
const num = x => nf('n', { maximumFractionDigits: 0 }).format(+x || 0);
const usd = x => '\u200E$' + num(x); // LRM keeps the symbol attached to the number inside right-to-left text
const pctf = x => (x > 0 && x < 0.5 ? '\u200E<1' : num(x)) + '%';
const pct = (a, b) => (+b ? (+a || 0) / +b * 100 : 0);          // real percentage (can exceed 100)
const cap = p => Math.max(0, Math.min(100, p));                  // for bars and rings
const text = (id, x) => { const el = $(id); if (el) el.textContent = x; };
const html = (id, x) => { const el = $(id); if (el) el.innerHTML = x; };
const name = x => x ? x[lang === 'ar' ? 'Arabic' : 'English'] || x.Arabic || x.English || '' : '';
const cfg = k => { const x = D?.Config.find(r => r.Key === k); return x ? (lang === 'ar' ? x.Arabic : x.English) : ''; };
const cfgEn = k => D?.Config.find(r => r.Key === k)?.English;
const sum = (arr, k) => arr.reduce((s, x) => s + (+x[k] || 0), 0);
const level = p => p >= 100 ? 'done' : p >= 75 ? 'ok' : p >= 40 ? 'mid' : p > 0 ? 'low' : 'none';
const sectorOf = id => X.sec.get(id);
const colorOf = id => sectorOf(id)?.Color || '#009fe3';
const govOf = id => X.gov.get(id);
const secList = () => D.Sectors.filter(s => s.ID !== 'ALL');
const govList = () => D.Governorates.filter(g => g.ID !== 'ALL');

function parseDate(v) {
  if (typeof v === 'number') return new Date(Math.round((v - 25569) * 864e5));
  const m = String(v || '').match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
  if (!m) { const d = new Date(v); return isNaN(d) ? null : d; }
  const mi = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'].indexOf(m[2].slice(0, 3).toLowerCase());
  return mi < 0 ? null : new Date(+m[3], mi, +m[1]);
}

/* ---------- scoped data access ---------- */
// g / s default to the active filters; pass 'ALL' to ignore one of them.
const alloc = (g = gov, s = sector) => D.Allocations.filter(x => (g === 'ALL' || x.GovernorateID === g) && (s === 'ALL' || x.SectorID === s));
const actRows = (g = gov, s = sector) => D.Activity_Data.filter(x => (g === 'ALL' || x.GovernorateID === g) && (s === 'ALL' || X.act.get(x.ActivityID)?.SectorID === s));
const indRows = (g = gov, s = sector) => D.Indicator_Data.filter(x => (g === 'ALL' || x.GovernorateID === g) && (s === 'ALL' || X.ind.get(x.IndicatorID)?.SectorID === s));
const aggStatus = rows => {
  if (!rows.length) return 'Planned';
  const st = rows.map(r => String(r.Status || 'Planned'));
  return st.every(s => s === 'Completed') ? 'Completed' : st.every(s => s === 'Planned') ? 'Planned' : 'Ongoing';
};

/* ---------- animations ---------- */
function animateNumbers(root = document) {
  root.querySelectorAll('[data-count]').forEach(el => {
    const to = +el.dataset.count, key = el.dataset.k, fmt = el.dataset.fmt === 'usd' ? usd : el.dataset.fmt === 'pct' ? pctf : num;
    const from = key && prevValues.has(key) ? prevValues.get(key) : 0;
    if (key) prevValues.set(key, to);
    if (reduced || from === to) { el.textContent = fmt(to); return; }
    const t0 = performance.now(), dur = 900;
    const step = now => {
      const p = Math.min(1, (now - t0) / dur), e = 1 - Math.pow(1 - p, 3);
      el.textContent = fmt(from + (to - from) * e);
      if (p < 1) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  });
}
// Bars and rings start empty and grow on the next frame so the CSS transitions run
function animateBars(root = document) {
  requestAnimationFrame(() => requestAnimationFrame(() => {
    root.querySelectorAll('[data-w]').forEach(el => { el.style.width = el.dataset.w + '%'; });
    root.querySelectorAll('[data-off]').forEach(el => { el.style.strokeDashoffset = el.dataset.off; });
  }));
}
function ring(p, size, stroke, color) {
  const r = (size - stroke) / 2, c = 2 * Math.PI * r;
  return `<svg viewBox="0 0 ${size} ${size}" aria-hidden="true"><circle class="ring-bg" cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke-width="${stroke}"/>` +
    `<circle class="ring-fg" cx="${size / 2}" cy="${size / 2}" r="${r}" fill="none" stroke="${color}" stroke-width="${stroke}" stroke-linecap="round" stroke-dasharray="${c}" stroke-dashoffset="${c}" data-off="${c * (1 - cap(p) / 100)}"/></svg>`;
}

/* ---------- data loading ---------- */
function rows(w, n) { const s = w.Sheets[n]; return s ? XLSX.utils.sheet_to_json(s, { range: 2, defval: '' }) : []; }

function loadScript(src) {
  return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = rej; document.head.appendChild(s); });
}
async function ensureXLSX() {
  if (window.XLSX) return;
  await new Promise(r => setTimeout(r, 400));          // the deferred local copy may still be executing
  if (!window.XLSX) await loadScript(XLSX_FALLBACK);   // last resort: CDN copy
  if (!window.XLSX) throw new Error('XLSX library not loaded');
}
async function fetchWorkbook() {
  for (const f of FILES) {
    try {
      const ctl = new AbortController(), to = setTimeout(() => ctl.abort(), 15000);
      const r = await fetch(f + '?v=' + Date.now(), { cache: 'no-store', signal: ctl.signal });
      clearTimeout(to);
      if (r.ok) return XLSX.read(await r.arrayBuffer(), { type: 'array' });
    } catch (e) { /* try the next file name */ }
  }
  throw new Error('Data file not found');
}
function buildData(w) {
  const d = { __missing: SHEETS.filter(n => !w.Sheets[n]) };
  SHEETS.forEach(n => { d[n] = rows(w, n); });
  return d;
}
function prepare() {
  X = {
    sec: new Map(D.Sectors.map(s => [s.ID, s])), gov: new Map(D.Governorates.map(g => [g.ID, g])),
    ind: new Map(D.Indicators.map(i => [i.ID, i])), act: new Map(D.Activities.map(a => [a.ID, a])), warns: []
  };
  (D.__missing || []).forEach(n => X.warns.push({ k: 'missing', sheet: n }));
  const check = (sheet, list, field, map) => list.forEach(r => { if (r[field] !== '' && !map.has(r[field])) X.warns.push({ k: 'unknown', sheet, id: r[field] }); });
  check('Indicator_Data', D.Indicator_Data, 'GovernorateID', X.gov); check('Indicator_Data', D.Indicator_Data, 'IndicatorID', X.ind);
  check('Activity_Data', D.Activity_Data, 'GovernorateID', X.gov); check('Activity_Data', D.Activity_Data, 'ActivityID', X.act);
  check('Allocations', D.Allocations, 'GovernorateID', X.gov); check('Allocations', D.Allocations, 'SectorID', X.sec);
  check('Indicators', D.Indicators, 'SectorID', X.sec); check('Activities', D.Activities, 'SectorID', X.sec);
}

async function load() {
  $('refresh').classList.add('spin');
  setProgress(15);
  const timer = setInterval(() => setProgress(Math.min(90, +$('plBar').dataset.p + 8)), 120);
  try {
    await ensureXLSX();
    D = buildData(await fetchWorkbook());
    fromCache = false;
    store.set(CACHE_KEY, JSON.stringify(D));
  } catch (e) {
    const cached = store.get(CACHE_KEY);
    if (cached) { try { D = JSON.parse(cached); fromCache = true; } catch (err) { D = null; } }
    if (!D) fromCache = null;
  }
  clearInterval(timer);
  if (D) {
    prepare();
    if (!X.sec.has(sector)) sector = 'ALL';
    if (!X.gov.has(gov)) gov = 'ALL';
  }
  render();
  setProgress(100);
  setTimeout(() => $('preloader').classList.add('done'), 200);
  $('refresh').classList.remove('spin');
}
function setProgress(p) { $('plBar').dataset.p = p; $('plBar').style.width = p + '%'; text('plPct', Math.round(p) + '%'); }

/* ---------- rendering ---------- */
function render() {
  const t = T[lang];
  document.documentElement.lang = lang;
  document.documentElement.dir = lang === 'ar' ? 'rtl' : 'ltr';
  $('lang').textContent = lang === 'ar' ? 'EN' : 'ع';
  $('skip').textContent = t.skip;
  [['refresh', t.refresh], ['print', t.print], ['theme', t.theme], ['toTop', t.top]].forEach(([id, l]) => { $(id).setAttribute('aria-label', l); $(id).title = l; });
  document.querySelectorAll('[data-go]').forEach((a, i) => a.textContent = t.nav[i]);
  ['over', 'desc', 'duration', 'impactT', 'impactS', 'filterT', 'filterD', 'areasT', 'areasD', 'heatT', 'indT', 'indD', 'actT', 'actD', 'upT', 'upD', 'footer',
    'overallT', 'overallS', 'bySectorT', 'budgetT', 'budgetS', 'thAct', 'thSector', 'thTarget', 'thActual', 'thBudget', 'thStatus', 'exportCsv'].forEach(k => text(k, t[k]));
  text('startL', t.start); text('endL', t.end); text('elapsedL', t.elapsed);
  text('sectorLabel', t.sector); text('govLabel', t.gov); text('reset', t.reset); text('fbLabel', t.filtering); text('fbReset', t.clear);
  html('legendA', `<span><i></i>${t.legendT}</span><span><i class="fill"></i>${t.legendA}</span>`);
  $('actSearch').placeholder = t.search;
  $('actSearch').setAttribute('aria-label', t.search);
  html('actSort', [['budget', t.sortBudget], ['target', t.sortTarget], ['name', t.sortName]].map(([k, l]) => `<button type="button" data-sort="${k}" aria-pressed="${actSort === k}" class="${actSort === k ? 'active' : ''}">${l}</button>`).join(''));
  html('heatLegend', [['b0', '0%'], ['b1', '1–24%'], ['b2', '25–49%'], ['b3', '50–74%'], ['b4', '75–99%'], ['b5', t.heatFull]].map(([c, l]) => `<span><i class="sw ${c}"></i><bdi>${l}</bdi></span>`).join(''));

  const live = $('live');
  live.classList.toggle('cached', fromCache !== false);
  text('liveT', fromCache === false ? t.live : t.cached);
  const notice = $('notice');
  notice.hidden = fromCache === false;
  notice.textContent = fromCache === true ? t.noticeCache : t.noticeFail;
  if (!D) { $('warnings').hidden = true; return; }

  const pt = cfg('ProjectName');
  text('title', pt);
  document.title = 'RECIE | ' + pt;
  text('donor', t.donor + ': ' + cfg('Donor'));
  text('partner', t.partner + ': ' + cfg('Partner'));
  text('start', cfg('StartDate')); text('end', cfg('EndDate'));
  text('updated', t.updated + ': ' + cfg('LastUpdated'));
  renderTimeline();
  renderImpact();
  renderWarnings();

  html('sectorFilters', D.Sectors.map(x => `<button type="button" style="--c:${esc(x.Color || '#009fe3')}" data-sector="${esc(x.ID)}"><i></i>${esc(name(x))}</button>`).join(''));
  html('govFilters', D.Governorates.map(x => `<button type="button" data-gov="${esc(x.ID)}">${esc(name(x))}</button>`).join(''));
  renderData();
}

function renderWarnings() {
  const t = T[lang], w = X.warns, el = $('warnings');
  el.hidden = !w.length;
  if (!w.length) return;
  const line = x => x.k === 'missing' ? `${t.warnMissing}: <b>${esc(x.sheet)}</b>` : `<b>${esc(x.sheet)}</b>: ${t.warnUnknown} «${esc(x.id)}»`;
  el.innerHTML = `<b>${t.warnT}</b><ul>${w.slice(0, 4).map(x => `<li>${line(x)}</li>`).join('')}</ul>${w.length > 4 ? `<small>${t.warnMore(w.length - 4)}</small>` : ''}`;
}

function renderTimeline() {
  const t = T[lang], s = parseDate(cfgEn('StartDate')), e = parseDate(cfgEn('EndDate')), now = new Date();
  if (!s || !e) return;
  const p = cap((now - s) / (e - s) * 100);
  const st = now < s ? 'upcoming' : now > e ? 'completed' : 'ongoing';
  text('status', t[st]);
  $('status').dataset.st = st;
  $('elapsed').dataset.count = p; $('elapsed').dataset.fmt = 'pct'; $('elapsed').dataset.k = 'elapsed';
  $('elapsedBar').dataset.w = p;
  text('daysLeft', st === 'ongoing' ? t.daysLeft(num(Math.ceil((e - now) / 864e5))) : st === 'upcoming' ? t.daysTo(num(Math.ceil((s - now) / 864e5))) : '');
}

function renderImpact() {
  const t = T[lang], a = D.Allocations;
  const vals = [[sum(a, 'People_Target'), ''], [sum(a, 'People_Actual'), ''], [sum(D.Activity_Data, 'Budget_USD'), 'usd'],
    [govList().length, ''], [secList().length, ''], [D.Activities.length, '']];
  html('impact', vals.map(([v, f], i) => `<div class="impact-item"><b data-count="${v}" data-fmt="${f}" data-k="imp${i}">0</b><span>${t.impact[i]}</span></div>`).join(''));
}

function renderData() {
  const t = T[lang];
  const mark = (id, attr, val) => $(id).querySelectorAll('button').forEach(b => { const on = b.dataset[attr] === val; b.classList.toggle('active', on); b.setAttribute('aria-pressed', on); });
  mark('sectorFilters', 'sector', sector); mark('govFilters', 'gov', gov);
  html('context', `${t.context}: <b>${esc(name(sectorOf(sector)))}</b> · <b>${esc(name(govOf(gov)))}</b>`);
  const filtered = sector !== 'ALL' || gov !== 'ALL';
  $('reset').hidden = !filtered;
  renderFilterBar(filtered);
  syncUrl();

  const secColor = sector === 'ALL' ? '#009fe3' : colorOf(sector);
  const a = alloc();
  const pT = sum(a, 'People_Target'), pA = sum(a, 'People_Actual'), aT = sum(a, 'Activities_Target'), aA = sum(a, 'Activities_Actual');

  html('kpis', [[pT, t.peopleTarget, null], [pA, t.peopleActual, pct(pA, pT)], [aT, t.activitiesTarget, null], [aA, t.activitiesActual, pct(aA, aT)]].map(([v, l, p], i) =>
    `<article class="kpi" style="--c:${i % 2 ? secColor : '#075e85'}"><span>${l}</span><b data-count="${v}" data-k="kpi${i}">0</b>` +
    (p === null ? `<small>${t.remaining}: ${num(Math.max(0, i === 0 ? pT - pA : aT - aA))}</small>` : `<small>${pctf(p)} ${t.ofTarget}</small><div class="bar"><i data-w="${cap(p)}"></i></div>`) + `</article>`).join(''));

  const op = pct(pA, pT);
  html('overallRing', ring(op, 160, 14, secColor) + `<div class="ring-txt"><b data-count="${op}" data-fmt="pct" data-k="ring">0</b><small>${num(pA)} / ${num(pT)}</small></div>`);

  renderProfile();
  renderSectorChart();
  renderBudget();
  renderHeat();
  renderGovCards();
  renderIndicators();
  renderActivities();
  renderNews();
  animateNumbers();
  animateBars();
}

function renderFilterBar(filtered) {
  const t = T[lang], bar = $('filterBar');
  bar.hidden = !filtered;
  if (!filtered) { $('fbChips').innerHTML = ''; return; }
  const chip = (kind, label, c) => `<button type="button" class="fb-chip" data-clear="${kind}" style="--c:${esc(c)}" aria-label="${esc(t.clear)}: ${esc(label)}"><span>${esc(label)}</span><svg aria-hidden="true"><use href="#i-x"/></svg></button>`;
  html('fbChips', (sector !== 'ALL' ? chip('sector', name(sectorOf(sector)), colorOf(sector)) : '') + (gov !== 'ALL' ? chip('gov', name(govOf(gov)), '#075e85') : ''));
}

function renderProfile() {
  const t = T[lang], el = $('govProfile');
  if (gov === 'ALL') { el.hidden = true; el.innerHTML = ''; return; }
  const g = govOf(gov), a = alloc(gov, sector);
  const p = pct(sum(a, 'People_Actual'), sum(a, 'People_Target')), lv = level(p);
  const ar = actRows(gov, sector), started = ar.filter(r => String(r.Status || 'Planned') !== 'Planned' || +r.Actual > 0).length;
  const inds = D.Indicators.filter(i => sector === 'ALL' || i.SectorID === sector).map(i => indRows(gov, 'ALL').filter(r => r.IndicatorID === i.ID)).filter(r => sum(r, 'Target') > 0);
  const good = inds.filter(r => pct(sum(r, 'Actual'), sum(r, 'Target')) >= 75).length;
  const note = (lang === 'ar' ? g.Note_AR : g.Note_EN) || '';
  el.hidden = false;
  el.innerHTML = `<div class="pf-head"><h3>${esc(name(g))}</h3><span class="tag ${lv}">${t.tags[lv]}</span></div>` +
    (note ? `<p class="pf-note"><b>${t.profNote}:</b> ${esc(note)}</p>` : '') +
    `<div class="pf-stats">` +
    `<div><small>${t.profBudget}</small><b>${usd(sum(ar, 'Budget_USD'))}</b></div>` +
    `<div><small>${t.profInd}</small><b>${num(good)} ${t.of} ${num(inds.length)}</b></div>` +
    `<div><small>${t.profAct}</small><b>${num(started)} ${t.of} ${num(ar.length)}</b></div></div>`;
}

function renderSectorChart() {
  const data = secList().map(s => {
    const a = alloc(gov, s.ID);
    return { s, tg: sum(a, 'People_Target'), ac: sum(a, 'People_Actual') };
  });
  const max = Math.max(1, ...data.map(d => Math.max(d.tg, d.ac)));
  html('bySector', data.map(({ s, tg, ac }) =>
    `<button type="button" class="hrow ${sector === s.ID ? 'active' : ''}" data-sector="${esc(s.ID)}" style="--c:${esc(s.Color)}" data-tip="${esc(name(s))}: ${num(ac)} / ${num(tg)} (${pctf(pct(ac, tg))})">` +
    `<span class="lbl"><i></i><span>${esc(name(s))}</span></span>` +
    `<span class="htrack"><span class="tg" data-w="${tg / max * 100}"></span><span class="ac" data-w="${ac / max * 100}"></span></span>` +
    `<span class="val"><b>${num(ac)}</b> / ${num(tg)}</span></button>`).join(''));
}

function renderBudget() {
  const t = T[lang];
  const data = secList().map(s => ({ s, v: sum(actRows(gov, s.ID), 'Budget_USD') })).filter(d => d.v > 0);
  const total = data.reduce((x, d) => x + d.v, 0), r = 60, C = 2 * Math.PI * r;
  if (!total) { html('budget', `<p class="empty-note">${t.none}</p>`); return; }
  let off = 0;
  const segs = data.map(({ s, v }) => {
    const len = v / total * C, seg = `<circle class="seg ${sector === s.ID ? 'on' : ''}" data-sector="${esc(s.ID)}" cx="80" cy="80" r="${r}" fill="none" stroke="${esc(s.Color)}" stroke-width="22" stroke-dasharray="${Math.max(0, len - 2)} ${C}" stroke-dashoffset="${-off}" data-tip="${esc(name(s))}: ${usd(v)} (${pctf(v / total * 100)})"/>`;
    off += len; return seg;
  }).join('');
  const focus = data.find(d => d.s.ID === sector);
  html('budget', `<div class="donut ${focus ? 'has-focus' : ''}"><svg viewBox="0 0 160 160" aria-hidden="true"><circle cx="80" cy="80" r="${r}" fill="none" class="ring-bg" stroke-width="22"/>${segs}</svg>` +
    `<div class="center"><small>${focus ? esc(name(focus.s)) : t.total}</small><b data-count="${focus ? focus.v : total}" data-fmt="usd" data-k="donut">0</b>${focus ? `<small>${pctf(focus.v / total * 100)}</small>` : ''}</div></div>` +
    `<div class="dlegend">${data.map(({ s, v }) => `<button type="button" class="${sector === s.ID ? 'active' : ''}" data-sector="${esc(s.ID)}" style="--c:${esc(s.Color)}"><i></i><span>${esc(name(s))}</span><span class="dv">${usd(v)}</span><b>${pctf(v / total * 100)}</b></button>`).join('')}</div>`);
}

/* Governorate x sector matrix: % of people target reached, in five fixed bands */
const heatBin = (ac, tg) => { if (!tg || !ac) return 0; const p = ac / tg * 100; return p < 25 ? 1 : p < 50 ? 2 : p < 75 ? 3 : p < 100 ? 4 : 5; };
function renderHeat() {
  const t = T[lang], secs = secList(), govs = govList();
  const cell = (g, s, label) => {
    const a = alloc(g, s), tg = sum(a, 'People_Target'), ac = sum(a, 'People_Actual'), p = pct(ac, tg), bin = heatBin(ac, tg);
    const on = (gov !== 'ALL' || sector !== 'ALL') && gov === g && sector === s;
    const txt = tg ? pctf(p) : '—';
    return `<td><button type="button" class="cell b${bin}${p > 100 ? ' over' : ''}${on ? ' on' : ''}" data-cell data-gov="${esc(g)}" data-sector="${esc(s)}" aria-pressed="${on}" ` +
      `data-tip="${esc(label)}: ${num(ac)} / ${num(tg)}${p > 100 ? ' · ' + t.heatOver : ''}${tg ? '' : ' · ' + t.heatNone}">${txt}</button></td>`;
  };
  const head = `<thead><tr><td></td>${secs.map(s => `<th scope="col"><button type="button" class="hd${sector === s.ID ? ' on' : ''}" data-sector="${esc(s.ID)}" style="--c:${esc(s.Color)}"><i></i><span>${esc(name(s))}</span></button></th>`).join('')}` +
    `<th scope="col"><button type="button" class="hd all" data-sector="ALL"><span>${t.total}</span></button></th></tr></thead>`;
  const body = govs.map(g => `<tr><th scope="row"><button type="button" class="rh${gov === g.ID ? ' on' : ''}" data-gov="${esc(g.ID)}">${esc(name(g))}</button></th>` +
    secs.map(s => cell(g.ID, s.ID, `${name(g)} · ${name(s)}`)).join('') + cell(g.ID, 'ALL', `${name(g)} · ${t.total}`) + `</tr>`).join('');
  const foot = `<tfoot><tr><th scope="row"><span class="rh">${t.total}</span></th>${secs.map(s => cell('ALL', s.ID, `${t.total} · ${name(s)}`)).join('')}${cell('ALL', 'ALL', t.total)}</tr></tfoot>`;
  html('heat', head + `<tbody>${body}</tbody>` + foot);
}

function renderGovCards() {
  const t = T[lang];
  html('govCards', govList().map(g => {
    const a = alloc(g.ID, sector), tg = sum(a, 'People_Target'), ac = sum(a, 'People_Actual');
    const aT = sum(a, 'Activities_Target'), aA = sum(a, 'Activities_Actual'), bud = sum(actRows(g.ID, sector), 'Budget_USD');
    const p = pct(ac, tg), lv = level(p), c = sector === 'ALL' ? '#009fe3' : colorOf(sector);
    return `<button type="button" class="card gov ${gov === g.ID ? 'active' : ''} ${gov !== 'ALL' && gov !== g.ID ? 'dim' : ''}" data-gov="${esc(g.ID)}" aria-pressed="${gov === g.ID}">` +
      `<div class="gov-top"><h3>${esc(name(g))}</h3><span class="tag ${lv}">${t.tags[lv]}</span></div>` +
      `<div><small class="sub"><span>${t.reached}</span></small><div class="big">${num(ac)}</div></div>` +
      `<div class="bar" style="--c:${c}"><i data-w="${cap(p)}"></i></div>` +
      `<small class="sub"><span>${pctf(p)} ${t.of} ${num(tg)}</span><span>${t.budgetShort}: ${usd(bud)}</span></small>` +
      `<small class="sub"><span>${num(aA)} / ${num(aT)} ${t.activitiesShort}</span></small></button>`;
  }).join(''));
}

function renderIndicators() {
  const t = T[lang], list = D.Indicators.filter(x => sector === 'ALL' || x.SectorID === sector);
  html('indicatorCards', list.length ? list.map(x => {
    const rs = D.Indicator_Data.filter(r => r.IndicatorID === x.ID), scoped = rs.filter(r => gov === 'ALL' || r.GovernorateID === gov);
    const tg = sum(scoped, 'Target'), ac = sum(scoped, 'Actual'), p = pct(ac, tg), c = esc(colorOf(x.SectorID)), lv = level(p), unit = esc(x[lang === 'ar' ? 'Unit_AR' : 'Unit_EN']);
    const breakdown = gov === 'ALL' && rs.length ? `<details class="by-gov"><summary>${t.byGov}</summary><div class="bg-list">` +
      govList().map(g => {
        const r = rs.filter(q => q.GovernorateID === g.ID); if (!r.length) return '';
        const gt = sum(r, 'Target'), ga = sum(r, 'Actual'), gp = pct(ga, gt);
        return `<button type="button" class="bg-row" data-gov="${esc(g.ID)}" data-tip="${esc(name(g))}: ${pctf(gp)}"><span class="n">${esc(name(g))}</span><span class="bg-bar"><i data-w="${cap(gp)}"></i></span><span class="v"><b>${num(ga)}</b> / ${num(gt)}</span></button>`;
      }).join('') + `</div></details>` : '';
    return `<article class="card ind" style="--c:${c}">` +
      `<div class="ind-top"><div class="ind-t"><span class="badge">${esc(name(sectorOf(x.SectorID)))}</span><h3>${esc(name(x))}</h3></div>` +
      `<div class="mini-ring">${ring(p, 58, 6, c)}<b>${pctf(p)}</b></div></div>` +
      `<div class="bar"><i data-w="${cap(p)}"></i></div>` +
      `<div class="metric"><span>${t.actual}: <b>${num(ac)}</b> ${unit}</span><span>${t.target}: <b>${num(tg)}</b></span></div>` +
      `<span class="tag ${lv}" style="width:fit-content">${t.tags[lv]}</span>${breakdown}</article>`;
  }).join('') : `<article class="card empty">${t.none}</article>`);
}

function renderActivities() {
  const t = T[lang], q = actQuery.trim().toLowerCase();
  const scopeBudget = sum(actRows(gov, 'ALL'), 'Budget_USD') || 1;
  const acts = D.Activities.filter(a => sector === 'ALL' || a.SectorID === sector).map(a => {
    const rs = D.Activity_Data.filter(r => r.ActivityID === a.ID && (gov === 'ALL' || r.GovernorateID === gov));
    return { a, rs, budget: sum(rs, 'Budget_USD'), target: sum(rs, 'Target'), actual: sum(rs, 'Actual'), status: aggStatus(rs) };
  }).filter(x => x.rs.length && (!q || (x.a.Arabic + ' ' + x.a.English).toLowerCase().includes(q)))
    .sort((x, y) => actSort === 'name' ? name(x.a).localeCompare(name(y.a), lang) : actSort === 'target' ? y.target - x.target : y.budget - x.budget);
  lastActRows = acts;

  html('activityRows', acts.length ? acts.map(({ a, rs, budget, target, actual, status }) => {
    const sh = budget / scopeBudget * 100, c = esc(colorOf(a.SectorID)), ap = pct(actual, target);
    const toggle = gov === 'ALL' ? `<button type="button" class="chev" data-toggle aria-expanded="false" aria-label="${esc(t.expand)}: ${esc(name(a))}"><svg aria-hidden="true"><use href="#i-chev"/></svg></button>` : '';
    const main = `<tr class="act-row" style="--c:${c}"><td class="a-name">${toggle}<span>${esc(name(a))}</span></td>` +
      `<td data-l="${t.thSector}"><span class="badge">${esc(name(sectorOf(a.SectorID)))}</span></td>` +
      `<td class="num" data-l="${t.thTarget}">${num(target)}</td>` +
      `<td class="num" data-l="${t.thActual}"><b>${num(actual)}</b> <small>${pctf(ap)}</small></td>` +
      `<td class="share-td" data-l="${t.thBudget}"><div class="share"><span class="num"><b>${usd(budget)}</b> <small>· ${pctf(sh)} ${t.share}</small></span><div class="bar"><i data-w="${cap(sh)}"></i></div></div></td>` +
      `<td data-l="${t.thStatus}"><span class="st ${esc(status.toLowerCase())}">${esc(t.status[status] || status)}</span></td></tr>`;
    if (gov !== 'ALL') return main;
    const sub = `<tr class="sub-row" hidden style="--c:${c}"><td colspan="6"><div class="sub-grid">` + govList().map(g => {
      const r = rs.filter(q2 => q2.GovernorateID === g.ID); if (!r.length) return '';
      const st = aggStatus(r);
      return `<div class="sub-r"><button type="button" class="link" data-gov="${esc(g.ID)}">${esc(name(g))}</button><span>${usd(sum(r, 'Budget_USD'))}</span><span>${num(sum(r, 'Actual'))} / ${num(sum(r, 'Target'))}</span><span class="st ${esc(st.toLowerCase())}">${esc(t.status[st] || st)}</span></div>`;
    }).join('') + `</div></td></tr>`;
    return main + sub;
  }).join('') : `<tr><td colspan="6" class="empty">${t.none}</td></tr>`);
  html('activityTotal', `<td>${t.total} (${num(acts.length)})</td><td></td><td class="num">${num(acts.reduce((s, x) => s + x.target, 0))}</td><td class="num">${num(acts.reduce((s, x) => s + x.actual, 0))}</td><td class="num">${usd(acts.reduce((s, x) => s + x.budget, 0))}</td><td></td>`);
}

function renderNews() {
  const t = T[lang], ar = lang === 'ar', u = D.Updates.filter(x => String(x.Published).trim().toLowerCase() === 'yes');
  html('news', u.length ? u.map(x => {
    const url = /^https?:\/\//i.test(x.URL) ? esc(x.URL) : '#';
    return `<a class="card news" href="${url}" target="_blank" rel="noopener noreferrer"><time>${esc(x[ar ? 'Date_AR' : 'Date_EN'])}</time>` +
      `<h3>${esc(x[ar ? 'Title_AR' : 'Title_EN'])}</h3><p>${esc(x[ar ? 'Summary_AR' : 'Summary_EN'])}</p><span class="more">${t.more}</span></a>`;
  }).join('') : `<article class="card empty">${t.noNews}</article>`);
}

function syncUrl() {
  const p = new URLSearchParams();
  if (sector !== 'ALL') p.set('sector', sector);
  if (gov !== 'ALL') p.set('gov', gov);
  if (lang !== 'ar') p.set('lang', lang);
  history.replaceState(null, '', location.pathname + (p.toString() ? '?' + p : '') + location.hash);
}

function exportCsv() {
  const t = T[lang], q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const head = [t.thAct, t.thSector, t.gov, t.thTarget, t.thActual, t.thBudget + ' (USD)', t.thStatus];
  const body = lastActRows.map(({ a, budget, target, actual, status }) => [name(a), name(sectorOf(a.SectorID)), name(govOf(gov)), target, actual, budget, t.status[status] || status]);
  const csv = '\ufeff' + [head, ...body].map(r => r.map(q).join(',')).join('\r\n');
  const link = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })), download: `recie-activities-${gov}-${sector}.csv` });
  document.body.appendChild(link); link.click(); link.remove(); setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

/* ---------- interactions ---------- */
document.addEventListener('click', e => {
  if (!D) return;
  const el = e.target.closest('[data-toggle], [data-clear], [data-cell], [data-sort], [data-sector], [data-gov]');
  if (!el) return;
  const d = el.dataset;
  if ('toggle' in d) {
    const row = el.closest('tr').nextElementSibling, open = row.hidden;
    row.hidden = !open; el.setAttribute('aria-expanded', open); el.classList.toggle('open', open);
    animateBars(row);
    return;
  }
  if ('clear' in d) { if (d.clear !== 'gov') sector = 'ALL'; if (d.clear !== 'sector') gov = 'ALL'; renderData(); return; }
  if ('sort' in d) {
    actSort = d.sort;
    $('actSort').querySelectorAll('button').forEach(b => { b.classList.toggle('active', b === el); b.setAttribute('aria-pressed', b === el); });
    renderActivities(); animateBars($('activityRows')); return;
  }
  if ('cell' in d) {
    const same = gov === d.gov && sector === d.sector;
    gov = same ? 'ALL' : d.gov; sector = same ? 'ALL' : d.sector;
  } else if ('sector' in d && !('gov' in d)) {
    sector = el.closest('#sectorFilters') ? d.sector : (sector === d.sector ? 'ALL' : d.sector);
  } else if ('gov' in d) {
    gov = el.closest('#govFilters') ? d.gov : (gov === d.gov ? 'ALL' : d.gov);
  }
  renderData();
  if (el.closest('#areas') && innerWidth < 850) $('dashboard').scrollIntoView({ behavior: reduced ? 'auto' : 'smooth' });
});
$('actSearch').addEventListener('input', e => { actQuery = e.target.value; if (D) { renderActivities(); animateBars($('activityRows')); } });
$('exportCsv').onclick = () => { if (D) exportCsv(); };
$('lang').onclick = () => { lang = lang === 'ar' ? 'en' : 'ar'; store.set('lang', lang); prevValues.clear(); render(); };
$('theme').onclick = () => {
  const v = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = v; store.set('theme', v);
};
$('refresh').onclick = () => load();
$('print').onclick = () => print();
$('toTop').onclick = () => scrollTo({ top: 0, behavior: reduced ? 'auto' : 'smooth' });
addEventListener('beforeprint', () => document.querySelectorAll('details.by-gov').forEach(d => { d.open = true; }));

// Tooltip for chart elements (mouse/pen only; values are also visible as text on touch screens)
const tip = $('tooltip');
document.addEventListener('pointermove', e => {
  const el = e.pointerType !== 'touch' && e.target.closest?.('[data-tip]');
  if (!el) { tip.classList.remove('show'); return; }
  tip.textContent = el.dataset.tip;
  const w = tip.offsetWidth, h = tip.offsetHeight;
  tip.style.left = Math.min(innerWidth - w - 8, Math.max(8, e.clientX - w / 2)) + 'px';
  tip.style.top = Math.max(8, e.clientY - h - 14) + 'px';
  tip.classList.add('show');
}, { passive: true });

// Scroll progress, back-to-top and scroll-spy for the nav
const navLinks = [...document.querySelectorAll('[data-go]')];
function onScroll() {
  const h = document.documentElement, max = h.scrollHeight - h.clientHeight;
  $('scrollBar').style.width = (max > 0 ? h.scrollTop / max * 100 : 0) + '%';
  $('toTop').classList.toggle('show', h.scrollTop > 600);
  let current = navLinks[0];
  for (const a of navLinks) { const sec = $(a.dataset.go); if (sec && sec.getBoundingClientRect().top <= 150) current = a; }
  if (innerHeight + h.scrollTop >= h.scrollHeight - 4) current = navLinks[navLinks.length - 1];
  if (!current.classList.contains('active')) {
    navLinks.forEach(a => { a.classList.toggle('active', a === current); if (a === current) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current'); });
    const nav = $('nav'), nr = nav.getBoundingClientRect(), ar = current.getBoundingClientRect();
    if (ar.left < nr.left || ar.right > nr.right) nav.scrollBy({ left: ar.left + ar.width / 2 - (nr.left + nr.width / 2), behavior: 'smooth' });
  }
}
let ticking = false;
addEventListener('scroll', () => { if (!ticking) { ticking = true; requestAnimationFrame(() => { onScroll(); ticking = false; }); } }, { passive: true });

render();
onScroll();
load();
