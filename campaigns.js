// Campaigns: what is being planned, set up, waiting for approval, running and ended, per client and platform.
// A board with a column per stage, a timeline, a table of what is on air now, and how the ended campaigns did.
// The same campaigns can be kept in a Google Sheets tab ("קמפיינים"), which a script keeps in step with this screen.
// They are stored next to the workspace's settings, as small documents named "cmp_…".
import { t, locale, _dicts } from './i18n.js';

export const CAMP_STAGES = ['plan', 'setup', 'wait', 'ok', 'live', 'paused', 'done'];
export const CAMP_PLATFORMS = [['Google Ads', '#1D7AFC'], ['Meta', '#8F7EE7'], ['TikTok', '#DA62AC'], ['Taboola', '#F38A3F'], ['ChatGPT', '#22A06B']];
// the objectives are kept by the names the sheet uses
export const CAMP_OBJECTIVES = ['מודעות', 'צפיות', 'תנועה לאתר', 'מעורבות', 'לידים', 'מכירות', 'התקנות'];
export const CAMP_RESULTS = ['open', 'yes', 'no'];   // did it reach its target: not ended yet, yes, no
const CAMP_MODES = ['board', 'time', 'live', 'res'];
const OTHER_COLOR = '#8590A2';
const CAMP_ID = /^cmp_[A-Za-z0-9]{1,40}$/;
const DAYS = 70, DAY_W = 22, ENDED_SHOWN = 8, ATTN_SHOWN = 4;

Object.assign(_dicts.he, {
  'side.camps':'קמפיינים', 'view.campsOf':'{name}: קמפיינים', 'cmp.new':'קמפיין חדש', 'cmp.edit':'עריכת קמפיין', 'cmp.one':'קמפיין',
  'cmp.mode.board':'לוח', 'cmp.mode.time':'ציר זמן', 'cmp.mode.live':'באוויר', 'cmp.mode.res':'תוצאות', 'cmp.mode':'תצוגה', 'cmp.all':'כל הפלטפורמות', 'cmp.platFilter':'סינון לפי פלטפורמה',
  'cmp.st.plan':'תכנון', 'cmp.st.setup':'בהקמה', 'cmp.st.wait':'ממתין לאישור', 'cmp.st.ok':'מאושר', 'cmp.st.live':'רץ', 'cmp.st.paused':'מושהה', 'cmp.st.done':'הסתיים',
  'cmp.f.name':'שם הקמפיין', 'cmp.f.client':'לקוח', 'cmp.f.platform':'פלטפורמה', 'cmp.f.stage':'שלב', 'cmp.f.start':'תאריך עלייה', 'cmp.f.end':'תאריך סיום',
  'cmp.f.net':'תקציב נטו (₪)', 'cmp.f.gross':'תקציב ברוטו (₪)', 'cmp.f.link':'קישור (תוכנית מדיה, דוח…)', 'cmp.f.notes':'הערות', 'cmp.other':'אחר…', 'cmp.otherPh':'שם הפלטפורמה',
  'cmp.f.objective':'אובג׳קטיב', 'cmp.objPh':'שם האובג׳קטיב', 'cmp.f.target':'יעד סופי', 'cmp.targetPh':'למשל: 500 לידים, ROAS 4', 'cmp.f.result':'הצלחנו?',
  'cmp.f.after':'אחרי הקמפיין', 'cmp.afterPh':'סיכום, מסקנות ומה עושים הלאה', 'cmp.resultNote':'נשאר "עוד לא נגמר" עד שהקמפיין מסתיים. מהתשובות נבנה גרף ההצלחה.',
  'cmp.o.0':'מודעות', 'cmp.o.1':'צפיות', 'cmp.o.2':'תנועה לאתר', 'cmp.o.3':'מעורבות', 'cmp.o.4':'לידים', 'cmp.o.5':'מכירות', 'cmp.o.6':'התקנות',
  'cmp.r.open':'עוד לא נגמר', 'cmp.r.yes':'כן', 'cmp.r.no':'לא', 'cmp.rs.open':'עוד לא סומן', 'cmp.rs.yes':'הצליח', 'cmp.rs.no':'לא הצליח',
  'cmp.k.live':'תקציב נטו של קמפיינים רצים', 'cmp.k.next':'תקציב נטו של קמפיינים שעוד לא עלו', 'cmp.k.gross':'ברוטו {v}', 'cmp.k.up':'עולים ב־7 הימים הקרובים', 'cmp.k.end':'מסתיימים ב־7 הימים הקרובים',
  'cmp.k.win':'אחוז הצלחה', 'cmp.k.winOf':'{a} מתוך {b} שסומנו', 'cmp.k.winNone':'עדיין אין תוצאות', 'cmp.k.attn':'דורשים תשומת לב', 'cmp.k.urgent':'{n} דחוף',
  'cmp.lv.0':'דחוף', 'cmp.lv.1':'בקרוב', 'cmp.lv.2':'לידיעה', 'cmp.a.more':'הצגת כל {n}', 'cmp.a.less':'הצגת פחות',
  'cmp.h.today':'עולה היום', 'cmp.h.tomorrow':'עולה מחר', 'cmp.h.in':'עולה בעוד {n} ימים', 'cmp.h.late':'תאריך העלייה עבר', 'cmp.h.left':'נותרו {n} ימים', 'cmp.h.left1':'נותר יום אחד',
  'cmp.h.endsToday':'מסתיים היום', 'cmp.h.noEnd':'ללא תאריך סיום', 'cmp.h.ended':'הסתיים ב־{d}', 'cmp.h.noDates':'בלי תאריכים',
  'cmp.a.title':'דורש תשומת לב', 'cmp.a.late':'תאריך העלייה ({d}) עבר, והקמפיין עדיין ב"{st}"', 'cmp.a.soon':'עולה בעוד {n} ימים ועדיין ב"{st}"',
  'cmp.a.soon1':'עולה בעוד יום אחד ועדיין ב"{st}"', 'cmp.a.soon0':'אמור לעלות היום ועדיין ב"{st}"', 'cmp.a.noStart':'מאושר, אבל בלי תאריך עלייה', 'cmp.a.paused':'הקמפיין מושהה',
  'cmp.a.pausedEnd':'מושהה, ותאריך הסיום ({d}) כבר עבר', 'cmp.a.grossLow':'התקציב ברוטו קטן מהנטו', 'cmp.a.ends':'מסתיים בעוד {n} ימים', 'cmp.a.ends1':'מסתיים בעוד יום אחד',
  'cmp.a.ends0':'מסתיים היום', 'cmp.a.noResult':'הסתיים ב־{d}, ועוד לא סומן אם הצלחנו', 'cmp.a.noResult0':'הסתיים, ועוד לא סומן אם הצלחנו', 'cmp.a.lack':'חסר: {list}',
  'cmp.a.noAfter':'הסתיים, וחסר סיכום ב"אחרי הקמפיין"', 'cmp.l.budget':'תקציב', 'cmp.l.platform':'פלטפורמה', 'cmp.l.objective':'אובג׳קטיב', 'cmp.l.target':'יעד סופי',
  'cmp.c.campaign':'קמפיין', 'cmp.c.client':'לקוח', 'cmp.c.platform':'פלטפורמה', 'cmp.c.objective':'אובג׳קטיב', 'cmp.c.dates':'תאריכים', 'cmp.c.progress':'התקדמות',
  'cmp.c.net':'נטו', 'cmp.c.gross':'ברוטו', 'cmp.c.target':'יעד סופי', 'cmp.c.result':'הצלחנו?', 'cmp.c.after':'אחרי הקמפיין', 'cmp.c.ended':'הסתיים',
  'cmp.live.none':'אין קמפיינים באוויר כרגע', 'cmp.live.total':'סה״כ: {n} קמפיינים באוויר', 'cmp.target':'יעד: {v}', 'cmp.netGross':'נטו {n} · ברוטו {g}',
  'cmp.res.title':'כמה קמפיינים הצליחו', 'cmp.res.sub':'מתוך הקמפיינים שהסתיימו, לפי אובג׳קטיב', 'cmp.res.all':'כל הקמפיינים', 'cmp.res.noObj':'ללא אובג׳קטיב',
  'cmp.res.none':'עדיין אין קמפיינים שהסתיימו', 'cmp.res.line':'{y} הצליחו · {n} לא', 'cmp.res.open':'{o} לא סומנו', 'cmp.res.rate':'{p}% הצלחה', 'cmp.res.ended':'קמפיינים שהסתיימו',
  'cmp.empty':'עדיין אין קמפיינים', 'cmp.emptySub':'הוסיפו קמפיין כאן, או בלשונית "קמפיינים" בגיליון.', 'cmp.emptyCol':'אין קמפיינים בשלב הזה', 'cmp.noDated':'אין קמפיינים עם תאריכים בתקופה הזו',
  'cmp.moveTo':'העברה לשלב', 'cmp.del':'מחיקת הקמפיין', 'cmp.delSure':'ללחוץ שוב למחיקה', 'cmp.deleted':'הקמפיין נמחק', 'cmp.nameReq':'צריך שם לקמפיין',
  'cmp.datesBad':'תאריך הסיום מוקדם מתאריך העלייה', 'cmp.noClient':'ללא לקוח', 'cmp.showAll':'הצגת כל {n} שהסתיימו', 'cmp.sum':'{live} רצים · {wait} ממתינים לאישור',
  'cmp.openLink':'פתיחת הקישור', 'cmp.autoNote':'"מאושר" הופך ל"רץ" בתאריך העלייה, ו"רץ" הופך ל"הסתיים" אחרי תאריך הסיום.', 'cmp.today':'היום', 'cmp.noBoards':'קודם צריך לוח של לקוח'
});
Object.assign(_dicts.en, {
  'side.camps':'Campaigns', 'view.campsOf':'{name}: campaigns', 'cmp.new':'New campaign', 'cmp.edit':'Edit campaign', 'cmp.one':'Campaign',
  'cmp.mode.board':'Board', 'cmp.mode.time':'Timeline', 'cmp.mode.live':'On air', 'cmp.mode.res':'Results', 'cmp.mode':'View', 'cmp.all':'All platforms', 'cmp.platFilter':'Filter by platform',
  'cmp.st.plan':'Planning', 'cmp.st.setup':'Setting up', 'cmp.st.wait':'Waiting for approval', 'cmp.st.ok':'Approved', 'cmp.st.live':'Running', 'cmp.st.paused':'Paused', 'cmp.st.done':'Ended',
  'cmp.f.name':'Campaign name', 'cmp.f.client':'Client', 'cmp.f.platform':'Platform', 'cmp.f.stage':'Stage', 'cmp.f.start':'Launch date', 'cmp.f.end':'End date',
  'cmp.f.net':'Net budget (₪)', 'cmp.f.gross':'Gross budget (₪)', 'cmp.f.link':'Link (media plan, report…)', 'cmp.f.notes':'Notes', 'cmp.other':'Other…', 'cmp.otherPh':'Platform name',
  'cmp.f.objective':'Objective', 'cmp.objPh':'Objective name', 'cmp.f.target':'Final target', 'cmp.targetPh':'For example: 500 leads, ROAS 4', 'cmp.f.result':'Did it succeed?',
  'cmp.f.after':'After the campaign', 'cmp.afterPh':'Summary, conclusions and what comes next', 'cmp.resultNote':'Stays “Not ended yet” until the campaign ends. The success chart is built from the answers.',
  'cmp.o.0':'Awareness', 'cmp.o.1':'Views', 'cmp.o.2':'Traffic', 'cmp.o.3':'Engagement', 'cmp.o.4':'Leads', 'cmp.o.5':'Sales', 'cmp.o.6':'Installs',
  'cmp.r.open':'Not ended yet', 'cmp.r.yes':'Yes', 'cmp.r.no':'No', 'cmp.rs.open':'Not marked yet', 'cmp.rs.yes':'Succeeded', 'cmp.rs.no':'Did not succeed',
  'cmp.k.live':'Net budget of running campaigns', 'cmp.k.next':'Net budget of campaigns not launched yet', 'cmp.k.gross':'Gross {v}', 'cmp.k.up':'Launching in the next 7 days', 'cmp.k.end':'Ending in the next 7 days',
  'cmp.k.win':'Success rate', 'cmp.k.winOf':'{a} of {b} marked', 'cmp.k.winNone':'No results yet', 'cmp.k.attn':'Need attention', 'cmp.k.urgent':'{n} urgent',
  'cmp.lv.0':'Urgent', 'cmp.lv.1':'Soon', 'cmp.lv.2':'Note', 'cmp.a.more':'Show all {n}', 'cmp.a.less':'Show fewer',
  'cmp.h.today':'Launches today', 'cmp.h.tomorrow':'Launches tomorrow', 'cmp.h.in':'Launches in {n} days', 'cmp.h.late':'Launch date has passed', 'cmp.h.left':'{n} days left', 'cmp.h.left1':'1 day left',
  'cmp.h.endsToday':'Ends today', 'cmp.h.noEnd':'No end date', 'cmp.h.ended':'Ended {d}', 'cmp.h.noDates':'No dates',
  'cmp.a.title':'Needs attention', 'cmp.a.late':'The launch date ({d}) has passed and the campaign is still “{st}”', 'cmp.a.soon':'Launches in {n} days and is still “{st}”',
  'cmp.a.soon1':'Launches tomorrow and is still “{st}”', 'cmp.a.soon0':'Should launch today and is still “{st}”', 'cmp.a.noStart':'Approved, but has no launch date', 'cmp.a.paused':'The campaign is paused',
  'cmp.a.pausedEnd':'Paused, and its end date ({d}) has passed', 'cmp.a.grossLow':'The gross budget is lower than the net', 'cmp.a.ends':'Ends in {n} days', 'cmp.a.ends1':'Ends tomorrow',
  'cmp.a.ends0':'Ends today', 'cmp.a.noResult':'Ended {d}, and not marked yet as succeeded or not', 'cmp.a.noResult0':'Ended, and not marked yet as succeeded or not', 'cmp.a.lack':'Missing: {list}',
  'cmp.a.noAfter':'Ended, and “After the campaign” has no summary', 'cmp.l.budget':'budget', 'cmp.l.platform':'platform', 'cmp.l.objective':'objective', 'cmp.l.target':'final target',
  'cmp.c.campaign':'Campaign', 'cmp.c.client':'Client', 'cmp.c.platform':'Platform', 'cmp.c.objective':'Objective', 'cmp.c.dates':'Dates', 'cmp.c.progress':'Progress',
  'cmp.c.net':'Net', 'cmp.c.gross':'Gross', 'cmp.c.target':'Final target', 'cmp.c.result':'Succeeded?', 'cmp.c.after':'After the campaign', 'cmp.c.ended':'Ended',
  'cmp.live.none':'No campaigns on air right now', 'cmp.live.total':'Total: {n} campaigns on air', 'cmp.target':'Target: {v}', 'cmp.netGross':'Net {n} · gross {g}',
  'cmp.res.title':'How many campaigns succeeded', 'cmp.res.sub':'Of the campaigns that ended, by objective', 'cmp.res.all':'All campaigns', 'cmp.res.noObj':'No objective',
  'cmp.res.none':'No campaigns have ended yet', 'cmp.res.line':'{y} succeeded · {n} did not', 'cmp.res.open':'{o} not marked', 'cmp.res.rate':'{p}% success', 'cmp.res.ended':'Ended campaigns',
  'cmp.empty':'No campaigns yet', 'cmp.emptySub':'Add a campaign here, or in the “קמפיינים” tab of the sheet.', 'cmp.emptyCol':'No campaigns in this stage', 'cmp.noDated':'No campaigns with dates in this period',
  'cmp.moveTo':'Move to stage', 'cmp.del':'Delete campaign', 'cmp.delSure':'Click again to delete', 'cmp.deleted':'Campaign deleted', 'cmp.nameReq':'The campaign needs a name',
  'cmp.datesBad':'The end date is before the launch date', 'cmp.noClient':'No client', 'cmp.showAll':'Show all {n} ended', 'cmp.sum':'{live} running · {wait} waiting for approval',
  'cmp.openLink':'Open the link', 'cmp.autoNote':'“Approved” becomes “Running” on the launch date, and “Running” becomes “Ended” after the end date.', 'cmp.today':'Today', 'cmp.noBoards':'Create a client board first'
});

const isISO = s=>typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
const pad = n=>String(n).padStart(2, '0');
const utc = d=>{ const [y, m, x] = d.split('-').map(Number); return Date.UTC(y, m - 1, x); };
const addDays = (d, n)=>{ const x = new Date(utc(d) + n * 864e5); return `${x.getUTCFullYear()}-${pad(x.getUTCMonth() + 1)}-${pad(x.getUTCDate())}`; };
const dayDiff = (a, b)=>Math.round((utc(b) - utc(a)) / 864e5);
const clean = s=>String(s ?? '').replace(/\s+/g, ' ').trim();
export const platColor = name=>(CAMP_PLATFORMS.find(p=>p[0] === name) || [0, OTHER_COLOR])[1];
const num = v=>typeof v === 'number' && isFinite(v) && v >= 0 ? v : null;
export const netOf = c=>num(c.net) ?? num(c.budget);   // "budget" is what the net budget was called at first
export const grossOf = c=>num(c.gross);
export const resultOf = c=>CAMP_RESULTS.includes(c.result) ? c.result : 'open';
const objLabel = name=>{ const i = CAMP_OBJECTIVES.indexOf(name); return i >= 0 ? t('cmp.o.' + i) : name; };
const dm = d=>Number(d.slice(8)) + '/' + Number(d.slice(5, 7));
function normLink(u){
  let s = String(u || '').trim(); if (!s) return '';
  if (!/^[a-z][a-z0-9+.-]*:/i.test(s)) s = 'https://' + s;
  const m = s.match(/^(https?):\/\/([^\s/?#]+\.[^\s/?#]+)(\S*)$/i);
  return m ? m[1].toLowerCase() + '://' + m[2].toLowerCase() + (m[3] ? (m[3][0] === '/' ? m[3] : '/' + m[3]) : '/') : '';
}
// the stage follows the dates: approved → running on the launch date, running → ended after the end date
export function stageNow(c, today){
  let st = CAMP_STAGES.includes(c.stage) ? c.stage : 'plan';
  if (st === 'ok' && isISO(c.start) && c.start <= today) st = 'live';
  if (st === 'live' && isISO(c.end) && c.end < today) st = 'done';
  return st;
}
// Every campaign that needs a look, most urgent first: [campaign, what to check, level]. Level 0 = wrong now, 1 = about
// to be, 2 = worth knowing. The same rules draw the "דורש תשומת לב" table of the sheet's dashboard.
export function attention(list, today){
  const out = [];
  list.forEach(c=>{
    const st = c.stage, s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '', pts = [], stName = t('cmp.st.' + st);
    const before = ['plan', 'setup', 'wait'].includes(st), toS = s ? dayDiff(today, s) : null, toE = e ? dayDiff(today, e) : null, net = netOf(c), gross = grossOf(c);
    if (before && s && toS < 0) pts.push([0, t('cmp.a.late', {d:dm(s), st:stName})]);
    if (st === 'paused' && e && toE < 0) pts.push([0, t('cmp.a.pausedEnd', {d:dm(e)})]);
    if (net !== null && gross !== null && gross < net) pts.push([0, t('cmp.a.grossLow')]);
    if (before && s && toS >= 0 && toS <= 3) pts.push([1, t(toS === 0 ? 'cmp.a.soon0' : toS === 1 ? 'cmp.a.soon1' : 'cmp.a.soon', {n:toS, st:stName})]);
    if (st === 'ok' && !s) pts.push([1, t('cmp.a.noStart')]);
    if (st === 'live' && e && toE >= 0 && toE <= 3) pts.push([1, t(toE === 0 ? 'cmp.a.ends0' : toE === 1 ? 'cmp.a.ends1' : 'cmp.a.ends', {n:toE})]);
    if (st === 'done' && resultOf(c) === 'open') pts.push([1, e ? t('cmp.a.noResult', {d:dm(e)}) : t('cmp.a.noResult0')]);
    const lack = [];
    if ((st === 'ok' || st === 'live') && net === null && gross === null) lack.push('budget');
    if (['wait', 'ok', 'live'].includes(st)) ['platform', 'objective', 'target'].forEach(k=>{ if (!clean(c[k])) lack.push(k); });
    if (lack.length) pts.push([lack[0] === 'budget' ? 1 : 2, t('cmp.a.lack', {list:lack.map(k=>t('cmp.l.' + k)).join(', ')})]);
    if (st === 'paused' && !(e && toE < 0)) pts.push([2, t('cmp.a.paused')]);
    if (st === 'done' && resultOf(c) !== 'open' && !String(c.after || '').trim() && e && dayDiff(e, today) <= 30) pts.push([2, t('cmp.a.noAfter')]);
    if (!pts.length) return;
    pts.sort((x, y)=>x[0] - y[0]);
    out.push([c, pts.slice(0, 2).map(p=>p[1]).join(' · ') + (pts.length > 2 ? ` (+${pts.length - 2})` : ''), pts[0][0]]);
  });
  return out.sort((a, b)=>a[2] - b[2] || String(a[0].start || '9').localeCompare(String(b[0].start || '9')) || clean(a[0].name).localeCompare(clean(b[0].name)));
}

export function createCampaigns(api){
  const { el, esc, I, P, savePrefs } = api;
  let all = [], showEnded = false, showAttn = false, dragId = null;
  const fixed = new Set();
  const money = n=>new Intl.NumberFormat(locale(), {style:'currency', currency:'ILS', maximumFractionDigits:0}).format(n);
  const sum = arr=>arr.reduce((s, c)=>s + (netOf(c) || 0), 0), sumGross = arr=>arr.reduce((s, c)=>s + (grossOf(c) || 0), 0);
  const get = id=>all.find(c=>c.id === id);
  const mode = ()=>CAMP_MODES.includes(P.campMode) ? P.campMode : 'board';

  // what the server sent. The stage shown is the one the dates give; someone who can edit also saves it
  function setData(docs){
    const today = api.today(), ops = [];
    all = docs.filter(d=>CAMP_ID.test(d.id) && d.data && d.data.kind === 'campaign' && !d.data.archived).map(d=>{
      const c = {...d.data, id:d.id}, st = stageNow(c, today);
      if (st !== c.stage && api.canEdit() && !fixed.has(c.id + st)){ fixed.add(c.id + st); ops.push({...c, stage:st, updated:Date.now()}); }
      c.stage = st;
      return c;
    });
    if (ops.length) write(ops);
  }
  function write(list){
    return api.commit(list.map(c=>{ const {id, ...data} = c; return {type:'set', col:'meta', id, data:JSON.parse(JSON.stringify(data))}; }));
  }
  function patch(id, change){
    const c = get(id); if (!c || !api.canEdit()) return;
    Object.assign(c, change, {updated:Date.now(), by:api.userName()});
    c.stage = stageNow(c, api.today());
    write([c]); api.rerender();
  }
  const scope = ()=>{ const b = api.scopeBoard(); return all.filter(c=>(!b || c.boardId === b.id) && (!P.campPlat || clean(c.platform) === P.campPlat)); };

  /* ---------- pieces ---------- */
  const platChip = name=>name ? `<span class="cg-plat"><i class="pf-dot" style="--sw:${platColor(name)}"></i><span data-bidi>${esc(name)}</span></span>` : '';
  const resChip = c=>`<span class="cg-chip rs-${resultOf(c)}">${esc(t('cmp.rs.' + resultOf(c)))}</span>`;
  const budgetTip = c=>grossOf(c) !== null && netOf(c) !== null ? t('cmp.netGross', {n:money(netOf(c)), g:money(grossOf(c))}) : '';
  // how far a running campaign is: its share of days gone, and what is left
  function progress(c, today){
    const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '';
    if (!s || !e || e < s) return '';
    const pct = Math.max(4, Math.min(100, Math.round((dayDiff(s, today) + 1) / (dayDiff(s, e) + 1) * 100)));
    return `<div class="kc-prog" role="progressbar" aria-valuenow="${pct}" aria-valuemin="0" aria-valuemax="100"><i style="width:${pct}%"></i></div>`;
  }
  function hint(c, today){
    const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '', st = c.stage;
    if (st === 'done') return [e ? t('cmp.h.ended', {d:api.fmtDay(e)}) : '', ''];
    if (st === 'live'){
      if (!e) return [t('cmp.h.noEnd'), 'mute'];
      const n = dayDiff(today, e);
      return [n <= 0 ? t('cmp.h.endsToday') : n === 1 ? t('cmp.h.left1') : t('cmp.h.left', {n}), n <= 2 ? 'warn' : ''];
    }
    if (st === 'paused') return ['', ''];
    if (!s) return [t('cmp.h.noDates'), 'mute'];
    const n = dayDiff(today, s);
    if (n < 0) return [t('cmp.h.late'), 'late'];
    return [n === 0 ? t('cmp.h.today') : n === 1 ? t('cmp.h.tomorrow') : t('cmp.h.in', {n}), n <= 3 ? 'warn' : ''];
  }
  function card(c, today, showClient){
    const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '', b = api.boardById(c.boardId), [h, hc] = hint(c, today), bud = netOf(c), obj = clean(c.objective), tgt = clean(c.target);
    const prog = c.stage === 'live' ? progress(c, today) : '';
    const dates = s || e ? `<div class="kc-dates">${I.cal}<span>${esc(s ? api.fmtDay(s) : '?')}${e ? ' – ' + esc(api.fmtDay(e)) : ''}</span></div>` : '';
    return `<article class="cg-card cs-${c.stage}" data-act="cmp-open" data-id="${c.id}" tabindex="0" role="button"${api.canEdit() ? ' draggable="true"' : ''}>
      <div class="kc-top">${platChip(clean(c.platform))}${obj ? `<span class="kc-obj" data-bidi>${esc(objLabel(obj))}</span>` : ''}<span class="sp"></span>
        ${c.link ? `<a class="kc-link" href="${esc(c.link)}" target="_blank" rel="noopener noreferrer" title="${esc(t('cmp.openLink'))}" aria-label="${esc(t('cmp.openLink'))}">${I.link}</a>` : ''}${showClient && b ? api.bAv(b) : ''}
        ${api.canEdit() ? `<button class="btn icon subtle sm kc-move" data-act="cmp-move" data-id="${c.id}" aria-haspopup="menu" aria-label="${esc(t('cmp.moveTo'))}" title="${esc(t('cmp.moveTo'))}">${I.more}</button>` : ''}</div>
      <h3 class="kc-name"><bdi>${esc(c.name)}</bdi></h3>
      ${showClient && b ? `<div class="kc-client" data-bidi>${esc(b.name)}</div>` : ''}
      ${dates}${prog}${tgt ? `<div class="kc-target" data-bidi title="${esc(tgt)}">${esc(t('cmp.target', {v:tgt}))}</div>` : ''}
      <div class="kc-foot"><span class="kc-hint ${hc}">${esc(h)}</span><span class="sp"></span>${c.stage === 'done' ? resChip(c) : ''}${bud !== null ? `<b class="kc-budget" title="${esc(budgetTip(c))}">${esc(money(bud))}</b>` : ''}</div>
    </article>`;
  }
  function bar(){
    const used = [...new Set(all.map(c=>clean(c.platform)).filter(Boolean))];
    const plats = [...CAMP_PLATFORMS.map(p=>p[0]), ...used.filter(n=>!CAMP_PLATFORMS.some(p=>p[0] === n)).sort()];
    if (P.campPlat && !plats.includes(P.campPlat)) P.campPlat = '';
    return `<div class="cg-bar">
      <div class="cg-seg" role="group" aria-label="${esc(t('cmp.mode'))}">${CAMP_MODES.map(m=>`<button data-act="cmp-mode" data-m="${m}" aria-pressed="${mode() === m}">${{board:I.board, time:I.chart, live:I.mega, res:I.check}[m]}${esc(t('cmp.mode.' + m))}</button>`).join('')}</div>
      <div class="cg-plats" role="group" aria-label="${esc(t('cmp.platFilter'))}">
        <button class="qf" data-act="cmp-plat" data-p="" aria-pressed="${!P.campPlat}">${esc(t('cmp.all'))}</button>
        ${plats.map(n=>`<button class="qf" data-act="cmp-plat" data-p="${esc(n)}" aria-pressed="${P.campPlat === n}"><i class="pf-dot" style="--sw:${platColor(n)}"></i><span data-bidi>${esc(n)}</span></button>`).join('')}
      </div><span class="sp"></span>
      ${api.canEdit() ? `<button class="btn primary" data-act="cmp-new">${I.plus}<span>${esc(t('cmp.new'))}</span></button>` : ''}
    </div>`;
  }
  function kpis(list, today, attn){
    const in7 = addDays(today, 7), by = st=>list.filter(c=>c.stage === st), live = by('live');
    const before = list.filter(c=>['plan', 'setup', 'wait', 'ok'].includes(c.stage));
    const marked = by('done').filter(c=>resultOf(c) !== 'open'), won = marked.filter(c=>resultOf(c) === 'yes'), urgent = attn.filter(a=>a[2] === 0).length;
    const gross = arr=>sumGross(arr) ? t('cmp.k.gross', {v:money(sumGross(arr))}) : '';
    const tiles = [
      [money(sum(live)), t('cmp.k.live'), 'live', gross(live)],
      [money(sum(before)), t('cmp.k.next'), '', gross(before)],
      [before.filter(c=>isISO(c.start) && c.start >= today && c.start <= in7).length, t('cmp.k.up'), '', ''],
      [live.filter(c=>isISO(c.end) && c.end >= today && c.end <= in7).length, t('cmp.k.end'), '', ''],
      [marked.length ? Math.round(won.length / marked.length * 100) + '%' : '—', t('cmp.k.win'), '', marked.length ? t('cmp.k.winOf', {a:won.length, b:marked.length}) : t('cmp.k.winNone')],
      [attn.length, t('cmp.k.attn'), attn.length ? (urgent ? 'bad' : 'warn') : 'good', urgent ? t('cmp.k.urgent', {n:urgent}) : '']
    ];
    return `<div class="cg-kpis">${tiles.map(([v, l, k, sub])=>`<div class="cg-kpi ${k}"><b>${esc(v)}</b><span>${esc(l)}</span>${sub ? `<small>${esc(sub)}</small>` : ''}</div>`).join('')}</div>` +
      (attn.length ? `<div class="cg-attn"><h3>${I.clock}${esc(t('cmp.a.title'))}<span class="cg-n">${attn.length}</span></h3>${(showAttn ? attn : attn.slice(0, ATTN_SHOWN)).map(([c, why, lv])=>`<button class="cg-attn-i" data-act="cmp-open" data-id="${c.id}">
        <span class="cg-lv l${lv}">${esc(t('cmp.lv.' + lv))}</span><b><bdi>${esc(c.name)}</bdi></b><span class="cg-chip cs-${c.stage}">${esc(t('cmp.st.' + c.stage))}</span><span class="why">${esc(why)}</span></button>`).join('')}
        ${attn.length > ATTN_SHOWN ? `<button class="btn subtle sm cg-attn-more" data-act="cmp-attn" aria-expanded="${showAttn}">${esc(showAttn ? t('cmp.a.less') : t('cmp.a.more', {n:attn.length}))}</button>` : ''}</div>` : '');
  }
  function boardView(list, today, showClient){
    const sortUp = (a, b)=>String(a.start || '9').localeCompare(String(b.start || '9')) || a.name.localeCompare(b.name);
    return `<div class="cg-board">${CAMP_STAGES.map(st=>{
      let items = list.filter(c=>c.stage === st).sort(st === 'done' ? (a, b)=>String(b.end || b.start || '').localeCompare(String(a.end || a.start || '')) : sortUp);
      const n = items.length, cut = st === 'done' && !showEnded && n > ENDED_SHOWN;
      if (cut) items = items.slice(0, ENDED_SHOWN);
      return `<section class="cg-col cs-${st}" data-st="${st}" aria-label="${esc(t('cmp.st.' + st))}">
        <header class="cg-col-h"><i class="cg-dot"></i><span>${esc(t('cmp.st.' + st))}</span><span class="cg-n">${n || ''}</span>
          ${n && sum(items) ? `<span class="sp"></span><span class="cg-col-sum">${esc(money(sum(list.filter(c=>c.stage === st))))}</span>` : ''}</header>
        <div class="cg-cards">${items.map(c=>card(c, today, showClient)).join('') || `<p class="empty-hint">${esc(t('cmp.emptyCol'))}</p>`}
          ${cut ? `<button class="btn subtle sm" data-act="cmp-ended">${esc(t('cmp.showAll', {n}))}</button>` : ''}</div>
      </section>`; }).join('')}</div>`;
  }
  function timeView(list, today, showClient){
    const d = new Date(today + 'T12:00'), d0 = addDays(today, -d.getDay() - 7), dEnd = addDays(d0, DAYS - 1);
    const rows = list.filter(c=>{
      const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '';
      if (!s && !e) return false;
      const from = s || e, to = e || (c.stage === 'done' ? from : '9999-12-31');
      return from <= dEnd && to >= d0;
    }).sort((a, b)=>String(a.start || a.end).localeCompare(String(b.start || b.end)) || a.name.localeCompare(b.name));
    const weeks = Array.from({length:DAYS / 7}, (_, w)=>`<div class="ct-week">${esc(api.fmtDay(addDays(d0, w * 7)))}</div>`).join('');
    const days = Array.from({length:DAYS}, (_, k)=>{ const day = addDays(d0, k); return `<i class="${day === today ? 'on' : ''}">${Number(day.slice(8))}</i>`; }).join('');
    const tIx = dayDiff(d0, today);
    const body = rows.map(c=>{
      const s = isISO(c.start) ? c.start : c.end, e = isISO(c.end) ? c.end : (c.stage === 'done' ? s : null), b = api.boardById(c.boardId);
      const a = Math.max(0, dayDiff(d0, s)), z = Math.min(DAYS - 1, e ? dayDiff(d0, e) : DAYS - 1);
      const sub = [showClient && b ? b.name : '', clean(c.platform)].filter(Boolean).join(' · ');
      return `<button class="ct-label" data-act="cmp-open" data-id="${c.id}"><b><bdi>${esc(c.name)}</bdi></b><small><i class="pf-dot" style="--sw:${platColor(clean(c.platform))}"></i><span data-bidi>${esc(sub)}</span></small></button>
        <div class="ct-track"><button class="ct-bar cs-${c.stage}${e ? '' : ' open'}${dayDiff(d0, s) < 0 ? ' cut' : ''}" data-act="cmp-open" data-id="${c.id}" style="--a:${a};--n:${z - a + 1}"
          title="${esc(c.name + ' · ' + t('cmp.st.' + c.stage) + ' · ' + api.fmtDay(s) + (e ? ' – ' + api.fmtDay(e) : '') + (budgetTip(c) ? ' · ' + budgetTip(c) : ''))}"><span>${esc(dm(s) + (e && e !== s ? ' – ' + dm(e) : ''))} · ${esc(t('cmp.st.' + c.stage))}${netOf(c) !== null ? ' · ' + esc(money(netOf(c))) : ''}</span></button></div>`;
    }).join('');
    return `<div class="cg-time" style="--d:${DAY_W}px;--days:${DAYS};--today:${tIx}">
      <div class="ct-label ct-corner">${esc(t('cmp.one'))}</div>
      <div class="ct-head"><div class="ct-weeks">${weeks}</div><div class="ct-days">${days}</div></div>
      ${body || `<p class="empty-hint ct-none">${esc(t('cmp.noDated'))}</p>`}
    </div>`;
  }

  // on air now: only what is running, the soonest to end first
  function liveView(list, today, showClient){
    const rows = list.filter(c=>c.stage === 'live').sort((a, b)=>String(a.end || '9').localeCompare(String(b.end || '9')) || a.name.localeCompare(b.name));
    if (!rows.length) return `<div class="cg-panel"><p class="empty-hint">${esc(t('cmp.live.none'))}</p></div>`;
    const heads = ['campaign', showClient ? 'client' : '', 'platform', 'objective', 'dates', 'progress', 'net', 'gross', 'target'].filter(Boolean);
    return `<div class="cg-panel cg-tbl-wrap"><table class="cg-tbl">
      <thead><tr>${heads.map(h=>`<th scope="col" class="${h === 'net' || h === 'gross' ? 'num' : ''}">${esc(t('cmp.c.' + h))}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(c=>{
        const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '', b = api.boardById(c.boardId), [h, hc] = hint(c, today);
        return `<tr class="cg-row cs-live" data-act="cmp-open" data-id="${c.id}" tabindex="0">
          <th scope="row"><bdi>${esc(c.name)}</bdi></th>${showClient ? `<td data-bidi>${esc(b ? b.name : '')}</td>` : ''}
          <td>${platChip(clean(c.platform))}</td><td data-bidi>${esc(objLabel(clean(c.objective)))}</td>
          <td class="nw">${esc(s ? dm(s) : '?')} – ${esc(e ? dm(e) : '…')}</td>
          <td><div class="cg-pr">${progress(c, today)}<span class="kc-hint ${hc}">${esc(h)}</span></div></td>
          <td class="num">${netOf(c) !== null ? esc(money(netOf(c))) : ''}</td><td class="num">${grossOf(c) !== null ? esc(money(grossOf(c))) : ''}</td>
          <td class="tgt" data-bidi>${esc(clean(c.target))}</td></tr>`; }).join('')}</tbody>
      <tfoot><tr><th scope="row" colspan="${heads.length - 3}">${esc(t('cmp.live.total', {n:rows.length}))}</th><td class="num">${esc(money(sum(rows)))}</td><td class="num">${sumGross(rows) ? esc(money(sumGross(rows))) : ''}</td><td></td></tr></tfoot>
    </table></div>`;
  }
  // how the ended campaigns did: one bar for all of them and one per objective, then the campaigns themselves
  function resView(list, today, showClient){
    const done = list.filter(c=>c.stage === 'done');
    if (!done.length) return `<div class="cg-panel"><p class="empty-hint">${esc(t('cmp.res.none'))}</p></div>`;
    const groups = new Map();
    done.forEach(c=>{ const o = clean(c.objective); if (!groups.has(o)) groups.set(o, []); groups.get(o).push(c); });
    const lines = [[t('cmp.res.all'), done], ...[...groups].sort((a, b)=>b[1].length - a[1].length || a[0].localeCompare(b[0])).map(([o, arr])=>[o ? objLabel(o) : t('cmp.res.noObj'), arr])];
    const most = Math.max(...lines.map(l=>l[1].length));
    const chart = lines.map(([name, arr], i)=>{
      const n = {yes:0, no:0, open:0}; arr.forEach(c=>n[resultOf(c)]++);
      const rated = n.yes + n.no, text = t('cmp.res.line', {y:n.yes, n:n.no}) + (n.open ? ' · ' + t('cmp.res.open', {o:n.open}) : '');
      return `<div class="rz-row${i ? '' : ' all'}"><span class="rz-name" data-bidi>${esc(name)}</span>
        <div class="rz-track"><div class="rz-bar" style="width:${arr.length / most * 100}%" role="img" aria-label="${esc(name + ': ' + text)}">${CAMP_RESULTS.slice(1).concat('open').map(k=>n[k]
          ? `<i class="rs-${k}" style="flex:${n[k]}" title="${esc(t('cmp.rs.' + k) + ': ' + n[k])}"></i>` : '').join('')}</div></div>
        <span class="rz-val">${esc(text)}${rated ? ` · <b>${esc(t('cmp.res.rate', {p:Math.round(n.yes / rated * 100)}))}</b>` : ''}</span></div>`;
    }).join('');
    const ended = done.slice().sort((a, b)=>String(b.end || b.start || '').localeCompare(String(a.end || a.start || '')) || a.name.localeCompare(b.name));
    const heads = ['campaign', showClient ? 'client' : '', 'platform', 'objective', 'ended', 'net', 'target', 'result', 'after'].filter(Boolean);
    return `<div class="cg-panel cg-res"><header class="rz-head"><h3>${esc(t('cmp.res.title'))}</h3><span>${esc(t('cmp.res.sub'))}</span><span class="sp"></span>
        <div class="rz-legend">${['yes', 'no', 'open'].map(k=>`<span class="rs-${k}"><i></i>${esc(t('cmp.rs.' + k))}</span>`).join('')}</div></header>
      <div class="rz-chart">${chart}</div></div>
      <div class="cg-panel cg-tbl-wrap"><table class="cg-tbl"><caption>${esc(t('cmp.res.ended'))}</caption>
      <thead><tr>${heads.map(h=>`<th scope="col" class="${h === 'net' ? 'num' : ''}">${esc(t('cmp.c.' + h))}</th>`).join('')}</tr></thead>
      <tbody>${ended.map(c=>{ const b = api.boardById(c.boardId), e = isISO(c.end) ? c.end : '';
        return `<tr class="cg-row" data-act="cmp-open" data-id="${c.id}" tabindex="0"><th scope="row"><bdi>${esc(c.name)}</bdi></th>${showClient ? `<td data-bidi>${esc(b ? b.name : '')}</td>` : ''}
          <td>${platChip(clean(c.platform))}</td><td data-bidi>${esc(objLabel(clean(c.objective)))}</td><td class="nw">${esc(e ? api.fmtDay(e) : '')}</td>
          <td class="num" title="${esc(budgetTip(c))}">${netOf(c) !== null ? esc(money(netOf(c))) : ''}</td><td class="tgt" data-bidi>${esc(clean(c.target))}</td>
          <td>${resChip(c)}</td><td class="aft"><div data-bidi>${esc(String(c.after || '').trim())}</div></td></tr>`; }).join('')}</tbody></table></div>`;
  }

  /* ---------- the screen ---------- */
  function render(){
    const today = api.today(), list = scope(), b = api.scopeBoard(), attn = attention(list, today);
    if (!all.length){
      el.innerHTML = `<div class="cg">${bar()}<div class="empty-state"><b>${esc(t('cmp.empty'))}</b>${esc(t('cmp.emptySub'))}</div></div>`;
      return '';
    }
    const view = {board:boardView, time:timeView, live:liveView, res:resView}[mode()];
    el.innerHTML = `<div class="cg">${bar()}${kpis(list, today, attn)}${view(list, today, !b)}</div>`;
    return t('cmp.sum', {live:list.filter(c=>c.stage === 'live').length, wait:list.filter(c=>c.stage === 'wait').length});
  }

  /* ---------- the dialog ---------- */
  function open(id){
    const c = id ? get(id) : null, edit = api.canEdit(), boards = api.boards();
    if (id && !c) return;
    if (!c && !edit) return;
    if (!c && !boards.length) return api.toast(t('cmp.noBoards'));
    const cur = c || {name:'', boardId:(api.scopeBoard() || boards[0]).id, platform:P.campPlat || '', objective:'', stage:'plan', start:'', end:'', net:null, gross:null, target:'', result:'open', after:'', link:'', notes:''};
    const plat = clean(cur.platform), known = CAMP_PLATFORMS.some(p=>p[0] === plat), dis = edit ? '' : ' disabled';
    const obj = clean(cur.objective), objKnown = CAMP_OBJECTIVES.includes(obj);
    const dlg = api.showModal(`<form novalidate style="display:contents">
      ${api.head(c ? (edit ? t('cmp.edit') : t('cmp.one')) : t('cmp.new'))}
      <div class="c-body cg-form">
        <label class="fld"><span>${esc(t('cmp.f.name'))} <span class="req" aria-hidden="true">*</span></span><input name="name" data-bidi autocomplete="off" maxlength="200" required value="${esc(cur.name)}"${dis}></label>
        <div class="grid2">
          <label class="fld">${esc(t('cmp.f.client'))}<select name="board"${dis}>${!api.boardById(cur.boardId) ? `<option value="" selected>${esc(t('cmp.noClient'))}</option>` : ''}${boards.map(x=>`<option value="${x.id}"${x.id === cur.boardId ? ' selected' : ''}>${esc(x.name)}</option>`).join('')}</select></label>
          <label class="fld">${esc(t('cmp.f.platform'))}<select name="platform"${dis}><option value=""></option>${CAMP_PLATFORMS.map(p=>`<option${p[0] === plat ? ' selected' : ''}>${esc(p[0])}</option>`).join('')}<option value="*"${plat && !known ? ' selected' : ''}>${esc(t('cmp.other'))}</option></select>
            <input name="platformOther" data-bidi maxlength="40" placeholder="${esc(t('cmp.otherPh'))}" value="${esc(known ? '' : plat)}"${plat && !known ? '' : ' hidden'}${dis}></label>
        </div>
        <div class="grid2">
          <label class="fld">${esc(t('cmp.f.objective'))}<select name="objective"${dis}><option value=""></option>${CAMP_OBJECTIVES.map((o, i)=>`<option value="${esc(o)}"${o === obj ? ' selected' : ''}>${esc(t('cmp.o.' + i))}</option>`).join('')}<option value="*"${obj && !objKnown ? ' selected' : ''}>${esc(t('cmp.other'))}</option></select>
            <input name="objectiveOther" data-bidi maxlength="40" placeholder="${esc(t('cmp.objPh'))}" value="${esc(objKnown ? '' : obj)}"${obj && !objKnown ? '' : ' hidden'}${dis}></label>
          <label class="fld">${esc(t('cmp.f.target'))}<input name="target" data-bidi autocomplete="off" maxlength="200" placeholder="${esc(t('cmp.targetPh'))}" value="${esc(cur.target || '')}"${dis}></label>
        </div>
        <fieldset class="fld cg-stages"><legend>${esc(t('cmp.f.stage'))}</legend><div class="cg-stage-row">${CAMP_STAGES.map(st=>`<label class="cg-chip cs-${st}"><input type="radio" name="stage" value="${st}"${st === cur.stage ? ' checked' : ''}${dis}><span>${esc(t('cmp.st.' + st))}</span></label>`).join('')}</div>
          <small>${esc(t('cmp.autoNote'))}</small></fieldset>
        <div class="grid2">
          <label class="fld">${esc(t('cmp.f.start'))}<input type="date" name="start" value="${isISO(cur.start) ? cur.start : ''}"${dis}></label>
          <label class="fld">${esc(t('cmp.f.end'))}<input type="date" name="end" value="${isISO(cur.end) ? cur.end : ''}"${dis}></label>
        </div>
        <div class="grid2">
          <label class="fld">${esc(t('cmp.f.net'))}<input type="number" name="net" min="0" step="1" inputmode="decimal" value="${netOf(cur) ?? ''}"${dis}></label>
          <label class="fld">${esc(t('cmp.f.gross'))}<input type="number" name="gross" min="0" step="1" inputmode="decimal" value="${grossOf(cur) ?? ''}"${dis}></label>
        </div>
        <label class="fld">${esc(t('cmp.f.link'))}<span class="cg-link-row"><input name="link" dir="ltr" inputmode="url" autocomplete="off" placeholder="https://" value="${esc(cur.link || '')}"${dis}>
          ${cur.link ? `<a class="btn icon subtle" href="${esc(cur.link)}" target="_blank" rel="noopener noreferrer" title="${esc(t('cmp.openLink'))}" aria-label="${esc(t('cmp.openLink'))}">${I.open}</a>` : ''}</span></label>
        <label class="fld">${esc(t('cmp.f.notes'))}<textarea name="notes" rows="3" data-bidi maxlength="2000"${dis}>${esc(cur.notes || '')}</textarea></label>
        <fieldset class="fld cg-stages cg-after"><legend>${esc(t('cmp.f.result'))}</legend><div class="cg-stage-row">${CAMP_RESULTS.map(r=>`<label class="cg-chip rs-${r}"><input type="radio" name="result" value="${r}"${r === resultOf(cur) ? ' checked' : ''}${dis}><span>${esc(t('cmp.r.' + r))}</span></label>`).join('')}</div>
          <small>${esc(t('cmp.resultNote'))}</small></fieldset>
        <label class="fld">${esc(t('cmp.f.after'))}<textarea name="after" rows="2" data-bidi maxlength="2000" placeholder="${esc(t('cmp.afterPh'))}"${dis}>${esc(cur.after || '')}</textarea></label>
        <em class="err cg-err" hidden></em>
      </div>
      <div class="m-foot">${c && edit ? `<button type="button" class="btn danger-subtle" data-act="cmp-del" data-id="${c.id}">${I.trash}<span>${esc(t('cmp.del'))}</span></button>` : ''}<span class="sp"></span>
        <button type="button" class="btn subtle" data-act="close-modal">${esc(edit ? t('btn.cancel') : t('btn.close'))}</button>${edit ? `<button type="submit" class="btn primary">${esc(c ? t('btn.save') : t('btn.create'))}</button>` : ''}</div>
    </form>`, 'campaign', 'narrow');
    dlg.setAttribute('aria-label', c ? t('cmp.edit') : t('cmp.new'));
    const f = dlg.querySelector('form'), err = dlg.querySelector('.cg-err');
    f.platform.addEventListener('change', ()=>{ f.platformOther.hidden = f.platform.value !== '*'; if (!f.platformOther.hidden) f.platformOther.focus(); });
    f.objective.addEventListener('change', ()=>{ f.objectiveOther.hidden = f.objective.value !== '*'; if (!f.objectiveOther.hidden) f.objectiveOther.focus(); });
    f.addEventListener('submit', e=>{
      e.preventDefault();
      if (!api.canEdit()) return;
      const name = clean(f.name.value).slice(0, 200), start = f.start.value, end = f.end.value;
      const fail = !name ? t('cmp.nameReq') : isISO(start) && isISO(end) && end < start ? t('cmp.datesBad') : '';
      err.textContent = fail; err.hidden = !fail;
      if (fail){ (name ? f.end : f.name).focus(); return; }
      const amount = inp=>inp.value === '' ? null : num(Math.round(Number(inp.value) * 100) / 100), text = inp=>String(inp.value || '').replace(/\r\n?/g, '\n').trim().slice(0, 2000);
      const net = amount(f.net), now = Date.now();
      // "budget" is the name the net budget had at first: kept equal to it, for a sheet whose script was not updated yet
      const data = {kind:'campaign', name, boardId:f.board.value, platform:f.platform.value === '*' ? clean(f.platformOther.value).slice(0, 40) : f.platform.value,
        objective:f.objective.value === '*' ? clean(f.objectiveOther.value).slice(0, 40) : f.objective.value,
        stage:f.stage.value || 'plan', start:isISO(start) ? start : '', end:isISO(end) ? end : '', net, gross:amount(f.gross), budget:net,
        target:clean(f.target.value).slice(0, 200), result:CAMP_RESULTS.includes(f.result.value) ? f.result.value : 'open', after:text(f.after),
        link:normLink(f.link.value), notes:text(f.notes), archived:false,
        created:c ? (c.created || now) : now, updated:now, by:api.userName()};
      data.stage = stageNow(data, api.today());
      // on top of what the campaign already holds, so a field this screen does not know is kept
      const next = {...(c || {}), ...data, id:c ? c.id : 'cmp_' + api.uid().replace(/[^A-Za-z0-9]/g, '')};
      if (c) all[all.indexOf(c)] = next; else all.push(next);
      write([next]); api.closeModal(); api.rerender();
    });
    if (!c && !api.coarse()) setTimeout(()=>f.name.focus(), 60);
  }

  /* ---------- clicks ---------- */
  function onClick(act, b, e){
    const id = b.dataset.id;
    switch (act){
      case 'cmp-new': open(null); break;
      case 'cmp-open': if (e.target.closest('a')) break; open(id); break;
      case 'cmp-mode': P.campMode = b.dataset.m; savePrefs(); api.rerender(); break;
      case 'cmp-plat': P.campPlat = b.dataset.p || ''; savePrefs(); api.rerender(); break;
      case 'cmp-ended': showEnded = true; api.rerender(); break;
      case 'cmp-attn': showAttn = !showAttn; api.rerender(); break;
      case 'cmp-move': {
        e.stopPropagation();
        const c = get(id); if (!c) break;
        api.openPop(b, `<div class="menu-h">${esc(t('cmp.moveTo'))}</div>` + CAMP_STAGES.map(st=>`<button class="menu-i" role="menuitem" data-act="cmp-stage" data-id="${id}" data-st="${st}">
          <span class="cg-chip cs-${st}">${esc(t('cmp.st.' + st))}</span>${st === c.stage ? `<span class="chk-mark">${I.check}</span>` : ''}</button>`).join(''), id);
        break; }
      case 'cmp-stage': api.closePop(); patch(id, {stage:b.dataset.st}); break;
      case 'cmp-del': {
        if (!b.classList.contains('armed')){ b.classList.add('armed'); b.querySelector('span').textContent = t('cmp.delSure'); break; }
        const c = get(id); if (!c) break;
        all = all.filter(x=>x !== c);
        write([{...c, archived:true, updated:Date.now(), by:api.userName()}]);
        api.closeModal(); api.toast(t('cmp.deleted')); api.rerender();
        break; }
    }
  }
  el.addEventListener('keydown', e=>{ if ((e.key === 'Enter' || e.key === ' ') && e.target.matches('.cg-card,.cg-row')){ e.preventDefault(); open(e.target.dataset.id); } });
  // drag a card to another stage (mouse; on touch the card's menu does the same)
  el.addEventListener('dragstart', e=>{ const c = e.target.closest?.('.cg-card'); if (!c || !api.canEdit()) return; dragId = c.dataset.id; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', dragId); c.classList.add('dragging'); });
  el.addEventListener('dragend', ()=>{ dragId = null; el.querySelectorAll('.dragging,.over').forEach(x=>x.classList.remove('dragging', 'over')); });
  el.addEventListener('dragover', e=>{ const col = dragId && e.target.closest('.cg-col'); if (!col) return; e.preventDefault(); e.dataTransfer.dropEffect = 'move';
    if (!col.classList.contains('over')){ el.querySelectorAll('.cg-col.over').forEach(x=>x.classList.remove('over')); col.classList.add('over'); } });
  el.addEventListener('drop', e=>{ const col = dragId && e.target.closest('.cg-col'); if (!col) return; e.preventDefault(); const id = dragId, c = get(id); dragId = null; if (c && c.stage !== col.dataset.st) patch(id, {stage:col.dataset.st}); else api.rerender(); });

  return { setData, render, onClick, open, count:()=>all.length };
}
