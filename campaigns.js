// Campaigns: what is being planned, set up, waiting for approval, running and ended, per client and platform.
// A board with a column per stage, and a timeline. The same campaigns can be kept in a Google Sheets tab
// ("קמפיינים"), which a script keeps in step with this screen.
// They are stored next to the workspace's settings, as small documents named "cmp_…".
import { t, locale, _dicts } from './i18n.js';

export const CAMP_STAGES = ['plan', 'setup', 'wait', 'ok', 'live', 'paused', 'done'];
export const CAMP_PLATFORMS = [['Google Ads', '#1D7AFC'], ['Meta', '#8F7EE7'], ['TikTok', '#DA62AC'], ['Taboola', '#F38A3F'], ['ChatGPT', '#22A06B']];
const OTHER_COLOR = '#8590A2';
const CAMP_ID = /^cmp_[A-Za-z0-9]{1,40}$/;
const DAYS = 70, DAY_W = 22, ENDED_SHOWN = 8;

Object.assign(_dicts.he, {
  'side.camps':'קמפיינים', 'view.campsOf':'{name}: קמפיינים', 'cmp.new':'קמפיין חדש', 'cmp.edit':'עריכת קמפיין', 'cmp.one':'קמפיין',
  'cmp.mode.board':'לוח', 'cmp.mode.time':'ציר זמן', 'cmp.mode':'תצוגה', 'cmp.all':'כל הפלטפורמות', 'cmp.platFilter':'סינון לפי פלטפורמה',
  'cmp.st.plan':'תכנון', 'cmp.st.setup':'בהקמה', 'cmp.st.wait':'ממתין לאישור', 'cmp.st.ok':'מאושר', 'cmp.st.live':'רץ', 'cmp.st.paused':'מושהה', 'cmp.st.done':'הסתיים',
  'cmp.f.name':'שם הקמפיין', 'cmp.f.client':'לקוח', 'cmp.f.platform':'פלטפורמה', 'cmp.f.stage':'שלב', 'cmp.f.start':'תאריך עלייה', 'cmp.f.end':'תאריך סיום',
  'cmp.f.budget':'תקציב (₪)', 'cmp.f.link':'קישור (תוכנית מדיה, דוח…)', 'cmp.f.notes':'הערות', 'cmp.other':'אחר…', 'cmp.otherPh':'שם הפלטפורמה',
  'cmp.k.live':'תקציב קמפיינים רצים', 'cmp.k.next':'תקציב בדרך', 'cmp.k.up':'עולים ב־7 הימים הקרובים', 'cmp.k.end':'מסתיימים ב־7 הימים הקרובים', 'cmp.k.attn':'דורשים תשומת לב',
  'cmp.h.today':'עולה היום', 'cmp.h.tomorrow':'עולה מחר', 'cmp.h.in':'עולה בעוד {n} ימים', 'cmp.h.late':'תאריך העלייה עבר', 'cmp.h.left':'נותרו {n} ימים', 'cmp.h.left1':'נותר יום אחד',
  'cmp.h.endsToday':'מסתיים היום', 'cmp.h.noEnd':'ללא תאריך סיום', 'cmp.h.ended':'הסתיים ב־{d}', 'cmp.h.noDates':'בלי תאריכים',
  'cmp.a.title':'דורש תשומת לב', 'cmp.a.late':'תאריך העלייה ({d}) עבר, והקמפיין עדיין ב"{st}"', 'cmp.a.soon':'עולה בעוד {n} ימים ועדיין ממתין לאישור',
  'cmp.a.soon0':'אמור לעלות היום ועדיין ממתין לאישור', 'cmp.a.noStart':'מאושר, אבל בלי תאריך עלייה', 'cmp.a.paused':'הקמפיין מושהה',
  'cmp.empty':'עדיין אין קמפיינים', 'cmp.emptySub':'הוסיפו קמפיין כאן, או בלשונית "קמפיינים" בגיליון.', 'cmp.emptyCol':'אין קמפיינים בשלב הזה', 'cmp.noDated':'אין קמפיינים עם תאריכים בתקופה הזו',
  'cmp.moveTo':'העברה לשלב', 'cmp.del':'מחיקת הקמפיין', 'cmp.delSure':'ללחוץ שוב למחיקה', 'cmp.deleted':'הקמפיין נמחק', 'cmp.nameReq':'צריך שם לקמפיין',
  'cmp.datesBad':'תאריך הסיום מוקדם מתאריך העלייה', 'cmp.noClient':'ללא לקוח', 'cmp.showAll':'הצגת כל {n} שהסתיימו', 'cmp.sum':'{live} רצים · {wait} ממתינים לאישור',
  'cmp.openLink':'פתיחת הקישור', 'cmp.autoNote':'"מאושר" הופך ל"רץ" בתאריך העלייה, ו"רץ" הופך ל"הסתיים" אחרי תאריך הסיום.', 'cmp.today':'היום', 'cmp.noBoards':'קודם צריך לוח של לקוח'
});
Object.assign(_dicts.en, {
  'side.camps':'Campaigns', 'view.campsOf':'{name}: campaigns', 'cmp.new':'New campaign', 'cmp.edit':'Edit campaign', 'cmp.one':'Campaign',
  'cmp.mode.board':'Board', 'cmp.mode.time':'Timeline', 'cmp.mode':'View', 'cmp.all':'All platforms', 'cmp.platFilter':'Filter by platform',
  'cmp.st.plan':'Planning', 'cmp.st.setup':'Setting up', 'cmp.st.wait':'Waiting for approval', 'cmp.st.ok':'Approved', 'cmp.st.live':'Running', 'cmp.st.paused':'Paused', 'cmp.st.done':'Ended',
  'cmp.f.name':'Campaign name', 'cmp.f.client':'Client', 'cmp.f.platform':'Platform', 'cmp.f.stage':'Stage', 'cmp.f.start':'Launch date', 'cmp.f.end':'End date',
  'cmp.f.budget':'Budget (₪)', 'cmp.f.link':'Link (media plan, report…)', 'cmp.f.notes':'Notes', 'cmp.other':'Other…', 'cmp.otherPh':'Platform name',
  'cmp.k.live':'Budget of running campaigns', 'cmp.k.next':'Budget on the way', 'cmp.k.up':'Launching in the next 7 days', 'cmp.k.end':'Ending in the next 7 days', 'cmp.k.attn':'Need attention',
  'cmp.h.today':'Launches today', 'cmp.h.tomorrow':'Launches tomorrow', 'cmp.h.in':'Launches in {n} days', 'cmp.h.late':'Launch date has passed', 'cmp.h.left':'{n} days left', 'cmp.h.left1':'1 day left',
  'cmp.h.endsToday':'Ends today', 'cmp.h.noEnd':'No end date', 'cmp.h.ended':'Ended {d}', 'cmp.h.noDates':'No dates',
  'cmp.a.title':'Needs attention', 'cmp.a.late':'The launch date ({d}) has passed and the campaign is still “{st}”', 'cmp.a.soon':'Launches in {n} days and is still waiting for approval',
  'cmp.a.soon0':'Should launch today and is still waiting for approval', 'cmp.a.noStart':'Approved, but has no launch date', 'cmp.a.paused':'The campaign is paused',
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
const budgetOf = c=>typeof c.budget === 'number' && isFinite(c.budget) && c.budget >= 0 ? c.budget : null;
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
// what needs a look, most urgent first: [campaign, why]
export function attention(list, today){
  const out = [];
  list.forEach(c=>{
    const st = c.stage, s = isISO(c.start) ? c.start : '';
    if (['plan', 'setup', 'wait'].includes(st) && s && s < today) out.push([c, t('cmp.a.late', {d:s.slice(8) * 1 + '/' + s.slice(5, 7) * 1, st:t('cmp.st.' + st)}), 0]);
    else if (st === 'wait' && s && dayDiff(today, s) <= 3) out.push([c, dayDiff(today, s) === 0 ? t('cmp.a.soon0') : t('cmp.a.soon', {n:dayDiff(today, s)}), 1]);
    else if (st === 'ok' && !s) out.push([c, t('cmp.a.noStart'), 2]);
    else if (st === 'paused') out.push([c, t('cmp.a.paused'), 3]);
  });
  return out.sort((a, b)=>a[2] - b[2] || String(a[0].start || '').localeCompare(String(b[0].start || '')));
}

export function createCampaigns(api){
  const { el, esc, I, P, savePrefs } = api;
  let all = [], showEnded = false, dragId = null;
  const fixed = new Set();
  const money = n=>new Intl.NumberFormat(locale(), {style:'currency', currency:'ILS', maximumFractionDigits:0}).format(n);
  const sum = arr=>arr.reduce((s, c)=>s + (budgetOf(c) || 0), 0);
  const get = id=>all.find(c=>c.id === id);
  const mode = ()=>P.campMode === 'time' ? 'time' : 'board';

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
    const s = isISO(c.start) ? c.start : '', e = isISO(c.end) ? c.end : '', b = api.boardById(c.boardId), [h, hc] = hint(c, today), bud = budgetOf(c);
    let prog = '';
    if (c.stage === 'live' && s && e && e >= s){
      const pct = Math.max(4, Math.min(100, Math.round((dayDiff(s, today) + 1) / (dayDiff(s, e) + 1) * 100)));
      prog = `<div class="kc-prog" role="progressbar" aria-valuenow="${pct}" aria-valuemin="0" aria-valuemax="100"><i style="width:${pct}%"></i></div>`;
    }
    const dates = s || e ? `<div class="kc-dates">${I.cal}<span>${esc(s ? api.fmtDay(s) : '?')}${e ? ' – ' + esc(api.fmtDay(e)) : ''}</span></div>` : '';
    return `<article class="cg-card cs-${c.stage}" data-act="cmp-open" data-id="${c.id}" tabindex="0" role="button"${api.canEdit() ? ' draggable="true"' : ''}>
      <div class="kc-top">${platChip(clean(c.platform))}<span class="sp"></span>
        ${c.link ? `<a class="kc-link" href="${esc(c.link)}" target="_blank" rel="noopener noreferrer" title="${esc(t('cmp.openLink'))}" aria-label="${esc(t('cmp.openLink'))}">${I.link}</a>` : ''}${showClient && b ? api.bAv(b) : ''}
        ${api.canEdit() ? `<button class="btn icon subtle sm kc-move" data-act="cmp-move" data-id="${c.id}" aria-haspopup="menu" aria-label="${esc(t('cmp.moveTo'))}" title="${esc(t('cmp.moveTo'))}">${I.more}</button>` : ''}</div>
      <h3 class="kc-name"><bdi>${esc(c.name)}</bdi></h3>
      ${showClient && b ? `<div class="kc-client" data-bidi>${esc(b.name)}</div>` : ''}
      ${dates}${prog}
      <div class="kc-foot"><span class="kc-hint ${hc}">${esc(h)}</span><span class="sp"></span>${bud !== null ? `<b class="kc-budget">${esc(money(bud))}</b>` : ''}</div>
    </article>`;
  }
  function bar(){
    const used = [...new Set(all.map(c=>clean(c.platform)).filter(Boolean))];
    const plats = [...CAMP_PLATFORMS.map(p=>p[0]), ...used.filter(n=>!CAMP_PLATFORMS.some(p=>p[0] === n)).sort()];
    if (P.campPlat && !plats.includes(P.campPlat)) P.campPlat = '';
    return `<div class="cg-bar">
      <div class="cg-seg" role="group" aria-label="${esc(t('cmp.mode'))}">${['board', 'time'].map(m=>`<button data-act="cmp-mode" data-m="${m}" aria-pressed="${mode() === m}">${m === 'board' ? I.board : I.chart}${esc(t('cmp.mode.' + m))}</button>`).join('')}</div>
      <div class="cg-plats" role="group" aria-label="${esc(t('cmp.platFilter'))}">
        <button class="qf" data-act="cmp-plat" data-p="" aria-pressed="${!P.campPlat}">${esc(t('cmp.all'))}</button>
        ${plats.map(n=>`<button class="qf" data-act="cmp-plat" data-p="${esc(n)}" aria-pressed="${P.campPlat === n}"><i class="pf-dot" style="--sw:${platColor(n)}"></i><span data-bidi>${esc(n)}</span></button>`).join('')}
      </div><span class="sp"></span>
      ${api.canEdit() ? `<button class="btn primary" data-act="cmp-new">${I.plus}<span>${esc(t('cmp.new'))}</span></button>` : ''}
    </div>`;
  }
  function kpis(list, today, attn){
    const in7 = addDays(today, 7), by = st=>list.filter(c=>c.stage === st);
    const before = list.filter(c=>['plan', 'setup', 'wait', 'ok'].includes(c.stage));
    const tiles = [
      [money(sum(by('live'))), t('cmp.k.live'), 'live'],
      [money(sum(before)), t('cmp.k.next'), ''],
      [before.filter(c=>isISO(c.start) && c.start >= today && c.start <= in7).length, t('cmp.k.up'), ''],
      [by('live').filter(c=>isISO(c.end) && c.end >= today && c.end <= in7).length, t('cmp.k.end'), ''],
      [attn.length, t('cmp.k.attn'), attn.length ? 'bad' : 'good']
    ];
    return `<div class="cg-kpis">${tiles.map(([v, l, k])=>`<div class="cg-kpi ${k}"><b>${esc(v)}</b><span>${esc(l)}</span></div>`).join('')}</div>` +
      (attn.length ? `<div class="cg-attn"><h3>${I.clock}${esc(t('cmp.a.title'))}</h3>${attn.map(([c, why])=>`<button class="cg-attn-i" data-act="cmp-open" data-id="${c.id}">
        <span class="cg-chip cs-${c.stage}">${esc(t('cmp.st.' + c.stage))}</span><b><bdi>${esc(c.name)}</bdi></b><span class="why">${esc(why)}</span></button>`).join('')}</div>` : '');
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
          title="${esc(c.name + ' · ' + t('cmp.st.' + c.stage) + ' · ' + api.fmtDay(s) + (e ? ' – ' + api.fmtDay(e) : ''))}"><span data-bidi>${esc(t('cmp.st.' + c.stage))}${budgetOf(c) !== null ? ' · ' + esc(money(budgetOf(c))) : ''}</span></button></div>`;
    }).join('');
    return `<div class="cg-time" style="--d:${DAY_W}px;--days:${DAYS};--today:${tIx}">
      <div class="ct-label ct-corner">${esc(t('cmp.one'))}</div>
      <div class="ct-head"><div class="ct-weeks">${weeks}</div><div class="ct-days">${days}</div></div>
      ${body || `<p class="empty-hint ct-none">${esc(t('cmp.noDated'))}</p>`}
    </div>`;
  }

  /* ---------- the screen ---------- */
  function render(){
    const today = api.today(), list = scope(), b = api.scopeBoard(), attn = attention(list, today);
    if (!all.length){
      el.innerHTML = `<div class="cg">${bar()}<div class="empty-state"><b>${esc(t('cmp.empty'))}</b>${esc(t('cmp.emptySub'))}</div></div>`;
      return '';
    }
    el.innerHTML = `<div class="cg">${bar()}${kpis(list, today, attn)}${mode() === 'time' ? timeView(list, today, !b) : boardView(list, today, !b)}</div>`;
    return t('cmp.sum', {live:list.filter(c=>c.stage === 'live').length, wait:list.filter(c=>c.stage === 'wait').length});
  }

  /* ---------- the dialog ---------- */
  function open(id){
    const c = id ? get(id) : null, edit = api.canEdit(), boards = api.boards();
    if (id && !c) return;
    if (!c && !edit) return;
    if (!c && !boards.length) return api.toast(t('cmp.noBoards'));
    const cur = c || {name:'', boardId:(api.scopeBoard() || boards[0]).id, platform:P.campPlat || '', stage:'plan', start:'', end:'', budget:null, link:'', notes:''};
    const plat = clean(cur.platform), known = CAMP_PLATFORMS.some(p=>p[0] === plat), dis = edit ? '' : ' disabled';
    const dlg = api.showModal(`<form novalidate style="display:contents">
      ${api.head(c ? (edit ? t('cmp.edit') : t('cmp.one')) : t('cmp.new'))}
      <div class="c-body cg-form">
        <label class="fld"><span>${esc(t('cmp.f.name'))} <span class="req" aria-hidden="true">*</span></span><input name="name" data-bidi autocomplete="off" maxlength="200" required value="${esc(cur.name)}"${dis}></label>
        <div class="grid2">
          <label class="fld">${esc(t('cmp.f.client'))}<select name="board"${dis}>${!api.boardById(cur.boardId) ? `<option value="" selected>${esc(t('cmp.noClient'))}</option>` : ''}${boards.map(x=>`<option value="${x.id}"${x.id === cur.boardId ? ' selected' : ''}>${esc(x.name)}</option>`).join('')}</select></label>
          <label class="fld">${esc(t('cmp.f.platform'))}<select name="platform"${dis}><option value=""></option>${CAMP_PLATFORMS.map(p=>`<option${p[0] === plat ? ' selected' : ''}>${esc(p[0])}</option>`).join('')}<option value="*"${plat && !known ? ' selected' : ''}>${esc(t('cmp.other'))}</option></select>
            <input name="platformOther" data-bidi maxlength="40" placeholder="${esc(t('cmp.otherPh'))}" value="${esc(known ? '' : plat)}"${plat && !known ? '' : ' hidden'}${dis}></label>
        </div>
        <fieldset class="fld cg-stages"><legend>${esc(t('cmp.f.stage'))}</legend><div class="cg-stage-row">${CAMP_STAGES.map(st=>`<label class="cg-chip cs-${st}"><input type="radio" name="stage" value="${st}"${st === cur.stage ? ' checked' : ''}${dis}><span>${esc(t('cmp.st.' + st))}</span></label>`).join('')}</div>
          <small>${esc(t('cmp.autoNote'))}</small></fieldset>
        <div class="grid3">
          <label class="fld">${esc(t('cmp.f.start'))}<input type="date" name="start" value="${isISO(cur.start) ? cur.start : ''}"${dis}></label>
          <label class="fld">${esc(t('cmp.f.end'))}<input type="date" name="end" value="${isISO(cur.end) ? cur.end : ''}"${dis}></label>
          <label class="fld">${esc(t('cmp.f.budget'))}<input type="number" name="budget" min="0" step="1" inputmode="decimal" value="${budgetOf(cur) ?? ''}"${dis}></label>
        </div>
        <label class="fld">${esc(t('cmp.f.link'))}<span class="cg-link-row"><input name="link" dir="ltr" inputmode="url" autocomplete="off" placeholder="https://" value="${esc(cur.link || '')}"${dis}>
          ${cur.link ? `<a class="btn icon subtle" href="${esc(cur.link)}" target="_blank" rel="noopener noreferrer" title="${esc(t('cmp.openLink'))}" aria-label="${esc(t('cmp.openLink'))}">${I.open}</a>` : ''}</span></label>
        <label class="fld">${esc(t('cmp.f.notes'))}<textarea name="notes" rows="3" data-bidi maxlength="2000"${dis}>${esc(cur.notes || '')}</textarea></label>
        <em class="err cg-err" hidden></em>
      </div>
      <div class="m-foot">${c && edit ? `<button type="button" class="btn danger-subtle" data-act="cmp-del" data-id="${c.id}">${I.trash}<span>${esc(t('cmp.del'))}</span></button>` : ''}<span class="sp"></span>
        <button type="button" class="btn subtle" data-act="close-modal">${esc(edit ? t('btn.cancel') : t('btn.close'))}</button>${edit ? `<button type="submit" class="btn primary">${esc(c ? t('btn.save') : t('btn.create'))}</button>` : ''}</div>
    </form>`, 'campaign', 'narrow');
    dlg.setAttribute('aria-label', c ? t('cmp.edit') : t('cmp.new'));
    const f = dlg.querySelector('form'), err = dlg.querySelector('.cg-err');
    f.platform.addEventListener('change', ()=>{ f.platformOther.hidden = f.platform.value !== '*'; if (!f.platformOther.hidden) f.platformOther.focus(); });
    f.addEventListener('submit', e=>{
      e.preventDefault();
      if (!api.canEdit()) return;
      const name = clean(f.name.value).slice(0, 200), start = f.start.value, end = f.end.value;
      const fail = !name ? t('cmp.nameReq') : isISO(start) && isISO(end) && end < start ? t('cmp.datesBad') : '';
      err.textContent = fail; err.hidden = !fail;
      if (fail){ (name ? f.end : f.name).focus(); return; }
      const bud = f.budget.value === '' ? null : Math.round(Number(f.budget.value) * 100) / 100, now = Date.now();
      const data = {kind:'campaign', name, boardId:f.board.value, platform:f.platform.value === '*' ? clean(f.platformOther.value).slice(0, 40) : f.platform.value,
        stage:f.stage.value || 'plan', start:isISO(start) ? start : '', end:isISO(end) ? end : '', budget:bud !== null && isFinite(bud) && bud >= 0 ? bud : null,
        link:normLink(f.link.value), notes:String(f.notes.value || '').replace(/\r\n?/g, '\n').trim().slice(0, 2000), archived:false,
        created:c ? (c.created || now) : now, updated:now, by:api.userName()};
      data.stage = stageNow(data, api.today());
      const next = {...data, id:c ? c.id : 'cmp_' + api.uid().replace(/[^A-Za-z0-9]/g, '')};
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
  el.addEventListener('keydown', e=>{ if ((e.key === 'Enter' || e.key === ' ') && e.target.classList.contains('cg-card')){ e.preventDefault(); open(e.target.dataset.id); } });
  // drag a card to another stage (mouse; on touch the card's menu does the same)
  el.addEventListener('dragstart', e=>{ const c = e.target.closest?.('.cg-card'); if (!c || !api.canEdit()) return; dragId = c.dataset.id; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', dragId); c.classList.add('dragging'); });
  el.addEventListener('dragend', ()=>{ dragId = null; el.querySelectorAll('.dragging,.over').forEach(x=>x.classList.remove('dragging', 'over')); });
  el.addEventListener('dragover', e=>{ const col = dragId && e.target.closest('.cg-col'); if (!col) return; e.preventDefault(); e.dataTransfer.dropEffect = 'move';
    if (!col.classList.contains('over')){ el.querySelectorAll('.cg-col.over').forEach(x=>x.classList.remove('over')); col.classList.add('over'); } });
  el.addEventListener('drop', e=>{ const col = dragId && e.target.closest('.cg-col'); if (!col) return; e.preventDefault(); const id = dragId, c = get(id); dragId = null; if (c && c.stage !== col.dataset.st) patch(id, {stage:col.dataset.st}); else api.rerender(); });

  return { setData, render, onClick, open, count:()=>all.length };
}
