import { createStore, isConfigured } from './store.js';
import { t, tn, getLang, setLang, locale, _dicts } from './i18n.js';
import { writeXlsx, readXlsx, isoToSerial, msToSerial, serialToIso, ST, MAX_CELL } from './xlsx.js';
import { writeTracker, TRACKER } from './tracker.js';

/* ================= Constants ================= */
const ALL = '*';
const CATS = ['todo','doing','wait','done'];
const PRIOS = ['highest','high','medium','low','lowest'];
const PRANK = Object.fromEntries(PRIOS.map((k,i)=>[k,i]));
const catName = c=>t('cat.'+c);
const prioName = p=>t('prio.'+p);
const PALETTE = [['sky','#4BADE8'],['blue','#1D7AFC'],['teal','#2898BD'],['green','#22A06B'],['lime','#63BA3C'],['yellow','#E2B203'],['orange','#F38A3F'],['red','#E5493A'],['magenta','#DA62AC'],['purple','#8F7EE7'],['grey','#8590A2']];
const PLATFORMS = [['Google Ads','#1D7AFC'],['Meta','#8F7EE7'],['Microsoft Ads','#2898BD'],['TikTok','#DA62AC'],['Taboola / Outbrain','#F38A3F'],['other','#8590A2']];
const EXTRA = [['sky','#4BADE8'],['green','#22A06B'],['lime','#63BA3C'],['yellow','#E2B203'],['red','#E5493A']];
const platLabel = n=>n === 'other' ? t('plat.other') : n;
function colorName(h){
  const p = PLATFORMS.find(x=>x[1]===h); if (p) return platLabel(p[0]);
  const c = PALETTE.find(x=>x[1]===h); return c ? t('color.'+c[0]) : h;
}
const GLYPHS = {
  check:'<path d="M4.6 8.3l2.2 2.2 4.6-4.8" stroke="#fff" stroke-width="1.8" fill="none" stroke-linecap="round" stroke-linejoin="round"/>',
  bookmark:'<path d="M5.5 4h5a.5.5 0 0 1 .5.5V12l-3-2.2L5 12V4.5a.5.5 0 0 1 .5-.5z" fill="#fff"/>',
  dot:'<circle cx="8" cy="8" r="3" fill="#fff"/>',
  ring:'<circle cx="8" cy="8" r="3.1" fill="none" stroke="#fff" stroke-width="1.8"/>',
  bolt:'<path d="M9 2.8L4.8 9h3l-.8 4.2L11.2 7h-3z" fill="#fff"/>',
  star:'<path d="M8 3.2l1.4 2.9 3.2.4-2.3 2.2.6 3.1L8 10.3l-2.9 1.5.6-3.1-2.3-2.2 3.2-.4z" fill="#fff"/>',
  flag:'<path d="M5.2 3.6v9" stroke="#fff" stroke-width="1.6" stroke-linecap="round"/><path d="M5.8 4h5.4l-1.4 2.2 1.4 2.3H5.8z" fill="#fff"/>',
  doc:'<path d="M5 3.5h4l2 2v7H5z" fill="#fff"/><path d="M6.6 8h2.8M6.6 10h2.8" stroke-width="1" stroke="currentColor"/>',
  eye:'<path d="M2.8 8S4.6 4.8 8 4.8 13.2 8 13.2 8 11.4 11.2 8 11.2 2.8 8 2.8 8z" fill="none" stroke="#fff" stroke-width="1.5"/><circle cx="8" cy="8" r="1.4" fill="#fff"/>',
  question:'<path d="M6.2 6.3a1.9 1.9 0 1 1 2.6 1.8c-.5.2-.8.6-.8 1.1v.4" fill="none" stroke="#fff" stroke-width="1.7" stroke-linecap="round"/><circle cx="8" cy="11.7" r="1" fill="#fff"/>',
  up:'<path d="M8 12V4.4M4.9 7.4 8 4.3l3.1 3.1" fill="none" stroke="#fff" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>',
  heart:'<path d="M8 12.2S3.6 9.6 3.6 6.6A2.2 2.2 0 0 1 8 5.8a2.2 2.2 0 0 1 4.4.8c0 3-4.4 5.6-4.4 5.6z" fill="#fff"/>'
};
/* Starting setups. Text follows the language chosen when the workspace is created. */
const COLS_PPC = ()=>[['todo','todo'],['doing','doing'],['approval','wait'],['fix','doing'],['live','doing'],['done','done'],['notdone','done']].map(([k,cat])=>({id:uid(), name:t('tpl.col.'+k), cat}));
const COLS_BASIC = ()=>[['todo','todo'],['doing','doing'],['done','done']].map(([k,cat])=>({id:uid(), name:t('tpl.col.'+k), cat}));
const TYPES_PPC = ()=>[['brief','#8F7EE7','flag'],['analysis','#2898BD','eye'],['mediaplan','#D97008','ring'],['launch','#22A06B','bolt'],['creative','#E774BB','heart'],['optim','#1D7AFC','up'],['report','#626F86','doc'],['tracking','#E5493A','dot']].map(([id,color,glyph])=>({id, name:t('tpl.type.'+id), color, glyph}));
const TYPES_BASIC = ()=>[['task','#4BADE8','check'],['story','#63BA3C','bookmark'],['bug','#E5493A','dot']].map(([id,color,glyph])=>({id, name:t('tpl.type.'+id), color, glyph}));
const PPC_CLIENTS = [['קניון עופר','OFR','#2898BD','kanyon-ofer'],['אלדן','ELD','#22A06B','eldan'],['משביר לצרכן','MSB','#44546F','mashbir'],['ליימן שליסל','LSH','#8F7EE7','leiman-schlissel']];

/* ================= Helpers ================= */
const $ = (s,r=document)=>r.querySelector(s), $$ = (s,r=document)=>[...r.querySelectorAll(s)];
const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
const isMobile = ()=>matchMedia('(max-width:760px)').matches;
const sideOverlay = ()=>matchMedia('(max-width:1024px)').matches;
const coarse = ()=>matchMedia('(pointer:coarse)').matches;
const rtl = ()=>document.documentElement.dir === 'rtl';
function uid(){ return Math.random().toString(36).slice(2,9)+Date.now().toString(36).slice(-4); }
const clone = o=>JSON.parse(JSON.stringify(o));
function stable(v){
  if (Array.isArray(v)) return '['+v.map(stable).join(',')+']';
  if (v && typeof v === 'object') return '{'+Object.keys(v).filter(k=>v[k]!==undefined).sort().map(k=>JSON.stringify(k)+':'+stable(v[k])).join(',')+'}';
  return JSON.stringify(v === undefined ? null : v);
}
const pad = n=>String(n).padStart(2,'0');
const iso = d=>`${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
const today = ()=>iso(new Date());
const addDays = n=>{const d=new Date(); d.setDate(d.getDate()+n); return iso(d);};
const isISO = s=>typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s);
const daysBetween = (a,b)=>Math.round((new Date(b+'T12:00') - new Date(a+'T12:00'))/864e5);
const esc = s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const lower = s=>String(s||'').trim().toLowerCase();
const isEmail = s=>/^[a-z0-9._%+-]+@[a-z0-9.-]+[.][a-z]{2,}$/.test(String(s).trim().toLowerCase()); // same pattern as firestore.rules
// toLocaleDateString builds a new Intl formatter on every call, which made big boards slow to
// render (most of a render with ~1000 tasks). Reuse formatters and remember formatted days.
const fmtCache = new Map(), dayCache = new Map();
function fmtr(kind, o){
  const k = locale() + '|' + kind;
  let f = fmtCache.get(k); if (!f){ f = new Intl.DateTimeFormat(locale(), o); fmtCache.set(k, f); }
  return f;
}
function fmtDay(d){
  const y = new Date().getFullYear(), k = locale() + '|' + y + '|' + d;
  let out = dayCache.get(k);
  if (out === undefined){
    const dt = new Date(d+'T00:00');
    out = isNaN(dt) ? String(d ?? '') : dt.getFullYear() !== y ? fmtr('dmy', {day:'numeric',month:'short',year:'numeric'}).format(dt) : fmtr('dm', {day:'numeric',month:'short'}).format(dt);
    if (dayCache.size > 5000) dayCache.clear();
    dayCache.set(k, out);
  }
  return out;
}
const fmtTime = ts=>fmtr('time', {day:'numeric',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit'}).format(new Date(ts));
function ago(ts){
  const s = (Date.now()-ts)/1000;
  if (s < 45) return t('ago.now');
  const m = Math.round(s/60); if (m < 60) return tn('ago.min', m);
  const h = Math.round(m/60); if (h < 24) return tn('ago.hour', h);
  const d = Math.round(h/24); if (d < 7) return d===1 ? t('ago.yesterday') : tn('ago.day', d);
  return fmtTime(ts);
}
function hl(s){
  const q = query.trim(); if (!q) return esc(s);
  const ls = s.toLowerCase(), lq = q.toLowerCase(); let out = '', i = 0, j;
  while ((j = ls.indexOf(lq,i)) > -1){ out += esc(s.slice(i,j))+'<mark>'+esc(s.slice(j,j+q.length))+'</mark>'; i = j+q.length; }
  return out+esc(s.slice(i));
}
const linkify = s=>s.replace(/(https?:\/\/[^\s<]+)/g,'<a href="$1" target="_blank" rel="noopener" dir="ltr">$1</a>');
const normLabel = s=>String(s).replace(/,/g,' ').replace(/\s+/g,' ').trim().slice(0,40);
const sameText = (a,b)=>normLabel(a).toLowerCase() === normLabel(b).toLowerCase();
function autosize(el){ if(!el) return; el.style.height='auto'; el.style.height=(el.scrollHeight+(el.offsetHeight-el.clientHeight))+'px'; }
const isHex = c=>typeof c === 'string' && /^#[0-9a-f]{6}$/i.test(c);
const isImg = s=>typeof s === 'string' && s.length < 200000 && /^data:image\/(png|jpeg|webp|gif);base64,[A-Za-z0-9+/]+={0,2}$/.test(s);
// Every id that ends up in HTML attributes must match this. Data from the database is checked on arrival.
const SAFE_ID = /^[A-Za-z0-9_-]{1,64}$/, REQ_ID = /^[A-Za-z0-9_@.%+-]{1,200}$/;
const sid = (v, fallback='')=>{ v = String(v ?? ''); return SAFE_ID.test(v) ? v : fallback; };
const REQ_STATUS = ['pending','approved','declined'];
const HEB_LAT = {'א':'A','ב':'B','ג':'G','ד':'D','ה':'H','ו':'V','ז':'Z','ח':'H','ט':'T','י':'Y','כ':'K','ך':'K','ל':'L','מ':'M','ם':'M','נ':'N','ן':'N','ס':'S','ע':'A','פ':'P','ף':'P','צ':'Z','ץ':'Z','ק':'K','ר':'R','ש':'S','ת':'T'};
function keyFromName(n){
  // Hebrew names get a key too (פנים → PNY), by spelling the letters in Latin
  const w = String(n).trim().replace(/[\u05D0-\u05EA]/g, ch=>HEB_LAT[ch] || '').toUpperCase().replace(/[^A-Z0-9\s]/g,'').split(/\s+/).filter(Boolean);
  if (!w.length) return '';
  return (w.length === 1 ? w[0].slice(0,3) : w.map(x=>x[0]).join('').slice(0,4)).replace(/^[0-9]+/,'');
}
function normUrl(v){
  v = String(v||'').trim(); if (!v) return null;
  if (!/^[a-z][a-z0-9+.-]*:/i.test(v)) v = 'https://' + v;
  try { const u = new URL(v); return /^https?:$/.test(u.protocol) && u.hostname.includes('.') ? u.href : null; } catch(e){ return null; }
}
function hostOf(u){ try { return new URL(u).hostname.replace(/^www\./,''); } catch(e){ return u; } }
const initialOf = n=>esc((String(n).trim()[0]||'?').toUpperCase());
function hue(s){ let h = 0; for (const c of String(s)) h = (h*31 + c.charCodeAt(0)) % 360; return h; }
// Profile photos. Your own photo comes straight from Google sign-in, so any Google image host is fine
// (this is what worked before). Photos stored on tasks by other people must be on Google's profile-photo
// hosts (lh3…): other googleusercontent hosts can serve content whose owner sees who loaded it.
function photoHost(s){ try { const u = new URL(s); return u.protocol === 'https:' ? u.hostname.toLowerCase() : ''; } catch(e){ return ''; } }
const isOwnPhoto = s=>typeof s === 'string' && s.length <= 2048 && /^[a-z0-9.-]+\.googleusercontent\.com$/.test(photoHost(s));
const isPhoto = s=>typeof s === 'string' && s.length <= 2048 && /^lh[0-9]{1,2}\.googleusercontent\.com$/.test(photoHost(s));
function avatar(p, cls=''){
  p = p || {};
  const n = p.name || p.email || '?', parts = String(n).trim().split(/\s+/);
  const ini = esc(((parts[0]||'?')[0] + (parts[1]?.[0]||'')).toUpperCase()), tip = esc(p.name||p.email||''), bg = `hsl(${hue(p.email||n)} 45% 45%)`;
  if (isPhoto(p.photo) || (user && p.photo && p.photo === user.photo && isOwnPhoto(p.photo))) return `<img class="avatar ${cls}" src="${esc(p.photo)}" alt="" referrerpolicy="no-referrer"${cls === 'xs' ? ' loading="lazy"' : ''} title="${tip}" data-ini="${ini}" data-bg="${bg}">`;
  return `<span class="avatar ${cls}" style="background:${bg}" title="${tip}">${ini}</span>`;
}
// a profile photo that fails to load (old link, offline) falls back to initials
document.addEventListener('error', e=>{
  const im = e.target;
  if (!(im instanceof HTMLImageElement) || !im.classList.contains('avatar')) return;
  const sp = document.createElement('span'); sp.className = im.className; sp.title = im.title; sp.style.background = im.dataset.bg || ''; sp.textContent = im.dataset.ini || '?';
  im.replaceWith(sp);
}, true);

/* ================= Icons ================= */
const ic = d=>`<svg class="i" viewBox="0 0 24 24" aria-hidden="true">${d}</svg>`;
const I = {
  search:ic('<circle cx="11" cy="11" r="6.5"/><path d="M20 20l-4.2-4.2"/>'),
  plus:ic('<path d="M12 5v14M5 12h14"/>'),
  x:ic('<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>'),
  more:ic('<circle cx="5.5" cy="12" r="1.7" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.7" fill="currentColor" stroke="none"/><circle cx="18.5" cy="12" r="1.7" fill="currentColor" stroke="none"/>'),
  check:ic('<path d="M5 12.5l4.5 4.5L19 7.5"/>'),
  cal:ic('<rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4"/>'),
  repeat:ic('<path d="M17 3l3 3-3 3"/><path d="M4 11V9a3 3 0 0 1 3-3h13"/><path d="M7 21l-3-3 3-3"/><path d="M20 13v2a3 3 0 0 1-3 3H4"/>'),
  chart:ic('<path d="M4 20V10M10 20V4M16 20v-7M21 20H3"/>'),
  sliders:ic('<path d="M4 7h9M17 7h3M4 17h3M11 17h9"/><circle cx="15" cy="7" r="2"/><circle cx="9" cy="17" r="2"/>'),
  eyeOff:ic('<path d="M3 3l18 18M10.6 6.1A8.8 8.8 0 0 1 12 6c6 0 9.5 6 9.5 6a16 16 0 0 1-2.9 3.5M6.6 6.7C4 8.4 2.5 12 2.5 12S6 18 12 18c1.7 0 3.2-.4 4.4-1.1M9.9 9.9a3 3 0 0 0 4.2 4.2"/>'),
  sub:ic('<rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/><path d="M7.5 11v4.5A1.5 1.5 0 0 0 9 17h4"/>'),
  comment:ic('<path d="M5 5h14a1 1 0 0 1 1 1v10a1 1 0 0 1-1 1h-8l-4 3v-3H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1z"/>'),
  menu:ic('<path d="M4 7h16M4 12h16M4 17h16"/>'),
  board:ic('<rect x="4" y="4" width="4.5" height="16" rx="1"/><rect x="9.75" y="4" width="4.5" height="11" rx="1"/><rect x="15.5" y="4" width="4.5" height="7" rx="1"/>'),
  list:ic('<path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01"/>'),
  funnel:ic('<path d="M4 5h16l-6.2 7.4V19l-3.6-1.8v-4.8z"/>'),
  table:ic('<rect x="3.5" y="4.5" width="17" height="15" rx="2"/><path d="M3.5 9.5h17M3.5 14.5h17M9.5 9.5v10"/>'),
  down:ic('<path d="M12 4v11M7 10l5 5 5-5M5 20h14"/>'),
  up:ic('<path d="M12 15V4M7 9l5-5 5 5M5 20h14"/>'),
  trash:ic('<path d="M5 7h14M10 11v6M14 11v6M6.5 7l.8 12a2 2 0 0 0 2 1.9h5.4a2 2 0 0 0 2-1.9l.8-12M9.5 7V4.5h5V7"/>'),
  sun:ic('<circle cx="12" cy="12" r="4"/><path d="M12 2.5v2M12 19.5v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2.5 12h2M19.5 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/>'),
  moon:ic('<path d="M19.5 14.5A7.5 7.5 0 0 1 9.5 4.5a7.5 7.5 0 1 0 10 10z"/>'),
  palette:ic('<path d="M12 3a9 9 0 1 0 0 18c1.1 0 1.8-.8 1.8-1.7 0-.5-.2-.9-.5-1.2-.3-.3-.5-.7-.5-1.2 0-1 .8-1.7 1.8-1.7H17a4 4 0 0 0 4-4C21 6.6 17 3 12 3z"/><circle cx="7.5" cy="11" r="1.2"/><circle cx="10" cy="7.2" r="1.2"/><circle cx="14.5" cy="7.2" r="1.2"/>'),
  kbd:ic('<rect x="3" y="6" width="18" height="12" rx="2"/><path d="M7 10h.01M11 10h.01M15 10h.01M7.5 14h9"/>'),
  chev:ic('<path d="M7 10l5 5 5-5"/>'),
  copy:ic('<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5.5A1.5 1.5 0 0 0 14.5 4h-9A1.5 1.5 0 0 0 4 5.5v9A1.5 1.5 0 0 0 5.5 16H8"/>'),
  open:ic('<path d="M14 4h6v6M20 4l-9 9M18 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V7a1 1 0 0 1 1-1h5"/>'),
  clock:ic('<circle cx="12" cy="12" r="8"/><path d="M12 8v4l2.5 2"/>'),
  arrow:ic('<path d="M12 19V5M6 11l6-6 6 6"/>'),
  aup:ic('<path d="M7 14l5-5 5 5"/>'),
  adown:ic('<path d="M7 10l5 5 5-5"/>'),
  gear:ic('<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/>'),
  shapes:ic('<rect x="4" y="4" width="7" height="7" rx="1.5"/><circle cx="16.5" cy="7.5" r="3.5"/><path d="M7.5 13.5l4 6.5h-8z"/><rect x="13" y="13" width="7" height="7" rx="3.5"/>'),
  archive:ic('<rect x="3.5" y="4" width="17" height="4.5" rx="1"/><path d="M5 8.5V19a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1V8.5M10 12.5h4"/>'),
  restore:ic('<path d="M4 12a8 8 0 1 0 2.4-5.7M4 4v4.5h4.5"/>'),
  link:ic('<path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/>'),
  grid:ic('<rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/>'),
  users:ic('<circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M18 14.5a6.5 6.5 0 0 1 3.5 5.5"/>'),
  bell:ic('<path d="M6 16V11a6 6 0 1 1 12 0v5l1.5 2h-15zM10 20.5a2 2 0 0 0 4 0"/>'),
  mail:ic('<rect x="3" y="5" width="18" height="14" rx="2"/><path d="M3.5 6.5l8.5 6.5 8.5-6.5"/>'),
  globe:ic('<circle cx="12" cy="12" r="8.5"/><path d="M3.5 12h17M12 3.5c2.5 2.6 3.7 5.4 3.7 8.5s-1.2 5.9-3.7 8.5c-2.5-2.6-3.7-5.4-3.7-8.5s1.2-5.9 3.7-8.5z"/>'),
  logout:ic('<path d="M14 4h4a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-4M10 16l-4-4 4-4M6 12h10"/>'),
  image:ic('<rect x="3.5" y="4.5" width="17" height="15" rx="2"/><circle cx="9" cy="10" r="1.8"/><path d="M20.5 16l-5-5-8.5 8.5"/>'),
  eye:ic('<path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12z"/><circle cx="12" cy="12" r="3"/>'),
  key:ic('<circle cx="8" cy="15" r="4"/><path d="M11 12l8-8M16 7l2 2M14 9l2 2"/>'),
  tick:'<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 8.5l3 3 6-6.5"/></svg>'
};
const APP_MARK = '<svg viewBox="0 0 16 16" aria-hidden="true"><rect x="2" y="2.5" width="3" height="11" rx="1" fill="#fff"/><rect x="6.5" y="2.5" width="3" height="7.5" rx="1" fill="#fff" opacity=".85"/><rect x="11" y="2.5" width="3" height="4.5" rx="1" fill="#fff" opacity=".7"/></svg>';
const swTick = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.5 8.5l3 3 6-6.5"/></svg>';
function typeSVG(ty){
  const name = esc(ty?.name || ''), color = ty?.color || '#4BADE8', g = GLYPHS[ty?.glyph] || GLYPHS.check;
  return `<svg class="t-ico" viewBox="0 0 16 16" role="img" aria-label="${name}" style="color:${color}"><title>${name}</title><rect width="16" height="16" rx="3" fill="${color}"/>${g}</svg>`;
}
const typeIcon = id=>typeSVG(typeOf(id));
function prioIcon(p){
  const up='#E2483D', mid='#E2B203', dn='#1D7AFC';
  const paths = {highest:[up,'M4 8l4-3 4 3M4 12l4-3 4 3'], high:[up,'M4 10l4-3.5 4 3.5'], medium:[mid,'M3.5 6.5h9M3.5 9.5h9'],
    low:[dn,'M4 6l4 3.5 4-3.5'], lowest:[dn,'M4 4l4 3 4-3M4 8l4 3 4-3']}[p] || [mid,'M3.5 6.5h9M3.5 9.5h9'];
  const n = esc(t('prio.label', {p:prioName(p||'medium')}));
  return `<svg class="p-ico" viewBox="0 0 16 16" role="img" aria-label="${n}"><title>${n}</title><path d="${paths[1]}" fill="none" stroke="${paths[0]}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
}
const lozC = c=>`<span class="loz ${c.cat}" data-bidi>${esc(c.name)}</span>`;
function bAv(b, cls=''){
  if (!b) return '';
  if (isImg(b.logo)) return `<span class="b-av logo ${cls}" title="${esc(b.name)}"><img src="${esc(b.logo)}" alt=""></span>`;
  return `<span class="b-av ${cls}" style="--bc:${b.color}" title="${esc(b.name)}">${initialOf(b.name)}</span>`;
}
function swatches(act, sel, opts={}){
  const list = opts.list || PALETTE;
  return `<div class="sw-row" role="group" aria-label="${esc(opts.label||t('f.color'))}">${opts.none ? `<button type="button" class="sw none${!sel?' on':''}${opts.sm?' sm':''}" data-act="${act}" data-c="" aria-label="${esc(t('color.none'))}" aria-pressed="${!sel}" title="${esc(t('color.none'))}"></button>` : ''}${
    list.map(([n,h])=>`<button type="button" class="sw${sel===h?' on':''}${opts.sm?' sm':''}" style="--sw:${h}" data-act="${act}" data-c="${h}" aria-label="${esc(t('color.'+n))}" aria-pressed="${sel===h}" title="${esc(t('color.'+n))}">${swTick}</button>`).join('')}</div>`;
}
function platformPicker(act, sel){
  return `<div class="plat-wrap"><div class="plat" role="group" aria-label="${esc(t('f.platform'))}">
      <button type="button" class="plat-b none${!sel?' on':''}" data-act="${act}" data-c="" aria-pressed="${!sel}">${esc(t('plat.none'))}</button>
      ${PLATFORMS.map(([n,h])=>`<button type="button" class="plat-b${sel===h?' on':''}" style="--sw:${h}" data-act="${act}" data-c="${h}" aria-pressed="${sel===h}"><i></i><span data-bidi>${esc(platLabel(n))}</span></button>`).join('')}
    </div><div class="plat-more"><span>${esc(t('plat.more'))}</span>${swatches(act, sel, {list:EXTRA, sm:true, label:t('plat.more')})}</div></div>`;
}
function markPicked(btn){
  const wrap = btn.closest('.plat-wrap') || btn.parentElement;
  $$('[data-c]', wrap).forEach(s=>{ const on = s.dataset.c === btn.dataset.c; s.classList.toggle('on', on); s.setAttribute('aria-pressed', on); });
}

/* ================= Preferences (this device) ================= */
const PREF_KEY = 'workbench.prefs';
const P = (()=>{ try { return {view:'board', theme:null, side:true, active:{}, ...JSON.parse(localStorage.getItem(PREF_KEY)||'{}')}; } catch(e){ return {view:'board', theme:null, side:true, active:{}}; } })();
const savePrefs = ()=>{ try { localStorage.setItem(PREF_KEY, JSON.stringify(P)); } catch(e){} };

/* ================= Session state ================= */
let store = null, user = null, wsList = null, wsId = null, ws = null;
let incoming = [], outgoing = [], mayCreate = true;
let unsubWs = null, unsubData = null, unsubIn = null, unsubOut = null;
let S = {boards:[], issues:[], types:[], archive:null};
// Monthly "archive what's done" prompt: shown from `day` of each month until handled (stored in the workspace settings)
const normArchive = a=>({on:a?.on !== false, day:Math.min(28, Math.max(1, parseInt(a?.day, 10) || 1)), handled:/^\d{4}-\d{2}$/.test(a?.handled || '') ? a.handled : ''});
const ym = (d=new Date())=>`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
let loaded = {boards:false, issues:false, meta:false};
const synced = {boards:new Map(), issues:new Map(), meta:null};
let screen = 'loading'; // loading | signin | welcome | app
let awaitOpen = null;

const role = ()=>!ws ? null : ws.ownerUid === user?.uid ? 'owner' : (ws.roles||{})[user?.email] || 'viewer';
const canEdit = ()=>role() === 'owner' || role() === 'editor';
const isOwner = ()=>role() === 'owner';
const me = ()=>({uid:user.uid, name:user.name, email:user.email, ...(isPhoto(user.photo) ? {photo:user.photo} : {})});

const get = id=>S.issues.find(i=>i.id===id);
const activeKey = ()=>P.active[wsId];
const isAll = ()=>activeKey() === ALL;
const boardById = id=>S.boards.find(b=>b.id===id);
const curBoard = ()=>boardById(activeKey()) || boardById(P.lastBoard?.[wsId]) || S.boards[0] || null;
const boardOf = i=>boardById(i.boardId) || S.boards[0];
const colOf = i=>{ const b = boardOf(i); return b.columns.find(c=>c.id===i.status) || b.columns[0]; };
const colById = (b,id)=>b.columns.find(c=>c.id===id);
const isDone = i=>colOf(i).cat === 'done';
const typeOf = id=>S.types.find(x=>x.id===id) || S.types[0] || {id:'', name:'', color:'#4BADE8', glyph:'check'};
const key = i=>`${boardOf(i).key}-${i.num}`;
const boardIssues = (b=curBoard())=>b ? S.issues.filter(i=>i.boardId===b.id && !i.archived) : [];
const scopeIssues = ()=>{ const b = curBoard(); return S.issues.filter(i=>!i.archived && (isAll() || (b && i.boardId===b.id))); };

/* ---------- normalizing documents ---------- */
function normBoard(id, d){
  return {id, name:String(d.name||'Board'), key:String(d.key||'B').toUpperCase().replace(/[^A-Z0-9]/g,'').slice(0,10)||'B',
    color:isHex(d.color)?d.color:'#1D7AFC', logo:isImg(d.logo)?d.logo:null, seq:Number(d.seq)||0, order:Number.isFinite(d.order)?d.order:0,
    ...(d.alias && typeof d.alias === 'string' ? {alias:d.alias.slice(0,100)} : {}), ...(d.tc ? {tc:Number(d.tc) || 1} : {}),
    columns:(Array.isArray(d.columns)&&d.columns.length?d.columns:COLS_BASIC()).map(c=>({id:sid(c?.id, 'c'+uid()), name:String(c.name||''), cat:CATS.includes(c.cat)?c.cat:'todo'}))};
}
const normBy = b=>b && typeof b === 'object' ? {uid:String(b.uid||''), name:String(b.name||''), email:String(b.email||''), ...(isPhoto(b.photo) ? {photo:b.photo} : {})} : null;
/* ---------- recurring tasks ----------
   A recurring task carries its rule (every day / week / 2 weeks / month) and the date the next copy is due.
   When that date arrives (checked whenever someone who can edit has the site open) the site makes the copy
   in the board's first column, and the rule moves on to the new copy. Copies get a fixed id per date, so
   two people opening the site at the same time can't create it twice. Missed periods aren't backfilled:
   after a long break only the latest copy is made. */
const RECUR = ['day','week','2week','month'];
function normRecur(r){
  if (!r || typeof r !== 'object' || !RECUR.includes(r.every)) return null;
  const o = {every:r.every, next:isISO(r.next) ? r.next : today()};
  if (r.every === 'week' || r.every === '2week') o.wd = Math.min(6, Math.max(0, Number(r.wd) || 0));
  if (r.every === 'month') o.md = Math.min(31, Math.max(1, Number(r.md) || 1));
  return o;
}
const isoAdd = (d, n)=>{ const x = new Date(d+'T12:00'); x.setDate(x.getDate()+n); return iso(x); };
function monthDay(y, m, md){ const last = new Date(y, m+1, 0, 12).getDate(); return iso(new Date(y, m, Math.min(md, last), 12)); }
// the first date strictly after `after` that matches the rule
function recurFirst(r, after){
  if (r.every === 'day') return isoAdd(after, 1);
  if (r.every === 'week' || r.every === '2week'){ const d = new Date(after+'T12:00'); let k = (r.wd - d.getDay() + 7) % 7; if (k === 0) k = 7; return isoAdd(after, k); }
  const d = new Date(after+'T12:00'); let c = monthDay(d.getFullYear(), d.getMonth(), r.md);
  if (c <= after) c = monthDay(d.getFullYear(), d.getMonth()+1, r.md);
  return c;
}
// the occurrence after `d` (keeps two-week rules on their rhythm)
function recurAfter(r, d){
  if (r.every === 'day') return isoAdd(d, 1);
  if (r.every === 'week') return isoAdd(d, 7);
  if (r.every === '2week') return isoAdd(d, 14);
  const x = new Date(d+'T12:00'); return monthDay(x.getFullYear(), x.getMonth()+1, r.md);
}
function recurText(r){
  if (!r) return t('rc.none');
  if (r.every === 'day') return t('rc.day');
  if (r.every === 'month') return t('rc.monthOn', {d:r.md});
  const wd = fmtr('wdl1', {weekday:'long'}).format(new Date(2026, 10, 1 + r.wd, 12)); // 1 Nov 2026 is a Sunday
  return t(r.every === 'week' ? 'rc.weekOn' : 'rc.2weekOn', {wd});
}
function runRecurring(){
  if (!canEdit() || !loaded.boards || !loaded.issues || !loaded.meta) return 0;
  const tday = today(); let made = 0;
  for (const i of S.issues.slice()){
    const r = i.recur; if (!r || r.next > tday) continue;
    const b = boardOf(i); if (!b){ continue; }
    let d = r.next, guard = 0;
    while (recurAfter(r, d) <= tday && guard++ < 1000) d = recurAfter(r, d);
    const series = i.recurOf || i.id, id = ('r-' + series + '-' + d).slice(0, 140);
    let copy = get(id);
    if (!copy){
      const span = i.start && i.due ? daysBetween(i.start, i.due) : null;
      b.seq = Math.max(b.seq || 0, ...S.issues.filter(x=>x.boardId === b.id).map(x=>x.num)) + 1;
      const now = Date.now();
      copy = normIssue(id, {boardId:b.id, num:b.seq, type:i.type, title:i.title, status:b.columns[0].id, priority:i.priority, start:d,
        due:span != null && span >= 0 ? isoAdd(d, span) : '', desc:i.desc, labels:i.labels.slice(), color:i.color, points:i.points,
        links:i.links.map(l=>({...l, id:uid()})), subtasks:i.subtasks.map(x=>({id:uid(), text:x.text, done:false})),
        created:now, updated:now, creator:i.creator || null, recurOf:series,
        history:[{id:uid(), at:now, text:t('rc.made', {rule:recurText(r)}), by:''}]});
      copy.rank = undefined; S.issues.push(copy); made++;
    }
    copy.recur = {...r, next:recurAfter(r, d)}; copy.recurOf = series;
    delete i.recur; i.recurOf = series; i.updated = Date.now();
  }
  if (made){ save(); render(); toast(tn('rc.toast', made)); }
  return made;
}
setInterval(()=>{ if (screen === 'app') runRecurring(); }, 30 * 60000);
document.addEventListener('visibilitychange', ()=>{ if (document.visibilityState === 'visible' && screen === 'app') runRecurring(); });
function recurField(i, edit){
  const r = i.recur;
  if (!edit) return `<span class="ro-val${r ? '' : ' t-muted'}">${esc(recurText(r))}</span>`;
  const wdNames = Array.from({length:7}, (_,k)=>fmtr('wdl1', {weekday:'long'}).format(new Date(2026, 10, 1 + k, 12)));
  return `<div class="rc-edit"><select class="ctl" data-f="recur" aria-label="${esc(t('f.recur'))}">${['', ...RECUR].map(v=>`<option value="${v}"${(r?.every || '') === v ? ' selected' : ''}>${esc(t('rc.opt.' + (v || 'none')))}</option>`).join('')}</select>
    ${r && (r.every === 'week' || r.every === '2week') ? `<select class="ctl" data-f="recur-wd" aria-label="${esc(t('rc.wd'))}">${wdNames.map((n,k)=>`<option value="${k}"${r.wd === k ? ' selected' : ''}>${esc(n)}</option>`).join('')}</select>` : ''}
    ${r && r.every === 'month' ? `<select class="ctl" data-f="recur-md" aria-label="${esc(t('rc.md'))}">${Array.from({length:31}, (_,k)=>`<option value="${k+1}"${r.md === k+1 ? ' selected' : ''}>${esc(t('rc.dayN', {n:k+1}))}</option>`).join('')}</select>` : ''}
    ${r ? `<span class="rc-next">${I.repeat}${esc(t('rc.next', {date:fmtDay(r.next)}))}</span>` : ''}</div>`;
}
// new rule from the task: the weekday / day of month comes from the task's start date
function recurFrom(i, every){
  if (!RECUR.includes(every)) return null;
  const base = isISO(i.start) ? i.start : today(), dt = new Date(base+'T12:00');
  const r = {every, wd:dt.getDay(), md:dt.getDate()};
  r.next = recurFirst(r, base > today() ? base : today());
  return normRecur(r);
}

function normIssue(id, d){
  const now = Date.now();
  const o = {id, boardId:sid(d.boardId), num:Number(d.num)||0, type:sid(d.type), title:String(d.title??''), status:sid(d.status),
    priority:PRIOS.includes(d.priority)?d.priority:'medium', due:isISO(d.due)?d.due:'', start:isISO(d.start)?d.start:iso(new Date(Number(d.created)||now)),
    doneAt:Number(d.doneAt)||null, points:(d.points===''||d.points==null||isNaN(Number(d.points)))?null:Number(d.points),
    labels:Array.isArray(d.labels)?d.labels.map(l=>normLabel(l)).filter(Boolean):[], desc:String(d.desc??''), color:isHex(d.color)?d.color:null,
    links:(Array.isArray(d.links)?d.links:[]).filter(l=>l&&normUrl(l.url)).map(l=>({id:sid(l.id, uid()), url:normUrl(l.url), title:String(l.title||'')})),
    subtasks:Array.isArray(d.subtasks)?d.subtasks.filter(Boolean).map(s=>({id:sid(s.id, uid()), text:String(s.text||''), done:!!s.done})):[],
    comments:Array.isArray(d.comments)?d.comments.filter(Boolean).map(c=>({id:sid(c.id, uid()), text:String(c.text||''), at:Number(c.at)||now, edited:!!c.edited, by:normBy(c.by)})):[],
    history:Array.isArray(d.history)?d.history.filter(Boolean).slice(0,200).map(h=>({id:sid(h.id, uid()), at:Number(h.at)||now, text:String(h.text||''), by:h.by?String(h.by):''})):[],
    archived:!!d.archived, archivedAt:d.archived?(Number(d.archivedAt)||now):null, created:Number(d.created)||now, updated:Number(d.updated)||now,
    rank:Number.isFinite(d.rank)?d.rank:0};
  if (d.creator && typeof d.creator === 'object') o.creator = normBy(d.creator); // older tasks have none
  const rc = normRecur(d.recur); if (rc) o.recur = rc;
  if (typeof d.recurOf === 'string' && d.recurOf) o.recurOf = d.recurOf.slice(0, 120);
  return o;
}
const serBoard = b=>{ const {id, ...rest} = b; return rest; };
const serIssue = i=>{ const {id, ...rest} = i; return rest; };
const serMeta = ()=>({types:S.types.map(({id,name,color,glyph})=>({id,name,color,glyph})), archive:{...(S.archive || normArchive())}});

/* ---------- merging snapshots ---------- */
// base = the last version of each doc we know the server has. Used for three-way merges,
// so two people editing different parts of the same issue don't overwrite each other.
const base = {boards:new Map(), issues:new Map()};
const ID_ARRAYS = new Set(['comments','links','subtasks','history']);
function replaceInPlace(target, src){ for (const k of Object.keys(target)) if (!(k in src)) delete target[k]; Object.assign(target, src); }
function assignMissing(arr, field){
  for (let k = 0; k < arr.length; k++){
    if (Number.isFinite(arr[k][field])) continue;
    let prev = null, next = null;
    for (let m = k-1; m >= 0; m--) if (Number.isFinite(arr[m][field])){ prev = arr[m][field]; break; }
    for (let m = k+1; m < arr.length; m++) if (Number.isFinite(arr[m][field])){ next = arr[m][field]; break; }
    arr[k][field] = prev === null && next === null ? 0 : prev === null ? next - 1000 : next === null ? prev + 1000 : (prev + next) / 2;
  }
}
const byRank = f=>(a,b)=>(a[f] - b[f]) || (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
function merge3(local, old, inc){
  const out = {...local};
  for (const k of new Set([...Object.keys(inc), ...Object.keys(local)])){
    if (k === 'id') continue;
    const lv = stable(local[k]), bv = stable(old?.[k]), iv = stable(inc[k]);
    if (lv === bv) out[k] = clone(inc[k] ?? null);             // only changed there: take theirs
    else if (iv === bv || iv === lv) continue;                  // only changed here: keep mine
    else if (ID_ARRAYS.has(k) && Array.isArray(local[k]) && Array.isArray(inc[k])){
      const oldIds = new Set((old?.[k]||[]).map(x=>x.id)), mine = new Set(local[k].map(x=>x.id)), theirs = new Set(inc[k].map(x=>x.id));
      const added = inc[k].filter(x=>!mine.has(x.id) && !oldIds.has(x.id));                 // new there
      const kept = local[k].filter(x=>!(oldIds.has(x.id) && !theirs.has(x.id)));           // not deleted there
      out[k] = k === 'comments' || k === 'history' ? [...added, ...kept].sort((a,b)=>(b.at||0)-(a.at||0)) : [...kept, ...added];
    }
    // both changed the same field: keep mine (last write wins)
  }
  return out;
}
function merge(kind, docs){
  const norm = kind === 'boards' ? normBoard : normIssue, ser = kind === 'boards' ? serBoard : serIssue, f = kind === 'boards' ? 'order' : 'rank';
  assignMissing(S[kind], f);
  const local = new Map(S[kind].map(x=>[x.id,x])), seen = new Set(), out = [];
  let openChanged = false;
  for (const {id, data} of docs){
    if (!SAFE_ID.test(id) || !data || typeof data !== 'object') continue;
    seen.add(id);
    const inc = norm(id, data), incJson = stable(ser(inc)), loc = local.get(id);
    if (!loc){
      if (synced[kind].has(id)) continue;                     // deleted here; the delete is about to be written
      synced[kind].set(id, incJson); base[kind].set(id, clone(inc)); out.push(inc);
      if (id === openId) openChanged = true;
      continue;
    }
    if (incJson !== synced[kind].get(id)){                    // changed on the server since we last saw it
      const dirty = stable(ser(loc)) !== synced[kind].get(id);
      replaceInPlace(loc, dirty ? merge3(loc, base[kind].get(id), inc) : inc);
      if (id === openId) openChanged = true;
    }
    synced[kind].set(id, incJson); base[kind].set(id, clone(inc));
    out.push(loc);
  }
  for (const loc of S[kind]) if (!seen.has(loc.id)){
    if (synced[kind].has(loc.id)){ synced[kind].delete(loc.id); base[kind].delete(loc.id); if (loc.id === openId) openChanged = true; } // deleted elsewhere
    else out.push(loc);                                          // created here, not written yet
  }
  out.sort(byRank(f));
  S[kind] = out;
  if (out.some(o=>stable(ser(o)) !== synced[kind].get(o.id))) save();
  return openChanged;
}
function onBoards(docs){ merge('boards', docs); loaded.boards = true; dataArrived(); if (screen === 'app') scheduleRender(); }
function onIssues(docs){
  const ch = merge('issues', docs); loaded.issues = true; dataArrived();
  if (ch && modalKind === 'issue'){
    if (!get(openId)){ closeModal(true); toast(t('toast.issueGoneElsewhere')); }
    else if (!$('.dialog :focus:is(input,textarea)', layer)) refreshIssue();
  }
}
function normTypes(m){
  const types = Array.isArray(m?.types) && m.types.length ? m.types : TYPES_BASIC();
  return types.filter(Boolean).map(x=>({id:sid(x.id, 't'+uid()), name:String(x.name||''), color:isHex(x.color)?x.color:'#4BADE8', glyph:GLYPHS[x.glyph]?x.glyph:'check'}));
}
function cleanReqs(list){
  return (list||[]).filter(r=>r && REQ_ID.test(String(r.id)) && REQ_STATUS.includes(r.status)).map(r=>({
    id:String(r.id), fromUid:String(r.fromUid||''), fromEmail:String(r.fromEmail||''), fromName:String(r.fromName||''), toEmail:String(r.toEmail||''),
    status:r.status, role:r.role === 'viewer' ? 'viewer' : r.role === 'editor' ? 'editor' : null, wsId:sid(r.wsId) || null, wsName:String(r.wsName||''),
    created:Number(r.created)||0, decidedAt:Number(r.decidedAt)||0}));
}
function cleanWs(list){
  return (list||[]).filter(w=>w && SAFE_ID.test(String(w.id))).map(w=>{
    const roles = {}; for (const [e, r] of Object.entries(w.roles && typeof w.roles === 'object' ? w.roles : {})) if (r === 'editor' || r === 'viewer') roles[String(e)] = r;
    return {id:String(w.id), name:String(w.name||''), logo:isImg(w.logo) ? w.logo : null, ownerUid:String(w.ownerUid||''), ownerEmail:String(w.ownerEmail||''),
      ownerName:String(w.ownerName||''), roles, memberEmails:Array.isArray(w.memberEmails) ? w.memberEmails.map(String) : [], created:Number(w.created)||0, lang:w.lang === 'en' ? 'en' : 'he'};
  });
}
function onMeta(m){
  const types = normTypes(m), archive = normArchive(m?.archive), incJson = stable({types, archive});
  const dirty = synced.meta !== null && stable(serMeta()) !== synced.meta;
  if (!dirty){ S.types = types; S.archive = archive; }
  synced.meta = incJson;
  loaded.meta = true; dataArrived(); if (screen === 'app') scheduleRender();
}
function dataArrived(){ if (loaded.boards && loaded.issues && loaded.meta){ addTrackerStatuses(); renameClients(); showApp(); scheduleRender(); setTimeout(runRecurring, 0); } }
// once per board: boards made before these existed get the two statuses of the task-tracker sheet
// ("דורש תיקונים", "לא בוצע"). The board is marked (tc), so if it later removes them they stay removed.
// Kept on the boards, not in the settings doc, so it can never overwrite settings that haven't arrived yet.
function ensureTrackerCols(b){
  const has = w=>b.columns.some(c=>loose(c.name) === loose(w) || loose(TRK_COL_OUT[norm(c.name)] || '') === loose(w));
  let added = false;
  // her sheet's "ממתין לאישור": boards made from the basic set (to do / in progress / done) had no such column
  if (!b.columns.some(c=>c.cat === 'wait')){
    const fi = b.columns.findIndex(x=>loose(x.name) === loose('דורש תיקונים') || norm(x.name) === 'needs fixes'), di = b.columns.findIndex(x=>x.cat === 'done');
    b.columns.splice(fi >= 0 ? fi : di >= 0 ? di : b.columns.length, 0, {id:'c'+uid(), name:t('tpl.col.approval'), cat:'wait'}); added = true;
  }
  if (!has('דורש תיקונים')){
    const c = {id:'c'+uid(), name:t('tpl.col.fix'), cat:'doing'}, wi = b.columns.findIndex(x=>x.cat === 'wait'), di = b.columns.findIndex(x=>x.cat === 'done');
    b.columns.splice(wi >= 0 ? wi + 1 : di >= 0 ? di : b.columns.length, 0, c); added = true;
  }
  if (!has('לא בוצע')){ b.columns.push({id:'c'+uid(), name:t('tpl.col.notdone'), cat:'done'}); added = true; }
  return added;
}
const TC = 2; // version of the one-time column update (2: also adds a waiting-for-approval column)
/* ---------- renamed clients ----------
   When a client changes, every existing board of the old client is updated once, for everyone who opens the
   site and can edit: new name, code (if it's still the template one and free), colour and logo. The board
   keeps its tasks; the old names keep matching when a tracker sheet is imported. */
const CLIENT_RENAMES = [{from:'בית בלב', to:'משביר לצרכן', oldKey:'BTB', key:'MSB', oldColor:'#F38A3F', color:'#44546F', logo:'mashbir', also:['המשביר לצרכן','משביר','המשביר']}];
let renaming = false;
async function renameClients(){
  if (renaming || !canEdit()) return;
  const todo = CLIENT_RENAMES.map(r=>({r, bs:S.boards.filter(b=>loose(b.name) === loose(r.from))})).filter(x=>x.bs.length);
  if (!todo.length) return;
  renaming = true;
  try {
    for (const {r, bs} of todo){
      // all or nothing: without the new logo (offline) the board waits for the next visit
      let logo = null;
      try { const res = await fetch(`assets/clients/${r.logo}.png`); if (res.ok) logo = await readLogo(await res.blob()); } catch {}
      if (!logo) continue;
      bs.forEach(b=>{
        if (loose(b.name) !== loose(r.from)) return;
        b.name = r.to; b.logo = logo;
        if (b.key === r.oldKey && !S.boards.some(x=>x !== b && x.key === r.key)) b.key = r.key;
        if (String(b.color).toUpperCase() === r.oldColor) b.color = r.color;
      });
    }
    save(); scheduleRender();
  } finally { renaming = false; }
}
// a client name as written in a sheet → the name used on the site now
const isOldClient = w=>CLIENT_RENAMES.some(r=>loose(r.from) === loose(w));
function clientNow(w){ const l = loose(w); const r = CLIENT_RENAMES.find(r=>loose(r.from) === l || r.also.some(a=>loose(a) === l)); return r ? r.to : w; }

function addTrackerStatuses(){
  if (!canEdit()) return;
  const todo = S.boards.filter(b=>(b.tc || 0) < TC); if (!todo.length) return;
  todo.forEach(b=>{ ensureTrackerCols(b); b.tc = TC; });
  save();
}

/* ---------- writing ---------- */
let flushTimer = 0;
function save(){ clearTimeout(flushTimer); flushTimer = setTimeout(flush, 350); }
function fixRanks(arr, field, step=1000){
  let prev = -Infinity, need = false;
  for (let k = 0; k < arr.length; k++){
    const it = arr[k];
    if (typeof it[field] === 'number' && it[field] > prev){ prev = it[field]; continue; }
    let next = Infinity;
    for (let m = k+1; m < arr.length; m++){ const r = arr[m][field]; if (typeof r === 'number' && r > prev){ next = r; break; } }
    it[field] = next === Infinity ? (prev === -Infinity ? 0 : prev + step) : (prev === -Infinity ? next - step : (prev + next)/2);
    if (next !== Infinity && prev !== -Infinity && next - prev < 1e-6) need = true;
    prev = it[field];
  }
  if (need) arr.forEach((it,k)=>{ it[field] = k*step; });
}
function flush(){
  clearTimeout(flushTimer);
  if (!wsId || !canEdit() || !store) return;
  if (!loaded.boards || !loaded.issues || !loaded.meta){ save(); return; } // never write before everything has loaded
  assignMissing(S.issues, 'rank'); assignMissing(S.boards, 'order');
  const ops = [];
  for (const [kind, ser] of [['boards', serBoard], ['issues', serIssue]]){
    const ids = new Set();
    for (const o of S[kind]){
      ids.add(o.id);
      const data = ser(o), json = stable(data);
      if (synced[kind].get(o.id) !== json){ ops.push({type:'set', col:kind, id:o.id, data:JSON.parse(JSON.stringify(data))}); synced[kind].set(o.id, json); base[kind].set(o.id, clone(o)); }
    }
    for (const id of [...synced[kind].keys()]) if (!ids.has(id)){ ops.push({type:'del', col:kind, id}); synced[kind].delete(id); base[kind].delete(id); }
  }
  const mj = stable(serMeta());
  if (synced.meta !== mj){ ops.push({type:'set', col:'meta', id:'settings', data:serMeta()}); synced.meta = mj; }
  if (!ops.length) return;
  const target = wsId;
  store.commit(target, ops).catch(err=>{
    console.error(err);
    if (target !== wsId){ toast(t('toast.saveFailed')); return; }
    toast(err?.code === 'permission-denied' ? t('toast.noPermission') : t('toast.saveFailed'));
    openWorkspace(wsId, true);
  });
}
addEventListener('pagehide', flush);
document.addEventListener('visibilitychange', ()=>{ if (document.visibilityState === 'hidden') flush(); });

function touch(i, text){
  i.updated = Date.now();
  if (text){ i.history.unshift({id:uid(), at:Date.now(), text, by:user?.name || ''}); if (i.history.length > 200) i.history.length = 200; }
}

/* ---------- columns shared by name across boards (All boards view) ---------- */
function vColumns(){
  const map = new Map();
  S.boards.forEach(b=>{ const L = b.columns.length; b.columns.forEach((c,ix)=>{
    const k = c.name.trim().toLowerCase();
    if (!map.has(k)) map.set(k, {id:'v:'+k, name:c.name, cat:c.cat, pos:0, n:0, first:map.size});
    const v = map.get(k); v.pos += L > 1 ? ix/(L-1) : 0; v.n++;
  }); });
  return [...map.values()].map(v=>({...v, pos:v.pos/v.n})).sort((a,b)=>a.pos-b.pos || a.first-b.first);
}
const vKey = i=>'v:'+colOf(i).name.trim().toLowerCase();
function resolveStatus(b, status){
  if (!status || !status.startsWith('v:')) return colById(b,status) || null;
  const k = status.slice(2), v = vColumns().find(x=>x.id===status);
  return b.columns.find(c=>c.name.trim().toLowerCase()===k) || (v && b.columns.find(c=>c.cat===v.cat)) || null;
}

let qTimer = 0, query = '', qf = new Set(), typeFilter = null, platFilter = null, openId = null, tab = 'comments', cmtDraft = '', editingCid = null, lastPct = null;
let lastCreateType = null;
let composer = null, composerText = '', composerType = null, composerBoard = null, composerFocus = false;
let sortBy = {k:'status', dir:1};
let fx = {};

/* ================= Screens ================= */
const appEl = $('#app'), gateEl = $('#gate');
function showGate(html){
  screen = html ? 'gate' : screen;
  appEl.hidden = true; gateEl.hidden = false; gateEl.innerHTML = html;
  document.title = 'Workbench';
}
function showApp(){
  if (screen === 'app') return;
  screen = 'app'; gateEl.hidden = true; gateEl.innerHTML = ''; appEl.hidden = false;
}
function gateFrame(inner){
  return `<div class="gate-top"><button class="btn subtle" data-act="lang">${I.globe}<span>${getLang()==='he'?'English':'עברית'}</span></button>
      <button class="btn icon subtle" data-act="theme" aria-label="${esc(t('nav.theme'))}" title="${esc(t('nav.theme'))}" aria-haspopup="menu">${themeIcon()}</button></div>
    <div class="gate-card">
      <div class="gate-brand"><img class="gate-logo brand-ink" src="assets/mccann-valley.png" alt="McCann Valley"><span class="gate-app"><span class="logo">${APP_MARK}</span>Workbench</span></div>
      ${inner}
    </div>`;
}
function renderSignIn(){
  screen = 'signin';
  const demo = store?.mode === 'demo';
  showGate(gateFrame(`
    <h1>${esc(t('gate.welcome'))}</h1>
    <p class="gate-sub">${esc(t('gate.sub'))}</p>
    ${demo ? `<div class="demo-note">${esc(isConfigured ? t('gate.demoForced') : t('gate.demo'))}</div>
      <form class="demo-form" data-form="demo-signin" novalidate>
        <label class="fld">${esc(t('gate.demoName'))}<input name="name" autocomplete="name" data-bidi></label>
        <label class="fld">${esc(t('gate.demoEmail'))}<input name="email" type="email" inputmode="email" autocomplete="email" dir="ltr" required></label>
        <button class="btn primary big">${esc(t('gate.demoSignIn'))}</button>
      </form>`
    : `<button class="btn google big" data-act="signin">${I.key}<span>${esc(t('gate.google'))}</span></button>
       <p class="gate-fine">${esc(t('gate.fine'))}</p>`}
  `));
}
function renderLoading(msg){
  screen = 'loading';
  showGate(`<div class="gate-loading"><span class="spinner"></span><span>${esc(msg || t('gate.loading'))}</span></div>`);
}
function localImportCandidate(){
  try {
    const raw = JSON.parse(localStorage.getItem('workbench.board.v2'));
    if (!raw) return null;
    const boards = Array.isArray(raw.boards) ? raw.boards.length : 0, issues = Array.isArray(raw.issues) ? raw.issues.length : 0;
    return issues || boards ? {raw, boards, issues} : null;
  } catch(e){ return null; }
}
function renderWelcome(){
  screen = 'welcome';
  const owned = (wsList||[]).some(w=>w.ownerUid === user.uid);
  const canCancel = (wsList||[]).length > 0;
  const pending = outgoing.filter(r=>r.status === 'pending'), declined = outgoing.filter(r=>r.status === 'declined');
  const local = localImportCandidate();
  showGate(gateFrame(`
    <div class="who">${avatar(user,'lg')}<div><b data-bidi>${esc(user.name)}</b><span dir="ltr">${esc(user.email)}</span></div><button class="btn subtle sm" data-act="signout">${esc(t('acc.signout'))}</button></div>
    <h1>${esc(t('welcome.title'))}</h1>
    <p class="gate-sub">${esc(t('welcome.sub'))}</p>
    ${pending.length ? `<div class="req-status">${pending.map(r=>`<div class="req-row">${I.mail}<span>${t('welcome.waiting', {email:`<b dir="ltr">${esc(r.toEmail)}</b>`})}</span><button class="btn subtle sm" data-act="req-cancel" data-id="${esc(r.id)}">${esc(t('btn.cancelRequest'))}</button></div>`).join('')}</div>` : ''}
    ${declined.length ? `<div class="req-status declined">${declined.map(r=>`<div class="req-row">${I.x}<span>${t('welcome.declined', {email:`<b dir="ltr">${esc(r.toEmail)}</b>`})}</span><button class="btn subtle sm" data-act="req-cancel" data-id="${esc(r.id)}">${esc(t('btn.dismiss'))}</button></div>`).join('')}</div>` : ''}
    ${owned ? '' : !mayCreate ? `<div class="note-box">${I.key}<span>${esc(t('welcome.noCreate'))}</span></div>` : `<label class="fld ws-name-fld">${esc(t('welcome.wsName'))}<input id="wsNameIn" maxlength="100" data-bidi value="${esc(t('welcome.wsNameDefault', {name:(user.name||'').split(' ')[0]||user.email}))}" autocomplete="off"></label>
    <div class="choices">
      <button class="choice" data-act="ws-create" data-kind="template">
        <span class="ch-ico tpl">${I.board}</span><span class="ch-txt"><b>${esc(t('welcome.tpl'))}</b><span>${esc(t('welcome.tplSub'))}</span>
        <span class="ch-tags">${PPC_CLIENTS.map(([n])=>`<i dir="rtl">${esc(n)}</i>`).join('')}</span></span></button>
      <button class="choice" data-act="ws-create" data-kind="empty">
        <span class="ch-ico">${I.plus}</span><span class="ch-txt"><b>${esc(t('welcome.empty'))}</b><span>${esc(t('welcome.emptySub'))}</span></span></button>
      ${local ? `<button class="choice" data-act="ws-create" data-kind="local">
        <span class="ch-ico">${I.up}</span><span class="ch-txt"><b>${esc(t('welcome.local'))}</b><span>${esc(t('welcome.localSub', {b:local.boards, i:local.issues}))}</span></span></button>` : ''}
    </div>`}
    <div class="choice-req">
      <div class="cr-head"><span class="ch-ico">${I.users}</span><span class="ch-txt"><b>${esc(t('welcome.ask'))}</b><span>${esc(t('welcome.askSub'))}</span></span></div>
      <form class="cr-form" data-form="request" novalidate><input class="inp" name="email" type="email" inputmode="email" dir="ltr" placeholder="name@example.com" aria-label="${esc(t('welcome.askEmail'))}" autocomplete="off"><button class="btn">${esc(t('btn.sendRequest'))}</button></form>
      <em class="ed-err" id="reqErr" hidden></em>
    </div>
    ${canCancel ? `<button class="btn subtle back-btn" data-act="welcome-back">${esc(t('btn.backToBoards'))}</button>` : ''}
  `));
  const n = $('#wsNameIn'); if (n) n.dataset.touched = '';
}

/* ================= Auth & workspaces ================= */
async function onAuth(u){
  unsubWs?.(); unsubIn?.(); unsubOut?.(); unsubData?.();
  unsubWs = unsubIn = unsubOut = unsubData = null;
  wsList = null; wsId = null; ws = null; incoming = []; outgoing = [];
  user = u;
  if (!u){ renderSignIn(); return; }
  renderLoading();
  mayCreate = await store.canCreate(u.email);
  if (user !== u) return;
  unsubIn = store.watchIncoming(u.email, list=>{ incoming = cleanReqs(list); renderBell(); if (modalKind === 'wsSettings') renderWsSettingsBody(); }, onListenErr);
  unsubOut = store.watchOutgoing(u.uid, list=>{ outgoing = cleanReqs(list); if (screen === 'welcome') renderWelcome(); }, onListenErr);
  unsubWs = store.watchWorkspaces(u, list=>{
    const before = wsList;
    wsList = cleanWs(list).sort((a,b)=>(a.ownerUid===u.uid?0:1)-(b.ownerUid===u.uid?0:1) || String(a.name).localeCompare(String(b.name)));
    if (awaitOpen && wsList.some(w=>w.id===awaitOpen)){ const id = awaitOpen; awaitOpen = null; openWorkspace(id); return; }
    if (wsId && !wsList.some(w=>w.id===wsId)){ toast(t('toast.lostAccess')); wsId = null; ws = null; unsubData?.(); unsubData = null; if (screen === 'app') screen = 'loading'; }
    if (screen === 'welcome' && before){
      const newOne = wsList.find(w=>!before.some(b=>b.id===w.id));
      if (newOne && !wsId){ openWorkspace(newOne.id); return; }
      if (wsId) ws = wsList.find(w=>w.id===wsId);
      renderWelcome(); return;
    }
    if (wsId){ ws = wsList.find(w=>w.id===wsId); applyRole(); renderNavWs(); if (modalKind === 'wsSettings') renderWsSettingsBody(); scheduleRender(); return; }
    if (screen === 'loading' && awaitOpen !== null) return;
    if (!wsList.length){ renderWelcome(); return; }
    const pick = wsList.find(w=>w.id===P.lastWs) || wsList.find(w=>w.ownerUid===u.uid) || wsList[0];
    openWorkspace(pick.id);
  }, onDataErr);
}
function onListenErr(e){ console.error(e); if (e?.code === 'permission-denied') return; toast(t('toast.loadFailed')); }
function onDataErr(e){
  console.error(e);
  if (screen === 'app'){ if (e?.code !== 'permission-denied') toast(t('toast.loadFailed')); return; }
  screen = 'error';
  showGate(gateFrame(`<h1>${esc(t('gate.loadErrTitle'))}</h1><p class="gate-sub">${esc(e?.code === 'permission-denied' ? t('gate.loadErrPerm') : t('gate.loadErrNet'))}</p>
    <button class="btn primary big" data-act="retry">${esc(t('btn.retry'))}</button><button class="btn subtle back-btn" data-act="signout">${esc(t('acc.signout'))}</button>`));
}
function openWorkspace(id, reload){
  flush();
  unsubData?.(); unsubData = null;
  wsId = id; ws = (wsList||[]).find(w=>w.id===id) || null; P.lastWs = id; savePrefs();
  S = {boards:[], issues:[], types:[], archive:normArchive()}; synced.boards.clear(); synced.issues.clear(); synced.meta = null; base.boards.clear(); base.issues.clear();
  loaded = {boards:false, issues:false, meta:false}; composer = null; openId = null;
  if (modalKind) closeModal(true);
  if (!reload){ screen = 'loading'; renderLoading(t('gate.opening', {name:ws?.name||''})); }
  applyRole(); renderNavWs();
  unsubData = store.watchData(id, {boards:onBoards, issues:onIssues, meta:onMeta}, onDataErr);
}
function applyRole(){
  document.body.classList.toggle('ro', !!ws && !canEdit());
  const chip = $('#roChip'); if (chip) chip.hidden = !ws || canEdit();
}
async function createWorkspace(kind){
  const nameIn = $('#wsNameIn');
  let name = (nameIn?.value || '').trim() || t('welcome.wsNameDefault', {name:(user.name||'').split(' ')[0]||user.email});
  if (kind === 'template' && nameIn && !nameIn.dataset.touched) name = t('welcome.tplWsName');
  renderLoading(t('gate.creating'));
  awaitOpen = '';
  try {
    const id = await store.createWorkspace({name:name.slice(0,100), logo:null, ownerUid:user.uid, ownerEmail:user.email, ownerName:String(user.name||'').slice(0,200), roles:{}, memberEmails:[], created:Date.now(), lang:getLang()});
    const ops = [];
    let boards = [], issues = [], types;
    if (kind === 'template'){
      const logos = await Promise.all(PPC_CLIENTS.map(([,,,f])=>fetch(`assets/clients/${f}.png`).then(r=>r.ok ? r.blob() : null).then(b=>b && readLogo(b)).catch(()=>null)));
      boards = PPC_CLIENTS.map(([n,k,c],ix)=>({id:'b-'+uid(), name:n, key:k, color:c, logo:logos[ix] || null, seq:0, order:ix*1000, columns:COLS_PPC(), tc:TC}));
      types = TYPES_PPC();
    } else if (kind === 'local'){
      const conv = convertLocal(localImportCandidate()?.raw);
      boards = conv.boards; issues = conv.issues; types = conv.types;
    } else {
      types = TYPES_BASIC();
    }
    boards.forEach(b=>{ const {id:bid, ...d} = b; ops.push({type:'set', col:'boards', id:bid, data:d}); });
    issues.forEach(i=>{ const {id:iid, ...d} = i; ops.push({type:'set', col:'issues', id:iid, data:d}); });
    ops.push({type:'set', col:'meta', id:'settings', data:{types}});
    await store.commit(id, ops);
    P.active[id] = boards[0]?.id || ALL; savePrefs();
    if ((wsList||[]).some(w=>w.id===id)){ awaitOpen = null; openWorkspace(id); } else awaitOpen = id;
    if (kind === 'local') toast(t('toast.localImported', {n:issues.length}));
  } catch(e){ console.error(e); awaitOpen = null; toast(t('toast.createFailed')); renderWelcome(); }
}
function convertLocal(raw){
  const out = {boards:[], issues:[], types:TYPES_BASIC()};
  if (!raw) return out;
  const bmap = new Map();
  (raw.boards||[]).forEach((b,ix)=>{ const nb = normBoard('b-'+uid(), {...b, order:ix*1000}); bmap.set(String(b?.id), nb); out.boards.push(nb); });
  if (!out.boards.length && Array.isArray(raw.issues)){
    const nb = normBoard('b-'+uid(), {name:'Workbench', key:'WB', columns:[{id:'todo',name:t('tpl.col.todo'),cat:'todo'},{id:'doing',name:t('tpl.col.doing'),cat:'doing'},{id:'wait',name:t('cat.wait'),cat:'wait'},{id:'done',name:t('tpl.col.done'),cat:'done'}]});
    out.boards.push(nb); (raw.issues||[]).forEach(i=>{ bmap.set(String(i?.boardId||'_'), nb); });
  }
  if (Array.isArray(raw.types) && raw.types.length) out.types = raw.types.filter(Boolean).map(x=>({id:sid(x.id, 't'+uid()), name:String(x.name||''), color:isHex(x.color)?x.color:'#4BADE8', glyph:GLYPHS[x.glyph]?x.glyph:'check'}));
  const tset = new Set(out.types.map(x=>x.id));
  (raw.issues||[]).forEach((i,ix)=>{
    if (!i || typeof i !== 'object') return;
    const b = bmap.get(String(i.boardId)) || bmap.get('_') || out.boards[0]; if (!b) return;
    const n = normIssue('i-'+uid(), {...i, boardId:b.id, rank:ix*1000});
    if (!b.columns.some(c=>c.id===n.status)) n.status = b.columns[0].id;
    if (!tset.has(n.type)) n.type = out.types[0].id;
    b.seq = Math.max(b.seq, n.num);
    out.issues.push(n);
  });
  out.boards.forEach(b=>{ out.issues.filter(i=>i.boardId===b.id && !i.num).forEach(i=>{ i.num = ++b.seq; }); });
  return out;
}

/* ================= Rendering ================= */
const board = $('#board'), listview = $('#listview');
function matches(i){
  const q = query.trim().toLowerCase();
  if (q){
    const hay = [i.title, key(i), i.desc, boardOf(i).name, i.color ? colorName(i.color) : '', ...i.labels, ...i.subtasks.map(s=>s.text), ...i.links.map(l=>l.title+' '+l.url)].join('\n').toLowerCase();
    if (!hay.includes(q)) return false;
  }
  const tday = today(), done = isDone(i);
  if (qf.has('overdue') && !(i.due && i.due < tday && !done)) return false;
  if (qf.has('week') && !(i.due && i.due >= tday && i.due <= addDays(7) && !done)) return false;
  if (qf.has('high') && PRANK[i.priority] > 1) return false;
  if (typeFilter && i.type !== typeFilter) return false;
  if (platFilter === 'none' ? !!i.color : (platFilter && i.color !== platFilter)) return false;
  return true;
}
const filtering = ()=>!!query.trim() || qf.size > 0 || !!typeFilter || !!platFilter;
function dueClass(i){
  if (!i.due || isDone(i)) return '';
  if (i.due < today()) return 'over';
  if (i.due <= addDays(1)) return 'soon';
  return '';
}
function dueLabel(i){
  if (!isDone(i)){ if (i.due === today()) return t('due.today'); if (i.due === addDays(1)) return t('due.tomorrow'); }
  return fmtDay(i.due);
}
function ageInfo(i){
  const done = isDone(i), end = done && i.doneAt ? iso(new Date(i.doneAt)) : today(), n = daysBetween(i.start, end);
  const short = x=>t('age.short', {n:x});
  if (n < 0) return {n, short:t('age.in', {n:-n}), long:t('age.starts', {date:fmtDay(i.start), d:tn('days', -n)}), cls:'future'};
  if (done) return {n, short:short(n), long:t('age.took', {d:tn('days', n), from:fmtDay(i.start), to:fmtDay(end)}), cls:'done'};
  return {n, short:short(n), long:n === 0 ? t('age.startedToday') : t('age.running', {date:fmtDay(i.start), d:tn('days', n)}), cls:n >= 14 ? 'long' : ''};
}
function visibleLabels(i){ return isAll() ? i.labels : i.labels.filter(l=>!sameText(l, boardOf(i).name)); }
// Short one-paragraph preview of the description for cards (CSS clamps it to 2 lines)
const cardDesc = s=>{ s = String(s).replace(/\s+/g, ' ').trim(); return s.length > 160 ? s.slice(0, 157).trimEnd() + '…' : s; };
// Who opened the task. Tasks from before this was saved fall back to the name in their first history line.
const CREATED_TEXTS = new Set(Object.values(_dicts).map(d=>d['h.created']).filter(Boolean));
function creatorOf(i){
  if (i.creator && (i.creator.name || i.creator.email)) return i.creator;
  const h = i.history[i.history.length-1];
  return h && h.by && CREATED_TEXTS.has(h.text) ? {name:h.by} : null;
}
const isShared = ()=>!!ws && (Object.keys(ws.roles||{}).length > 0 || (ws.memberEmails||[]).length > 1);
function creatorAv(i){
  const c = creatorOf(i); if (!c) return '';
  const nm = c.name || c.email;
  return `<span class="c-by" title="${esc(t('card.createdBy', {name:nm}))}" aria-label="${esc(t('card.createdBy', {name:nm}))}">${avatar(c,'xs')}</span>`;
}
function cardHTML(i){
  const n = i.subtasks.length, d = i.subtasks.filter(s=>s.done).length, dc = dueClass(i), done = isDone(i), age = ageInfo(i), labels = visibleLabels(i);
  const cls = ['card']; if (i.color) cls.push('has-color');
  if (fx.new===i.id) cls.push('is-new'); if (fx.land===i.id) cls.push('is-land'); if (fx.done===i.id) cls.push('is-done-fx');
  return `<article class="${cls.join(' ')}" data-id="${i.id}" tabindex="0"${i.color?` style="--cc:${i.color}" title="${esc(colorName(i.color))}"`:''} aria-label="${esc(key(i)+': '+(i.title||t('issue.untitled')))}">
    <p class="c-title" data-bidi>${hl(i.title || t('issue.untitled'))}</p>
    ${i.desc.trim() ? `<p class="c-desc" data-bidi>${hl(cardDesc(i.desc))}</p>` : ''}
    <button class="c-menu" data-act="card-menu" data-id="${i.id}" aria-label="${esc(t('card.actions', {key:key(i)}))}" aria-haspopup="menu">${I.more}</button>
    ${labels.length ? `<div class="c-labels">${labels.map(l=>`<span class="lbl" data-bidi>${hl(l)}</span>`).join('')}</div>` : ''}
    <div class="c-dates">${i.due ? `<span class="due ${dc}" title="${esc(t('due.title', {date:fmtDay(i.due)}))}">${I.cal}${dc==='over'?esc(t('due.overduePrefix')):''}${esc(dueLabel(i))}</span>` : ''}<span class="age ${age.cls}" title="${esc(age.long)}">${I.clock}<bdi>${esc(age.short)}</bdi></span></div>
    <div class="c-foot">${isAll() ? bAv(boardOf(i),'xs') : ''}${typeIcon(i.type)}<span class="key${done?' done':''}">${hl(key(i))}</span>
      ${i.recur ? `<span class="meta-n rc-ico" title="${esc(t('rc.cardTip', {rule:recurText(i.recur), date:fmtDay(i.recur.next)}))}">${I.repeat}</span>` : ''}
      ${n ? `<span class="meta-n${d===n?' all':''}" title="${esc(t('card.subtasks', {d, n}))}">${I.sub}<bdi>${d}/${n}</bdi></span>` : ''}
      ${i.comments.length ? `<span class="meta-n" title="${esc(tn('card.comments', i.comments.length))}">${I.comment}${i.comments.length}</span>` : ''}
      ${i.links.length ? `<span class="meta-n" title="${esc(tn('card.links', i.links.length))}">${I.link}${i.links.length}</span>` : ''}
      <span class="c-right">${i.points!=null ? `<span class="pts" title="${esc(t('f.points'))}">${esc(i.points)}</span>` : ''}${prioIcon(i.priority)}${isShared() ? creatorAv(i) : ''}</span>
    </div>
  </article>`;
}
function composerHTML(){
  if (typeOf(composerType)?.id !== composerType) composerType = S.types[0]?.id;
  if (!boardById(composerBoard)) composerBoard = (boardById(P.lastBoard?.[wsId]) || S.boards[0])?.id;
  return `<form class="composer" data-form="inline">
    <textarea data-bidi rows="2" placeholder="${esc(t('composer.ph'))}" aria-label="${esc(t('composer.aria'))}">${esc(composerText)}</textarea>
    <div class="comp-row">
      <label class="type-sel" title="${esc(t('f.type'))}">${typeIcon(composerType)}${I.chev}
        <select aria-label="${esc(t('f.type'))}" data-c="type">${S.types.map(x=>`<option value="${x.id}"${x.id===composerType?' selected':''}>${esc(x.name)}</option>`).join('')}</select></label>
      ${isAll() ? `<label class="type-sel" title="${esc(t('f.board'))}">${bAv(boardById(composerBoard),'xs')}${I.chev}<select aria-label="${esc(t('f.board'))}" data-c="board">${S.boards.map(b=>`<option value="${b.id}"${b.id===composerBoard?' selected':''}>${esc(b.name)}</option>`).join('')}</select></label>` : ''}
      <span class="comp-hint">${esc(t('composer.hint'))}</span>
      <button type="submit" class="btn primary sm">${esc(t('btn.create'))}</button>
    </div>
  </form>`;
}
let prevCounts = {};
let lastCols = null;
function archivePromptCount(){
  if (!canEdit() || !S.archive?.on) return 0;
  const d = new Date(); if (d.getDate() < S.archive.day || S.archive.handled === ym(d)) return 0;
  return S.issues.filter(i=>!i.archived && isDone(i)).length;
}
function archivePromptHTML(){
  const n = archivePromptCount(); if (!n) return '';
  return `<div class="arch-prompt" role="status"><div class="ap-head">${I.archive}<b>${esc(t('ap.title'))}</b></div><p>${esc(tn('ap.text', n))}</p>
    <div class="ap-acts"><button class="btn primary sm" data-act="ap-yes">${esc(t('ap.yes'))}</button><button class="btn subtle sm" data-act="ap-no">${esc(t('ap.no'))}</button></div></div>`;
}
let boardPending = false; // a redraw that arrived mid-drag; done when the card is dropped
function renderBoard(){
  // redrawing under a dragged card would bring the card back into its column and drop the placeholder
  if (drag?.active || settling){ boardPending = true; return; }
  boardPending = false;
  const all = isAll(), b = curBoard();
  if (!b){
    lastCols = null;
    board.innerHTML = `<div class="empty-state wide">${I.board}<b>${esc(t('empty.noBoards'))}</b><span>${esc(canEdit() ? t('empty.noBoardsSub') : t('empty.noBoardsRo'))}</span>${canEdit() ? `<button class="btn primary" data-act="new-board">${I.plus}${esc(t('btn.createBoard'))}</button>` : ''}</div>`;
    return;
  }
  const scroll = {}; $$('.list',board).forEach(l=>scroll[l.dataset.status] = l.scrollTop);
  // keep typing in the inline composer when its column is redrawn (e.g. a save comes back from the server)
  const ca = document.activeElement, compSel = ca?.matches?.('.composer textarea') ? [ca.selectionStart, ca.selectionEnd] : null;
  const pool = scopeIssues(), vis = pool.filter(matches), f = filtering();
  const cols = all ? vColumns() : b.columns, inCol = all ? (i,c)=>vKey(i)===c.id : (i,c)=>i.status===c.id;
  if (composer && !cols.some(c=>c.id===composer)) composer = null;
  const edit = canEdit();
  const html = cols.map(c=>{
    const items = vis.filter(i=>inCol(i,c)), total = pool.filter(i=>inCol(i,c)).length;
    const empty = !pool.length && c === cols[0] ? `<div class="empty-hint">${esc(all ? t('empty.allNone') : edit ? t('empty.board') : t('empty.boardRo'))}</div>` : '';
    return `<section class="col" data-col="${esc(c.id)}" data-cat="${c.cat}" aria-label="${esc(c.name)}">
      <header class="col-head"><span class="col-name" data-bidi title="${esc(c.name)}">${esc(c.name)}</span><span class="col-count" data-s="${esc(c.id)}"><bdi>${f ? esc(t('col.countOf', {a:items.length, b:total})) : total}</bdi></span>${c.cat==='done'&&items.length?`<span class="col-done" title="${esc(catName('done'))}">${I.check}</span>`:''}${c.cat==='done'&&edit?`<button class="btn icon subtle sm col-arch" data-act="archive-col" data-status="${esc(c.id)}" aria-label="${esc(t('col.archiveAll', {name:c.name}))}" title="${esc(items.length ? tn('col.archiveN', items.length) : t('col.archiveNone'))}"${items.length?'':' disabled'}>${I.archive}</button>`:''}</header>
      ${c.cat==='done' && c === cols.find(x=>x.cat==='done') ? archivePromptHTML() : ''}
      <div class="list" data-status="${esc(c.id)}">${items.map(cardHTML).join('') || empty}</div>
      ${edit ? `<div class="col-foot">${composer===c.id ? composerHTML() : `<button class="add-inline" data-act="add-inline" data-status="${esc(c.id)}">${I.plus}${esc(t('btn.createIssue'))}</button>`}</div>` : '<div class="col-foot ro-foot"></div>'}
    </section>`;
  });
  // Only touch the columns that changed: keeps renders cheap and avoids re-triggering animations elsewhere.
  const layoutKey = wsId + '|' + (all ? ALL : b.id) + '|' + cols.map(c=>c.id).join(',');
  const secs = board.children;
  if (lastCols && lastCols.key === layoutKey && secs.length === html.length){
    html.forEach((h,ix)=>{ if (lastCols.html[ix] !== h){ const tpl = document.createElement('template'); tpl.innerHTML = h.trim(); secs[ix].replaceWith(tpl.content.firstElementChild); } });
  } else board.innerHTML = html.join('');
  lastCols = {key:layoutKey, html};
  $$('.list',board).forEach(l=>{ if (scroll[l.dataset.status] && l.scrollTop !== scroll[l.dataset.status]) l.scrollTop = scroll[l.dataset.status]; });
  const scopeKey = all ? ALL : b.id;
  cols.forEach(c=>{
    const k = wsId+scopeKey+c.id, n = pool.filter(i=>inCol(i,c)).length;
    if (prevCounts[k] !== undefined && n > prevCounts[k] && !reduce) $(`.col-count[data-s="${CSS.escape(c.id)}"]`,board)?.classList.add('bump');
    prevCounts[k] = n;
  });
  if (composer){
    const ta = $('.composer textarea',board); autosize(ta);
    if (compSel && document.activeElement !== ta && !composerFocus){ ta.focus({preventScroll:true}); try { ta.setSelectionRange(...compSel); } catch(e){} }
    if (composerFocus){ ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length); composerFocus = false;
      const list = $(`.list[data-status="${CSS.escape(composer)}"]`,board); if (list) list.scrollTop = list.scrollHeight; }
  }
  const nc = fx.new && $(`.card[data-id="${fx.new}"]`,board);
  if (nc) nc.scrollIntoView({block:'nearest', inline:'nearest', behavior:reduce?'auto':'smooth'});
}
function sortVal(i,k){
  switch(k){
    case 'type': return typeOf(i.type).name.toLowerCase(); case 'key': return boardOf(i).key + String(i.num).padStart(6,'0'); case 'title': return i.title.toLowerCase();
    case 'board': return boardOf(i).name.toLowerCase();
    case 'status': return isAll() ? vColumns().findIndex(v=>v.id===vKey(i)) : boardOf(i).columns.indexOf(colOf(i)); case 'priority': return PRANK[i.priority];
    case 'due': return i.due || '9999'; case 'days': return ageInfo(i).n; case 'points': return i.points==null ? Infinity : Number(i.points);
    case 'updated': return -i.updated; case 'archived': return -(i.archivedAt||0);
  }
}
function renderList(archive){
  const all = isAll(), b = curBoard();
  if (!b){ listview.innerHTML = `<div class="empty-state"><b>${esc(t('empty.noBoards'))}</b></div>`; return; }
  const scope = S.issues.map((i,ix)=>[i,ix]).filter(([i])=>!!i.archived === !!archive && (all || i.boardId===b.id));
  const vis = scope.filter(([i])=>matches(i));
  const sk = archive && !['type','key','title','board','status','archived'].includes(sortBy.k) ? 'archived' : sortBy.k;
  vis.sort((a,c)=>{ const x = sortVal(a[0],sk), y = sortVal(c[0],sk); return x===y ? a[1]-c[1] : (x<y?-1:1) * (sk===sortBy.k ? sortBy.dir : 1); });
  if (!vis.length){
    listview.innerHTML = archive
      ? `<div class="empty-state"><b>${esc(scope.length ? t('empty.archiveFiltered') : all ? t('empty.archiveAll') : t('empty.archiveBoard'))}</b>${esc(scope.length ? t('empty.tryFilters') : t('empty.archiveHow'))}</div>`
      : `<div class="empty-state"><b>${esc(scope.length ? t('empty.listFiltered') : all ? t('empty.listAll') : t('empty.listBoard'))}</b>${esc(scope.length ? t('empty.tryFilters') : canEdit() ? t('empty.listHow') : '')}</div>`;
    return;
  }
  const cols = archive
    ? [['type','th.type',''],['key','th.key',''],['title','th.summary',''],...(all?[['board','th.board','hide-m']]:[]),['status','th.status','hide-m'],['archived','th.archived','hide-m'],...(canEdit()?[['','','']]:[])]
    : [['type','th.type',''],['key','th.key',''],['title','th.summary',''],...(all?[['board','th.board','hide-m']]:[]),['status','th.status',''],['priority','th.priority','hide-m'],['due','th.due','hide-m'],['days','th.days','hide-m'],['updated','th.updated','hide-m']];
  const row = i=>{
    const cells = {
      type:`<td>${typeIcon(i.type)}</td>`,
      key:`<td><span class="key${isDone(i)?' done':''}">${hl(key(i))}</span></td>`,
      title:`<td class="t-sum" data-bidi>${i.color?`<span class="c-dot" style="--cc:${i.color}" title="${esc(colorName(i.color))}"></span>`:''}${hl(i.title||t('issue.untitled'))}${visibleLabels(i).map(l=>`<span class="lbl" data-bidi>${esc(l)}</span>`).join('')}</td>`,
      board:`<td class="hide-m"><span class="t-board">${bAv(boardOf(i),'xs')}<span data-bidi>${esc(boardOf(i).name)}</span></span></td>`,
      status:`<td${archive?' class="hide-m"':''}>${lozC(colOf(i))}</td>`,
      priority:`<td class="hide-m"><span class="c-pri">${prioIcon(i.priority)}${esc(prioName(i.priority))}</span></td>`,
      due:`<td class="hide-m">${i.due ? `<span class="due ${dueClass(i)}">${esc(dueLabel(i))}</span>` : `<span class="t-muted">${esc(t('none'))}</span>`}</td>`,
      days:(()=>{ const a = ageInfo(i); return `<td class="hide-m"><span class="age ${a.cls}" title="${esc(a.long)}">${I.clock}<bdi>${esc(a.short)}</bdi></span></td>`; })(),
      updated:`<td class="hide-m t-muted">${esc(ago(i.updated))}</td>`,
      archived:`<td class="hide-m t-muted">${esc(fmtDay(iso(new Date(i.archivedAt))))}</td>`,
      '':`<td class="t-act"><button class="btn sm" data-act="restore" data-id="${i.id}">${I.restore}${esc(t('btn.restore'))}</button></td>`
    };
    return `<tr data-id="${i.id}" tabindex="0">${cols.map(([k])=>cells[k]).join('')}</tr>`;
  };
  listview.innerHTML = `<table class="tbl"><thead><tr>${cols.map(([k,l,c])=>k ? `<th class="${c}"><button data-act="sort" data-k="${k}" class="${sk===k?'on':''}${sk===k&&sortBy.dir<0&&sk===sortBy.k?' desc':''}">${esc(t(l))}${I.arrow}</button></th>` : `<th><span class="sr">${esc(t('th.actions'))}</span></th>`).join('')}</tr></thead>
  <tbody>${vis.map(([i])=>row(i)).join('')}</tbody></table>`;
}
/* ================= Table view =================
   A spreadsheet-style grid of the same tasks: every cell edits the task directly, so it stays in
   sync with the board (and with everyone else) with no extra data. */
const sheetview = $('#sheetview');
let sheetSort = {k:null, dir:1}, sheetCache = null, sheetPending = false, sheetFocusVal = null, lastTouch = null;
// every column the table can show; the task name is always shown, the rest can be hidden (P.sheetHide, per device)
const sheetAllCols = ()=>['key','title','status','prio',...(isAll() ? ['board'] : []),'type','plat','start','due','company','desc','subs','by'];
const sheetHidden = ()=>new Set((Array.isArray(P.sheetHide) ? P.sheetHide : []).filter(k=>k !== 'title'));
function sheetCols(){ const h = sheetHidden(); return sheetAllCols().filter(k=>!h.has(k)); }
function sheetColsMenu(){
  const h = sheetHidden();
  return `<div class="menu-h">${esc(t('sheet.colsShow'))}</div><div class="sf-list dash-pick">${sheetAllCols().map(k=>{ const on = !h.has(k), lock = k === 'title';
      return `<button class="menu-i" role="menuitemcheckbox" aria-checked="${on}" data-act="s-col" data-k="${k}"${lock ? ' aria-disabled="true"' : ''}><span class="sf-box${on ? ' on' : ''}${lock ? ' lock' : ''}">${on ? I.check : ''}</span><span class="sf-l">${esc(t('x.col.' + k))}${lock ? `<small class="dash-only">${esc(t('sheet.colAlways'))}</small>` : ''}</span></button>`; }).join('')}</div>
    ${h.size ? `<div class="menu-sep"></div><div class="dash-pick-f"><button class="link-btn" data-act="s-colsall">${esc(t('sheet.colsAll'))}</button></div>` : ''}`;
}
function sheetColsRefresh(){
  render();
  const btn = $('[data-act="s-cols"]', sheetview);
  if (popAnchor && pop.dataset.id === 'scols' && btn){ popAnchor = btn; btn.setAttribute('aria-expanded', 'true'); const sc = $('.sf-list', pop)?.scrollTop || 0; pop.innerHTML = sheetColsMenu(); $('.sf-list', pop).scrollTop = sc; placePop(); }
}
function setSheetHidden(k, hide){ const h = sheetHidden(); hide ? h.add(k) : h.delete(k); P.sheetHide = [...h]; savePrefs(); }
const PRIO_TINT = {highest:'#C9372C', high:'#E56910', medium:'#B38600', low:'#1D7AFC', lowest:'#626F86'};
function sheetVal(i, k){
  switch(k){
    case 'key': return (boardOf(i)?.key || '') + String(i.num).padStart(6,'0');
    case 'title': return lower(i.title); case 'prio': return PRANK[i.priority];
    case 'status': { const b = boardOf(i); return b ? b.columns.findIndex(c=>c.id===i.status) : 0; }
    case 'board': return lower(boardOf(i)?.name); case 'type': return S.types.findIndex(x=>x.id===i.type);
    case 'plat': return i.color ? lower(platLabelOf(i.color)) : '￿';
    case 'start': return i.start || '￿'; case 'due': return i.due || '￿';
    case 'company': return lower(i.labels.join(',')) || '￿'; case 'desc': return lower(i.desc) || '￿';
    case 'subs': return i.subtasks.length ? i.subtasks.filter(x=>x.done).length / i.subtasks.length : 2;
    case 'by': { const c = creatorOf(i); return c ? lower(c.name || c.email) : '￿'; }
  }
  return 0;
}
function sheetCell(i, k, edit){
  const id = i.id, ro = !edit;
  const menu = (f, inner, cls='')=>ro ? `<span class="chip ${cls}">${inner}</span>` : `<button class="chip ${cls}" data-act="s-menu" data-sf="${f}" data-id="${id}" aria-haspopup="menu">${inner}${I.chev}</button>`;
  switch(k){
    case 'key': return `<td class="s-key"><button class="s-open" data-act="open-issue" data-id="${id}" title="${esc(t('sheet.open'))}">${typeIcon(i.type)}<span class="key${isDone(i)?' done':''}">${esc(key(i))}</span></button></td>`;
    case 'title': return `<td class="s-title">${ro ? `<span class="s-ro" data-bidi>${hl(i.title || t('issue.untitled'))}</span>` : `<input class="s-in" data-sf="title" data-id="${id}" value="${esc(i.title)}" data-bidi autocomplete="off" aria-label="${esc(t('x.col.title'))}">`}</td>`;
    case 'status': { const c = colOf(i); return `<td>${menu('status', `<span class="dot"></span><span data-bidi>${esc(c?.name || '')}</span>`, 'st '+(c?.cat||''))}</td>`; }
    case 'prio': return `<td>${menu('prio', `${prioIcon(i.priority)}<span>${esc(prioName(i.priority))}</span>`, 'pr')}</td>`.replace('class="chip pr"', `class="chip pr" style="--pc:${PRIO_TINT[i.priority]}"`);
    case 'board': return `<td>${menu('board', `${bAv(boardOf(i),'xs')}<span data-bidi>${esc(boardOf(i)?.name || '')}</span>`, 'plain')}</td>`;
    case 'type': return `<td>${menu('type', `${typeIcon(i.type)}<span data-bidi>${esc(typeOf(i.type).name)}</span>`, 'plain')}</td>`;
    case 'plat': return `<td>${menu('plat', i.color ? `<i class="pf-dot" style="--sw:${i.color}"></i><span>${esc(platLabelOf(i.color))}</span>` : `<span class="t-muted">—</span>`, i.color ? 'plat' : 'plain')}</td>`;
    case 'start': case 'due': {
      const f = k, v = i[f], cls = f === 'due' ? dueClass(i) : '';
      if (ro) return `<td class="s-date ${cls}">${v ? esc(fmtDay(v)) : '<span class="t-muted">—</span>'}</td>`;
      return `<td class="s-date ${cls}"><input type="date" class="s-in" data-sf="${f}" data-id="${id}" value="${esc(v)}" aria-label="${esc(t('x.col.'+f))}"></td>`; }
    case 'company': return `<td class="s-company">${ro ? `<span class="s-ro" data-bidi>${esc(i.labels.join(', '))}</span>` : `<input class="s-in" data-sf="labels" data-id="${id}" value="${esc(i.labels.join(', '))}" data-bidi autocomplete="off" placeholder="—" aria-label="${esc(t('x.col.company'))}">`}</td>`;
    case 'desc': return `<td class="s-desc">${ro ? `<span class="s-ro" data-bidi>${esc(cardDesc(i.desc))}</span>` : `<textarea class="s-in" data-sf="desc" data-id="${id}" rows="1" data-bidi placeholder="—" aria-label="${esc(t('x.col.desc'))}">${esc(i.desc)}</textarea>`}</td>`;
    case 'subs': { const n = i.subtasks.length, d = i.subtasks.filter(x=>x.done).length;
      return `<td class="s-subs">${n ? `<button class="s-prog${d===n?' all':''}" data-act="open-issue" data-id="${id}" title="${esc(t('card.subtasks', {d, n}))}"><span class="bar"><b style="width:${Math.round(d/n*100)}%"></b></span><bdi>${d}/${n}</bdi></button>` : '<span class="t-muted">—</span>'}</td>`; }
    case 'by': { const c = creatorOf(i); return `<td class="s-by">${c ? `${avatar(c,'xs')}<span data-bidi>${esc(c.name || c.email)}</span>` : '<span class="t-muted">—</span>'}</td>`; }
  }
  return '<td></td>';
}
/* ----- per-column filters in the table (like a spreadsheet filter) -----
   Lists of values with check marks for status, priority, client, type, platform, company, creator,
   date ranges and subtasks; "contains" text for the task name and the description. */
let sheetF = {};
const SF_KIND = {status:'list', prio:'list', board:'list', type:'list', plat:'list', company:'list', by:'list', start:'list', due:'list', subs:'list', title:'text', desc:'text'};
function sfDateFacets(v, due, done){
  if (!v) return ['none'];
  const tday = today(), out = [];
  if (due && v < tday && !done) out.push('over');
  if (v === tday) out.push('today');
  if (v >= tday && v <= addDays(7)) out.push('week');
  if (v.slice(0,7) === tday.slice(0,7)) out.push('month');
  out.push(v < tday ? 'past' : v > tday ? 'future' : 'todayx');
  return out;
}
function sfFacets(i, k){
  switch(k){
    case 'status': return [colOf(i)?.name || ''];
    case 'prio': return [i.priority];
    case 'board': return [i.boardId];
    case 'type': return [i.type];
    case 'plat': return [i.color || ''];
    case 'company': return i.labels.length ? i.labels.map(normLabel) : [''];
    case 'by': { const c = creatorOf(i); return [c ? (c.name || c.email) : '']; }
    case 'start': return sfDateFacets(i.start, false, isDone(i));
    case 'due': return sfDateFacets(i.due, true, isDone(i));
    case 'subs': { const n = i.subtasks.length, d = i.subtasks.filter(x=>x.done).length; return [!n ? 'none' : d === n ? 'done' : 'open']; }
  }
  return [];
}
function sfMatchOne(i, k){
  const f = sheetF[k]; if (!f) return true;
  if (SF_KIND[k] === 'text') return lower(k === 'title' ? i.title : i.desc).includes(lower(f));
  return sfFacets(i, k).some(v=>f.has(v));
}
const sheetMatch = (i, skip)=>Object.keys(sheetF).every(k=>k === skip || sfMatchOne(i, k));
const sheetFCount = ()=>Object.values(sheetF).filter(Boolean).length;
const sheetFKey = ()=>Object.entries(sheetF).filter(([,v])=>v).map(([k,v])=>k + '=' + (typeof v === 'string' ? v : [...v].sort().join('/'))).join(';');
function sfLabel(k, v){
  if (k === 'status') return v || '—';
  if (k === 'prio') return prioName(v);
  if (k === 'board') return boardById(v)?.name || '—';
  if (k === 'type') return typeOf(v).name;
  if (k === 'plat') return v ? platLabelOf(v) : t('filter.noPlatform');
  if (k === 'company') return v || t('sf.empty');
  if (k === 'by') return v || t('sf.unknown');
  if (k === 'start' || k === 'due') return t('sf.d.' + v);
  if (k === 'subs') return t('sf.s.' + v);
  return v;
}
function sfOptions(k, pool){
  const counts = new Map();
  pool.forEach(i=>sfFacets(i, k).forEach(v=>counts.set(v, (counts.get(v) || 0) + 1)));
  let order;
  if (k === 'prio') order = PRIOS;
  else if (k === 'status'){ const seen = []; (isAll() ? S.boards : [curBoard()]).forEach(b=>b?.columns.forEach(c=>{ if (!seen.includes(c.name)) seen.push(c.name); })); order = seen; }
  else if (k === 'board') order = S.boards.map(b=>b.id);
  else if (k === 'type') order = S.types.map(x=>x.id);
  else if (k === 'start') order = ['today','week','month','past','future'];
  else if (k === 'due') order = ['over','today','week','month','future','none'];
  else if (k === 'subs') order = ['open','done','none'];
  else order = [...counts.keys()].sort((a, b)=>(!a) - (!b) || sfLabel(k, a).localeCompare(sfLabel(k, b)));
  const f = sheetF[k];
  return order.filter(v=>counts.has(v) || f?.has(v)).map(v=>({v, n:counts.get(v) || 0}));
}
function sheetFilterMenu(k){
  const b = curBoard(), all = isAll(), name = t('x.col.' + k);
  const pool = S.issues.filter(i=>!i.archived && (all || i.boardId === b?.id)).filter(matches).filter(i=>sheetMatch(i, k));
  if (SF_KIND[k] === 'text') return `<div class="menu-h">${esc(t('sf.by', {col:name}))}</div><div class="pop-pad sf-text"><input class="inp" data-k="s-ftext" data-col="${k}" value="${esc(sheetF[k] || '')}" placeholder="${esc(t('sf.contains'))}" data-bidi autocomplete="off"></div>
    ${sheetF[k] ? `<div class="menu-sep"></div><button class="menu-i" role="menuitem" data-act="s-fclear" data-k="${k}">${I.x}${esc(t('sf.clearCol'))}</button>` : ''}${k !== 'title' ? `<div class="menu-sep"></div><button class="menu-i" role="menuitem" data-act="s-hidecol" data-k="${k}">${I.eyeOff}${esc(t('sheet.hideCol'))}</button>` : ''}`;
  const f = sheetF[k], opts = sfOptions(k, pool);
  const dot = v=>k === 'plat' ? `<i class="pf-dot${v ? '' : ' none'}"${v ? ` style="--sw:${v}"` : ''}></i>` : k === 'type' ? typeSVG(typeOf(v)) : k === 'prio' ? prioIcon(v) : k === 'board' ? bAv(boardById(v), 'xs') : '';
  const hideCol = `<div class="menu-sep"></div><button class="menu-i" role="menuitem" data-act="s-hidecol" data-k="${k}">${I.eyeOff}${esc(t('sheet.hideCol'))}</button>`;
  return `<div class="menu-h">${esc(t('sf.showOnly', {col:name}))}</div>
    <button class="menu-i" role="menuitem" data-act="s-fall" data-k="${k}"><span class="sf-box${!f ? ' on' : ''}">${!f ? I.check : ''}</span>${esc(t('sf.all'))}</button>
    <div class="sf-list">${opts.map(o=>`<button class="menu-i" role="menuitemcheckbox" aria-checked="${!!f?.has(o.v)}" data-act="s-fval" data-k="${k}" data-v="${esc(o.v)}"><span class="sf-box${f?.has(o.v) ? ' on' : ''}">${f?.has(o.v) ? I.check : ''}</span>${dot(o.v)}<span class="sf-l" data-bidi>${esc(sfLabel(k, o.v))}</span><span class="sf-n">${o.n}</span></button>`).join('') || `<p class="muted-p pop-pad">${esc(t('sf.none'))}</p>`}</div>${hideCol}`;
}
function sheetFRefresh(k){
  render();
  // the header was rebuilt: re-attach the open menu to the new filter button and redraw it
  const btn = $(`.s-filt[data-k="${CSS.escape(k)}"]`, sheetview);
  if (popAnchor && btn){ popAnchor = btn; btn.setAttribute('aria-expanded', 'true'); pop.innerHTML = sheetFilterMenu(k); placePop(); }
}
function sheetFBar(){
  const act = Object.entries(sheetF).filter(([,v])=>v), nh = sheetHidden().size;
  const cols = `<button class="btn subtle sm s-colsbtn" data-act="s-cols" aria-haspopup="true" aria-expanded="false">${I.table}<span>${esc(t('sheet.cols'))}</span>${nh ? `<span class="dash-hid">${esc(tn('sheet.hiddenN', nh))}</span>` : ''}</button>`;
  if (!act.length) return `<div class="s-fbar">${cols}</div>`;
  return `<div class="s-fbar">${cols}<span class="s-fsep" aria-hidden="true"></span>${I.funnel}${act.map(([k, v])=>`<span class="s-fchip"><b>${esc(t('x.col.' + k))}:</b> <span data-bidi>${esc(typeof v === 'string' ? `"${v}"` : [...v].map(x=>sfLabel(k, x)).join(', '))}</span><button class="btn icon subtle sm" data-act="s-fclear" data-k="${k}" aria-label="${esc(t('sf.clearCol'))}">${I.x}</button></span>`).join('')}
    ${act.length > 1 ? `<button class="link-btn" data-act="s-fclear">${esc(t('sf.clearAll'))}</button>` : ''}</div>`;
}

const SHEET_PAGE = 150; // rows are added as you scroll, so big boards open fast (500 rows at once took ~0.6s)
let sheetLimit = SHEET_PAGE, sheetIO = null;
function renderSheet(){
  const all = isAll(), b = curBoard(), edit = canEdit();
  if (!b){ sheetview.innerHTML = `<div class="empty-state"><b>${esc(t('empty.noBoards'))}</b></div>`; sheetCache = null; return; }
  const cols = sheetCols();
  const bIx = new Map(S.boards.map((x,ix)=>[x.id, x.order ?? ix]));
  const colIx = i=>{ const bb = boardOf(i); return bb ? bb.columns.findIndex(c=>c.id===i.status) : 0; };
  let vis = S.issues.filter(i=>!i.archived && (all || i.boardId===b.id)).filter(matches).filter(i=>sheetMatch(i)).map((i,ix)=>[i,ix]);
  if (sheetSort.k){ const k = sheetSort.k, d = sheetSort.dir, E = '￿';
    vis.sort((a,c)=>{ const x = sheetVal(a[0],k), y = sheetVal(c[0],k); if (x===y) return a[1]-c[1]; if (x===E || x===2 && k==='subs') return 1; if (y===E || y===2 && k==='subs') return -1; return (x<y?-1:1)*d; }); } // empty cells always last
  else vis.sort((a,c)=>((bIx.get(a[0].boardId)??0)-(bIx.get(c[0].boardId)??0)) || (colIx(a[0])-colIx(c[0])) || (a[0].rank-c[0].rank) || a[1]-c[1]);
  vis = vis.map(x=>x[0]);
  const scope = wsId + (all ? ALL : b.id);
  const layoutKey = [scope, cols.join(','), sheetSort.k, sheetSort.dir, edit, getLang(), query, [...qf].join(), typeFilter, platFilter, sheetFKey()].join('|');
  if (!sheetCache || sheetCache.key !== layoutKey) sheetLimit = SHEET_PAGE;
  if (fx.new){ const ix = vis.findIndex(i=>i.id===fx.new); if (ix >= sheetLimit) sheetLimit = ix + 20; } // a task just created is always shown
  const shown = vis.slice(0, sheetLimit), more = vis.length - shown.length;
  const rows = shown.map(i=>[i.id, `<tr data-id="${i.id}">${cols.map(k=>sheetCell(i,k,edit)).join('')}</tr>`]);
  const a = document.activeElement, foc = a && sheetview.contains(a) && a.matches('input,textarea') ? a.closest('tr')?.dataset.id || 'add' : null;
  const tb = $('tbody', sheetview), ids = rows.map(r=>r[0]), order = ids.join(',');
  const moreRow = ()=>more ? `<tr class="s-more"><td colspan="${cols.length}">${esc(tn('sheet.more', more))}</td></tr>` : '';
  if (sheetCache && sheetCache.key === layoutKey && tb && (sheetCache.order === order || (sheetCache.order && order.startsWith(sheetCache.order + ',')))){
    // same rows (maybe more of them): swap only the rows that changed, never the one being edited
    for (const [id, h] of rows){
      if (!sheetCache.html.has(id)) continue;
      if (sheetCache.html.get(id) === h) continue;
      if (id === foc){ sheetPending = true; continue; }
      const tr = $(`tr[data-id="${CSS.escape(id)}"]`, tb); if (!tr) continue;
      const tpl = document.createElement('template'); tpl.innerHTML = h; tr.replaceWith(tpl.content.firstElementChild);
      sheetCache.html.set(id, h);
    }
    const extra = rows.filter(r=>!sheetCache.html.has(r[0]));
    $('.s-more', tb)?.remove();
    if (extra.length){ tb.insertAdjacentHTML('beforeend', extra.map(r=>r[1]).join('')); extra.forEach(([id,h])=>sheetCache.html.set(id, h)); sheetCache.order = order; }
    tb.insertAdjacentHTML('beforeend', moreRow());
    watchMore();
    return;
  }
  if (foc && foc !== 'add' && sheetCache && tb){ sheetPending = true; return; } // rows added/removed/re-ordered: rebuild once the edit is done
  const addVal = foc === 'add' ? a.value : null;
  const same = sheetCache && sheetCache.scope === scope, sx = same ? sheetview.scrollLeft : 0, sy = same ? sheetview.scrollTop : 0;
  const head = cols.map(k=>`<th class="h-${k}"><div class="th-in"><button data-act="s-sort" data-k="${k}" class="th-sort ${sheetSort.k===k?'on':''}${sheetSort.k===k&&sheetSort.dir<0?' desc':''}">${esc(t('x.col.'+k))}${I.arrow}</button>${SF_KIND[k] ? `<button class="s-filt${sheetF[k] ? ' on' : ''}" data-act="s-filter" data-k="${k}" aria-haspopup="menu" aria-label="${esc(t('sf.by', {col:t('x.col.'+k)}))}" title="${esc(t('sf.by', {col:t('x.col.'+k)}))}">${I.funnel}</button>` : ''}</div></th>`).join('');
  const fbar = sheetFBar();
  const empty = !rows.length ? `<tr class="s-empty"><td colspan="${cols.length}">${esc(filtering() || sheetFCount() ? t('empty.listFiltered') : t('sheet.empty'))}</td></tr>` : '';
  // the add bar lives outside the table: a sticky table footer cell drew as a blank box over the name column in iOS Safari
  const add = edit ? `<form data-form="sheet-add" class="s-addbar">${I.plus}<input data-k="sheet-add" placeholder="${esc(all ? t('sheet.addIn', {board:(boardById(P.lastBoard?.[wsId]) || S.boards[0]).name}) : t('sheet.add'))}" data-bidi autocomplete="off" enterkeyhint="done" aria-label="${esc(t('sheet.add'))}"></form>` : '';
  sheetview.innerHTML = fbar + `<table class="sheet${all?' all':''}${cols.includes('key') ? '' : ' nokey'}"><thead><tr>${head}</tr></thead><tbody>${rows.map(r=>r[1]).join('')}${empty}${moreRow()}</tbody></table>${add}`;
  sheetview.scrollLeft = sx; sheetview.scrollTop = sy;
  if (addVal != null){ const ad = $('[data-k="sheet-add"]', sheetview); if (ad){ ad.value = addVal; ad.focus({preventScroll:true}); } }
  sheetCache = {key:layoutKey, order, html:new Map(rows), scope};
  sheetPending = false;
  watchMore();
}
function watchMore(){
  sheetIO?.disconnect();
  const m = $('.s-more', sheetview); if (!m) return;
  sheetIO = new IntersectionObserver(es=>{ if (es.some(e=>e.isIntersecting)){ sheetIO.disconnect(); sheetLimit += SHEET_PAGE; requestAnimationFrame(()=>{ if (P.view === 'table') renderSheet(); }); } }, {root:sheetview, rootMargin:'600px'});
  sheetIO.observe(m);
}
function sheetMenu(i, f){
  const opt = (v, inner, on)=>`<button class="menu-i" role="menuitem" data-act="s-set" data-sf="${f}" data-v="${esc(v)}">${inner}${chk(on)}</button>`;
  if (f === 'status') return `<div class="menu-h">${esc(t('menu.changeStatus'))}</div>` + boardOf(i).columns.map(c=>opt(c.id, lozC(c), c.id===i.status)).join('');
  if (f === 'prio') return `<div class="menu-h">${esc(t('x.col.prio'))}</div>` + PRIOS.map(p=>opt(p, `${prioIcon(p)}<span>${esc(prioName(p))}</span>`, p===i.priority)).join('');
  if (f === 'type') return `<div class="menu-h">${esc(t('x.col.type'))}</div>` + S.types.map(x=>opt(x.id, `${typeSVG(x)}<span data-bidi>${esc(x.name)}</span>`, x.id===i.type)).join('');
  if (f === 'board') return `<div class="menu-h">${esc(t('x.col.board'))}</div>` + S.boards.map(x=>opt(x.id, `${bAv(x,'xs')}<span data-bidi>${esc(x.name)}</span>`, x.id===i.boardId)).join('');
  if (f === 'plat') return `<div class="menu-h">${esc(t('x.col.plat'))}</div>` + opt('', `<i class="pf-dot none"></i><span>${esc(t('filter.noPlatform'))}</span>`, !i.color)
    + PLATFORMS.map(([n,h])=>opt(h, `<i class="pf-dot" style="--sw:${h}"></i><span>${esc(platLabel(n))}</span>`, i.color===h)).join('')
    + (i.color && !PLATFORMS.some(x=>x[1]===i.color) ? opt(i.color, `<i class="pf-dot" style="--sw:${i.color}"></i><span>${esc(colorName(i.color))}</span>`, true) : '');
  return '';
}
function sheetSet(i, f, v){
  if (f === 'status'){ if (v !== i.status) moveIssue(i.id, v, null, {toast:false}); return; }
  if (f === 'board'){ if (v !== i.boardId) moveToBoard(i, v); return; }
  if (f === 'plat'){ if ((v||null) !== (i.color||null)) setColor(i, v); return; }
  if (f === 'prio'){ if (v === i.priority || !PRIOS.includes(v)) return; i.priority = v; touch(i, t('h.priority', {name:prioName(v)})); }
  if (f === 'type'){ if (v === i.type || !S.types.some(x=>x.id===v)) return; i.type = v; touch(i, t('h.type', {name:typeOf(v).name})); }
  save(); render();
}
// one history line per field while it's being edited (Safari's date picker fires "change" on open)
function touchOnce(i, f, text){
  const now = Date.now(), h = lastTouch && lastTouch.id === i.id && lastTouch.f === f && now - lastTouch.at < 120000 && i.history.find(x=>x.id === lastTouch.hid);
  if (h){ h.text = text; h.at = now; i.updated = now; lastTouch.at = now; return; }
  touch(i, text); lastTouch = {id:i.id, f, hid:i.history[0]?.id, at:now};
}
function sheetCommit(el){
  if (!canEdit()) return;
  const i = get(el.dataset.id), f = el.dataset.sf; if (!i) return;
  if (f === 'title'){
    const v = el.value.trim();
    if (!v){ el.value = i.title; return; }
    if (v === i.title) return;
    i.title = v.slice(0, 500); touchOnce(i, f, t('h.summary'));
  } else if (f === 'desc'){
    if (el.value === i.desc) return;
    i.desc = el.value; touchOnce(i, f, t('h.desc'));
  } else if (f === 'labels'){
    const next = [...new Map(el.value.split(',').map(normLabel).filter(Boolean).map(l=>[lower(l), l])).values()];
    const added = next.filter(l=>!i.labels.some(x=>sameText(x,l))), removed = i.labels.filter(l=>!next.some(x=>sameText(x,l)));
    if (!added.length && !removed.length) return;
    i.labels = next; added.forEach(l=>touch(i, t('h.companyAdd', {name:l}))); removed.forEach(l=>touch(i, t('h.companyRemove', {name:l})));
  } else if (f === 'due'){
    const v = isISO(el.value) ? el.value : '';
    if (v === i.due) return;
    i.due = v; touchOnce(i, f, v ? t('h.due', {date:fmtDay(v)}) : t('h.dueNone'));
  } else if (f === 'start'){
    const v = isISO(el.value) ? el.value : today();
    if (v === i.start){ el.value = v; return; }
    i.start = v; touchOnce(i, f, t('h.start', {date:fmtDay(v)}));
  } else return;
  save(); render();
}
sheetview.addEventListener('focusin', e=>{ const el = e.target; if (el.matches('.s-in')){ sheetFocusVal = el.value; if (el.tagName === 'TEXTAREA') autosize(el); } });
sheetview.addEventListener('change', e=>{ if (e.target.matches('.s-in')) sheetCommit(e.target); });
sheetview.addEventListener('input', e=>{ if (e.target.matches('textarea.s-in')) autosize(e.target); });
sheetview.addEventListener('focusout', e=>{
  const el = e.target;
  if (el.matches('textarea.s-in')) el.style.height = '';
  if (sheetPending) setTimeout(()=>{ const a = document.activeElement; if (sheetPending && !(a && sheetview.contains(a) && a.matches('input,textarea'))){ sheetPending = false; render(); } }, 0);
});
sheetview.addEventListener('keydown', e=>{
  const el = e.target; if (!el.matches('.s-in') || e.isComposing) return;
  if (e.key === 'Escape'){ e.preventDefault(); e.stopPropagation(); if (sheetFocusVal != null) el.value = sheetFocusVal; el.blur(); return; }
  if (e.key === 'Enter' && !(el.tagName === 'TEXTAREA' && (e.shiftKey || e.altKey))){
    // Enter saves and moves down like a spreadsheet (Shift+Enter makes a new line in the description)
    e.preventDefault();
    const tr = el.closest('tr'), f = el.dataset.sf, dir = e.shiftKey && el.tagName !== 'TEXTAREA' ? 'previousElementSibling' : 'nextElementSibling';
    sheetCommit(el); sheetFocusVal = el.value;
    let n = tr[dir]; while (n && !n.dataset.id) n = n[dir];
    const next = n && $(`.s-in[data-sf="${f}"]`, n);
    if (next){ next.focus(); if (next.select && next.type !== 'date') next.select(); } else el.blur();
  }
});
sheetview.addEventListener('submit', e=>{
  e.preventDefault();
  const form = e.target; if (form.dataset.form !== 'sheet-add' || !canEdit()) return;
  const inp = $('input', form), title = inp.value.trim(); if (!title) return;
  const b = isAll() ? (boardById(P.lastBoard?.[wsId]) || S.boards[0]) : curBoard(); if (!b) return;
  inp.value = '';
  const i = createIssue({title, boardId:b.id, status:b.columns[0].id, type:S.types[0]?.id, labels:[normLabel(b.name)]});
  toast(isAll() ? t('toast.createdIn', {key:key(i), board:b.name}) : t('toast.created', {key:key(i)}), {action:t('btn.view'), fn:()=>openIssue(i.id)});
  requestAnimationFrame(()=>{ $('[data-k="sheet-add"]', sheetview)?.focus({preventScroll:true}); const tr = $(`tr[data-id="${i.id}"]`, sheetview); tr?.scrollIntoView({block:'nearest'}); tr?.classList.add('s-new'); });
});

/* ================= Calendar view =================
   Tasks sit on their due date (דד ליין). Drag one to another day to move its due date, or from
   "no due date" onto a day to give it one. On phones the month shows dots; tap a day to list it. */
const calview = $('#calview');
let calAnchor = null, calSel = null, calDrag = null, calSuppress = false, calCap = {all:3, more:2};
const isoOf = d=>iso(d);
const addD = (isoStr, n)=>{ const d = new Date(isoStr+'T12:00'); d.setDate(d.getDate()+n); return iso(d); };
function calRange(){
  const mode = P.calMode === 'week' ? 'week' : 'month', a = new Date((calAnchor || today())+'T12:00');
  if (mode === 'week'){ const st = new Date(a); st.setDate(a.getDate() - a.getDay()); return {mode, start:iso(st), days:7, month:a.getMonth()}; }
  const first = new Date(a.getFullYear(), a.getMonth(), 1, 12), st = new Date(first); st.setDate(1 - first.getDay());
  const dim = new Date(a.getFullYear(), a.getMonth()+1, 0).getDate(), weeks = Math.ceil((first.getDay() + dim)/7);
  return {mode, start:iso(st), days:weeks*7, month:a.getMonth(), year:a.getFullYear()};
}
function calTitle(r){
  if (r.mode === 'month') return fmtr('my', {month:'long', year:'numeric'}).format(new Date(r.year, r.month, 1, 12));
  const a = new Date(r.start+'T12:00'), z = new Date(addD(r.start, 6)+'T12:00');
  try { return fmtr('range', {day:'numeric', month:'short', year:'numeric'}).formatRange(a, z); } catch(e){ return fmtDay(r.start) + ' – ' + fmtDay(addD(r.start, 6)); }
}
function calChip(i, cls=''){
  const done = isDone(i), over = dueClass(i) === 'over', c = colOf(i);
  const tip = `${key(i)} · ${i.title || t('issue.untitled')} · ${c?.name || ''}`;
  return `<button class="cal-chip${done?' done':''}${over?' over':''} ${cls}" data-act="open-issue" data-id="${i.id}" title="${esc(tip)}" style="--cc:${i.color || typeOf(i.type).color}">${isAll() ? bAv(boardOf(i),'xs') : ''}<span class="cc-t" data-bidi>${esc(i.title || t('issue.untitled'))}</span></button>`;
}
function calSort(a, b){ return (isDone(a) - isDone(b)) || (PRANK[a.priority] - PRANK[b.priority]) || a.title.localeCompare(b.title); }
let calPending = false;
function renderCal(){
  if (calDrag?.active){ calPending = true; return; }
  calPending = false;
  const all = isAll(), b = curBoard(), edit = canEdit();
  if (!b){ calview.innerHTML = `<div class="empty-state"><b>${esc(t('empty.noBoards'))}</b></div>`; return; }
  if (!calAnchor) calAnchor = today();
  const r = calRange(), pool = scopeIssues().filter(matches), byDay = new Map();
  for (const i of pool) if (i.due){ let l = byDay.get(i.due); if (!l) byDay.set(i.due, l = []); l.push(i); }
  byDay.forEach(l=>l.sort(calSort));
  const undated = pool.filter(i=>!i.due && !isDone(i)).sort(calSort);
  const tday = today(), mobile = isMobile();
  if (!calSel || calSel < r.start || calSel > addD(r.start, r.days-1)) calSel = tday >= r.start && tday <= addD(r.start, r.days-1) ? tday : (r.mode === 'month' ? iso(new Date(r.year, r.month, 1, 12)) : r.start);
  const wd = fmtr('wd', {weekday:'short'}), wdl = fmtr('wdl', {weekday:'long', day:'numeric', month:'long'});
  const heads = Array.from({length:7}, (_,k)=>`<div class="cal-wd">${esc(wd.format(new Date(addD(r.start,k)+'T12:00')))}</div>`).join('');
  const capOf = n=>r.mode !== 'month' ? 999 : n <= calCap.all ? n : Math.max(1, calCap.more);
  const cells = Array.from({length:r.days}, (_,k)=>{
    const d = addD(r.start, k), dt = new Date(d+'T12:00'), list = byDay.get(d) || [], out = r.mode === 'month' && dt.getMonth() !== r.month;
    const open = list.filter(i=>!isDone(i)).length, overN = list.filter(i=>dueClass(i)==='over').length;
    const dots = list.slice(0,4).map(i=>`<i style="--cc:${i.color || typeOf(i.type).color}"${isDone(i)?' class="done"':''}></i>`).join('') + (list.length > 4 ? '<b>+</b>' : '');
    return `<div class="cal-day${out?' out':''}${d===tday?' today':''}${d<tday?' past':''}${d===calSel?' sel':''}${overN?' has-over':''}" data-day="${d}" role="gridcell" aria-label="${esc(wdl.format(dt))}${list.length ? ', ' + esc(tn('issues', list.length)) : ''}">
      <div class="cd-head"><button class="cd-num" data-act="cal-sel" data-day="${d}" title="${esc(wdl.format(dt))}">${r.mode === 'week' ? `<span class="cd-wd">${esc(wd.format(dt))}</span>` : ''}${dt.getDate()}</button>
        ${edit ? `<button class="cd-add" data-act="cal-add" data-day="${d}" aria-label="${esc(t('cal.addOn', {date:fmtDay(d)}))}" title="${esc(t('cal.addOn', {date:fmtDay(d)}))}">${I.plus}</button>` : ''}</div>
      <div class="cd-list">${(cap=>list.slice(0,cap).map(i=>calChip(i)).join('') + (list.length > cap ? `<button class="cd-more" data-act="cal-day" data-day="${d}">${esc(tn('cal.more', list.length - cap))}</button>` : ''))(capOf(list.length))}</div>
      <div class="cd-dots" aria-hidden="true">${dots}</div>
    </div>`;
  }).join('');
  const selList = byDay.get(calSel) || [];
  const dayPanel = `<section class="cal-daylist" aria-live="polite"><div class="cdl-head"><b>${esc(wdl.format(new Date(calSel+'T12:00')))}</b>${edit ? `<button class="btn sm" data-act="cal-add" data-day="${calSel}">${I.plus}${esc(t('cal.add'))}</button>` : ''}</div>
    ${selList.length ? `<div class="cdl-list">${selList.map(i=>calChip(i,'row')).join('')}</div>` : `<p class="muted-p">${esc(t('cal.none'))}</p>`}</section>`;
  const und = `<aside class="cal-undated${P.calUndated ? ' open' : ''}" data-day="">
      <div class="cu-head"><b>${esc(t('cal.undated'))}</b><span class="sec-meta">${undated.length}</span></div>
      ${edit && undated.length ? `<p class="cu-hint">${esc(mobile ? t('cal.undatedHintTouch') : t('cal.undatedHint'))}</p>` : ''}
      <div class="cu-list">${undated.length ? undated.map(i=>calChip(i,'row')).join('') : `<p class="muted-p">${esc(t('cal.undatedNone'))}</p>`}</div></aside>`;
  calview.innerHTML = `<div class="cal-bar">
      <div class="cal-nav"><button class="btn sm" data-act="cal-today">${esc(t('cal.today'))}</button>
        <button class="btn icon subtle sm cal-prev" data-act="cal-step" data-n="-1" aria-label="${esc(t('cal.prev'))}" title="${esc(t('cal.prev'))}">${I.chev}</button>
        <button class="btn icon subtle sm cal-next" data-act="cal-step" data-n="1" aria-label="${esc(t('cal.next'))}" title="${esc(t('cal.next'))}">${I.chev}</button>
        <h2 class="cal-title">${esc(calTitle(r))}</h2><span class="cal-hint">${I.cal}${esc(t('cal.byDue'))}</span></div>
      <div class="cal-tools"><div class="seg" role="group"><button data-act="cal-mode" data-m="month" aria-pressed="${r.mode==='month'}">${esc(t('cal.month'))}</button><button data-act="cal-mode" data-m="week" aria-pressed="${r.mode==='week'}">${esc(t('cal.week'))}</button></div>
        <button class="btn sm" data-act="cal-ics" title="${esc(t('cal.icsTip'))}">${I.down}${esc(t('cal.ics'))}</button>
        <button class="btn sm cal-und-btn" data-act="cal-undated" aria-pressed="${!!P.calUndated}">${esc(t('cal.undated'))}<span class="sec-meta">${undated.length}</span></button></div>
    </div>
    <div class="cal-body${P.calUndated ? ' with-und' : ''}">
      <div class="cal-main"><div class="cal-grid ${r.mode}" role="grid"><div class="cal-wds">${heads}</div><div class="cal-days" style="--weeks:${r.days/7}">${cells}</div></div>${dayPanel}</div>
      ${und}
    </div>`;
  // fit as many tasks per day as the month cells have room for
  if (r.mode === 'month' && !mobile){
    const cell = $('.cal-day', calview), head = cell && $('.cd-head', cell);
    if (cell){ const room = cell.clientHeight - head.offsetHeight - 10, fit = {all:Math.max(1, Math.floor((room + 2) / 24)), more:Math.max(1, Math.floor((room - 14) / 24))};
      if (fit.all !== calCap.all || fit.more !== calCap.more){ calCap = fit; renderCal(); } }
  }
}
let calResize = 0;
addEventListener('resize', ()=>{ clearTimeout(calResize); calResize = setTimeout(()=>{ if (screen === 'app' && P.view === 'cal' && !calDrag) renderCal(); }, 150); });
/* ----- calendar file (.ics) for Outlook / Google / Apple calendars -----
   Open tasks with a due date, as all-day events. Stable UIDs, so a calendar app that re-imports
   the file can match the same task instead of adding a copy. */
function icsText(s){ return String(s ?? '').replace(/\\/g,'\\\\').replace(/;/g,'\;').replace(/,/g,'\\,').replace(/\r?\n/g,'\\n'); }
function icsFold(line){
  // lines are limited to 75 bytes (UTF-8); continue with CRLF + space, never splitting a character
  const enc = new TextEncoder(); if (enc.encode(line).length <= 75) return line;
  const out = []; let cur = '', n = 0, lim = 75;
  for (const ch of line){ const b = enc.encode(ch).length; if (n + b > lim){ out.push(cur); cur = ' '; n = 1; lim = 75; } cur += ch; n += b; }
  out.push(cur); return out.join('\r\n');
}
function exportIcs(){
  const list = scopeIssues().filter(matches).filter(i=>i.due && !isDone(i)).sort((a,b)=>a.due.localeCompare(b.due));
  if (!list.length) return toast(t('cal.icsNone'));
  const stamp = new Date().toISOString().replace(/[-:]/g,'').replace(/\.\d+/,'');
  const d8 = s=>s.replace(/-/g,'');
  const calName = `Workbench – ${isAll() ? (ws?.name || '') : (curBoard()?.name || '')}`;
  const L = ['BEGIN:VCALENDAR','VERSION:2.0','PRODID:-//McCann Valley//Workbench//HE','CALSCALE:GREGORIAN','METHOD:PUBLISH', 'X-WR-CALNAME:' + icsText(calName)];
  for (const i of list){
    const b = boardOf(i), c = colOf(i);
    const desc = [`${t('x.col.board')}: ${b?.name || ''}`, `${t('x.col.status')}: ${c?.name || ''}`, `${t('x.col.prio')}: ${prioName(i.priority)}`,
      i.color ? `${t('x.col.plat')}: ${platLabelOf(i.color)}` : '', i.desc ? '\n' + i.desc.slice(0, 1500) : ''].filter(Boolean).join('\n');
    L.push('BEGIN:VEVENT', `UID:${i.id}@workbench.mccann-valley`, `DTSTAMP:${stamp}`, `DTSTART;VALUE=DATE:${d8(i.due)}`, `DTEND;VALUE=DATE:${d8(addD(i.due, 1))}`,
      `SEQUENCE:${Math.max(0, Math.floor((i.updated - 1.7e12) / 1000))}`, 'SUMMARY:' + icsText(`${key(i)} · ${i.title || t('issue.untitled')}`),
      'DESCRIPTION:' + icsText(desc), 'CATEGORIES:' + icsText(b?.name || 'Workbench'), 'TRANSP:TRANSPARENT', 'END:VEVENT');
  }
  L.push('END:VCALENDAR');
  const blob = new Blob([L.map(icsFold).join('\r\n') + '\r\n'], {type:'text/calendar;charset=utf-8'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `workbench-calendar-${today()}.ics`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 1500);
  toast(tn('cal.icsDone', list.length), {ms:9000});
}
function calMoveDue(i, day){
  if (!canEdit()) return;
  const v = isISO(day) ? day : '';
  if (v === i.due) return;
  i.due = v; touch(i, v ? t('h.due', {date:fmtDay(v)}) : t('h.dueNone'));
  fx.land = i.id; save(); render();
  toast(v ? t('cal.moved', {key:key(i), date:fmtDay(v)}) : t('cal.cleared', {key:key(i)}), {action:t('btn.view'), fn:()=>openIssue(i.id)});
}
// drag a chip to a day (mouse: just drag; touch: press and hold)
calview.addEventListener('pointerdown', e=>{
  if (e.button !== 0 || !canEdit()) return;
  const chip = e.target.closest('.cal-chip'); if (!chip) return;
  calDrag = {chip, id:chip.dataset.id, sx:e.clientX, sy:e.clientY, x:e.clientX, y:e.clientY, touch:e.pointerType !== 'mouse', active:false, timer:0};
  if (calDrag.touch) calDrag.timer = setTimeout(()=>{ if (calDrag && !calDrag.active) calBegin(); }, 300);
});
function calBegin(){
  const d = calDrag;
  if (!d.chip.isConnected){ const c = $(`.cal-chip[data-id="${CSS.escape(d.id)}"]`, calview); if (!c){ calDrag = null; return; } d.chip = c; }
  const r = d.chip.getBoundingClientRect(); if (!r.width){ calDrag = null; return; }
  d.active = true; d.ox = d.sx - r.left; d.oy = d.sy - r.top;
  if (d.touch && navigator.vibrate) navigator.vibrate(10);
  const g = d.chip.cloneNode(true); g.classList.add('cal-ghost'); g.style.width = r.width+'px'; g.style.transform = `translate(${r.left}px,${r.top}px)`;
  document.body.appendChild(g); d.ghost = g; d.chip.classList.add('dragging-src'); document.body.classList.add('dragging');
}
window.addEventListener('pointermove', e=>{
  const d = calDrag; if (!d) return;
  d.x = e.clientX; d.y = e.clientY;
  if (!d.active){ const dist = Math.hypot(d.x-d.sx, d.y-d.sy); if (d.touch){ if (dist > 8){ clearTimeout(d.timer); calDrag = null; } } else if (dist > 5) calBegin(); return; }
  d.ghost.style.transform = `translate(${d.x-d.ox}px,${d.y-d.oy}px)`;
  const el = document.elementFromPoint(d.x, d.y), tgt = el?.closest?.('#calview [data-day]');
  if (tgt !== d.tgt){ d.tgt?.classList.remove('drop-on'); d.tgt = tgt; tgt?.classList.add('drop-on'); }
});
document.addEventListener('touchmove', e=>{ if (calDrag?.active) e.preventDefault(); }, {passive:false});
function calEnd(cancel){
  const d = calDrag; if (!d) return; calDrag = null; clearTimeout(d.timer);
  if (!d.active) return;
  calSuppress = true; setTimeout(()=>{ calSuppress = false; }, 80);
  d.ghost.remove(); d.chip.classList.remove('dragging-src'); document.body.classList.remove('dragging'); d.tgt?.classList.remove('drop-on');
  if (cancel || !d.tgt){ if (calPending) renderCal(); return; }
  const i = get(d.id); if (i) calMoveDue(i, d.tgt.dataset.day); if (calPending) renderCal();
}
window.addEventListener('pointerup', ()=>calEnd(false));
window.addEventListener('pointercancel', ()=>calEnd(true));
calview.addEventListener('click', e=>{ if (calSuppress){ e.stopPropagation(); e.preventDefault(); return; }
  const day = e.target.closest('.cal-day'); if (day && !e.target.closest('button') && isMobile()){ calSel = day.dataset.day; renderCal(); } }, true);
calview.addEventListener('contextmenu', e=>{ if (e.target.closest('.cal-chip')) e.preventDefault(); });

/* ================= Dashboard (per client, or every client side by side) =================
   A set of widgets, each of which can be switched off (P.dashHide). The layout adapts to what is shown:
   cards flow in rows and stretch to fill them, and with only a few charts on screen the charts get taller.
   Archived tasks count for history (completed, on time, speed). */
const dashview = $('#dashview');
const DASH_RANGES = [4, 8, 12, 26];
// order = order on screen. kind: 'kpi' (the number tiles), 'half' (a time chart), 'wide' (full row), '' (a regular card)
const DASH_W = [
  {id:'kpis', kind:'kpi'}, {id:'weekly', kind:'half'}, {id:'backlog', kind:'half'},
  {id:'clients', all:true}, {id:'compare', kind:'wide', all:true},
  {id:'deadlines'}, {id:'cycle'}, {id:'weekday'},
  {id:'over'}, {id:'soon'}, {id:'status'}, {id:'plat'}, {id:'type'}, {id:'prio'}];
const dashHidden = ()=>new Set(Array.isArray(P.dashHide) ? P.dashHide : []);
function weekStart(isoDay){ const d = new Date(isoDay+'T12:00'); d.setDate(d.getDate() - d.getDay()); return iso(d); }
const dayOf = ms=>iso(new Date(ms));
// when a task stopped being open: finished, or archived without finishing (null = still open)
const closedAt = i=>i.doneAt || (isDone(i) ? i.updated : null) || (i.archived ? i.archivedAt : null);
function dashStats(list, weeks){
  const tday = today(), from = isoAdd(weekStart(tday), -7*(weeks-1)), live = list.filter(i=>!i.archived);
  const open = live.filter(i=>!isDone(i));
  const doneIn = list.filter(i=>i.doneAt && dayOf(i.doneAt) >= from);
  const withDue = doneIn.filter(i=>i.due);
  const onTime = withDue.filter(i=>dayOf(i.doneAt) <= i.due).length;
  const spans = doneIn.map(i=>daysBetween(i.start, dayOf(i.doneAt))).filter(n=>n >= 0);
  return {open, over:open.filter(i=>i.due && i.due < tday), soon:open.filter(i=>i.due && i.due >= tday && i.due <= addDays(7)),
    doneIn, done:doneIn.length, onTime:withDue.length ? Math.round(onTime / withDue.length * 100) : null, onTimeN:withDue.length,
    spans, avg:spans.length ? Math.round(spans.reduce((a, b)=>a + b, 0) / spans.length * 10) / 10 : null,
    created:list.filter(i=>dayOf(i.created) >= from).length, from};
}
const dashWeeks = weeks=>Array.from({length:weeks}, (_, k)=>isoAdd(weekStart(today()), -7*(weeks-1-k)));
// a round top for the y axis, so the middle gridline is a whole number too
const niceTop = max=>max <= 2 ? 2 : max <= 4 ? 4 : Math.ceil(max / 4) * 4;
const tile = (label, value, sub, cls='')=>`<div class="kpi ${cls}"><span class="kpi-l">${esc(label)}</span><b class="kpi-v">${value}</b>${sub ? `<span class="kpi-s">${sub}</span>` : ''}</div>`;
const legend = items=>`<div class="viz-legend">${items.map(([cls, label])=>`<span><i class="sw ${cls}"></i>${esc(label)}</span>`).join('')}</div>`;
function card(w, title, body, o={}){
  return `<section class="viz-card${w.kind ? ' ' + w.kind : ''}" data-w="${w.id}"><div class="viz-head"><h3>${esc(title)}${o.meta != null ? ` <span class="sec-meta">${o.meta}</span>` : ''}</h3>${o.legend || ''}
    <button class="btn icon subtle sm viz-x" data-act="dash-hide" data-w="${w.id}" aria-label="${esc(t('dash.hideIt', {name:title}))}" title="${esc(t('dash.hide'))}">${I.eyeOff}</button></div>${body}</section>`;
}
const yGrid = top=>`<div class="wk-grid">${[top, top/2, 0].map(v=>`<span style="bottom:${v / top * 100}%"><em>${v}</em></span>`).join('')}</div>`;
// x labels are fitted to the space after drawing: dates thin out (always keeping the current week),
// categories switch to their short name when the long one doesn't fit
let xCtx = null;
function fitX(){
  $$('.wk', dashview).forEach(wk=>{
    const xs = $$('.wk-x', wk); if (!xs.length) return;
    const colW = xs[0].parentElement.getBoundingClientRect().width; if (!colW) return;
    xCtx ||= document.createElement('canvas').getContext('2d');
    const cs = getComputedStyle(xs[0]); xCtx.font = `${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    const wOf = s=>xCtx.measureText(s).width;
    if (xs[0].hasAttribute('data-time')){
      const need = Math.max(...xs.map(x=>wOf(x.dataset.x))) + 10, n = xs.length;
      const step = [1,2,3,4,6,8,13,26].find(s=>s * colW >= need) || n;
      xs.forEach((x, k)=>{ x.textContent = (n - 1 - k) % step === 0 ? x.dataset.x : ''; });
    } else {
      const long = Math.max(...xs.map(x=>wOf(x.dataset.x))) + 6 <= colW;
      xs.forEach(x=>{ x.textContent = long ? x.dataset.x : x.dataset.xs; });
    }
  });
}
new ResizeObserver(()=>{ if (!dashview.hidden) requestAnimationFrame(fitX); }).observe(dashview);
function hbars(w, rows, title, empty){
  // one series → one colour; the value is written at the end of each bar
  const max = Math.max(1, ...rows.map(r=>r.n));
  return card(w, title, rows.length ? `<div class="hb">${rows.map(r=>`<div class="hb-row" data-tip="${esc(r.label)}: ${r.n}">
      <span class="hb-l" data-bidi>${r.ico || ''}${esc(r.label)}</span><span class="hb-track"><span class="hb-bar" style="width:${Math.max(2, r.n / max * 100)}%"></span></span><span class="hb-n">${r.n}</span></div>`).join('')}</div>` : `<p class="muted-p">${esc(empty)}</p>`);
}
// vertical bars, one series (a bar may be highlighted with .s2, explained in the legend)
function colChart(w, title, cols, o={}){
  const top = niceTop(Math.max(1, ...cols.map(c=>c.n))), n = cols.length;
  const body = cols.some(c=>c.n) ? `<div class="wk one" style="--n:${n}">${yGrid(top)}
      <div class="wk-cols">${cols.map(c=>`<div class="wk-col" data-tip="${esc(c.tip || c.label)}: ${c.n}">
        <div class="wk-bars"><span class="b ${c.cls || 's1'}" style="height:${c.n / top * 100}%"></span></div><span class="wk-x" data-x="${esc(c.label)}" data-xs="${esc(c.short || c.label)}" title="${esc(c.label)}"></span></div>`).join('')}</div></div>`
    : `<p class="muted-p">${esc(o.empty)}</p>`;
  return card(w, title, body, {legend:o.legend, meta:o.meta});
}
function weekChart(w, list, weeks){
  const ws = dashWeeks(weeks), idx = d=>ws.indexOf(weekStart(d));
  const made = ws.map(()=>0), done = ws.map(()=>0);
  list.forEach(i=>{ const a = idx(dayOf(i.created)); if (a >= 0) made[a]++; if (i.doneAt){ const c = idx(dayOf(i.doneAt)); if (c >= 0) done[c]++; } });
  const top = niceTop(Math.max(1, ...made, ...done));
  return card(w, t('dash.weekly'), `<div class="wk" style="--n:${weeks}">${yGrid(top)}
      <div class="wk-cols">${ws.map((d, k)=>`<div class="wk-col" data-tip="${esc(t('dash.weekOf', {date:fmtDay(d)}))} · ${esc(t('dash.created'))}: ${made[k]} · ${esc(t('dash.completed'))}: ${done[k]}">
        <div class="wk-bars"><span class="b s1" style="height:${made[k] / top * 100}%"></span><span class="b s2" style="height:${done[k] / top * 100}%"></span></div>
        <span class="wk-x" data-x="${esc(fmtDay(d))}" data-time></span></div>`).join('')}</div>
    </div>`, {legend:legend([['s1', t('dash.created')], ['s2', t('dash.completed')]])});
}
// how many tasks were open (and how many of those overdue) at the end of each week — is the pile growing?
function backlogChart(w, list, weeks){
  const ws = dashWeeks(weeks), now = Date.now(), tday = today();
  const pts = ws.map(d=>{
    const endDay = isoAdd(d, 6), T = endDay >= tday ? now : new Date(endDay+'T23:59:59').getTime(), dayT = endDay >= tday ? tday : endDay;
    let open = 0, over = 0;
    list.forEach(i=>{ if (i.created > T) return; const c = closedAt(i); if (c && c <= T) return; open++; if (i.due && i.due < dayT) over++; });
    return {d, open, over};
  });
  const top = niceTop(Math.max(1, ...pts.map(p=>p.open))), n = ws.length, r = rtl();
  const x = k=>{ const v = (k + .5) / n * 100; return (r ? 100 - v : v).toFixed(2); };
  const line = key=>pts.map((p, k)=>`${x(k)},${(100 - p[key] / top * 100).toFixed(2)}`).join(' ');
  return card(w, t('dash.backlog'), `<div class="wk ln" style="--n:${n}">${yGrid(top)}
      <svg class="ln-svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><polyline class="ln s1" points="${line('open')}" vector-effect="non-scaling-stroke"/><polyline class="ln s2" points="${line('over')}" vector-effect="non-scaling-stroke"/></svg>
      <div class="wk-cols">${pts.map((p, k)=>`<div class="wk-col" data-tip="${esc(k === n - 1 ? t('dash.now') : t('dash.endOf', {date:fmtDay(isoAdd(p.d, 6))}))} · ${esc(t('dash.open'))}: ${p.open} · ${esc(t('dash.over'))}: ${p.over}">
        <div class="ln-area"><span class="pt s1" style="bottom:${p.open / top * 100}%"></span><span class="pt s2" style="bottom:${p.over / top * 100}%"></span></div>
        <span class="wk-x" data-x="${esc(fmtDay(p.d))}" data-time></span></div>`).join('')}</div>
    </div>`, {legend:legend([['s1', t('dash.open')], ['s2', t('dash.over')]])});
}
// open work per client, the overdue part shown separately
function clientsChart(w, rows){
  const max = Math.max(1, ...rows.map(r=>r.open)), tday = today();
  return card(w, t('dash.clients'), rows.some(r=>r.open) ? `<div class="hb">${rows.map(r=>`<div class="hb-row" data-tip="${esc(r.b.name)} · ${esc(t('dash.open'))}: ${r.open} · ${esc(t('dash.over'))}: ${r.over}">
      <button class="hb-l link-btn" data-act="switch-board" data-id="${r.b.id}" data-bidi>${bAv(r.b, 'xs')}<span>${esc(r.b.name)}</span></button>
      <span class="hb-track st">${r.open - r.over ? `<span class="hb-bar" style="width:${(r.open - r.over) / max * 100}%"></span>` : ''}${r.over ? `<span class="hb-bar s2" style="width:${r.over / max * 100}%"></span>` : ''}</span><span class="hb-n">${r.open}</span></div>`).join('')}</div>` : `<p class="muted-p">${esc(t('dash.noOpen'))}</p>`,
    {legend:legend([['s1', t('dash.onTrack')], ['s2', t('dash.over')]])});
}
function taskList(w, list, title, empty){
  return card(w, title, list.length ? `<ul class="dl">${list.slice(0, 8).map(i=>`<li><button class="dl-i" data-act="open-issue" data-id="${i.id}">${typeIcon(i.type)}<span class="dl-t" data-bidi>${esc(i.title || t('issue.untitled'))}</span>${isAll() ? `<span class="dl-b" data-bidi>${esc(boardOf(i)?.name || '')}</span>` : ''}<span class="due ${dueClass(i)}">${esc(dueLabel(i))}</span></button></li>`).join('')}</ul>${list.length > 8 ? `<p class="muted-p">${esc(tn('dash.more', list.length - 8))}</p>` : ''}` : `<p class="muted-p">${esc(empty)}</p>`, {meta:list.length});
}
function dashCustomMenu(){
  const hid = dashHidden(), all = isAll();
  return `<div class="menu-h">${esc(t('dash.showWhat'))}</div><div class="sf-list dash-pick">${DASH_W.map(w=>{ const on = !hid.has(w.id), off = w.all && !all;
      return `<button class="menu-i" role="menuitemcheckbox" aria-checked="${on}" data-act="dash-toggle" data-w="${w.id}"${off ? ' aria-describedby="dash-allnote"' : ''}><span class="sf-box${on ? ' on' : ''}">${on ? I.check : ''}</span><span class="sf-l">${esc(t('dash.w.' + w.id))}${off ? `<small class="dash-only">${esc(t('dash.allOnly'))}</small>` : ''}</span></button>`; }).join('')}</div>
    <div class="menu-sep"></div><div class="dash-pick-f"><button class="link-btn" data-act="dash-showall">${esc(t('dash.showAll'))}</button><button class="link-btn" data-act="dash-hideall">${esc(t('dash.hideAll'))}</button></div>`;
}
function dashRefresh(){
  renderDash();
  const btn = $('[data-act="dash-custom"]', dashview);
  if (popAnchor && pop.dataset.id === 'dash' && btn){ popAnchor = btn; btn.setAttribute('aria-expanded', 'true'); const sc = $('.sf-list', pop)?.scrollTop || 0; pop.innerHTML = dashCustomMenu(); $('.sf-list', pop).scrollTop = sc; placePop(); }
}
function renderDash(){
  const all = isAll(), b = curBoard();
  if (!b){ dashview.innerHTML = `<div class="empty-state"><b>${esc(t('empty.noBoards'))}</b></div>`; return; }
  const weeks = DASH_RANGES.includes(P.dashWeeks) ? P.dashWeeks : 8, hid = dashHidden();
  const W = DASH_W.filter(w=>!hid.has(w.id) && (!w.all || all)), on = new Set(W.map(w=>w.id));
  const list = S.issues.filter(i=>(all || i.boardId === b.id) && boardOf(i)).filter(matches);
  const st = dashStats(list, weeks);
  const byDue = (a, c)=>(a.due || '').localeCompare(c.due || '');
  const count = (arr, key)=>{ const m = new Map(); arr.forEach(i=>{ const k = key(i); m.set(k, (m.get(k) || 0) + 1); }); return m; };
  const boards = ()=>S.boards.slice().sort((a, c)=>(a.order ?? 0) - (c.order ?? 0));
  const tday = today(), wk0 = weekStart(tday);
  const make = {
    weekly:w=>weekChart(w, list, weeks),
    backlog:w=>backlogChart(w, list, weeks),
    clients:w=>clientsChart(w, boards().map(x=>{ const o = st.open.filter(i=>i.boardId === x.id); return {b:x, open:o.length, over:o.filter(i=>i.due && i.due < tday).length}; }).sort((a, c)=>c.open - a.open)),
    compare:w=>{
      const rows = boards().map(x=>({b:x, s:dashStats(list.filter(i=>i.boardId === x.id), weeks)}));
      const maxOpen = Math.max(1, ...rows.map(r=>r.s.open.length));
      return card(w, t('dash.compare'), `<div class="cmp-wrap"><table class="cmp">
        <thead><tr><th>${esc(t('x.col.board'))}</th><th>${esc(t('dash.open'))}</th><th>${esc(t('dash.over'))}</th><th>${esc(t('dash.soon'))}</th><th>${esc(t('dash.done'))}</th><th>${esc(t('dash.onTime'))}</th><th>${esc(t('dash.avg'))}</th></tr></thead>
        <tbody>${rows.map(({b:x, s:q})=>`<tr><td><button class="link-btn cmp-b" data-act="switch-board" data-id="${x.id}">${bAv(x,'xs')}<span data-bidi>${esc(x.name)}</span></button></td>
          <td><span class="cmp-n"><span class="cmp-bar" style="width:${q.open.length / maxOpen * 100}%"></span><b>${q.open.length}</b></span></td>
          <td class="${q.over.length ? 'is-over' : ''}">${q.over.length ? `${I.clock}${q.over.length}` : '0'}</td><td>${q.soon.length}</td><td>${q.done}</td>
          <td>${q.onTime == null ? '—' : q.onTime + '%'}</td><td>${q.avg == null ? '—' : esc(tn('dash.days', q.avg))}</td></tr>`).join('')}</tbody></table></div>`);
    },
    deadlines:w=>{
      // open tasks by when they are due: overdue, this week, the next weeks, later, no date
      const c = [0,0,0,0,0,0,0];
      st.open.forEach(i=>{ if (!i.due){ c[6]++; return; } if (i.due < tday){ c[0]++; return; } const k = Math.round(daysBetween(wk0, weekStart(i.due)) / 7); c[Math.min(5, 1 + k)]++; });
      const keys = ['over','thisWeek','nextWeek','in2','in3','later','noDate'];
      return colChart(w, t('dash.deadlines'), keys.map((k, ix)=>({label:t('dash.dl.' + k), short:t('dash.dls.' + k), n:c[ix], cls:ix === 0 ? 's2' : ix === 6 ? 's0' : 's1'})), {empty:t('dash.noOpen'), meta:st.open.length});
    },
    cycle:w=>{
      // how long finished tasks took, start → done
      const B = [[1,'d1'],[3,'d3'],[7,'d7'],[14,'d14'],[30,'d30'],[Infinity,'dMore']], c = B.map(()=>0);
      st.spans.forEach(n=>{ c[B.findIndex(([m])=>n <= m)]++; });
      return colChart(w, t('dash.cycle'), B.map(([, k], ix)=>({label:t('dash.cy.' + k), short:t('dash.cys.' + k), n:c[ix]})), {empty:t('dash.noneDone'), meta:st.spans.length});
    },
    weekday:w=>{
      const c = [0,0,0,0,0,0,0]; st.doneIn.forEach(i=>{ c[new Date(i.doneAt).getDay()]++; });
      const f = fmtr('wds', {weekday:'short'}), fn = fmtr('wdn', {weekday:'narrow'}), fl = fmtr('wdl1', {weekday:'long'}), d0 = new Date(2026, 0, 4, 12); // a Sunday
      return colChart(w, t('dash.weekday'), c.map((n, k)=>{ const d = new Date(d0); d.setDate(d0.getDate() + k); return {label:f.format(d), short:fn.format(d), tip:fl.format(d), n}; }), {empty:t('dash.noneDone'), meta:st.done});
    },
    over:w=>taskList(w, st.over.slice().sort(byDue), t('dash.overList'), t('dash.overNone')),
    soon:w=>taskList(w, st.soon.slice().sort(byDue), t('dash.soonList'), t('dash.soonNone')),
    status:w=>{ const cols = []; (all ? S.boards : [b]).forEach(x=>x.columns.forEach(c=>{ if (c.cat !== 'done' && !cols.includes(c.name)) cols.push(c.name); }));
      const by = count(st.open, i=>colOf(i)?.name || ''); return hbars(w, cols.map(n=>({label:n, n:by.get(n) || 0})).filter(r=>r.n), t('dash.byStatus'), t('dash.noOpen')); },
    plat:w=>{ const by = count(st.open, i=>i.color || '');
      return hbars(w, [...by].sort((a, c)=>c[1] - a[1]).map(([h, n])=>({label:h ? platLabelOf(h) : t('filter.noPlatform'), n, ico:`<i class="pf-dot${h ? '' : ' none'}"${h ? ` style="--sw:${h}"` : ''}></i>`})), t('dash.byPlat'), t('dash.noOpen')); },
    type:w=>{ const by = count(st.open, i=>i.type);
      return hbars(w, S.types.map(x=>({label:x.name, n:by.get(x.id) || 0, ico:typeSVG(x)})).filter(r=>r.n).sort((a, c)=>c.n - a.n), t('dash.byType'), t('dash.noOpen')); },
    prio:w=>{ const by = count(st.open, i=>i.priority);
      return hbars(w, PRIOS.map(p=>({label:prioName(p), n:by.get(p) || 0, ico:prioIcon(p)})).filter(r=>r.n), t('dash.byPrio'), t('dash.noOpen')); },
  };
  const rangeSel = `<div class="seg" role="group" aria-label="${esc(t('dash.range'))}">${DASH_RANGES.map(w=>`<button data-act="dash-range" data-w="${w}" aria-pressed="${w === weeks}">${esc(tn('dash.weeks', w))}</button>`).join('')}</div>`;
  const shownN = W.length, hiddenN = DASH_W.filter(w=>!w.all || all).length - shownN;
  const custom = `<button class="btn subtle sm dash-cust" data-act="dash-custom" aria-haspopup="true" aria-expanded="false">${I.sliders}<span>${esc(t('dash.custom'))}</span>${hiddenN ? `<span class="dash-hid">${esc(tn('dash.hiddenN', hiddenN))}</span>` : ''}</button>`;
  const charts = W.filter(w=>w.kind !== 'kpi');
  // fewer charts on screen → more room for each; two time charts side by side only when both are on
  const size = charts.length <= 2 ? 'roomy' : charts.length <= 5 ? 'mid' : '';
  dashview.className = 'dashview' + (size ? ' ' + size : '');
  dashview.innerHTML = `<div class="dash-bar"><span class="dash-scope">${esc(t('dash.period', {date:fmtDay(st.from)}))}</span><div class="dash-ctl">${custom}${rangeSel}</div></div>
    ${on.has('kpis') ? `<div class="kpis">
      ${tile(t('dash.open'), st.open.length)}
      ${tile(t('dash.over'), st.over.length, st.over.length ? esc(t('dash.overSub')) : '', st.over.length ? 'bad' : '')}
      ${tile(t('dash.soon'), st.soon.length)}
      ${tile(t('dash.done'), st.done, esc(tn('dash.createdN', st.created)))}
      ${tile(t('dash.onTime'), st.onTime == null ? '—' : st.onTime + '%', st.onTimeN ? esc(tn('dash.ofDue', st.onTimeN)) : esc(t('dash.noDue')))}
      ${tile(t('dash.avg'), st.avg == null ? '—' : esc(String(st.avg)), st.avg == null ? '' : esc(t('dash.daysUnit')))}
    </div>` : ''}
    ${charts.length ? `<div class="viz-grid" data-n="${charts.length}">${charts.map(w=>make[w.id](w.kind === 'half' && !(on.has('weekly') && on.has('backlog')) ? {...w, kind:'wide'} : w)).join('')}</div>`
      : !on.has('kpis') ? `<div class="empty-state dash-none"><b>${esc(t('dash.allHidden'))}</b><button class="btn primary" data-act="dash-showall">${esc(t('dash.showAll'))}</button></div>` : ''}`;
  fitX();
}
// hover / tap tooltip for the charts
const vizTip = document.createElement('div'); vizTip.className = 'viz-tip'; vizTip.hidden = true; document.body.appendChild(vizTip);
dashview.addEventListener('pointermove', e=>{
  const el = e.target.closest('[data-tip]'); if (!el){ vizTip.hidden = true; return; }
  vizTip.textContent = el.dataset.tip; vizTip.hidden = false;
  const w = vizTip.offsetWidth, h = vizTip.offsetHeight;
  vizTip.style.left = Math.max(8, Math.min(innerWidth - w - 8, e.clientX - w / 2)) + 'px';
  vizTip.style.top = Math.max(8, e.clientY - h - 14) + 'px';
});
dashview.addEventListener('pointerleave', ()=>{ vizTip.hidden = true; });
addEventListener('scroll', ()=>{ vizTip.hidden = true; }, true);

function renderSide(){
  const all = isAll(), b = curBoard(), edit = canEdit();
  const openCount = id=>S.issues.filter(i=>(id===ALL || i.boardId===id) && !i.archived && !isDone(i)).length || '';
  $('#boardList').innerHTML = (S.boards.length ? `<div class="board-item${all?' on':''}" role="button" tabindex="0" data-act="switch-board" data-id="${ALL}" aria-current="${all}">
      <span class="b-av all">${I.grid}</span><span class="bi-name">${esc(t('nav.allBoards'))}</span><span class="bi-n">${openCount(ALL)}</span><span class="bi-sp"></span></div>` : '') +
    S.boards.map(x=>`<div class="board-item${!all && b && x.id===b.id?' on':''}" role="button" tabindex="0" data-act="switch-board" data-id="${x.id}" aria-current="${!all && b && x.id===b.id}">
      ${bAv(x)}<span class="bi-name" data-bidi>${esc(x.name)}</span>
      <span class="bi-n">${openCount(x.id)}</span>
      ${edit ? `<span class="bi-menu" role="button" tabindex="0" data-act="board-menu" data-id="${x.id}" aria-label="${esc(t('board.actions', {name:x.name}))}" aria-haspopup="menu">${I.more}</span>` : '<span class="bi-sp"></span>'}</div>`).join('');
  const av = $('#projAv');
  if (all || !b){
    av.className = 'proj-av'; av.innerHTML = I.grid; av.style.setProperty('--bc', '#44546F');
    $('#projName').textContent = all ? t('nav.allBoards') : (ws?.name || '');
    const n = S.issues.filter(i=>!i.archived).length;
    $('#projSub').textContent = t('side.allSub', {b:tn('boards', S.boards.length), i:tn('issues', n)});
  } else {
    if (isImg(b.logo)){ av.className = 'proj-av logo'; av.innerHTML = `<img src="${esc(b.logo)}" alt="">`; }
    else { av.className = 'proj-av'; av.textContent = (b.name.trim()[0]||'B').toUpperCase(); av.style.setProperty('--bc', b.color); }
    $('#projName').textContent = b.name;
    $('#projSub').textContent = t('side.boardSub', {key:b.key, i:tn('issues', boardIssues(b).length)});
  }
  $('#crumbWs').textContent = ws?.name || '';
  $('#crumbBoard').textContent = all ? t('nav.allBoards') : (b?.name || '');
  $('#sideSettings').hidden = all || !b || !edit; $('#gearBtn').hidden = all || !b || !edit;
  $('#sideTypes').hidden = !edit; $('#newBoardBtn').hidden = !edit; $('#sideImport').hidden = !edit;
  $('#sideWs').hidden = !isOwner();
  document.title = `${all ? t('nav.allBoards') : (b?.name || ws?.name || '')} - Workbench`;
}
function renderNavWs(){
  const btn = $('#wsBtn'); if (!btn || !ws) return;
  btn.innerHTML = `${isImg(ws.logo) ? `<span class="ws-logo"><img src="${esc(ws.logo)}" alt=""></span>` : ''}<span class="ws-name" data-bidi>${esc(ws.name)}</span>${I.chev}`;
  const r = role();
  $('#roChip').hidden = canEdit();
  $('#roChip').textContent = t('role.viewOnly');
  btn.title = r === 'owner' ? '' : t('ws.sharedBy', {name:ws.ownerName || ws.ownerEmail});
  renderBell();
  const av = $('#meBtn'); if (av && user) av.innerHTML = avatar(user);
}
function renderBell(){
  const n = incoming.filter(r=>r.status==='pending').length, b = $('#bellBtn');
  if (!b) return;
  b.hidden = !n; $('#bellN').textContent = n || '';
  b.title = tn('req.pendingN', n);
}
function render(){
  if (screen !== 'app') return;
  const v = P.view, all = isAll(), b = curBoard(), name = all ? t('nav.allBoards') : (b?.name || ws?.name || '');
  board.hidden = v !== 'board'; listview.hidden = v !== 'list' && v !== 'archive'; sheetview.hidden = v !== 'table'; calview.hidden = v !== 'cal'; dashview.hidden = v !== 'dash';
  const title = v === 'list' ? t('view.listOf', {name}) : v === 'archive' ? t('view.archiveOf', {name}) : v === 'table' ? t('view.tableOf', {name}) : v === 'cal' ? t('view.calOf', {name}) : v === 'dash' ? t('view.dashOf', {name}) : name;
  $('#viewTitle').innerHTML = `${!all && b && isImg(b.logo) ? `<span class="h1-logo"><img src="${esc(b.logo)}" alt=""></span>` : ''}<span data-bidi>${esc(title)}</span>`;
  $$('.side-item[data-view]').forEach(x=>x.classList.toggle('on', x.dataset.view===v));
  if (typeFilter && !S.types.some(x=>x.id===typeFilter)) typeFilter = null;
  if (v === 'board') renderBoard(); else if (v === 'table') renderSheet(); else if (v === 'cal') renderCal(); else if (v === 'dash') renderDash(); else renderList(v === 'archive');
  renderSide();
  const pool = scopeIssues(), open = pool.filter(i=>!isDone(i)), tday = today();
  const over = open.filter(i=>i.due && i.due < tday).length;
  const arch = S.issues.filter(i=>i.archived && (all || (b && i.boardId===b.id))).length;
  $('#summary').textContent = v === 'archive' ? tn('summary.archived', arch) : pool.length ? tn('summary.open', open.length) + (over ? t('summary.overdueSuffix', {n:over}) : '') : '';
  $$('.qf[data-qf]').forEach(x=>x.setAttribute('aria-pressed', qf.has(x.dataset.qf)));
  const tf = $('#typeF'); tf.setAttribute('aria-pressed', !!typeFilter);
  tf.innerHTML = esc(typeFilter ? t('filter.typeIs', {name:typeOf(typeFilter).name}) : t('filter.type')) + I.chev;
  const pf = $('#platF'); pf.setAttribute('aria-pressed', !!platFilter);
  pf.innerHTML = (platFilter === 'none' ? esc(t('filter.noPlatform')) : platFilter ? `<i class="pf-dot" style="--sw:${platFilter}"></i>${esc(colorName(platFilter))}` : esc(t('filter.platform'))) + I.chev;
  $('#clearF').hidden = !filtering() && !(v === 'table' && sheetFCount());
  $('#createBtn').hidden = !canEdit() || !S.boards.length;
  fx = {};
}
let rq = 0; const scheduleRender = ()=>{ cancelAnimationFrame(rq); rq = requestAnimationFrame(render); };
function switchBoard(id){
  if (id !== ALL && !boardById(id)) return;
  if (activeKey() !== id){ P.active[wsId] = id; if (id !== ALL){ P.lastBoard = P.lastBoard || {}; P.lastBoard[wsId] = id; } composer = null; composerText = ''; savePrefs(); }
  render(); shell.classList.remove('side-open');
  board.scrollLeft = 0;
}

/* ================= Moves ================= */
function moveIssue(id, status, beforeId, opts={}){
  if (!canEdit()) return;
  const i = get(id); if (!i) return;
  const b = boardOf(i), was = colOf(i), to = resolveStatus(b, status);
  if (!to){
    const v = vColumns().find(x=>x.id===status);
    toast(t('toast.noSuchColumn', {board:b.name, col:v ? v.name : ''}));
    render(); return;
  }
  const arr = S.issues; arr.splice(arr.indexOf(i),1);
  const k = beforeId ? arr.findIndex(x=>x.id===beforeId) : -1;
  k < 0 ? arr.push(i) : arr.splice(k,0,i);
  i.rank = undefined;
  if (was.id !== to.id){
    i.status = to.id; touch(i, t('h.status', {from:was.name, to:to.name}));
    if (to.cat === 'done' && was.cat !== 'done'){ fx.done = id; i.doneAt = Date.now(); } else fx.land = id;
    if (to.cat !== 'done') i.doneAt = null;
    if (opts.toast) toast(t('toast.moved', {key:key(i), col:to.name}));
  } else { fx.land = id; i.updated = Date.now(); }
  save(); render();
  if (openId === id) refreshIssue();
}
function moveToBoard(i, boardId){
  const from = boardOf(i), to = boardById(boardId); if (!to || from.id === to.id) return;
  const oldKey = key(i), oldCol = colOf(i);
  to.seq = Math.max(to.seq||0, ...S.issues.filter(x=>x.boardId===to.id).map(x=>x.num)) + 1; i.num = to.seq; i.boardId = to.id;
  const nc = to.columns.find(c=>c.name.trim().toLowerCase()===oldCol.name.trim().toLowerCase()) || to.columns.find(c=>c.cat===oldCol.cat) || to.columns[0];
  i.status = nc.id; if (nc.cat !== 'done') i.doneAt = null; else if (!i.doneAt) i.doneAt = Date.now();
  i.labels = i.labels.map(l=>sameText(l, from.name) ? normLabel(to.name) : l);
  touch(i, t('h.movedBoard', {from:from.name, key:oldKey, to:to.name}));
  S.issues.splice(S.issues.indexOf(i),1); S.issues.push(i); i.rank = undefined;
  save(); render(); refreshIssue();
  toast(t('toast.movedBoard', {board:to.name, key:key(i)}), isAll() ? {} : {action:t('btn.goToBoard'), fn:()=>{ closeModal(); switchBoard(to.id); }});
}
function createIssue(d){
  const b = boardById(d.boardId) || curBoard();
  const col = resolveStatus(b, d.status) || b.columns[0];
  b.seq = Math.max(b.seq||0, ...S.issues.filter(x=>x.boardId===b.id).map(x=>x.num)) + 1;
  const now = Date.now();
  const i = normIssue('i-'+uid(), {priority:'medium', start:today(), ...d, boardId:b.id, num:b.seq, status:col.id, doneAt:col.cat==='done' ? now : null,
    created:now, updated:now, creator:me(), history:[{id:uid(), at:now, text:t('h.created'), by:user.name}]});
  i.rank = undefined;
  S.issues.push(i); fx.new = i.id; save(); render();
  return i;
}
function deleteIssue(id){
  const ix = S.issues.findIndex(i=>i.id===id); if (ix < 0) return;
  const [i] = S.issues.splice(ix,1);
  if (openId === id) closeModal(true);
  save(); render();
  toast(t('toast.deleted', {key:key(i)}), {action:t('btn.undo'), fn:()=>{ if (!boardById(i.boardId)) return toast(t('toast.boardGone')); S.issues.splice(Math.min(ix,S.issues.length),0,i); i.rank = undefined; fx.new = i.id; save(); render(); }});
}
function setColor(i, c){
  i.color = isHex(c) ? c : null;
  touch(i, i.color ? t('h.platform', {name:colorName(i.color)}) : t('h.platformNone'));
  save(); render();
}
function archiveIssues(list, msg){
  if (!list.length) return;
  const now = Date.now();
  list.forEach(i=>{ i.archived = true; i.archivedAt = now; touch(i, t('h.archived')); });
  if (list.some(i=>i.id===openId)) closeModal();
  save(); render();
  toast(msg, {action:t('btn.undo'), fn:()=>{ list.forEach(i=>{ i.archived = false; i.archivedAt = null; touch(i, t('h.restored')); }); save(); render(); }});
}
function restoreIssue(i){
  i.archived = false; i.archivedAt = null; touch(i, t('h.restored')); fx.new = i.id;
  save(); render(); refreshIssue();
  toast(t('toast.restored', {key:key(i), col:colOf(i).name}));
}

/* ================= Drag & drop ================= */
let drag = null, settling = false, suppressClick = false;
board.addEventListener('click', e=>{
  const card = e.target.closest('.card');
  if (!card || e.target.closest('button') || suppressClick || settling) return;
  openIssue(card.dataset.id);
});
board.addEventListener('pointerdown', e=>{
  if (settling || e.button !== 0 || !canEdit()) return;
  const card = e.target.closest('.card'); if (!card || e.target.closest('button')) return;
  drag = {card, id:card.dataset.id, sx:e.clientX, sy:e.clientY, x:e.clientX, y:e.clientY, touch:e.pointerType!=='mouse', active:false, timer:0};
  if (drag.touch) drag.timer = setTimeout(()=>{ if (drag && !drag.active) begin(); }, 280);
});
window.addEventListener('pointermove', e=>{
  if (!drag) return;
  drag.x = e.clientX; drag.y = e.clientY;
  if (!drag.active){
    const dist = Math.hypot(drag.x-drag.sx, drag.y-drag.sy);
    if (drag.touch){ if (dist > 8){ clearTimeout(drag.timer); drag = null; } }
    else if (dist > 4) begin();
  }
});
window.addEventListener('pointerup', ()=>end(false));
window.addEventListener('pointercancel', ()=>end(true));
document.addEventListener('touchmove', e=>{ if (drag && drag.active) e.preventDefault(); }, {passive:false});
board.addEventListener('contextmenu', e=>{ if (e.target.closest('.card')) e.preventDefault(); });
function begin(){
  lastCols = null;
  // the board may have been redrawn since the press (a sync from the server): use the card that is on screen now
  if (!drag.card.isConnected){ const c = $(`.card[data-id="${CSS.escape(drag.id)}"]`, board); if (!c){ drag = null; return; } drag.card = c; }
  const card = drag.card, r = card.getBoundingClientRect();
  if (!r.width){ drag = null; return; }
  closePop();
  drag.active = true; drag.ox = drag.sx - r.left; drag.oy = drag.sy - r.top;
  if (drag.touch && navigator.vibrate) navigator.vibrate(10);
  const g = card.cloneNode(true); g.classList.remove('is-new','is-land','is-done-fx'); g.classList.add('ghost'); g.removeAttribute('tabindex'); g.style.width = r.width+'px';
  g.style.transform = `translate(${r.left}px,${r.top}px)`;
  document.body.appendChild(g); drag.ghost = g;
  const ph = document.createElement('div'); ph.className = 'ph'; ph.style.height = r.height+'px';
  card.after(ph); drag.ph = ph; card.classList.add('is-hidden');
  drag.col = card.closest('.col'); drag.col.classList.add('over');
  document.body.classList.add('dragging');
  const tick = ()=>{ if (!drag || !drag.active) return; autoscroll(); follow(); drag.raf = requestAnimationFrame(tick); };
  drag.raf = requestAnimationFrame(tick);
}
function nextCard(el){ let n = el.nextElementSibling; while (n && (!n.classList.contains('card') || n.classList.contains('is-hidden'))) n = n.nextElementSibling; return n; }
function flip(fn){
  if (reduce){ fn(); return; }
  const els = $$('.card:not(.is-hidden)', board), r0 = new Map(els.map(e=>[e,e.getBoundingClientRect()]));
  fn();
  for (const e of els){
    const a = r0.get(e), b = e.getBoundingClientRect(), dx = a.left-b.left, dy = a.top-b.top;
    if (Math.abs(dx) > .5 || Math.abs(dy) > .5)
      e.animate([{transform:`translate(${dx}px,${dy}px)`},{transform:'translate(0,0)'}], {duration:200, easing:'cubic-bezier(.2,.8,.2,1)'});
  }
}
function follow(){
  const {x, y, ghost, ph} = drag;
  ghost.style.transform = `translate(${x-drag.ox}px,${y-drag.oy}px)`;
  let col = document.elementFromPoint(x,y)?.closest('.col');
  if (!col){ const br = board.getBoundingClientRect(); if (y > br.top && y < br.bottom) col = $$('.col',board).find(c=>{ const r = c.getBoundingClientRect(); return x >= r.left && x <= r.right; }); }
  if (col && col !== drag.col){ drag.col?.classList.remove('over'); col.classList.add('over'); drag.col = col; }
  if (!col) return;
  const list = $('.list',col), lr = list.getBoundingClientRect(), ly = y - lr.top + list.scrollTop;
  let before = null;
  for (const c of $$('.card:not(.is-hidden)', list)){ if (ly < c.offsetTop + c.offsetHeight/2){ before = c; break; } }
  const same = ph.parentElement === list && nextCard(ph) === before;
  if (!same){ $('.empty-hint',list)?.remove(); flip(()=>{ before ? list.insertBefore(ph,before) : list.appendChild(ph); }); }
}
function autoscroll(){
  const {x, y} = drag, E = 60, br = board.getBoundingClientRect();
  const hd = drag.touch ? 9 : 5; // gentler on phones, where one column is almost the whole screen
  if (x < br.left + E) board.scrollLeft -= Math.ceil((br.left+E-x)/hd);
  else if (x > br.right - E) board.scrollLeft += Math.ceil((x-br.right+E)/hd);
  if (drag.col){
    // speeds up the longer the card is held at the edge, so long columns (100+ tasks) stay usable
    const list = $('.list',drag.col), lr = list.getBoundingClientRect(), now = performance.now();
    const dir = y < lr.top + 40 ? -1 : (y > lr.bottom - 40 && y < lr.bottom + 80) ? 1 : 0;
    if (dir !== drag.edgeDir || list !== drag.edgeList){ drag.edgeDir = dir; drag.edgeList = list; drag.edgeT = now; }
    if (dir){
      const base = dir < 0 ? Math.ceil((lr.top+40-y)/4) : Math.ceil((y-lr.bottom+40)/4), boost = 1 + Math.min((now - drag.edgeT)/350, 7);
      list.scrollTop += dir * Math.min(Math.ceil(base * boost), 90);
    }
  }
}
function end(cancel){
  if (!drag) return;
  clearTimeout(drag.timer); cancelAnimationFrame(drag.raf);
  const d = drag; drag = null;
  if (!d.active) return;
  suppressClick = true; setTimeout(()=>{ suppressClick = false; }, 60);
  document.body.classList.remove('dragging');
  $$('.col.over',board).forEach(c=>c.classList.remove('over'));
  if (cancel){ d.ghost.remove(); d.ph.remove(); d.card.classList.remove('is-hidden'); if (boardPending) renderBoard(); return; }
  const list = d.ph.parentElement, status = list?.dataset.status, beforeId = nextCard(d.ph)?.dataset.id;
  if (!status || !get(d.id)){ d.ghost.remove(); d.ph.remove(); d.card.classList.remove('is-hidden'); lastCols = null; renderBoard(); return; }
  if (reduce){ d.ghost.remove(); moveIssue(d.id, status, beforeId); if (boardPending) renderBoard(); return; }
  settling = true;
  const r = d.ph.getBoundingClientRect();
  d.ghost.style.transition = 'transform .18s cubic-bezier(.2,.8,.2,1), rotate .18s, scale .18s, box-shadow .18s';
  d.ghost.classList.add('dropping');
  d.ghost.style.transform = `translate(${r.left}px,${r.top}px)`;
  setTimeout(()=>{ d.ghost.remove(); settling = false; moveIssue(d.id, status, beforeId); if (boardPending) renderBoard();
    if (isMobile()) $(`.col[data-col="${CSS.escape(status)}"]`,board)?.scrollIntoView({behavior:'smooth', inline:'center', block:'nearest'}); }, 180);
}
board.addEventListener('keydown', e=>{
  const card = e.target.closest('.card'); if (!card || e.target !== card) return;
  if (e.key === 'Enter' || e.key === ' '){ e.preventDefault(); openIssue(card.dataset.id); }
  if (e.shiftKey && (e.key === 'ArrowRight' || e.key === 'ArrowLeft') && canEdit()){
    e.preventDefault();
    const i = get(card.dataset.id), forward = (e.key === 'ArrowRight') !== rtl(), step = forward ? 1 : -1;
    const cs = isAll() ? vColumns() : boardOf(i).columns, cur = isAll() ? cs.findIndex(v=>v.id===vKey(i)) : cs.indexOf(colOf(i)), ix = cur + step;
    if (ix < 0 || ix >= cs.length) return;
    moveIssue(i.id, cs[ix].id, null);
    $(`.card[data-id="${i.id}"]`,board)?.focus();
  }
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp'){
    e.preventDefault();
    const sib = e.key === 'ArrowDown' ? card.nextElementSibling : card.previousElementSibling;
    if (sib?.classList.contains('card')) sib.focus();
  }
});
listview.addEventListener('click', e=>{ if (e.target.closest('button')) return; const tr = e.target.closest('tbody tr'); if (tr) openIssue(tr.dataset.id); });
listview.addEventListener('keydown', e=>{ const tr = e.target.closest('tbody tr'); if (tr && e.target === tr && (e.key==='Enter'||e.key===' ')){ e.preventDefault(); openIssue(tr.dataset.id); } });

/* ================= Inline composer ================= */
function openComposer(status){ composer = status; composerText = ''; composerFocus = true; render(); }
function closeComposer(){ if (!composer) return; composer = null; composerText = ''; render(); }
board.addEventListener('input', e=>{ if (e.target.matches('.composer textarea')){ composerText = e.target.value; autosize(e.target); } });
board.addEventListener('change', e=>{
  const el = e.target; if (!el.matches('.composer select')) return;
  if (el.dataset.c === 'type'){ composerType = el.value; el.closest('.type-sel').firstElementChild.outerHTML = typeIcon(composerType); }
  if (el.dataset.c === 'board'){ composerBoard = el.value; el.closest('.type-sel').firstElementChild.outerHTML = bAv(boardById(composerBoard),'xs'); }
  $('.composer textarea',board).focus();
});
board.addEventListener('keydown', e=>{
  if (!e.target.matches('.composer textarea')) return;
  if (e.key === 'Enter' && !e.shiftKey && !e.isComposing){ e.preventDefault(); e.target.form.requestSubmit(); }
  if (e.key === 'Escape'){ e.stopPropagation(); closeComposer(); }
});
board.addEventListener('submit', e=>{
  e.preventDefault();
  if (!canEdit()) return;
  const txt = composerText.trim(); if (!txt){ $('.composer textarea',board).focus(); return; }
  const b = isAll() ? boardById(composerBoard) : curBoard();
  if (!resolveStatus(b, composer)){ toast(t('toast.noColumnHere', {board:b.name})); return; }
  composerText = ''; composerFocus = true;
  createIssue({title:txt, type:composerType, boardId:b.id, status:composer, labels:[normLabel(b.name)]});
});
document.addEventListener('pointerdown', e=>{
  if (composer && !e.target.closest('.composer') && !e.target.closest('[data-act="add-inline"]') && !e.target.closest('#layer') && !composerText.trim()) closeComposer();
}, true);

/* ================= Popover ================= */
const pop = $('#pop'); let popAnchor = null;
// typing in a table column's "contains" filter
let sfTimer = 0;
pop.addEventListener('input', e=>{
  const el = e.target; if (el.dataset?.k !== 's-ftext') return;
  const k = el.dataset.col, v = el.value.trim();
  clearTimeout(sfTimer); sheetF[k] = v || undefined;
  sfTimer = setTimeout(()=>{ render(); const btn = $(`.s-filt[data-k="${CSS.escape(k)}"]`, sheetview); if (popAnchor && btn){ popAnchor = btn; btn.setAttribute('aria-expanded','true'); placePop(); } el.focus(); }, 160);
});
function openPop(anchor, html, id){
  if (popAnchor === anchor){ closePop(); return; }
  closePop();
  pop.innerHTML = html; pop.dataset.id = id; pop.hidden = false; popAnchor = anchor;
  anchor.setAttribute('aria-expanded','true');
  placePop();
  requestAnimationFrame(()=>{ pop.classList.add('on'); $('.menu-i, .plat-b, .sw, .glyph-btn',pop)?.focus({preventScroll:true}); });
}
function placePop(){
  if (!popAnchor) return;
  const r = popAnchor.getBoundingClientRect(), w = pop.offsetWidth, h = pop.offsetHeight;
  const want = rtl() ? r.right - w : r.left;
  const left = Math.max(8, Math.min(want, innerWidth - w - 8));
  let top = r.bottom + 4;
  const up = top + h > innerHeight - 8 && r.top - h - 4 >= 8;
  if (up) top = r.top - h - 4; else top = Math.max(8, Math.min(top, innerHeight - h - 8));
  pop.classList.toggle('up', up);
  pop.style.left = left+'px'; pop.style.top = top+'px';
}
function closePop(){
  if (!popAnchor) return;
  popAnchor.setAttribute('aria-expanded','false'); popAnchor = null;
  pop.classList.remove('on'); setTimeout(()=>{ if (!popAnchor) pop.hidden = true; }, 150);
}
document.addEventListener('pointerdown', e=>{ if (popAnchor && !pop.contains(e.target) && !popAnchor.contains(e.target)) closePop(); });
addEventListener('resize', ()=>{ closePop(); if (!sideOverlay()) shell.classList.remove('side-open'); }); // the drawer only exists on narrow screens
pop.addEventListener('keydown', e=>{
  const items = $$('.menu-i, .plat-b, .sw, .glyph-btn',pop), ix = items.indexOf(document.activeElement);
  const nextKey = rtl() ? 'ArrowLeft' : 'ArrowRight', prevKey = rtl() ? 'ArrowRight' : 'ArrowLeft';
  if (e.key === 'ArrowDown' || e.key === nextKey){ e.preventDefault(); items[(ix+1)%items.length].focus(); }
  if (e.key === 'ArrowUp' || e.key === prevKey){ e.preventDefault(); items[(ix-1+items.length)%items.length].focus(); }
  if (e.key === 'Tab'){ closePop(); }
});
const chk = on=>on ? `<span class="chk-mark">${I.check}</span>` : '';
function statusMenu(i, withExtras){
  const b = boardOf(i), edit = canEdit();
  if (!edit) return `<button class="menu-i" role="menuitem" data-act="open-issue">${I.open}${esc(t('menu.open'))}</button>
    <button class="menu-i" role="menuitem" data-act="copy-key">${I.copy}${esc(t('menu.copyKey'))}</button>`;
  return `<div class="menu-h">${esc(withExtras ? t('menu.moveTo') : t('menu.changeStatus'))}</div>${b.columns.filter(c=>c.id!==i.status).map(c=>`<button class="menu-i" role="menuitem" data-act="set-status" data-status="${c.id}">${lozC(c)}</button>`).join('')}
  ${withExtras ? `<div class="menu-sep"></div><div class="menu-h">${esc(t('f.platform'))}</div><div class="pop-pad">${platformPicker('set-color', i.color)}</div>
  <div class="menu-sep"></div>
  <button class="menu-i" role="menuitem" data-act="open-issue">${I.open}${esc(t('menu.open'))}</button>
  <button class="menu-i" role="menuitem" data-act="copy-key">${I.copy}${esc(t('menu.copyKey'))}</button>
  <button class="menu-i" role="menuitem" data-act="archive-issue">${I.archive}${esc(t('menu.archive'))}</button>
  <button class="menu-i danger" role="menuitem" data-act="del-issue">${I.trash}${esc(t('menu.delete'))}</button>` : ''}`;
}
function wsMenu(){
  const owned = (wsList||[]).some(w=>w.ownerUid===user.uid);
  return `<div class="menu-h">${esc(t('ws.your'))}</div>${(wsList||[]).map(w=>{
      const r = w.ownerUid===user.uid ? 'owner' : (w.roles||{})[user.email] || 'viewer';
      return `<button class="menu-i ws-i" role="menuitem" data-act="switch-ws" data-id="${esc(w.id)}">${isImg(w.logo) ? `<span class="ws-logo sm"><img src="${esc(w.logo)}" alt=""></span>` : `<span class="b-av" style="--bc:#44546F">${initialOf(w.name)}</span>`}
        <span class="ws-i-txt"><b data-bidi>${esc(w.name)}</b><small>${esc(r==='owner' ? t('role.owner') : t('ws.sharedRole', {role:t('role.'+r), name:w.ownerName||w.ownerEmail}))}</small></span>${chk(w.id===wsId)}</button>`;
    }).join('')}
    <div class="menu-sep"></div>
    ${owned ? '' : `<button class="menu-i" role="menuitem" data-act="welcome">${I.plus}${esc(t('ws.createOwn'))}</button>`}
    <button class="menu-i" role="menuitem" data-act="request-access">${I.users}${esc(t('ws.requestAccess'))}</button>
    ${isOwner() ? `<button class="menu-i" role="menuitem" data-act="ws-settings">${I.gear}${esc(t('ws.settings'))}</button>` : ''}`;
}
function meMenu(){
  return `<div class="me-head">${avatar(user,'lg')}<div><b data-bidi>${esc(user.name)}</b><span dir="ltr">${esc(user.email)}</span></div></div>
    <div class="menu-sep"></div>
    ${isOwner() ? `<button class="menu-i" role="menuitem" data-act="ws-settings">${I.users}${esc(t('ws.shareMenu'))}</button>` : ''}
    <button class="menu-i" role="menuitem" data-act="request-access">${I.mail}${esc(t('ws.requestAccess'))}</button>
    <button class="menu-i" role="menuitem" data-act="lang">${I.globe}${getLang()==='he'?'English':'עברית'}</button>
    <div class="menu-sep"></div>${themePicker()}
    <div class="menu-sep"></div>
    <button class="menu-i" role="menuitem" data-act="signout">${I.logout}${esc(t('acc.signout'))}</button>`;
}

/* ================= Modals ================= */
const layer = $('#layer'), shell = $('#shell');
let modalKind = null, lastFocus = null, openFocusId = null;
function showModal(html, kind, size=''){
  closePop();
  if (!modalKind) lastFocus = document.activeElement;
  layer.innerHTML = `<div class="backdrop"><div class="dialog ${size}" role="dialog" aria-modal="true" tabindex="-1">${html}</div></div>`;
  modalKind = kind;
  // Making the page inert restyles every card (slow with hundreds of tasks), so let the dialog paint first.
  if (!shell.inert) requestAnimationFrame(()=>setTimeout(()=>{ if (modalKind){ shell.inert = true; $('.nav').inert = true; } }, 0));
  const bd = $('.backdrop',layer); let downOnBd = false;
  bd.addEventListener('pointerdown', e=>{ downOnBd = e.target === bd; });
  bd.addEventListener('click', e=>{ if (e.target === bd && downOnBd) closeModal(); });
  $$('textarea',layer).forEach(autosize);
  return $('.dialog',layer);
}
function closeModal(instant){
  if (!modalKind) return;
  closePop();
  if (modalKind === 'issue'){
    const i = get(openId);
    if (i && !i.title.trim() && canEdit()){ i.title = t('issue.untitled'); save(); render(); }
    openId = null; editingCid = null;
  }
  bDraft = null; tDraft = null; confirmFn = null;
  modalKind = null; shell.inert = false; $('.nav').inert = false;
  const bd = $('.backdrop',layer);
  if (!bd) return;
  if (instant || reduce){ layer.innerHTML = ''; }
  else { bd.classList.add('closing'); bd.style.pointerEvents = 'none'; setTimeout(()=>{ if (bd.isConnected && bd.classList.contains('closing')) bd.remove(); }, 170); }
  if (lastFocus && lastFocus.isConnected) lastFocus.focus({preventScroll:true});
  else { const c = openFocusId && $(`.card[data-id="${openFocusId}"]`,board); c?.focus({preventScroll:true}); }
}
function rerender(el, html){
  const a = document.activeElement, ak = a?.dataset?.k, sel = ak && a.selectionStart != null ? [a.selectionStart, a.selectionEnd] : null;
  el.innerHTML = html;
  if (ak){ const n = $(`[data-k="${CSS.escape(ak)}"]`, layer); if (n){ n.focus({preventScroll:true}); if (sel) try{ n.setSelectionRange(...sel); }catch(e){} } }
}
const head = (title, extra='')=>`<div class="m-head"><h2>${esc(title)}</h2>${extra}<button type="button" class="btn icon subtle" data-act="close-modal" aria-label="${esc(t('btn.close'))}">${I.x}</button></div>`;

/* ----- Confirm ----- */
let confirmFn = null;
function openConfirm({title, text, ok, fn}){
  confirmFn = fn;
  const dlg = showModal(`${head(title)}<p class="confirm-p">${text}</p>
    <div class="m-foot" style="border:0"><span class="sp"></span><button class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button><button class="btn danger" data-act="confirm-ok">${esc(ok)}</button></div>`, 'confirm', 'tiny');
  dlg.setAttribute('aria-label', title);
  setTimeout(()=>$('.m-foot [data-act="close-modal"]',dlg)?.focus(), 30);
}

/* ----- Issue view ----- */
function openIssue(id){
  const i = get(id); if (!i) return;
  if (composer && !composerText.trim()){ composer = null; render(); }
  openId = id; openFocusId = id; tab = 'comments'; cmtDraft = ''; editingCid = null; lastPct = 0;
  const dlg = showModal(issueHTML(i), 'issue');
  dlg.setAttribute('aria-label', `${key(i)}: ${i.title}`);
  afterIssueRender();
  dlg.focus({preventScroll:true});
}
function pctOf(i){ const n = i.subtasks.length; return n ? Math.round(i.subtasks.filter(s=>s.done).length / n * 100) : 0; }
function linksHTML(i, edit){
  return `<ul class="links">${i.links.map(l=>`<li data-lid="${l.id}">${I.link}<a href="${esc(l.url)}" target="_blank" rel="noopener noreferrer" data-bidi>${esc(l.title || hostOf(l.url))}</a><span class="l-host">${l.title ? esc(hostOf(l.url)) : ''}</span>${edit ? `<button class="btn icon subtle sm l-del" data-act="link-del" aria-label="${esc(t('links.remove'))}" title="${esc(t('links.remove'))}">${I.x}</button>` : ''}</li>`).join('')}</ul>`;
}
function issueHTML(i){
  const n = i.subtasks.length, pct = pctOf(i), col = colOf(i), b = boardOf(i), age = ageInfo(i), edit = canEdit(), dis = edit ? '' : ' disabled';
  const companies = [...new Set([...S.boards.map(x=>normLabel(x.name)), ...S.issues.flatMap(x=>x.labels)])].filter(l=>!i.labels.includes(l)).sort();
  return `
  <div class="m-head">
    <div class="m-crumb">${bAv(b)}<span class="crumb-b" data-bidi>${esc(b.name)}</span><span class="crumb-sep">/</span>${typeIcon(i.type)}<button class="link-btn key" data-act="copy-key" title="${esc(t('menu.copyKey'))}">${esc(key(i))}</button>${i.archived?`<span class="loz arch">${esc(t('issue.archived'))}</span>`:''}${!edit?`<span class="loz ro-loz">${esc(t('role.viewOnly'))}</span>`:''}</div>
    <div class="m-actions">
      ${edit ? (i.archived ? `<button class="btn sm" data-act="restore">${I.restore}${esc(t('btn.restore'))}</button>` : `<button class="btn icon subtle" data-act="archive-issue" aria-label="${esc(t('menu.archive'))}" title="${esc(t('menu.archive'))}">${I.archive}</button>`)
        + `<button class="btn icon subtle" data-act="del-issue" aria-label="${esc(t('menu.delete'))}" title="${esc(t('menu.delete'))}">${I.trash}</button>` : ''}
      <button class="btn icon subtle" data-act="close-modal" aria-label="${esc(t('btn.close'))}" title="${esc(t('btn.closeEsc'))}">${I.x}</button>
    </div>
  </div>
  <div class="m-body">
    <div class="m-title"><textarea class="sum" data-k="sum" data-f="title" rows="1" data-bidi aria-label="${esc(t('f.summary'))}" placeholder="${esc(t('f.summary'))}"${edit?'':' readonly'}>${esc(i.title)}</textarea></div>
    <div class="m-main">
      <section class="m-sec" style="margin-top:12px">
        <h3>${esc(t('f.desc'))}</h3>
        <textarea class="desc" rows="1" data-k="desc" data-f="desc" data-bidi aria-label="${esc(t('f.desc'))}" placeholder="${esc(edit ? t('f.descPh') : t('f.descNone'))}"${edit?'':' readonly'}>${esc(i.desc)}</textarea>
      </section>
      <section class="m-sec">
        <div class="sec-head"><h3>${esc(t('f.links'))}</h3>${i.links.length ? `<span class="sec-meta">${i.links.length}</span>` : ''}</div>
        ${linksHTML(i, edit)}
        ${edit ? `<form class="link-add" data-form="link" novalidate>
          <input class="inp" data-k="l-url" inputmode="url" autocapitalize="off" autocorrect="off" spellcheck="false" dir="ltr" placeholder="${esc(t('links.ph'))}" aria-label="${esc(t('links.url'))}" autocomplete="off">
          <input class="inp" data-k="l-title" data-bidi placeholder="${esc(t('links.textPh'))}" aria-label="${esc(t('links.text'))}" autocomplete="off">
          <button class="btn">${esc(t('links.add'))}</button>
        </form>
        <em class="ed-err" id="linkErr" hidden>${esc(t('links.err'))}</em>` : (i.links.length ? '' : `<p class="muted-p">${esc(t('links.none'))}</p>`)}
      </section>
      <section class="m-sec">
        <div class="sec-head"><h3>${esc(t('f.subtasks'))}</h3>${n ? `<span class="sec-meta" id="stMeta">${esc(t('st.pct', {n:pct}))}</span>` : ''}</div>
        ${n ? `<div class="pbar"><b></b></div>` : ''}
        <ul class="sts">${i.subtasks.map(s=>stHTML(s, edit)).join('')}</ul>
        ${edit ? `<form class="st-add" data-form="subtask"><button type="submit" class="st-plus" tabindex="-1" aria-label="${esc(t('st.addBtn'))}">${I.plus}</button><input data-k="st-add" placeholder="${esc(t('st.add'))}" aria-label="${esc(t('st.new'))}" data-bidi autocomplete="off" enterkeyhint="done"><button type="submit" class="btn primary sm st-go">${esc(t('st.addBtn'))}</button></form>` : (n ? '' : `<p class="muted-p">${esc(t('st.none'))}</p>`)}
      </section>
      <section class="m-sec">
        <div class="sec-head"><h3>${esc(t('f.activity'))}</h3>
          <div class="tabs" role="tablist">
            <button class="tab" role="tab" aria-selected="${tab==='comments'}" data-act="tab" data-tab="comments">${esc(t('tab.comments'))}${i.comments.length?` (${i.comments.length})`:''}</button>
            <button class="tab" role="tab" aria-selected="${tab==='history'}" data-act="tab" data-tab="history">${esc(t('tab.history'))}</button>
          </div></div>
        ${tab === 'comments' ? commentsHTML(i, edit) : historyHTML(i)}
      </section>
    </div>
    <aside class="m-side">
      <div class="side-top"><button class="status-btn ${col.cat}" data-act="status-menu" aria-haspopup="menu" aria-label="${esc(t('f.statusIs', {name:col.name}))}"${dis}><span data-bidi>${esc(col.name)}</span>${edit?I.chev:''}</button>
        <span class="age big ${age.cls}" title="${esc(age.long)}">${I.clock}${esc(age.long)}</span></div>
      <div class="details">
        <h3 class="det-head">${esc(t('f.details'))}</h3>
        <dl class="det">
          <dt>${esc(t('f.type'))}</dt><dd><span class="with-ico">${typeIcon(i.type)}<select class="ctl" data-f="type" aria-label="${esc(t('f.type'))}"${dis}>${S.types.map(x=>`<option value="${x.id}"${x.id===i.type?' selected':''}>${esc(x.name)}</option>`).join('')}</select></span></dd>
          <dt>${esc(t('f.priority'))}</dt><dd><span class="with-ico">${prioIcon(i.priority)}<select class="ctl" data-f="priority" aria-label="${esc(t('f.priority'))}"${dis}>${PRIOS.map(k=>`<option value="${k}"${k===i.priority?' selected':''}>${esc(prioName(k))}</option>`).join('')}</select></span></dd>
          <dt>${esc(t('f.start'))}</dt><dd>${edit ? `<input class="ctl" type="date" data-f="start" value="${esc(i.start)}" aria-label="${esc(t('f.start'))}">` : `<span class="ro-val">${esc(fmtDay(i.start))}</span>`}</dd>
          <dt>${esc(t('f.due'))}</dt><dd>${edit ? `<input class="ctl" type="date" data-f="due" value="${esc(i.due)}" aria-label="${esc(t('f.due'))}">` : `<span class="ro-val${i.due?'':' t-muted'}">${esc(i.due ? fmtDay(i.due) : t('none'))}</span>`}${i.due && edit ? `<button class="btn icon subtle sm" data-act="clear-due" aria-label="${esc(t('f.clearDue'))}" title="${esc(t('f.clearDue'))}">${I.x}</button>` : ''}</dd>
          <dt>${esc(t('f.company'))}</dt><dd><div class="labels-edit${edit?'':' ro'}">${i.labels.map(l=>`<span class="chip" data-bidi>${esc(l)}${edit?`<button data-act="label-del" data-l="${esc(l)}" aria-label="${esc(t('company.remove', {name:l}))}">${I.x}</button>`:''}</span>`).join('')}${edit ? `<input class="lbl-in" data-k="lbl" data-bidi placeholder="${i.labels.length?'':esc(t('company.add'))}" list="lblList" aria-label="${esc(t('company.add'))}" autocomplete="off">` : (i.labels.length ? '' : `<span class="t-muted">${esc(t('none'))}</span>`)}</div>
            <datalist id="lblList">${companies.map(l=>`<option value="${esc(l)}">`).join('')}</datalist></dd>
          <dt class="full">${esc(t('f.platform'))}</dt><dd class="full">${edit ? platformPicker('set-color', i.color) : (i.color ? `<span class="plat-b on" style="--sw:${i.color}"><i></i>${esc(colorName(i.color))}</span>` : `<span class="t-muted">${esc(t('plat.none'))}</span>`)}</dd>
          <dt>${esc(t('f.recur'))}</dt><dd class="rc-dd">${recurField(i, edit)}</dd>
          <dt>${esc(t('f.points'))}</dt><dd>${edit ? `<input class="ctl" type="number" min="0" step="0.5" inputmode="decimal" data-f="points" value="${i.points??''}" placeholder="${esc(t('none'))}" aria-label="${esc(t('f.points'))}">` : `<span class="ro-val${i.points==null?' t-muted':''}">${esc(i.points ?? t('none'))}</span>`}</dd>
          ${S.boards.length > 1 && edit ? `<dt>${esc(t('f.board'))}</dt><dd><select class="ctl" data-f="board" aria-label="${esc(t('f.board'))}">${S.boards.map(x=>`<option value="${x.id}"${x.id===i.boardId?' selected':''}>${esc(x.name)}</option>`).join('')}</select></dd>` : ''}
        </dl>
      </div>
      <p class="m-stamp">${(()=>{ const c = creatorOf(i); return c ? `<span class="m-by">${avatar(c,'xs')}<span>${esc(t('stamp.createdBy', {date:fmtTime(i.created), name:c.name||c.email}))}</span></span>` : esc(t('stamp.created', {date:fmtTime(i.created)})); })()}<br>${esc(t('stamp.updated', {ago:ago(i.updated)}))}${i.archived ? `<br>${esc(t('stamp.archived', {date:fmtTime(i.archivedAt)}))}` : ''}</p>
    </aside>
  </div>`;
}
const stHTML = (s, edit=true, cls='')=>`<li class="${s.done?'done':''} ${cls}" data-sid="${s.id}">
  <button class="chk" data-act="st-toggle" aria-pressed="${s.done}" aria-label="${esc(s.done?t('st.undone'):t('st.done'))}"${edit?'':' disabled'}>${I.tick}</button>
  <input class="st-text" value="${esc(s.text)}" data-bidi aria-label="${esc(t('st.aria'))}" autocomplete="off"${edit?'':' readonly'}>
  ${edit ? `<button class="btn icon subtle sm st-del" data-act="st-del" aria-label="${esc(t('st.delete'))}" title="${esc(t('st.delete'))}">${I.x}</button>` : ''}</li>`;
function commentsHTML(i, edit){
  const mine = c=>!c.by || c.by.uid === user.uid || isOwner();
  return `${edit ? `<div class="cmt-new">${avatar(user,'lg')}<div class="cmt-box">
      <textarea data-k="cmt" rows="1" data-bidi placeholder="${esc(t('cmt.ph'))}" aria-label="${esc(t('cmt.ph'))}">${esc(cmtDraft)}</textarea>
      <div class="cmt-actions"><button class="btn primary sm" data-act="cmt-save">${esc(t('btn.save'))}</button><button class="btn subtle sm" data-act="cmt-cancel">${esc(t('btn.cancel'))}</button><span class="hint">${esc(t('cmt.hint'))}</span></div>
    </div></div>` : ''}
    ${!i.comments.length && !edit ? `<p class="muted-p">${esc(t('cmt.none'))}</p>` : ''}
    <ol class="cmts">${i.comments.map(c=>{ const by = c.by || {name:ws?.ownerName||'', email:ws?.ownerEmail||''};
      return `<li data-cid="${c.id}" class="${fx.cmt===c.id?'in':''}">${avatar(by,'lg')}<div class="cmt-body">
      <div class="cmt-meta"><b data-bidi>${esc(by.name||by.email)}</b><time title="${esc(fmtTime(c.at))}">${esc(ago(c.at))}</time>${c.edited?`<span>${esc(t('cmt.edited'))}</span>`:''}</div>
      ${editingCid === c.id
        ? `<textarea class="cmt-edit" rows="1" data-k="cmt-edit" data-bidi aria-label="${esc(t('cmt.editAria'))}">${esc(c.text)}</textarea><div class="cmt-actions"><button class="btn primary sm" data-act="cmt-update">${esc(t('btn.save'))}</button><button class="btn subtle sm" data-act="cmt-edit-cancel">${esc(t('btn.cancel'))}</button></div>`
        : `<p data-bidi>${linkify(esc(c.text))}</p>${edit && mine(c) ? `<div class="cmt-links"><button class="link-btn" data-act="cmt-edit">${esc(t('btn.edit'))}</button><button class="link-btn" data-act="cmt-del">${esc(t('btn.delete'))}</button></div>` : ''}`}
    </div></li>`; }).join('')}</ol>`;
}
function historyHTML(i){
  if (!i.history.length) return `<p class="muted-p">${esc(t('hist.none'))}</p>`;
  return `<ol class="hist">${i.history.map(h=>`<li><span class="h-ico">${I.clock}</span><div><p data-bidi>${h.by ? `<b data-bidi>${esc(h.by)}</b>${getLang()==='he'?':':''} ` : ''}${esc(h.text)}</p><time title="${esc(fmtTime(h.at))}">${esc(ago(h.at))}</time></div></li>`).join('')}</ol>`;
}
function setBar(pct){
  const b = $('.pbar b',layer); if (!b) { lastPct = pct; return; }
  b.style.width = (lastPct ?? 0) + '%';
  requestAnimationFrame(()=>requestAnimationFrame(()=>{ b.style.width = pct + '%'; }));
  lastPct = pct;
}
function afterIssueRender(){
  $$('textarea',layer).forEach(autosize);
  const i = get(openId); if (i) setBar(pctOf(i));
}
function refreshIssue(){
  if (modalKind !== 'issue') return;
  const i = get(openId); if (!i) return closeModal();
  const dlg = $('.dialog',layer), body = $('.m-body',dlg), st = body?.scrollTop || 0;
  rerender(dlg, issueHTML(i));
  $('.m-body',dlg).scrollTop = st;
  afterIssueRender();
  fx.cmt = null;
}

/* ----- Create dialog ----- */
function openCreate(status, opts={}){
  if (!canEdit()) return toast(t('toast.viewOnly'));
  if (!S.boards.length) return openBoardSettings(null);
  const b = isAll() ? (boardById(P.lastBoard?.[wsId]) || S.boards[0]) : curBoard();
  if (typeOf(lastCreateType)?.id !== lastCreateType) lastCreateType = S.types[0]?.id;
  const col = resolveStatus(b, status) || b.columns[0];
  const companies = [...new Set([...S.boards.map(x=>normLabel(x.name)), ...S.issues.flatMap(x=>x.labels)])].sort();
  const dlg = showModal(`<form data-form="create" novalidate style="display:contents">
    ${head(t('create.title'))}
    <div class="c-body">
      <p class="req-note">${t('create.req', {star:'<span class="req">*</span>'})}</p>
      <div class="grid2">
        <label class="fld">${esc(t('f.board'))}<select name="board" data-prev="${esc(b.name)}">${S.boards.map(x=>`<option value="${x.id}"${x.id===b.id?' selected':''}>${esc(x.name)}</option>`).join('')}</select></label>
        <label class="fld">${esc(t('f.status'))}<select name="status">${b.columns.map(c=>`<option value="${c.id}"${c.id===col.id?' selected':''}>${esc(c.name)}</option>`).join('')}</select></label>
      </div>
      <div class="grid2">
        <label class="fld">${esc(t('f.type'))}<span class="with-ico">${typeIcon(lastCreateType)}<select name="type">${S.types.map(x=>`<option value="${x.id}"${x.id===lastCreateType?' selected':''}>${esc(x.name)}</option>`).join('')}</select></span></label>
        <label class="fld">${esc(t('f.priority'))}<span class="with-ico">${prioIcon('medium')}<select name="priority">${PRIOS.map(k=>`<option value="${k}"${k==='medium'?' selected':''}>${esc(prioName(k))}</option>`).join('')}</select></span></label>
      </div>
      <label class="fld" id="sumFld"><span>${esc(t('f.summary'))} <span class="req" aria-hidden="true">*</span></span><input name="title" data-bidi autocomplete="off" required><em class="err" hidden>${esc(t('create.sumReq'))}</em></label>
      <label class="fld">${esc(t('f.desc'))}<textarea name="desc" rows="3" data-bidi placeholder="${esc(t('create.descPh'))}"></textarea></label>
      <div class="grid3">
        <label class="fld">${esc(t('f.start'))}<input type="date" name="start" value="${today()}"></label>
        <label class="fld">${esc(t('f.due'))}<input type="date" name="due" value="${isISO(opts.due) ? opts.due : ''}"></label>
        <label class="fld">${esc(t('f.recur'))}<select name="recur">${['', ...RECUR].map(v=>`<option value="${v}">${esc(t('rc.opt.' + (v || 'none')))}</option>`).join('')}</select></label>
        <label class="fld">${esc(t('f.points'))}<input type="number" name="points" min="0" step="0.5" inputmode="decimal" placeholder="${esc(t('none'))}"></label>
      </div>
      <label class="fld">${esc(t('f.company'))}<input name="labels" data-bidi list="lblList2" value="${esc(normLabel(b.name))}" placeholder="${esc(t('company.commaPh'))}" autocomplete="off"><datalist id="lblList2">${companies.map(l=>`<option value="${esc(l)}">`).join('')}</datalist></label>
      <div class="fld">${esc(t('f.platform'))}<input type="hidden" name="color" value="">${platformPicker('create-color', null)}</div>
    </div>
    <div class="m-foot"><label class="chkline"><input type="checkbox" name="another">${esc(t('create.another'))}</label><span class="sp"></span>
      <button type="button" class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button><button type="submit" class="btn primary">${esc(t('btn.create'))}</button></div>
  </form>`, 'create', 'narrow');
  dlg.setAttribute('aria-label', t('create.title'));
  if (!coarse()) setTimeout(()=>$('[name="title"]',dlg)?.focus(), 60);
}
function submitCreate(form){
  const f = new FormData(form), title = String(f.get('title')||'').trim();
  const fld = $('#sumFld',form);
  if (!title){ fld.classList.add('invalid'); $('.err',fld).hidden = false; $('[name="title"]',form).focus(); return; }
  const pts = String(f.get('points')||'').trim(), start = isISO(f.get('start')) ? f.get('start') : today();
  const another = !!f.get('another'), ws0 = wsId;
  const d = {title, boardId:f.get('board'), type:f.get('type'), status:f.get('status'), priority:f.get('priority'), due:isISO(f.get('due')) ? f.get('due') : '', start,
    points: pts === '' || isNaN(Number(pts)) ? null : Number(pts), desc:String(f.get('desc')||'').trim(), color:isHex(f.get('color')) ? f.get('color') : null,
    labels:[...new Set(String(f.get('labels')||'').split(',').map(normLabel).filter(Boolean))]};
  const recurPick = String(f.get('recur') || '');
  const go = ()=>{
    if (wsId !== ws0 || !canEdit()) return;
    const i = createIssue(d);
    if (recurPick){ i.recur = recurFrom(i, recurPick); i.recurOf = i.id; touch(i, t('h.recur', {rule:recurText(i.recur)})); save(); render(); }
    lastCreateType = i.type; P.lastBoard = P.lastBoard || {}; P.lastBoard[wsId] = i.boardId; savePrefs();
    const onOther = !isAll() && i.boardId !== curBoard()?.id;
    const view = ()=>{ if (!isAll() && i.boardId !== curBoard()?.id) switchBoard(i.boardId); openIssue(i.id); };
    if (!another) openFocusId = i.id;
    toast(onOther ? t('toast.createdIn', {key:key(i), board:boardOf(i).name}) : t('toast.created', {key:key(i)}), {action:t('btn.view'), fn:view});
  };
  if (another){ go(); form.title.value = ''; form.desc.value = ''; autosize(form.desc); form.title.focus(); return; }
  // Close the dialog at once and let that frame paint before the board re-renders. Before, the dialog
  // stayed up while the board rebuilt and then faded out, so it looked open after the task existed.
  closeModal(true);
  requestAnimationFrame(()=>setTimeout(go, 0));
}

/* ----- Logos ----- */
const logoInput = $('#logoFile'); let logoTarget = null;
function pickLogo(target){ logoTarget = target; logoInput.value = ''; logoInput.click(); }
async function readLogo(file){
  if (!file || !/^image\//.test(file.type)) throw new Error('type');
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((ok, bad)=>{ const im = new Image(); im.onload = ()=>ok(im); im.onerror = bad; im.src = url; });
    const w0 = img.naturalWidth || 256, h0 = img.naturalHeight || 256;
    for (const max of [256, 192, 128]){
      const sc = Math.min(1, max / Math.max(w0, h0)), w = Math.max(1, Math.round(w0*sc)), h = Math.max(1, Math.round(h0*sc));
      const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
      cv.getContext('2d').drawImage(img, 0, 0, w, h);
      let out = cv.toDataURL('image/png');
      const webp = cv.toDataURL('image/webp', .92);
      if (webp.startsWith('data:image/webp') && webp.length < out.length) out = webp;
      if (out.length < 150000) return out;
    }
    throw new Error('size');
  } finally { URL.revokeObjectURL(url); }
}
logoInput.addEventListener('change', async ()=>{
  const f = logoInput.files[0]; if (!f || !logoTarget) return;
  try {
    const data = await readLogo(f);
    if (logoTarget === 'board' && bDraft){ bDraft.logo = data; $('#bsBody',layer).innerHTML = boardSettingsBody(); }
    else if (logoTarget === 'ws' && isOwner()){ await store.updateWorkspace(wsId, {logo:data}); toast(t('toast.logoSaved')); }
  } catch(e){ toast(e.message === 'size' ? t('toast.logoBig') : t('toast.logoBad')); }
});

/* ----- Board settings ----- */
let bDraft = null;
function openBoardSettings(boardId){
  if (!canEdit()) return toast(t('toast.viewOnly'));
  const b = boardId ? boardById(boardId) : null;
  bDraft = b ? {isNew:false, id:b.id, name:b.name, key:b.key, color:b.color, logo:b.logo, keyTouched:true, columns:clone(b.columns), removed:{}, err:{}}
             : {isNew:true, id:'b-'+uid(), name:'', key:'', color:PALETTE[(S.boards.length*3+1)%PALETTE.length][1], logo:null, keyTouched:false, columns:COLS_PPC(), removed:{}, err:{}};
  const dlg = showModal(`<form data-form="board" novalidate style="display:contents">
    ${head(bDraft.isNew ? t('bs.createTitle') : t('bs.title'))}
    <div class="c-body" id="bsBody">${boardSettingsBody()}</div>
    <div class="m-foot">${!bDraft.isNew ? `<button type="button" class="btn danger-subtle" data-act="del-board" data-id="${bDraft.id}">${I.trash}${esc(t('bs.delete'))}</button>` : ''}<span class="sp"></span>
      <button type="button" class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button><button type="submit" class="btn primary">${esc(bDraft.isNew ? t('btn.createBoard') : t('btn.saveChanges'))}</button></div>
  </form>`, 'board', 'narrow');
  dlg.setAttribute('aria-label', bDraft.isNew ? t('bs.createTitle') : t('bs.title'));
  if (!coarse()) setTimeout(()=>$('[data-k="bname"]',dlg)?.focus(), 60);
}
const colIssueCount = cid=>S.issues.filter(i=>i.boardId===bDraft.id && i.status===cid).length;
function boardSettingsBody(){
  const d = bDraft;
  return `
    <div class="grid-bs">
      <label class="fld${d.err.name?' invalid':''}"><span>${esc(t('bs.name'))} <span class="req" aria-hidden="true">*</span></span><input data-k="bname" value="${esc(d.name)}" data-bidi placeholder="${esc(t('bs.namePh'))}" autocomplete="off">${d.err.name?`<em class="err">${esc(d.err.name)}</em>`:''}</label>
      <label class="fld${d.err.key?' invalid':''}"><span>${esc(t('bs.key'))} <span class="req" aria-hidden="true">*</span></span><input class="key-in" data-k="bkey" value="${esc(d.key)}" maxlength="10" placeholder="ABC" dir="ltr" autocapitalize="characters" autocomplete="off">${d.err.key?`<em class="err">${esc(d.err.key)}</em>`:''}</label>
    </div>
    <p class="req-note" id="keyHint" style="margin-top:-8px">${t('bs.keyHint', {ex:`<b class="key">${esc(d.key||'KEY')}-1</b>`})}</p>
    <div class="logo-row">
      <div class="fld">${esc(t('bs.logo'))}<div class="logo-edit"><span class="logo-prev">${isImg(d.logo) ? `<img src="${esc(d.logo)}" alt="">` : `<span class="b-av big" style="--bc:${d.color}">${initialOf(d.name||'?')}</span>`}</span>
        <button type="button" class="btn" data-act="bd-logo">${I.image}${esc(isImg(d.logo) ? t('logo.replace') : t('logo.upload'))}</button>${isImg(d.logo) ? `<button type="button" class="btn subtle" data-act="bd-logo-del">${esc(t('logo.remove'))}</button>` : ''}</div>
        <span class="ed-hint">${esc(t('logo.hint'))}</span></div>
    </div>
    <div class="fld">${esc(t('bs.color'))}${swatches('bd-color', d.color, {label:t('bs.color')})}</div>
    <div class="ed-sec">
      <div class="ed-h"><span>${esc(t('bs.columns'))}</span><span class="ed-hint">${esc(t('bs.colHint'))}</span></div>
      <ol class="ed-list" id="colList">${d.columns.map((c,ix)=>colRowHTML(c,ix)).join('')}</ol>
      ${d.err.cols?`<em class="ed-err">${esc(d.err.cols)}</em>`:''}
      <button type="button" class="btn subtle ed-add" data-act="bd-addcol">${I.plus}${esc(t('bs.addCol'))}</button>
    </div>`;
}
function colRowHTML(c, ix){
  const d = bDraft, n = d.isNew ? 0 : colIssueCount(c.id);
  if (c._confirm){
    const others = d.columns.filter(x=>x.id!==c.id);
    return `<li class="ed-row confirm" data-cid="${c.id}"><p>${esc(t('bs.removeColQ', {name:c.name||'—', n:tn('issues', n)}))}</p>
      <label>${esc(t('bs.moveTo'))}<select class="inp" data-k="cmove-${c.id}">${others.map(o=>`<option value="${o.id}"${o.id===c._target?' selected':''}>${esc(o.name||'—')}</option>`).join('')}</select></label>
      <button type="button" class="btn danger sm" data-act="bd-delcol-yes" data-k="cyes-${c.id}">${esc(t('bs.removeCol'))}</button><button type="button" class="btn subtle sm" data-act="bd-delcol-no">${esc(t('btn.keep'))}</button></li>`;
  }
  const last = d.columns.length - 1;
  return `<li class="ed-row" data-cid="${c.id}">
    <span class="ed-pos">${ix+1}</span>
    <input class="inp ce-name" data-k="cname-${c.id}" value="${esc(c.name)}" data-bidi placeholder="${esc(t('bs.colName'))}" aria-label="${esc(t('bs.colName'))}" autocomplete="off">
    <select class="inp" data-k="ccat-${c.id}" aria-label="${esc(t('bs.category'))}">${CATS.map(k=>`<option value="${k}"${k===c.cat?' selected':''}>${esc(catName(k))}</option>`).join('')}</select>
    <span class="loz ${c.cat}" data-loz>${esc(catName(c.cat))}</span>
    <button type="button" class="btn icon subtle sm" data-act="bd-colup" data-k="cup-${c.id}" aria-label="${esc(t('bs.earlier'))}" title="${esc(t('bs.earlier'))}"${ix===0?' disabled':''}>${I.aup}</button>
    <button type="button" class="btn icon subtle sm" data-act="bd-coldown" data-k="cdn-${c.id}" aria-label="${esc(t('bs.later'))}" title="${esc(t('bs.later'))}"${ix===last?' disabled':''}>${I.adown}</button>
    <button type="button" class="btn icon subtle sm" data-act="bd-delcol" data-k="cdel-${c.id}" aria-label="${esc(t('bs.removeCol'))}" title="${esc(d.columns.length===1?t('bs.oneCol'):t('bs.removeCol'))}"${d.columns.length===1?' disabled':''}>${I.trash}</button>
  </li>`;
}
function rerenderCols(animate){
  const list = $('#colList',layer); if (!list) return;
  const rows = $$('.ed-row',list);
  const doIt = ()=>rerender(list, bDraft.columns.map((c,ix)=>colRowHTML(c,ix)).join(''));
  if (animate){
    const r0 = new Map(rows.map(r=>[r.dataset.cid, r.getBoundingClientRect().top]));
    doIt();
    if (!reduce) $$('.ed-row',list).forEach(r=>{ const a = r0.get(r.dataset.cid); if (a === undefined) return; const dy = a - r.getBoundingClientRect().top;
      if (Math.abs(dy) > .5) r.animate([{transform:`translateY(${dy}px)`},{transform:'none'}],{duration:200, easing:'cubic-bezier(.2,.8,.2,1)'}); });
  } else doIt();
}
function saveBoard(){
  const d = bDraft; d.err = {};
  d.name = d.name.trim(); d.key = d.key.trim().toUpperCase();
  if (!d.name) d.err.name = t('bs.errName');
  if (!/^[A-Z][A-Z0-9]{0,9}$/.test(d.key)) d.err.key = t('bs.errKey');
  else if (S.boards.some(b=>b.id!==d.id && b.key===d.key)) d.err.key = t('bs.errKeyUsed');
  if (d.columns.some(c=>c._confirm)) d.err.cols = t('bs.errFinish');
  d.columns.forEach(c=>{ c.name = c.name.trim(); });
  if (d.columns.some(c=>!c.name)) d.err.cols = t('bs.errColName');
  if (Object.keys(d.err).length){
    $('#bsBody',layer).innerHTML = boardSettingsBody();
    $('.invalid input, .ed-row .inp:placeholder-shown',layer)?.focus();
    return;
  }
  const cols = d.columns.map(({id,name,cat})=>({id,name,cat}));
  if (d.isNew){
    S.boards.push({id:d.id, name:d.name, key:d.key, color:d.color, logo:d.logo||null, seq:0, order:undefined, columns:cols, tc:TC});
    P.active[wsId] = d.id; P.lastBoard = P.lastBoard || {}; P.lastBoard[wsId] = d.id; savePrefs();
    composer = null; save(); closeModal(); render();
    toast(t('toast.boardCreated', {name:d.name}));
    return;
  }
  const b = boardById(d.id), oldName = b.name;
  const exists = new Set(cols.map(c=>c.id));
  const resolve = id=>{ let x = id, guard = 0; while (!exists.has(x) && guard++ < 50) x = d.removed[x]; return exists.has(x) ? x : cols[0].id; };
  const names = Object.fromEntries([...b.columns, ...cols].map(c=>[c.id,c.name]));
  let moved = 0;
  S.issues.filter(i=>i.boardId===b.id && !exists.has(i.status)).forEach(i=>{
    const to = resolve(i.status); touch(i, t('h.colRemoved', {from:names[i.status]||'—', to:names[to]})); i.status = to; moved++;
  });
  if (oldName !== d.name) S.issues.forEach(i=>{ if (i.boardId === b.id) i.labels = i.labels.map(l=>sameText(l, oldName) ? normLabel(d.name) : l); });
  Object.assign(b, {name:d.name, key:d.key, color:d.color, logo:d.logo||null, columns:cols});
  S.issues.forEach(i=>{ if (i.boardId !== b.id) return; const done = isDone(i); if (done && !i.doneAt) i.doneAt = Date.now(); if (!done) i.doneAt = null; });
  save(); closeModal(); render();
  toast(moved ? t('toast.boardSavedMoved', {n:tn('issues', moved)}) : t('toast.boardSaved'));
}
function deleteBoard(id){
  const b = boardById(id); if (!b) return;
  const n = S.issues.filter(i=>i.boardId===id).length;
  openConfirm({title:t('bs.deleteQ'), ok:t('bs.delete'),
    text:esc(n ? t('bs.deleteTextN', {name:b.name, n:tn('issues', n)}) : t('bs.deleteText', {name:b.name})),
    fn:()=>{
      const bix = S.boards.indexOf(b), removed = S.issues.filter(i=>i.boardId===id), wasActive = activeKey() === id;
      S.boards.splice(bix,1); S.issues = S.issues.filter(i=>i.boardId!==id);
      if (wasActive) P.active[wsId] = S.boards[Math.max(0,bix-1)]?.id || ALL;
      savePrefs(); save(); closeModal(); render();
      toast(t('toast.boardDeleted', {name:b.name}), {action:t('btn.undo'), fn:()=>{ S.boards.splice(bix,0,b); S.issues.push(...removed); removed.forEach(i=>{ i.rank = undefined; }); if (wasActive) P.active[wsId] = b.id; save(); render(); }});
    }});
}

/* ----- Issue types ----- */
let tDraft = null;
function openTypes(){
  if (!canEdit()) return toast(t('toast.viewOnly'));
  tDraft = {types:clone(S.types), removed:{}, err:''};
  const dlg = showModal(`<form data-form="types" novalidate style="display:contents">
    ${head(t('types.title'))}
    <div class="c-body"><p class="req-note">${esc(t('types.hint'))}</p><div class="ed-sec" id="tyBody">${typesBody()}</div></div>
    <div class="m-foot"><span class="sp"></span><button type="button" class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button><button type="submit" class="btn primary">${esc(t('btn.saveChanges'))}</button></div>
  </form>`, 'types', 'narrow');
  dlg.setAttribute('aria-label', t('types.title'));
}
const typeCount = id=>S.issues.filter(i=>i.type===id).length;
function typesBody(){
  const d = tDraft;
  return `<ol class="ed-list" id="tyList">${d.types.map(ty=>{
    if (ty._confirm){
      const n = typeCount(ty.id), others = d.types.filter(x=>x.id!==ty.id);
      return `<li class="ed-row confirm" data-tid="${ty.id}"><p>${esc(t('types.removeQ', {name:ty.name||'—', n:tn('issues', n)}))}</p>
        <label>${esc(t('types.changeTo'))}<select class="inp" data-k="tmove-${ty.id}">${others.map(o=>`<option value="${o.id}"${o.id===ty._target?' selected':''}>${esc(o.name||'—')}</option>`).join('')}</select></label>
        <button type="button" class="btn danger sm" data-act="ty-del-yes" data-k="tyes-${ty.id}">${esc(t('types.remove'))}</button><button type="button" class="btn subtle sm" data-act="ty-del-no">${esc(t('btn.keep'))}</button></li>`;
    }
    const n = typeCount(ty.id);
    return `<li class="ed-row${ty._new?' in':''}" data-tid="${ty.id}">
      <button type="button" class="ty-ico" data-act="ty-style" data-k="tsty-${ty.id}" aria-label="${esc(t('types.style'))}" aria-haspopup="menu" title="${esc(t('types.style'))}">${typeSVG(ty)}</button>
      <input class="inp" data-k="tname-${ty.id}" value="${esc(ty.name)}" data-bidi placeholder="${esc(t('types.namePh'))}" aria-label="${esc(t('types.namePh'))}" autocomplete="off">
      <span class="ty-count">${esc(n ? tn('issues', n) : t('types.unused'))}</span>
      <button type="button" class="btn icon subtle sm" data-act="ty-del" data-k="tdel-${ty.id}" aria-label="${esc(t('types.remove'))}" title="${esc(d.types.length===1?t('types.oneType'):t('types.remove'))}"${d.types.length===1?' disabled':''}>${I.trash}</button></li>`;
  }).join('')}</ol>${d.err?`<em class="ed-err">${esc(d.err)}</em>`:''}
  <button type="button" class="btn subtle ed-add" data-act="ty-add">${I.plus}${esc(t('types.add'))}</button><button type="button" class="btn subtle ed-add" data-act="ty-preset">${I.board}${esc(t('types.preset'))}</button>`;
}
function typeStylePop(ty){
  return `<div class="menu-h">${esc(t('f.color'))}</div>${swatches('ty-color', ty.color, {sm:true})}
    <div class="menu-h">${esc(t('types.symbol'))}</div><div class="glyph-grid">${Object.keys(GLYPHS).map(g=>`<button type="button" class="glyph-btn${ty.glyph===g?' on':''}" data-act="ty-glyph" data-g="${g}" aria-label="${g}" aria-pressed="${ty.glyph===g}">${typeSVG({...ty, glyph:g, name:g})}</button>`).join('')}</div>`;
}
function saveTypes(){
  const d = tDraft; d.err = '';
  d.types.forEach(x=>{ x.name = x.name.trim(); });
  if (d.types.some(x=>x._confirm)) d.err = t('types.errFinish');
  else if (d.types.some(x=>!x.name)) d.err = t('types.errName');
  if (d.err){ $('#tyBody',layer).innerHTML = typesBody(); return; }
  const types = d.types.map(({id,name,color,glyph})=>({id,name,color,glyph}));
  const exists = new Set(types.map(x=>x.id));
  const resolve = id=>{ let x = id, g = 0; while (!exists.has(x) && g++ < 50) x = d.removed[x]; return exists.has(x) ? x : types[0].id; };
  const names = Object.fromEntries([...S.types, ...types].map(x=>[x.id,x.name]));
  S.issues.filter(i=>!exists.has(i.type)).forEach(i=>{ const to = resolve(i.type); touch(i, t('h.typeRemoved', {from:names[i.type], to:names[to]})); i.type = to; });
  S.types = types;
  save(); closeModal(); render(); toast(t('toast.typesSaved'));
}

/* ----- Workspace settings: name, logo, people, requests ----- */
function openWsSettings(){
  if (!isOwner()) return toast(t('toast.ownerOnly'));
  const dlg = showModal(`${head(t('ws.settings'))}<div class="c-body" id="wsBody"></div>
    <div class="m-foot"><span class="sp"></span><button type="button" class="btn primary" data-act="close-modal">${esc(t('btn.done'))}</button></div>`, 'wsSettings', 'narrow');
  dlg.setAttribute('aria-label', t('ws.settings'));
  renderWsSettingsBody();
}
function renderWsSettingsBody(){
  const body = $('#wsBody', layer); if (!body || !ws) return;
  const roles = ws.roles || {}, members = Object.keys(roles).sort();
  const pend = incoming.filter(r=>r.status==='pending');
  const siteUrl = location.origin + location.pathname;
  rerender(body, `
    ${pend.length ? `<section class="ws-sec req-sec"><h3>${I.bell}${esc(t('req.title'))} <span class="pill">${pend.length}</span></h3>
      <ul class="people">${pend.map(r=>`<li>${avatar({name:r.fromName, email:r.fromEmail, photo:r.fromPhoto})}<div class="p-txt"><b data-bidi>${esc(r.fromName||r.fromEmail)}</b><span dir="ltr">${esc(r.fromEmail)}</span><small>${esc(ago(r.created||Date.now()))}</small></div>
        <div class="p-acts"><button class="btn primary sm" data-act="req-approve" data-id="${esc(r.id)}" data-role="editor">${esc(t('req.approveEdit'))}</button><button class="btn sm" data-act="req-approve" data-id="${esc(r.id)}" data-role="viewer">${esc(t('req.approveView'))}</button><button class="btn subtle sm" data-act="req-decline" data-id="${esc(r.id)}">${esc(t('req.decline'))}</button></div></li>`).join('')}</ul></section>` : ''}
    <section class="ws-sec">
      <label class="fld">${esc(t('ws.name'))}<input class="inp" data-k="wsname" maxlength="100" value="${esc(ws.name)}" data-bidi autocomplete="off"></label>
      <div class="fld">${esc(t('ws.logo'))}<div class="logo-edit"><span class="logo-prev wide">${isImg(ws.logo) ? `<img src="${esc(ws.logo)}" alt="">` : `<span class="logo big">${APP_MARK}</span>`}</span>
        <button type="button" class="btn" data-act="ws-logo">${I.image}${esc(isImg(ws.logo) ? t('logo.replace') : t('logo.upload'))}</button>${isImg(ws.logo) ? `<button type="button" class="btn subtle" data-act="ws-logo-del">${esc(t('logo.remove'))}</button>` : ''}</div>
        <span class="ed-hint">${esc(t('ws.logoHint'))}</span></div>
    </section>
    <section class="ws-sec">
      <h3>${I.users}${esc(t('share.title'))}</h3>
      <p class="req-note">${esc(t('share.hint'))}</p>
      <form class="share-form" data-form="share" novalidate>
        <input class="inp" name="email" type="email" inputmode="email" dir="ltr" placeholder="name@example.com" aria-label="${esc(t('share.email'))}" autocomplete="off">
        <select class="inp" name="role" aria-label="${esc(t('share.role'))}"><option value="editor">${esc(t('role.editor'))}</option><option value="viewer">${esc(t('role.viewer'))}</option></select>
        <button class="btn primary">${esc(t('share.add'))}</button>
      </form>
      <em class="ed-err" id="shareErr" hidden></em>
      <ul class="people">
        <li>${avatar({name:ws.ownerName, email:ws.ownerEmail, photo:ws.ownerUid===user.uid?user.photo:''})}<div class="p-txt"><b data-bidi>${esc(ws.ownerName||ws.ownerEmail)}</b><span dir="ltr">${esc(ws.ownerEmail)}</span></div><span class="role-tag">${esc(t('role.owner'))}</span></li>
        ${members.map(m=>`<li>${avatar({email:m, name:m})}<div class="p-txt"><span dir="ltr" class="p-mail">${esc(m)}</span></div>
          <select class="inp role-sel" data-k="role-${esc(m)}" data-email="${esc(m)}" aria-label="${esc(t('share.role'))}"><option value="editor"${roles[m]==='editor'?' selected':''}>${esc(t('role.editor'))}</option><option value="viewer"${roles[m]==='viewer'?' selected':''}>${esc(t('role.viewer'))}</option></select>
          <a class="btn icon subtle sm" href="mailto:${esc(m)}?subject=${encodeURIComponent(t('mail.subject', {ws:ws.name}))}&body=${encodeURIComponent(t('mail.body', {ws:ws.name, url:siteUrl, email:m}))}" aria-label="${esc(t('share.mail'))}" title="${esc(t('share.mail'))}">${I.mail}</a>
          <button class="btn icon subtle sm" data-act="member-remove" data-email="${esc(m)}" aria-label="${esc(t('share.remove'))}" title="${esc(t('share.remove'))}">${I.x}</button></li>`).join('')}
      </ul>
      ${!members.length ? `<p class="muted-p">${esc(t('share.none'))}</p>` : ''}
      <div class="copy-link"><span dir="ltr">${esc(siteUrl)}</span><button class="btn sm" data-act="copy-site">${I.copy}${esc(t('share.copyLink'))}</button></div>
    </section>
    <section class="ws-sec">
      <h3>${I.archive}${esc(t('ap.setTitle'))}</h3>
      <p class="req-note">${esc(t('ap.setHint'))}</p>
      <label class="chk-row"><input type="checkbox" data-k="ap-on"${S.archive?.on ? ' checked' : ''}><span>${esc(t('ap.setOn'))}</span></label>
      <label class="fld ap-day">${esc(t('ap.setDay'))}<select class="inp" data-k="ap-day"${S.archive?.on ? '' : ' disabled'}>${Array.from({length:28}, (_,k)=>k+1).map(d=>`<option value="${d}"${d===S.archive?.day?' selected':''}>${d}</option>`).join('')}</select></label>
    </section>`);
}
async function setMember(email, role){
  try { await store.setMember(wsId, email, role); return true; }
  catch(e){ console.error(e); toast(t('toast.saveFailed')); return false; }
}
let wsNameTimer = 0;

/* ----- Request access ----- */
function openRequestAccess(){
  const mine = outgoing.slice().sort((a,b)=>(b.created||0)-(a.created||0));
  const dlg = showModal(`${head(t('ra.title'))}<div class="c-body">
      <p class="req-note">${esc(t('ra.hint'))}</p>
      <form class="cr-form" data-form="request" novalidate><input class="inp" name="email" type="email" inputmode="email" dir="ltr" placeholder="name@example.com" aria-label="${esc(t('welcome.askEmail'))}" autocomplete="off"><button class="btn primary">${esc(t('btn.sendRequest'))}</button></form>
      <em class="ed-err" id="reqErr" hidden></em>
      ${mine.length ? `<ul class="people">${mine.map(r=>`<li>${I.mail}<div class="p-txt"><span dir="ltr" class="p-mail">${esc(r.toEmail)}</span><small>${esc(ago(r.created||Date.now()))}</small></div><span class="role-tag ${esc(r.status)}">${esc(t('ra.status.'+r.status))}</span><button class="btn icon subtle sm" data-act="req-cancel" data-id="${esc(r.id)}" aria-label="${esc(t('btn.cancelRequest'))}" title="${esc(t('btn.cancelRequest'))}">${I.x}</button></li>`).join('')}</ul>` : ''}
    </div>`, 'request', 'narrow');
  dlg.setAttribute('aria-label', t('ra.title'));
}
async function sendRequest(form){
  const email = lower(new FormData(form).get('email')), err = $('#reqErr', form.parentElement);
  const fail = msg=>{ err.textContent = msg; err.hidden = false; form.email.focus(); };
  if (!isEmail(email)) return fail(t('ra.errEmail'));
  if (email === user.email) return fail(t('ra.errSelf'));
  if (outgoing.some(r=>r.toEmail===email && r.status==='pending')) return fail(t('ra.errDup'));
  try {
    await store.createRequest({fromUid:user.uid, fromEmail:user.email, fromName:String(user.name||'').slice(0,200), fromPhoto:String(user.photo||'').slice(0,500), toEmail:email, status:'pending', created:Date.now()});
    form.email.value = ''; err.hidden = true;
    toast(t('toast.requestSent', {email}));
    if (modalKind === 'request') openRequestAccess();
  } catch(e){ console.error(e); fail(t('toast.saveFailed')); }
}

/* ----- Shortcuts ----- */
function openShortcuts(){
  const k = [['C','sc.create'],['/','sc.search'],['B','sc.board'],['L','sc.list'],['T','sc.table'],['K','sc.cal'],['D','sc.dash'],['A','sc.all'],['Enter','sc.open'],['Shift + ← →','sc.move'],['↑ ↓','sc.nav'],['Ctrl + Enter','sc.comment'],['Esc','sc.close'],['?','sc.help']];
  showModal(`${head(t('sc.title'))}<dl class="kbd-list">${k.map(([a,b])=>`<dt dir="ltr">${a.split(' ').map(x=>x==='+'?' + ':`<kbd>${x}</kbd>`).join('')}</dt><dd>${esc(t(b))}</dd>`).join('')}</dl>`, 'help', 'tiny');
}

/* ----- Modal events ----- */
let focusVal = {};
layer.addEventListener('focusin', e=>{ const f = e.target.dataset?.f; if (f === 'title' || f === 'desc') focusVal[f] = e.target.value; });
layer.addEventListener('input', e=>{
  const el = e.target, k = el.dataset?.k || '';
  if (el.tagName === 'TEXTAREA') autosize(el);
  if (modalKind === 'create'){ if (el.name === 'title'){ const f = $('#sumFld',layer); f.classList.remove('invalid'); $('.err',f).hidden = true; } return; }
  if (modalKind === 'board'){
    if (k === 'bname'){ bDraft.name = el.value; if (!bDraft.keyTouched){ bDraft.key = keyFromName(el.value); $('[data-k="bkey"]',layer).value = bDraft.key; } }
    else if (k === 'bkey'){ bDraft.keyTouched = true; bDraft.key = el.value.toUpperCase().replace(/[^A-Z0-9]/g,''); }
    else if (k.startsWith('cname-')){ const c = bDraft.columns.find(c=>c.id===k.slice(6)); if (c) c.name = el.value; }
    if (k === 'bname' || k === 'bkey'){ const hint = $('#keyHint b',layer); if (hint) hint.textContent = `${bDraft.key||'KEY'}-1`; }
    return;
  }
  if (modalKind === 'types'){
    if (k.startsWith('tname-')){ const ty = tDraft.types.find(x=>x.id===k.slice(6)); if (ty){ ty.name = el.value; el.parentElement.querySelector('.ty-ico').innerHTML = typeSVG(ty); } }
    return;
  }
  if (modalKind === 'wsSettings'){
    if (k === 'wsname'){ clearTimeout(wsNameTimer); const v = el.value.trim().slice(0,100); wsNameTimer = setTimeout(()=>{ if (v && v !== ws.name) store.updateWorkspace(wsId, {name:v}).catch(()=>toast(t('toast.saveFailed'))); }, 600); }
    if (el.name === 'email') $('#shareErr',layer).hidden = true;
    return;
  }
  if (modalKind === 'request'){ $('#reqErr',layer).hidden = true; return; }
  if (!canEdit()) return;
  const i = get(openId); if (!i) return;
  if (el.dataset.f === 'title' || el.dataset.f === 'desc'){ i[el.dataset.f] = el.value; i.updated = Date.now(); save(); scheduleRender(); }
  else if (el.classList.contains('st-text')){ const s = i.subtasks.find(s=>s.id===el.closest('li').dataset.sid); if (s){ s.text = el.value; save(); scheduleRender(); } }
  else if (k === 'cmt') cmtDraft = el.value;
  else if (k === 'l-url') $('#linkErr',layer).hidden = true;
});
layer.addEventListener('change', e=>{
  const el = e.target, k = el.dataset?.k || '';
  if (modalKind === 'create'){
    if (el.name === 'type') el.previousElementSibling.outerHTML = typeIcon(el.value);
    if (el.name === 'priority') el.previousElementSibling.outerHTML = prioIcon(el.value);
    if (el.name === 'board'){
      const b = boardById(el.value), st = el.form.status, prevCol = st.selectedOptions[0]?.textContent || '';
      st.innerHTML = b.columns.map(c=>`<option value="${c.id}"${c.name===prevCol?' selected':''}>${esc(c.name)}</option>`).join('');
      const lab = el.form.labels; if (sameText(lab.value, el.dataset.prev) || !lab.value.trim()) lab.value = normLabel(b.name);
      el.dataset.prev = b.name;
    }
    return;
  }
  if (modalKind === 'board'){
    if (k.startsWith('ccat-')){ const c = bDraft.columns.find(c=>c.id===k.slice(5)); c.cat = el.value; const l = el.parentElement.querySelector('[data-loz]'); l.className = `loz ${c.cat}`; l.textContent = catName(c.cat); }
    if (k.startsWith('cmove-')){ const c = bDraft.columns.find(c=>c.id===k.slice(6)); c._target = el.value; }
    return;
  }
  if (modalKind === 'types'){ if (k.startsWith('tmove-')){ const ty = tDraft.types.find(x=>x.id===k.slice(6)); ty._target = el.value; } return; }
  if (modalKind === 'wsSettings'){
    if (k === 'ap-on' || k === 'ap-day'){
      S.archive = {...normArchive(S.archive), [k === 'ap-on' ? 'on' : 'day']:k === 'ap-on' ? el.checked : Number(el.value)};
      if (k === 'ap-on') { const sel = $('[data-k="ap-day"]', layer); if (sel) sel.disabled = !el.checked; }
      save(); render(); toast(t('ap.saved'));
    }
    if (el.classList.contains('role-sel')){ const em = el.dataset.email, r = el.value; setMember(em, r).then(ok=>{ if (ok) toast(t('toast.roleChanged', {email:em, role:t('role.'+r)})); }); }
    return;
  }
  if (!canEdit()) return;
  const i = get(openId); if (!i) return;
  const f = el.dataset.f;
  if (f === 'title' || f === 'desc'){
    if (focusVal[f] !== undefined && focusVal[f] !== el.value){ touch(i, f === 'title' ? t('h.summary') : t('h.desc')); save(); }
    focusVal[f] = el.value; return;
  }
  if (f === 'type'){ i.type = el.value; touch(i, t('h.type', {name:typeOf(i.type).name})); }
  else if (f === 'priority'){ i.priority = el.value; touch(i, t('h.priority', {name:prioName(i.priority)})); }
  else if (f === 'due' || f === 'start'){
    // Safari fills today's date and fires "change" the moment its picker opens. Rebuilding the
    // dialog then would close the picker, so while the field still has focus only patch around it.
    if (f === 'due') i.due = isISO(el.value) ? el.value : '';
    else i.start = isISO(el.value) ? el.value : today();
    const text = f === 'due' ? (i.due ? t('h.due', {date:fmtDay(i.due)}) : t('h.dueNone')) : t('h.start', {date:fmtDay(i.start)});
    const live = document.activeElement === el;
    const h = dateEdit && dateEdit.el === el && i.history.find(x=>x.id === dateEdit.hid);
    if (h){ h.text = text; h.at = Date.now(); i.updated = Date.now(); }
    else { touch(i, text); if (live) dateEdit = {el, hid:i.history[0].id}; }
    save(); render();
    if (live) patchIssueDates(i); else refreshIssue();
    return;
  }
  else if (f === 'recur' || f === 'recur-wd' || f === 'recur-md'){
    if (f === 'recur'){
      if (!el.value){ delete i.recur; touch(i, t('h.recurOff')); }
      else { i.recur = recurFrom(i, el.value); i.recurOf = i.recurOf || i.id; touch(i, t('h.recur', {rule:recurText(i.recur)})); }
    } else if (i.recur){
      const r = {...i.recur, [f === 'recur-wd' ? 'wd' : 'md']:Number(el.value)};
      r.next = recurFirst(r, (isISO(i.start) && i.start > today()) ? i.start : today());
      i.recur = normRecur(r); touch(i, t('h.recur', {rule:recurText(i.recur)}));
    }
    save(); render(); refreshIssue(); runRecurring(); return;
  }
  else if (f === 'points'){ const v = el.value.trim(); i.points = v === '' || isNaN(Number(v)) ? null : Number(v); touch(i, i.points == null ? t('h.pointsNone') : t('h.points', {n:i.points})); }
  else if (f === 'board'){ moveToBoard(i, el.value); return; }
  else return;
  save(); render(); refreshIssue();
});
let dateEdit = null;
function patchIssueDates(i){
  const dd = $('[data-f="due"]',layer)?.parentElement;
  if (dd){
    const btn = $('[data-act="clear-due"]',dd);
    if (i.due && !btn) dd.insertAdjacentHTML('beforeend', `<button class="btn icon subtle sm" data-act="clear-due" aria-label="${esc(t('f.clearDue'))}" title="${esc(t('f.clearDue'))}">${I.x}</button>`);
    else if (!i.due && btn) btn.remove();
  }
  const ag = $('.side-top .age',layer), age = ageInfo(i);
  if (ag){ ag.className = `age big ${age.cls}`; ag.title = age.long; ag.innerHTML = `${I.clock}${esc(age.long)}`; }
}
layer.addEventListener('focusout', e=>{
  if (e.target.type !== 'date' || !dateEdit || dateEdit.el !== e.target) return;
  dateEdit = null;
});
function addLabel(input){
  const i = get(openId), l = normLabel(input.value);
  input.value = '';
  if (!l || i.labels.some(x=>sameText(x,l))) return;
  i.labels.push(l); touch(i, t('h.companyAdd', {name:l})); save(); render(); refreshIssue();
  $('.lbl-in',layer)?.focus();
}
layer.addEventListener('keydown', e=>{
  const el = e.target;
  if (['create','board','types'].includes(modalKind)){
    if (e.key === 'Enter' && (e.ctrlKey||e.metaKey)){ e.preventDefault(); $('form',layer)?.requestSubmit(); }
    else if (e.key === 'Enter' && el.tagName === 'INPUT' && (modalKind === 'board' || modalKind === 'types') && (el.dataset.k||'').match(/^(cname|tname)-/)){ e.preventDefault(); }
    return;
  }
  if (modalKind !== 'issue' || !canEdit()) return;
  if (el.classList.contains('sum') && e.key === 'Enter' && !e.isComposing){ e.preventDefault(); el.blur(); }
  if (el.classList.contains('st-text') && e.key === 'Enter' && !e.isComposing){ e.preventDefault(); el.blur(); }
  if (el.classList.contains('lbl-in')){
    if ((e.key === 'Enter' || e.key === ',') && !e.isComposing){ e.preventDefault(); addLabel(el); }
    else if (e.key === 'Backspace' && !el.value){ const i = get(openId); const l = i.labels.pop(); if (l){ touch(i, t('h.companyRemove', {name:l})); save(); render(); refreshIssue(); } }
  }
  if (el.dataset.k === 'cmt' && e.key === 'Enter' && (e.ctrlKey||e.metaKey)){ e.preventDefault(); saveComment(); }
  if (el.dataset.k === 'cmt-edit'){
    if (e.key === 'Enter' && (e.ctrlKey||e.metaKey)){ e.preventDefault(); updateComment(el); }
    if (e.key === 'Escape'){ e.stopPropagation(); editingCid = null; refreshIssue(); }
  }
});
layer.addEventListener('focusout', e=>{ if (e.target.classList?.contains('lbl-in') && e.target.value.trim() && !e.relatedTarget?.closest?.('.labels-edit')) addLabel(e.target); });
layer.addEventListener('submit', e=>{
  e.preventDefault();
  const form = e.target;
  if (form.dataset.form === 'request') return sendRequest(form);
  if (form.dataset.form === 'share') return shareSubmit(form);
  if (!canEdit()) return;
  if (form.dataset.form === 'create') return submitCreate(form);
  if (form.dataset.form === 'board') return saveBoard();
  if (form.dataset.form === 'types') return saveTypes();
  const i = get(openId); if (!i) return;
  if (form.dataset.form === 'link'){
    const ui = $('[data-k="l-url"]',form), ti = $('[data-k="l-title"]',form), url = normUrl(ui.value);
    if (!url){ $('#linkErr',layer).hidden = false; ui.focus(); return; }
    const l = {id:uid(), url, title:ti.value.trim()};
    i.links.push(l); touch(i, t('h.linkAdd', {name:l.title || hostOf(url)})); save(); render();
    $('.links',layer).outerHTML = linksHTML(i, true);
    $(`.links li[data-lid="${l.id}"]`,layer)?.classList.add('in');
    const sh = $('.links',layer).previousElementSibling; sh.innerHTML = `<h3>${esc(t('f.links'))}</h3><span class="sec-meta">${i.links.length}</span>`;
    ui.value = ''; ti.value = ''; ui.focus();
    return;
  }
  if (form.dataset.form === 'subtask'){
    const inp = $('input',form), v = inp.value.trim(); if (!v) return;
    const s = {id:uid(), text:v, done:false};
    i.subtasks.push(s); touch(i, t('h.stAdd', {name:v})); save(); render();
    if (i.subtasks.length > 1){
      $('.sts',layer).insertAdjacentHTML('beforeend', stHTML(s, true, 'in'));
      $('#stMeta',layer).textContent = t('st.pct', {n:pctOf(i)}); setBar(pctOf(i));
    } else refreshIssue();
    inp.value = ''; $('[data-k="st-add"]',layer).focus();
  }
});
async function shareSubmit(form){
  const f = new FormData(form), email = lower(f.get('email')), r = f.get('role') === 'viewer' ? 'viewer' : 'editor', err = $('#shareErr',layer);
  const fail = m=>{ err.textContent = m; err.hidden = false; form.email.focus(); };
  if (!isEmail(email)) return fail(t('ra.errEmail'));
  if (email === lower(ws.ownerEmail)) return fail(t('share.errOwner'));
  const had = email in (ws.roles||{});
  if (!await setMember(email, r)) return;
  const pendingFrom = incoming.find(x=>x.status==='pending' && lower(x.fromEmail)===email);
  if (pendingFrom) store.updateRequest(pendingFrom.id, {status:'approved', role:r, wsId, wsName:ws.name, decidedAt:Date.now()}).catch(()=>{});
  form.email.value = '';
  toast(had ? t('toast.roleChanged', {email, role:t('role.'+r)}) : t('toast.shared', {email}), had ? {} : {action:t('share.mail'), fn:()=>{ location.href = `mailto:${email}?subject=${encodeURIComponent(t('mail.subject', {ws:ws.name}))}&body=${encodeURIComponent(t('mail.body', {ws:ws.name, url:location.origin+location.pathname, email}))}`; }});
}
function saveComment(){
  const i = get(openId), ta = $('[data-k="cmt"]',layer), v = ta.value.trim();
  if (!v){ ta.focus(); return; }
  const c = {id:uid(), text:v, at:Date.now(), by:me()};
  i.comments.unshift(c); touch(i); cmtDraft = ''; fx.cmt = c.id; save(); render(); refreshIssue();
  $('[data-k="cmt"]',layer)?.blur();
}
function updateComment(ta){
  const i = get(openId), c = i.comments.find(c=>c.id===editingCid), v = ta.value.trim();
  if (c && v){ c.text = v; c.edited = true; touch(i); save(); render(); }
  editingCid = null; refreshIssue();
}

/* ================= Global click actions ================= */
const EDIT_ACTS = new Set(['trk-apply','cal-add','s-menu','s-set','ap-yes','ap-no','imp-add','imp-replace','new-board','board-settings','del-board','types','create','add-inline','set-status','set-color','archive-issue','archive-col','restore','del-issue','clear-due','label-del','link-del','st-toggle','st-del','cmt-save','cmt-edit','cmt-del','cmt-update','import','status-menu']);
document.addEventListener('click', async e=>{
  const b = e.target.closest('[data-act]'); if (!b) return;
  const act = b.dataset.act;
  if (EDIT_ACTS.has(act) && screen === 'app' && !canEdit()){ toast(t('toast.viewOnly')); return; }
  const inPop = !!b.closest('#pop');
  const ctxId = inPop ? (b.dataset.id || pop.dataset.id) : (b.dataset.id || openId);
  const i = ctxId && get(ctxId);
  switch(act){
    /* account & workspaces */
    case 'signin': try { await store.signIn(); } catch(err){ console.error(err); toast(t('toast.signinFailed')); } break;
    case 'signout': closePop(); flush(); await store.signOut(); break;
    case 'retry': if (wsId) openWorkspace(wsId); else location.reload(); break;
    case 'lang': closePop(); setLang(getLang() === 'he' ? 'en' : 'he'); relocalize(); break;
    case 'theme': openPop(b, themePicker(), 'theme'); break;
    case 'set-theme': { const th = b.dataset.th; setTheme(THEMES.includes(th) ? th : 'light', true);
      $$('.theme-opt', pop).forEach(x=>x.setAttribute('aria-pressed', x.dataset.th === effectiveTheme()));
      if (screen !== 'app') setTimeout(()=>{ closePop(); rerenderGate(); }, 200); break; }
    case 'ws-menu': openPop(b, wsMenu(), ''); break;
    case 'me-menu': openPop(b, meMenu(), ''); break;
    case 'switch-ws': closePop(); if (b.dataset.id !== wsId) openWorkspace(b.dataset.id); break;
    case 'welcome': closePop(); closeModal(true); renderWelcome(); break;
    case 'welcome-back': if (wsId){ showApp(); scheduleRender(); } else if (wsList?.length) openWorkspace(wsList[0].id); break;
    case 'ws-create': createWorkspace(b.dataset.kind); break;
    case 'ws-settings': closePop(); shell.classList.remove('side-open'); openWsSettings(); break;
    case 'request-access': closePop(); shell.classList.remove('side-open'); openRequestAccess(); break;
    case 'bell': if (isOwner()) openWsSettings(); else if (!(wsList||[]).some(w=>w.ownerUid===user.uid)) toast(t('req.needOwn')); else { const own = wsList.find(w=>w.ownerUid===user.uid); openWorkspace(own.id); setTimeout(()=>{ if (isOwner()) openWsSettings(); }, 400); } break;
    case 'req-cancel': store.deleteRequest(b.dataset.id).then(()=>{ if (modalKind === 'request') openRequestAccess(); }).catch(()=>toast(t('toast.saveFailed'))); break;
    case 'req-approve': {
      const r = incoming.find(x=>x.id===b.dataset.id); if (!r || !isOwner()) break;
      const email = lower(r.fromEmail), role = b.dataset.role === 'viewer' ? 'viewer' : 'editor';
      if (!await setMember(email, role)) break;
      store.updateRequest(r.id, {status:'approved', role, wsId, wsName:ws.name, decidedAt:Date.now()}).catch(()=>{});
      toast(t('toast.approved', {email, role:t('role.'+role)}));
      break; }
    case 'req-decline': store.updateRequest(b.dataset.id, {status:'declined', decidedAt:Date.now()}).catch(()=>toast(t('toast.saveFailed'))); break;
    case 'member-remove': {
      const email = b.dataset.email;
      if (await setMember(email, null)) toast(t('toast.removed', {email})); break; }
    case 'copy-site': copyText(location.origin + location.pathname).then(ok=>toast(ok ? t('toast.linkCopied') : location.origin + location.pathname)); break;
    case 'ws-logo': pickLogo('ws'); break;
    case 'ws-logo-del': store.updateWorkspace(wsId, {logo:null}).catch(()=>toast(t('toast.saveFailed'))); break;

    /* layout */
    case 'toggle-side':
      if (sideOverlay()) shell.classList.toggle('side-open');
      else { P.side = !P.side; shell.classList.toggle('side-closed', !P.side); savePrefs(); }
      break;
    case 'view': e.preventDefault(); P.view = b.dataset.view; savePrefs(); render(); shell.classList.remove('side-open'); break;
    case 'switch-board': switchBoard(b.dataset.id); break;
    case 'board-menu': {
      e.stopPropagation();
      const bid = b.dataset.id;
      openPop(b, `<button class="menu-i" role="menuitem" data-act="board-settings" data-id="${bid}">${I.gear}${esc(t('bs.title'))}</button>
        <button class="menu-i danger" role="menuitem" data-act="del-board" data-id="${bid}">${I.trash}${esc(t('bs.delete'))}</button>`, bid);
      break; }
    case 'new-board': closePop(); shell.classList.remove('side-open'); openBoardSettings(null); break;
    case 'board-settings': {
      closePop(); shell.classList.remove('side-open');
      const bid = inPop ? pop.dataset.id : (b.dataset.id || (isAll() ? null : curBoard()?.id));
      if (bid) openBoardSettings(bid); else toast(t('toast.pickBoard'));
      break; }
    case 'del-board': closePop(); shell.classList.remove('side-open'); deleteBoard(inPop ? pop.dataset.id : b.dataset.id); break;
    case 'confirm-ok': { const fn = confirmFn; confirmFn = null; fn?.(); break; }
    case 'imp-add': applyImport('add'); break;
    case 'imp-replace': if (isOwner()) applyImport('replace'); break;
    case 'ap-yes': { const list = S.issues.filter(i=>!i.archived && isDone(i)); S.archive = {...S.archive, handled:ym()}; if (list.length) archiveIssues(list, tn('toast.archivedN', list.length)); else { save(); render(); } break; }
    case 'ap-no': S.archive = {...S.archive, handled:ym()}; save(); render(); toast(t('ap.later')); break;
    case 'types': closePop(); shell.classList.remove('side-open'); openTypes(); break;
    case 'create': openCreate(); break;
    case 'shortcuts': openShortcuts(); break;
    case 'export': exportData(); shell.classList.remove('side-open'); break;
    case 'export-tracker': exportTracker(); shell.classList.remove('side-open'); break;
    case 'trk-apply': applyTracker(); break;
    case 'import': $('#importFile').click(); shell.classList.remove('side-open'); break;

    /* filters */
    case 'qf': qf.has(b.dataset.qf) ? qf.delete(b.dataset.qf) : qf.add(b.dataset.qf); render(); break;
    case 'type-filter-menu':
      openPop(b, `<div class="menu-h">${esc(t('filter.showType'))}</div><button class="menu-i" role="menuitem" data-act="type-filter" data-t="">${esc(t('filter.allTypes'))}${chk(!typeFilter)}</button>
        ${S.types.map(x=>`<button class="menu-i" role="menuitem" data-act="type-filter" data-t="${x.id}">${typeSVG(x)}<span data-bidi>${esc(x.name)}</span>${chk(typeFilter===x.id)}</button>`).join('')}`, '');
      break;
    case 'type-filter': typeFilter = b.dataset.t || null; closePop(); render(); break;
    case 'plat-filter-menu':
      openPop(b, `<div class="menu-h">${esc(t('filter.showPlatform'))}</div><button class="menu-i" role="menuitem" data-act="plat-filter" data-p="">${esc(t('filter.allPlatforms'))}${chk(!platFilter)}</button>
        ${PLATFORMS.map(([n,h])=>`<button class="menu-i" role="menuitem" data-act="plat-filter" data-p="${h}"><i class="pf-dot" style="--sw:${h}"></i>${esc(platLabel(n))}${chk(platFilter===h)}</button>`).join('')}
        <button class="menu-i" role="menuitem" data-act="plat-filter" data-p="none"><i class="pf-dot none"></i>${esc(t('filter.noPlatform'))}${chk(platFilter==='none')}</button>`, '');
      break;
    case 'plat-filter': platFilter = b.dataset.p || null; closePop(); render(); break;
    case 'clear-filters': qf.clear(); typeFilter = null; platFilter = null; query = ''; $('#q').value = ''; sheetF = {}; render(); break;
    case 's-filter': openPop(b, sheetFilterMenu(b.dataset.k), 'sf:' + b.dataset.k); break;
    case 's-fval': { const k = b.dataset.k, v = b.dataset.v; const set = new Set(sheetF[k] || []); set.has(v) ? set.delete(v) : set.add(v); sheetF[k] = set.size ? set : undefined; sheetFRefresh(k); break; }
    case 's-fall': { sheetF[b.dataset.k] = undefined; sheetFRefresh(b.dataset.k); break; }
    case 's-cols': openPop(b, sheetColsMenu(), 'scols'); break;
    case 's-col': { const k = b.dataset.k; if (k === 'title') break; setSheetHidden(k, !sheetHidden().has(k)); sheetColsRefresh(); break; }
    case 's-colsall': P.sheetHide = []; savePrefs(); sheetColsRefresh(); break;
    case 's-hidecol': { const k = b.dataset.k; closePop(); setSheetHidden(k, true); render();
      toast(t('sheet.colHidden', {col:t('x.col.' + k)}), {action:t('btn.undo'), fn:()=>{ setSheetHidden(k, false); render(); }}); break; }
    case 's-fclear': { if (b.dataset.k) sheetF[b.dataset.k] = undefined; else sheetF = {}; closePop(); render(); break; }

    /* issues */
    case 'add-inline': openComposer(b.dataset.status); break;
    case 'card-menu': openPop(b, statusMenu(i, true), i.id); break;
    case 'status-menu': openPop(b, statusMenu(i, false), i.id); break;
    case 's-menu': if (i) openPop(b, sheetMenu(i, b.dataset.sf), i.id); break;
    case 'cal-ics': exportIcs(); break;
    case 'dash-range': P.dashWeeks = Number(b.dataset.w) || 8; savePrefs(); renderDash(); break;
    case 'dash-custom': openPop(b, dashCustomMenu(), 'dash'); break;
    case 'dash-toggle': case 'dash-hide': { const h = dashHidden(), w = b.dataset.w; h.has(w) && act === 'dash-toggle' ? h.delete(w) : h.add(w); P.dashHide = [...h]; savePrefs(); dashRefresh();
      if (act === 'dash-hide') toast(t('dash.hidden', {name:t('dash.w.' + w)}), {action:t('btn.undo'), fn:()=>{ const x = dashHidden(); x.delete(w); P.dashHide = [...x]; savePrefs(); renderDash(); }}); break; }
    case 'dash-showall': P.dashHide = []; savePrefs(); dashRefresh(); break;
    case 'dash-hideall': P.dashHide = DASH_W.map(w=>w.id); savePrefs(); dashRefresh(); break;
    case 'cal-today': calAnchor = today(); calSel = today(); renderCal(); break;
    case 'cal-step': { const r = calRange(), n = Number(b.dataset.n) || 0, a = new Date((calAnchor || today())+'T12:00');
      if (r.mode === 'week') a.setDate(a.getDate() + 7*n); else { a.setDate(1); a.setMonth(a.getMonth() + n); }
      calAnchor = iso(a); calSel = null; renderCal(); break; }
    case 'cal-mode': P.calMode = b.dataset.m === 'week' ? 'week' : 'month'; calAnchor = calSel || calAnchor; savePrefs(); renderCal(); break;
    case 'cal-undated': P.calUndated = !P.calUndated; savePrefs(); renderCal(); break;
    case 'cal-sel': calSel = b.dataset.day; if (P.calMode === 'week' || !isMobile()) { if (!isMobile()){ calAnchor = calSel; P.calMode = 'week'; savePrefs(); } } renderCal(); break;
    case 'cal-add': openCreate(null, {due:b.dataset.day}); break;
    case 'cal-day': { const d = b.dataset.day, list = scopeIssues().filter(matches).filter(x=>x.due===d).sort(calSort);
      openPop(b, `<div class="menu-h">${esc(fmtr('wdl', {weekday:'long', day:'numeric', month:'long'}).format(new Date(d+'T12:00')))}</div><div class="cal-pop">${list.map(x=>calChip(x,'row')).join('')}</div>`, ''); break; }
    case 's-set': closePop(); if (i) sheetSet(i, b.dataset.sf, b.dataset.v); break;
    case 's-sort': { const k = b.dataset.k; sheetSort = sheetSort.k !== k ? {k, dir:1} : sheetSort.dir > 0 ? {k, dir:-1} : {k:null, dir:1}; render(); break; }
    case 'set-status': {
      const fromModal = modalKind === 'issue';
      closePop(); moveIssue(i.id, b.dataset.status, null, {toast:!fromModal});
      if (fromModal) $('.status-btn',layer)?.classList.add('flash');
      else $(`.card[data-id="${i.id}"]`,board)?.focus({preventScroll:true});
      break; }
    case 'set-color': setColor(i, b.dataset.c); if (inPop) closePop(); if (modalKind === 'issue') refreshIssue(); break;
    case 'create-color': b.closest('form').color.value = b.dataset.c; markPicked(b); break;
    case 'open-issue': closePop(); openIssue(i.id); break;
    case 'copy-key': closePop(); copyText(key(i)).then(ok=>toast(ok ? t('toast.copied', {key:key(i)}) : key(i))); break;
    case 'del-issue': closePop(); deleteIssue(i.id); break;
    case 'archive-issue': closePop(); archiveIssues([i], t('toast.archived', {key:key(i)})); break;
    case 'archive-col': {
      const st = b.dataset.status, list = scopeIssues().filter(matches).filter(x=>st.startsWith('v:') ? vKey(x)===st : x.status===st);
      archiveIssues(list, tn('toast.archivedN', list.length));
      break; }
    case 'restore': restoreIssue(i); break;
    case 'close-modal': closeModal(); break;
    case 'clear-due': i.due = ''; touch(i, t('h.dueNone')); save(); render(); refreshIssue(); break;
    case 'label-del': { const l = b.dataset.l; i.labels = i.labels.filter(x=>x!==l); touch(i, t('h.companyRemove', {name:l})); save(); render(); refreshIssue(); break; }
    case 'link-del': {
      const li = b.closest('li'), l = i.links.find(x=>x.id===li.dataset.lid), ix = i.links.indexOf(l);
      i.links.splice(ix,1); touch(i, t('h.linkRemove', {name:l.title || hostOf(l.url)})); save(); render();
      li.classList.add('out'); setTimeout(()=>refreshIssue(), reduce?0:190);
      toast(t('toast.linkRemoved'), {action:t('btn.undo'), fn:()=>{ i.links.splice(ix,0,l); save(); render(); refreshIssue(); }});
      break; }
    case 'tab': tab = b.dataset.tab; refreshIssue(); break;
    case 'st-toggle': {
      const li = b.closest('li'), s = i.subtasks.find(s=>s.id===li.dataset.sid);
      s.done = !s.done; li.classList.toggle('done', s.done); b.setAttribute('aria-pressed', s.done); b.setAttribute('aria-label', s.done?t('st.undone'):t('st.done'));
      touch(i, t(s.done ? 'h.stDone' : 'h.stUndone', {name:s.text})); save(); render();
      const pct = pctOf(i); $('#stMeta',layer).textContent = t('st.pct', {n:pct}); setBar(pct);
      if (pct === 100 && !isDone(i)){
        const doneCol = boardOf(i).columns.find(c=>c.cat==='done');
        if (doneCol) toast(t('toast.allSubtasks'), {action:t('btn.moveTo', {col:doneCol.name}), fn:()=>moveIssue(i.id, doneCol.id, null)});
      }
      break; }
    case 'st-del': {
      const li = b.closest('li'), s = i.subtasks.find(s=>s.id===li.dataset.sid);
      i.subtasks = i.subtasks.filter(x=>x!==s); touch(i, t('h.stDelete', {name:s.text})); save(); render();
      li.classList.add('out'); setTimeout(()=>{ li.remove(); if (!i.subtasks.length) refreshIssue(); else { const p = pctOf(i); $('#stMeta',layer).textContent = t('st.pct', {n:p}); setBar(p); } }, reduce?0:190);
      break; }
    case 'cmt-save': saveComment(); break;
    case 'cmt-cancel': cmtDraft = ''; { const ta = $('[data-k="cmt"]',layer); ta.value = ''; autosize(ta); ta.blur(); } break;
    case 'cmt-edit': editingCid = b.closest('li').dataset.cid; refreshIssue(); { const ta = $('[data-k="cmt-edit"]',layer); ta?.focus(); ta?.setSelectionRange(ta.value.length, ta.value.length); } break;
    case 'cmt-edit-cancel': editingCid = null; refreshIssue(); break;
    case 'cmt-update': updateComment($('[data-k="cmt-edit"]',layer)); break;
    case 'cmt-del': {
      const cid = b.closest('li').dataset.cid, ix = i.comments.findIndex(c=>c.id===cid), [c] = i.comments.splice(ix,1);
      touch(i); save(); render(); refreshIssue();
      toast(t('toast.cmtDeleted'), {action:t('btn.undo'), fn:()=>{ i.comments.splice(ix,0,c); save(); render(); refreshIssue(); }});
      break; }
    case 'sort': { const k = b.dataset.k; sortBy = sortBy.k === k ? {k, dir:-sortBy.dir} : {k, dir:1}; render(); break; }

    /* board settings editor */
    case 'bd-logo': pickLogo('board'); break;
    case 'bd-logo-del': bDraft.logo = null; $('#bsBody',layer).innerHTML = boardSettingsBody(); break;
    case 'bd-color': bDraft.color = b.dataset.c; markPicked(b); { const pv = $('.logo-prev .b-av',layer); if (pv) pv.style.setProperty('--bc', bDraft.color); } break;
    case 'bd-addcol': {
      const c = {id:uid(), name:'', cat:'doing'};
      const doneIx = bDraft.columns.findIndex(x=>x.cat==='done');
      doneIx > 0 ? bDraft.columns.splice(doneIx,0,c) : bDraft.columns.push(c);
      rerenderCols(true);
      const row = $(`.ed-row[data-cid="${c.id}"]`,layer); row?.classList.add('in'); $('.ce-name',row)?.focus();
      break; }
    case 'bd-colup': case 'bd-coldown': {
      const cid = b.closest('.ed-row').dataset.cid, cs = bDraft.columns, ix = cs.findIndex(c=>c.id===cid), to = ix + (act==='bd-colup' ? -1 : 1);
      if (to < 0 || to >= cs.length) break;
      [cs[ix], cs[to]] = [cs[to], cs[ix]];
      rerenderCols(true);
      const same = $(`[data-k="${b.dataset.k}"]`,layer);
      if (!same || same.disabled) $(`[data-k="cname-${cid}"]`,layer)?.focus();
      break; }
    case 'bd-delcol': {
      const cid = b.closest('.ed-row').dataset.cid, c = bDraft.columns.find(x=>x.id===cid);
      if (bDraft.columns.length === 1) break;
      if (!bDraft.isNew && colIssueCount(cid) > 0){ c._confirm = true; c._target = bDraft.columns.find(x=>x.id!==cid).id; rerenderCols(false); $(`[data-k="cyes-${cid}"]`,layer)?.focus(); }
      else removeDraftCol(cid);
      break; }
    case 'bd-delcol-yes': { const cid = b.closest('.ed-row').dataset.cid, c = bDraft.columns.find(x=>x.id===cid); bDraft.removed[cid] = c._target; removeDraftCol(cid); break; }
    case 'bd-delcol-no': { const cid = b.closest('.ed-row').dataset.cid, c = bDraft.columns.find(x=>x.id===cid); delete c._confirm; rerenderCols(false); $(`[data-k="cdel-${cid}"]`,layer)?.focus(); break; }

    /* types editor */
    case 'ty-preset': {
      const preset = TYPES_PPC(), keep = new Set(preset.map(x=>x.id));
      tDraft.types.forEach(x=>{ if (!keep.has(x.id)) tDraft.removed[x.id] = preset[0].id; });
      tDraft.types = preset.map(x=>({...x, _new:true})); tDraft.err = '';
      rerender($('#tyBody',layer), typesBody()); tDraft.types.forEach(x=>{ delete x._new; }); break; }
    case 'ty-add': {
      const used = new Set(tDraft.types.map(x=>x.color));
      const ty = {id:'t-'+uid(), name:'', color:(PALETTE.find(([,h])=>!used.has(h))||PALETTE[0])[1], glyph:'ring', _new:true};
      tDraft.types.push(ty); rerender($('#tyBody',layer), typesBody()); delete ty._new;
      $(`[data-k="tname-${ty.id}"]`,layer)?.focus();
      break; }
    case 'ty-style': { const tid = b.closest('.ed-row').dataset.tid, ty = tDraft.types.find(x=>x.id===tid); openPop(b, typeStylePop(ty), tid); break; }
    case 'ty-color': case 'ty-glyph': {
      const tid = pop.dataset.id, ty = tDraft?.types.find(x=>x.id===tid); if (!ty) break;
      if (act === 'ty-color') ty.color = b.dataset.c; else ty.glyph = b.dataset.g;
      const focusSel = act === 'ty-color' ? `.sw[data-c="${ty.color}"]` : `.glyph-btn[data-g="${ty.glyph}"]`;
      pop.innerHTML = typeStylePop(ty); $(focusSel,pop)?.focus();
      const ico = $(`.ed-row[data-tid="${tid}"] .ty-ico`,layer); if (ico) ico.innerHTML = typeSVG(ty);
      break; }
    case 'ty-del': {
      const tid = b.closest('.ed-row').dataset.tid, ty = tDraft.types.find(x=>x.id===tid);
      if (tDraft.types.length === 1) break;
      if (typeCount(tid) > 0){ ty._confirm = true; ty._target = tDraft.types.find(x=>x.id!==tid).id; rerender($('#tyBody',layer), typesBody()); $(`[data-k="tyes-${tid}"]`,layer)?.focus(); }
      else removeDraftType(tid);
      break; }
    case 'ty-del-yes': { const tid = b.closest('.ed-row').dataset.tid, ty = tDraft.types.find(x=>x.id===tid); tDraft.removed[tid] = ty._target; removeDraftType(tid); break; }
    case 'ty-del-no': { const tid = b.closest('.ed-row').dataset.tid, ty = tDraft.types.find(x=>x.id===tid); delete ty._confirm; rerender($('#tyBody',layer), typesBody()); $(`[data-k="tdel-${tid}"]`,layer)?.focus(); break; }
  }
});
document.addEventListener('keydown', e=>{
  if ((e.key === 'Enter' || e.key === ' ') && e.target.matches?.('.board-item, .bi-menu')){ e.preventDefault(); e.target.click(); }
});
function animateOut(row, fn){
  if (!row || reduce){ fn(); return; }
  row.animate([{opacity:1, transform:'none'},{opacity:0, transform:`translateX(${rtl()?-12:12}px)`}],{duration:180, easing:'ease-in'}).onfinish = fn;
}
function removeDraftCol(cid){
  const row = $(`.ed-row[data-cid="${cid}"]`,layer);
  animateOut(row, ()=>{ bDraft.columns = bDraft.columns.filter(c=>c.id!==cid); bDraft.columns.forEach(c=>{ if (c._target === cid) c._target = bDraft.columns.find(x=>x.id!==c.id)?.id; }); rerenderCols(true); });
}
function removeDraftType(tid){
  const row = $(`.ed-row[data-tid="${tid}"]`,layer);
  animateOut(row, ()=>{ tDraft.types = tDraft.types.filter(x=>x.id!==tid); tDraft.types.forEach(x=>{ if (x._target === tid) x._target = tDraft.types.find(y=>y.id!==x.id)?.id; }); rerender($('#tyBody',layer), typesBody()); });
}

/* gate forms (sign-in demo, request access on the welcome screen) */
gateEl.addEventListener('submit', async e=>{
  e.preventDefault();
  const form = e.target;
  if (form.dataset.form === 'demo-signin'){
    const f = new FormData(form), email = lower(f.get('email'));
    if (!isEmail(email)){ form.email.focus(); form.email.classList.add('bad'); return; }
    await store.signIn({name:String(f.get('name')||'').trim(), email});
  }
  if (form.dataset.form === 'request') sendRequest(form);
});
gateEl.addEventListener('input', e=>{
  if (e.target.id === 'wsNameIn') e.target.dataset.touched = '1';
  if (e.target.name === 'email'){ e.target.classList.remove('bad'); const er = $('#reqErr'); if (er) er.hidden = true; }
});

/* ================= Keyboard shortcuts ================= */
document.addEventListener('keydown', e=>{
  if (e.key === 'Escape'){
    if (popAnchor){ const a = popAnchor; closePop(); a.focus(); return; }
    if (modalKind){ e.preventDefault(); closeModal(); return; }
    if (shell.classList.contains('side-open')){ shell.classList.remove('side-open'); return; }
    if (document.activeElement === $('#q') && $('#q').value){ $('#q').value = ''; query = ''; render(); return; }
    return;
  }
  if (screen !== 'app') return;
  const typing = e.target.closest?.('input,textarea,select,[contenteditable]');
  if (typing || e.ctrlKey || e.metaKey || e.altKey || modalKind || drag?.active) return;
  const k = e.key.toLowerCase(), code = e.code;
  if ((k === 'c' || code === 'KeyC') && !e.shiftKey){ e.preventDefault(); openCreate(); }
  else if (e.key === '/' || code === 'Slash' && !e.shiftKey){ e.preventDefault(); $('#q').focus(); $('#q').select(); }
  else if (k === 'b' || code === 'KeyB'){ P.view = 'board'; savePrefs(); render(); }
  else if (k === 'l' || code === 'KeyL'){ P.view = 'list'; savePrefs(); render(); }
  else if (k === 't' || code === 'KeyT'){ P.view = 'table'; savePrefs(); render(); }
  else if (k === 'k' || code === 'KeyK'){ P.view = 'cal'; savePrefs(); render(); }
  else if (k === 'd' || code === 'KeyD'){ P.view = 'dash'; savePrefs(); render(); }
  else if (k === 'a' || code === 'KeyA'){ if (S.boards.length) switchBoard(ALL); }
  else if (e.key === '?'){ openShortcuts(); }
});
// debounced: with hundreds of tasks a render on every keystroke made typing lag on phones
$('#q').addEventListener('input', e=>{ query = e.target.value; clearTimeout(qTimer); qTimer = setTimeout(render, query ? 140 : 0); });

/* ================= Toasts ================= */
function toast(msg, opt={}){
  const box = $('#toasts');
  while (box.children.length >= 3) box.firstElementChild.remove();
  const el = document.createElement('div'); el.className = 'toast'; el.setAttribute('role','status');
  el.innerHTML = `<span class="t-check">${I.check}</span><span class="t-msg" data-bidi></span>${opt.action?`<button class="btn sm">${esc(opt.action)}</button>`:''}`;
  $('.t-msg',el).textContent = msg;
  box.appendChild(el);
  let gone = false;
  const close = ()=>{ if (gone) return; gone = true; el.classList.add('out'); setTimeout(()=>el.remove(), reduce?0:200); };
  let moved = false; // a swipe that starts on the button must not press it
  if (opt.action) $('button',el).onclick = ()=>{ if (moved) return; close(); opt.fn(); };
  let timer = setTimeout(close, opt.ms || (opt.action ? 6000 : 3200));
  // pause only while a real mouse is over it: on phones a tap fires mouseenter but never mouseleave,
  // which kept the message up for good
  el.addEventListener('pointerenter', e=>{ if (e.pointerType === 'mouse') clearTimeout(timer); });
  el.addEventListener('pointerleave', e=>{ if (e.pointerType === 'mouse' && !gone){ clearTimeout(timer); timer = setTimeout(close, 2000); } });
  // swipe sideways (or up/down) to dismiss
  let sx = 0, sy = 0, dx = 0, dy = 0, swiping = false, pid = null;
  el.addEventListener('pointerdown', e=>{ pid = e.pointerId; sx = e.clientX; sy = e.clientY; dx = dy = 0; moved = false; swiping = true; el.style.transition = 'none'; });
  el.addEventListener('pointermove', e=>{ if (!swiping || e.pointerId !== pid) return; dx = e.clientX - sx; dy = e.clientY - sy;
    if (!moved && Math.max(Math.abs(dx), Math.abs(dy)) > 8){ moved = true; try { el.setPointerCapture(pid); } catch(_){} }
    if (!moved) return;
    const d = Math.abs(dx) >= Math.abs(dy) ? `translateX(${dx}px)` : `translateY(${dy}px)`; el.style.transform = d; el.style.opacity = String(Math.max(.2, 1 - Math.max(Math.abs(dx), Math.abs(dy)) / 220)); });
  const endSwipe = ()=>{ if (!swiping) return; swiping = false; el.style.transition = ''; setTimeout(()=>{ moved = false; }, 0);
    if (Math.max(Math.abs(dx), Math.abs(dy)) > 50){ clearTimeout(timer); gone = true; el.style.transition = 'transform .18s ease-in, opacity .18s';
      el.style.transform = Math.abs(dx) >= Math.abs(dy) ? `translateX(${dx > 0 ? 120 : -120}%)` : `translateY(${dy > 0 ? 120 : -120}%)`; el.style.opacity = '0'; setTimeout(()=>el.remove(), 190); }
    else { el.style.transform = ''; el.style.opacity = ''; } };
  el.addEventListener('pointerup', endSwipe); el.addEventListener('pointercancel', endSwipe);
}

/* ================= Clipboard, backup ================= */
async function copyText(s){
  try { await navigator.clipboard.writeText(s); return true; }
  catch(e){
    try { const ta = document.createElement('textarea'); ta.value = s; ta.style.position='fixed'; ta.style.opacity='0'; document.body.appendChild(ta); ta.select();
      const ok = document.execCommand('copy'); ta.remove(); return ok; } catch(e2){ return false; }
  }
}
/* ================= Backup as a spreadsheet (.xlsx) ================= */
// Visible sheets are for people (and can be edited in Google Sheets / Excel); the hidden "_data" sheet keeps
// everything else (ids, comments, history, logos, column setup) so a restore is lossless.
const XCOLS = [ // [key, width, style]
  ['title',44,ST.title],['prio',13,ST.center],['board',16,ST.center],['status',18,ST.center],['start',13,ST.date],['due',13,ST.date],
  ['type',17,ST.center],['plat',17,ST.center],['company',16,ST.wrap],['points',9,ST.num],['desc',48,ST.wrap],['subs',34,ST.wrap],
  ['links',34,ST.wrap],['cmts',40,ST.meta],['arch',10,ST.center],['key',11,ST.key],['by',18,ST.center],['created',17,ST.dtime],['updated',17,ST.dtime],['id',3,ST.meta]
];
const BCOLS = [['name',22,ST.title],['bkey',10,ST.key],['color',12,ST.center],['columns',60,ST.wrap],['open',10,ST.num],['total',10,ST.num],['logo',10,ST.center],['id',3,ST.meta]];
const TCOLS = [['name',24,ST.title],['color',12,ST.center],['glyph',12,ST.center],['count',10,ST.num],['id',3,ST.meta]];
const PRIO_CHIP = {highest:'#C9372C', high:'#E56910', medium:'#E2B203', low:'#1D7AFC', lowest:'#8590A2'};
const CAT_CHIP = {todo:['#DCDFE4','#172B4D'], doing:['#1D7AFC','#FFFFFF'], wait:['#F5CD47','#172B4D'], done:['#22A06B','#FFFFFF']};
const textOn = hex=>{ const n = parseInt(String(hex).slice(1), 16); const l = (0.299*(n>>16) + 0.587*((n>>8)&255) + 0.114*(n&255)) / 255; return l > 0.62 ? '#172B4D' : '#FFFFFF'; };
const hx = k=>t('x.col.'+k);
const yesNo = b=>b ? t('x.yes') : t('x.no');
function platLabelOf(hex){ if (!hex) return ''; const p = PLATFORMS.find(x=>x[1]===hex); if (p) return platLabel(p[0]); return colorName(hex); }
const CHUNK = 30000;
function exportData(){
  flush();
  const boardsSorted = S.boards.slice().sort((a,b)=>a.order-b.order);
  const bIx = new Map(boardsSorted.map((b,i)=>[b.id,i]));
  const colIx = i=>{ const b = boardOf(i); return b ? b.columns.findIndex(c=>c.id===i.status) : 0; };
  const issues = S.issues.slice().sort((a,b)=>(a.archived-b.archived) || (bIx.get(a.boardId)??99)-(bIx.get(b.boardId)??99) || colIx(a)-colIx(b) || (a.rank-b.rank));
  const statusName = i=>colOf(i)?.name || '';
  const lines = (arr, f)=>arr.map(f).join('\n');
  const head = cols=>cols.map(([k])=>({v:k==='id' ? 'id' : hx(k), s:ST.head}));
  const taskRows = issues.map(i=>{
    const cell = {
      title:i.title, prio:prioName(i.priority), board:boardOf(i)?.name || '', status:statusName(i),
      start:i.start ? {v:isoToSerial(i.start), s:ST.date} : '', due:i.due ? {v:isoToSerial(i.due), s:ST.date} : '',
      type:typeOf(i.type).name, plat:platLabelOf(i.color), company:i.labels.join(', '), points:i.points ?? '',
      desc:i.desc, subs:lines(i.subtasks, s=>(s.done ? '☑ ' : '☐ ') + s.text), links:lines(i.links, l=>l.title ? `${l.title} — ${l.url}` : l.url),
      cmts:lines(i.comments, c=>`${c.by?.name || ''} (${fmtTime(c.at)}): ${c.text}`), arch:yesNo(i.archived), key:key(i), by:(c=>c ? (c.name || c.email) : '')(creatorOf(i)),
      created:{v:msToSerial(i.created), s:ST.dtime}, updated:{v:msToSerial(i.updated), s:ST.dtime}, id:i.id
    };
    return XCOLS.map(([k,,s])=>{ const c = cell[k]; return c && typeof c === 'object' ? c : {v:c, s}; });
  });
  const statusCat = new Map(); boardsSorted.forEach(b=>b.columns.forEach(c=>{ if (!statusCat.has(c.name)) statusCat.set(c.name, c.cat); }));
  const xc = k=>XCOLS.findIndex(c=>c[0]===k);
  const platValues = [...PLATFORMS.map(([n,h])=>[platLabel(n),h]), ...EXTRA.map(([n,h])=>[t('color.'+n),h])];
  const tasks = {name:t('x.sheet.tasks'), rtl:getLang()==='he', freeze:1, headHeight:26, autoFilter:true, tab:'#5B2C83',
    cols:XCOLS.map(([k,w])=>({width:w, hidden:k==='id'})), rows:[head(XCOLS), ...taskRows],
    lists:[{col:xc('prio'), values:PRIOS.map(prioName)}, {col:xc('board'), values:boardsSorted.map(b=>b.name)}, {col:xc('status'), values:[...statusCat.keys()]},
      {col:xc('type'), values:S.types.map(x=>x.name)}, {col:xc('plat'), values:platValues.map(p=>p[0])}, {col:xc('arch'), values:[t('x.yes'), t('x.no')]}],
    chips:[{col:xc('prio'), values:PRIOS.map(p=>({text:prioName(p), bg:PRIO_CHIP[p], fg:textOn(PRIO_CHIP[p])}))},
      {col:xc('status'), values:[...statusCat].map(([n,c])=>({text:n, bg:CAT_CHIP[c][0], fg:CAT_CHIP[c][1]}))},
      {col:xc('board'), values:boardsSorted.map(b=>({text:b.name, bg:b.color, fg:textOn(b.color)}))},
      {col:xc('type'), values:S.types.map(x=>({text:x.name, bg:x.color, fg:textOn(x.color)}))},
      {col:xc('plat'), values:platValues.map(([n,h])=>({text:n, bg:h, fg:textOn(h)}))},
      {col:xc('arch'), values:[{text:t('x.yes'), bg:'#8590A2', fg:'#FFFFFF'}]}]};
  const boards = {name:t('x.sheet.boards'), rtl:getLang()==='he', freeze:1, headHeight:26, tab:'#2898BD', cols:BCOLS.map(([k,w])=>({width:w, hidden:k==='id'})),
    rows:[head(BCOLS), ...boardsSorted.map(b=>{ const all = S.issues.filter(i=>i.boardId===b.id);
      const cell = {name:b.name, bkey:b.key, color:colorName(b.color), columns:b.columns.map(c=>c.name).join(' → '), open:all.filter(i=>!i.archived && !isDone(i)).length, total:all.length, logo:yesNo(isImg(b.logo)), id:b.id};
      return BCOLS.map(([k,,s])=>({v:cell[k], s})); })],
    chips:[{col:2, values:PALETTE.map(([n,h])=>({text:t('color.'+n), bg:h, fg:textOn(h)}))}]};
  const types = {name:t('x.sheet.types'), rtl:getLang()==='he', freeze:1, headHeight:26, tab:'#8F7EE7', cols:TCOLS.map(([k,w])=>({width:w, hidden:k==='id'})),
    rows:[head(TCOLS), ...S.types.map(x=>{ const cell = {name:x.name, color:colorName(x.color), glyph:x.glyph, count:S.issues.filter(i=>i.type===x.id).length, id:x.id};
      return TCOLS.map(([k,,s])=>({v:cell[k], s})); })],
    chips:[{col:0, values:S.types.map(x=>({text:x.name, bg:x.color, fg:textOn(x.color)}))}]};
  const dataRows = [['workbench-backup', 1, new Date().toISOString(), ws?.name || '']];
  const pushJson = (kind, id, obj)=>{ const j = JSON.stringify(obj), row = [kind, id]; for (let p = 0; p < j.length; p += CHUNK) row.push(j.slice(p, p + CHUNK)); dataRows.push(row); };
  S.boards.forEach(b=>pushJson('board', b.id, b));
  S.types.forEach(x=>pushJson('type', x.id, x));
  S.issues.forEach(i=>pushJson('issue', i.id, i));
  pushJson('meta', 'settings', {archive:S.archive});
  const blob = writeXlsx({title:ws?.name || 'Workbench', rtl:getLang()==='he', accent:'#5B2C83',
    sheets:[tasks, boards, types, {name:'_data', hidden:true, rows:dataRows.map(r=>r.map(v=>({v})))}]});
  // ASCII file name: some browsers drop non-Latin names on downloads
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `workbench-backup-${today()}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 1500);
  toast(t('toast.exported'));
}

/* ----- reading a backup spreadsheet back ----- */
const norm = s=>String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
const both = k=>[_dicts.he?.[k], _dicts.en?.[k]].filter(Boolean).map(norm);
function headerMap(row, cols){
  const m = {}; (row||[]).forEach((h, ix)=>{ const n = norm(h); if (n === 'id') m.id = ix; else for (const [k] of cols) if (both('x.col.'+k).includes(n)) m[k] = ix; });
  return m;
}
const findSheet = (sheets, k)=>sheets.find(s=>both('x.sheet.'+k).includes(norm(s.name)));
function cellDate(v){
  if (typeof v === 'number' && v > 20000 && v < 80000) return serialToIso(v);
  const s = String(v ?? '').trim(); let m;
  if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  if ((m = /^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/.exec(s))) return `${m[3].length===2 ? '20'+m[3] : m[3]}-${m[2].padStart(2,'0')}-${m[1].padStart(2,'0')}`;
  return '';
}
const byName = (list, name)=>list.find(x=>norm(x.name)===norm(name));
function colorFrom(v){
  const s = norm(v); if (!s) return null; if (isHex(String(v).trim())) return String(v).trim().toUpperCase();
  for (const [n,h] of PLATFORMS) if (norm(n)===s || both('plat.'+n).includes(s)) return h;
  for (const [n,h] of [...PALETTE, ...EXTRA]) if (both('color.'+n).includes(s) || n===s) return h;
  return null;
}
function backupFromSheets(sheets){
  const data = sheets.find(s=>s.name==='_data'), J = {board:new Map(), type:new Map(), issue:new Map(), meta:new Map()};
  if (data && data.rows[0]?.[0] === 'workbench-backup'){
    data.rows.slice(1).forEach(r=>{ if (!r || !J[r[0]]) return; try { J[r[0]].set(String(r[1]), JSON.parse(r.slice(2).map(x=>x ?? '').join(''))); } catch(e){} });
  }
  const tSheet = findSheet(sheets, 'tasks'), bSheet = findSheet(sheets, 'boards'), tySheet = findSheet(sheets, 'types');
  if (!tSheet && !J.issue.size) throw new Error('bad');
  // boards
  let boards = [];
  if (bSheet){
    const hm = headerMap(bSheet.rows[0], BCOLS);
    bSheet.rows.slice(1).forEach(r=>{
      const name = String(r[hm.name] ?? '').trim(); if (!name) return;
      const base = J.board.get(String(r[hm.id] ?? '')) ;
      const b = base ? clone(base) : {id:'b-'+uid(), name, key:keyFromName(name), color:PALETTE[boards.length % PALETTE.length][1], columns:null};
      b.name = name; if (hm.bkey != null && String(r[hm.bkey] ?? '').trim()) b.key = String(r[hm.bkey]).trim();
      const col = hm.color != null ? colorFrom(r[hm.color]) : null; if (col) b.color = col;
      if (!b.columns && hm.columns != null){ const names = String(r[hm.columns] ?? '').split(/→|,|\||\n/).map(x=>x.trim()).filter(Boolean);
        b.columns = names.map((n, ix)=>({id:'c'+uid(), name:n, cat:ix===names.length-1 && names.length > 1 ? 'done' : ix===0 ? 'todo' : 'doing'})); }
      if (!b.columns || !b.columns.length) b.columns = COLS_PPC();
      boards.push(b);
    });
  } else boards = [...J.board.values()].map(clone);
  // types
  let types = [];
  if (tySheet){
    const hm = headerMap(tySheet.rows[0], TCOLS);
    tySheet.rows.slice(1).forEach(r=>{ const name = String(r[hm.name] ?? '').trim(); if (!name) return;
      const base = J.type.get(String(r[hm.id] ?? '')); const x = base ? clone(base) : {id:'t-'+uid(), name, color:PALETTE[types.length % PALETTE.length][1], glyph:'ring'};
      x.name = name; const col = hm.color != null ? colorFrom(r[hm.color]) : null; if (col) x.color = col;
      if (hm.glyph != null && GLYPHS[String(r[hm.glyph] ?? '').trim()]) x.glyph = String(r[hm.glyph]).trim();
      types.push(x); });
  } else types = [...J.type.values()].map(clone);
  const ensureBoard = name=>{ let b = byName(boards, name); if (!b){ b = {id:'b-'+uid(), name, key:keyFromName(name), color:PALETTE[boards.length % PALETTE.length][1], columns:COLS_PPC()}; boards.push(b); } return b; };
  const ensureType = name=>{ let x = byName(types, name); if (!x){ x = {id:'t-'+uid(), name, color:PALETTE[(types.length + 3) % PALETTE.length][1], glyph:'ring'}; types.push(x); } return x; };
  // tasks
  let issues = [];
  if (tSheet){
    const hm = headerMap(tSheet.rows[0], XCOLS), has = k=>hm[k] != null, val = (r,k)=>has(k) ? r[hm[k]] : undefined, str = (r,k)=>String(val(r,k) ?? '').trim();
    tSheet.rows.slice(1).forEach(r=>{
      if (!r || !str(r,'title')) return;
      const base = J.issue.get(str(r,'id')), i = base ? clone(base) : {num:0, comments:[], history:[], created:Date.now(), updated:Date.now()};
      i.title = str(r,'title');
      const b = has('board') && str(r,'board') ? ensureBoard(str(r,'board')) : (boards.find(x=>x.id===i.boardId) || boards[0] || ensureBoard(t('x.defaultBoard')));
      if (i.boardId !== b.id) i.num = 0;
      i.boardId = b.id;
      if (has('status')){ const sn = str(r,'status'); let c = sn ? byName(b.columns, sn) : null;
        if (!c && sn){ c = {id:'c'+uid(), name:sn, cat:both('cat.done').includes(norm(sn)) ? 'done' : 'todo'}; b.columns.push(c); }
        i.status = (c || b.columns.find(x=>x.id===i.status) || b.columns[0]).id; }
      else if (!b.columns.some(x=>x.id===i.status)) i.status = b.columns[0].id;
      if (has('type') && str(r,'type')) i.type = ensureType(str(r,'type')).id;
      if (has('prio')){ const p = PRIOS.find(k=>both('prio.'+k).includes(norm(val(r,'prio'))) || k===norm(val(r,'prio'))); i.priority = p || i.priority || 'medium'; }
      if (has('plat')) i.color = colorFrom(val(r,'plat'));
      if (has('company')) i.labels = str(r,'company').split(',').map(normLabel).filter(Boolean);
      if (has('start')) i.start = cellDate(val(r,'start')) || i.start;
      if (has('due')) i.due = cellDate(val(r,'due'));
      if (has('points')){ const n = Number(val(r,'points')); i.points = str(r,'points') !== '' && Number.isFinite(n) ? n : null; }
      if (has('desc')){ const d = String(val(r,'desc') ?? ''); if (!(base && base.desc.length > d.length && d.length >= MAX_CELL - 10 && base.desc.startsWith(d))) i.desc = d; } // long texts are cut in the visible cell
      if (has('subs')){ const old = i.subtasks || [];
        i.subtasks = str(r,'subs').split('\n').map(x=>x.trim()).filter(Boolean).map(line=>{
          const done = /^(☑|✔|✅|\[x\])/i.test(line), text = line.replace(/^(☑|☐|✔|✅|\[x\]|\[ \]|-|•)\s*/i, '').trim();
          const prev = old.find(s=>s.text===text); return {id:prev?.id || uid(), text, done}; }).filter(s=>s.text); }
      if (has('links')){ const old = i.links || [];
        i.links = str(r,'links').split('\n').map(x=>x.trim()).filter(Boolean).map(line=>{
          const parts = line.split(/\s+—\s+|\s+-\s+(?=\S*\.)/), url = normUrl(parts.at(-1)); if (!url) return null;
          const title = parts.length > 1 ? parts.slice(0, -1).join(' — ').trim() : ''; const prev = old.find(l=>l.url===url);
          return {id:prev?.id || uid(), url, title}; }).filter(Boolean); }
      if (has('arch')){ const a = both('x.yes').includes(norm(val(r,'arch'))) || val(r,'arch') === true; if (a && !i.archived) i.archivedAt = Date.now(); i.archived = a; }
      issues.push(i);
    });
  } else issues = [...J.issue.values()].map(clone);
  if (!boards.length) throw new Error('bad');
  const used = new Set(boards.map(x=>String(x.key||'').toUpperCase()).filter(Boolean)); let kn = 1;
  boards.forEach(x=>{ if (!x.key){ while (used.has('C'+kn)) kn++; x.key = 'C'+kn; used.add(x.key); } });
  return {boards, types:types.length ? types : TYPES_PPC(), issues, archive:J.meta.get('settings')?.archive};
}
async function readBackupFile(f){
  if (f.size > 25e6) throw new Error('too-big');
  if (/\.xlsx$/i.test(f.name) || /spreadsheetml/.test(f.type)){
    const {sheets} = await readXlsx(await f.arrayBuffer());
    const tr = findTracker(sheets); if (tr) return {tracker:tr};
    return backupFromSheets(sheets);
  }
  return JSON.parse(await f.text());
}
/* ================= Task-tracker sheet (Elizabeth's "מעקב משימות") =================
   Columns: task, priority, client, status, start, due, documents, notes. Clients are boards, statuses are
   columns. Import adds new rows and updates matching tasks (same client + same task name); export writes
   the same sheet back. Tasks that are on the site but not in the sheet are left alone. */
const TRK_HEAD = {title:['משימה','task'], prio:['עדיפות','priority'], client:['לקוח','client','customer'], status:['סטטוס','status'],
  start:['תאריך התחלה','start','start date'], due:['דד ליין','due','due date','deadline'], docs:['מסמכים','documents','docs'], notes:['הערות','notes']};
const TRK_PRIO_IN = {'גבוהה':'high', 'בינונית':'medium', 'נמוכה':'low', 'high':'high', 'medium':'medium', 'low':'low'};
const TRK_PRIO_OUT = {highest:'גבוהה', high:'גבוהה', medium:'בינונית', low:'נמוכה', lowest:'נמוכה'};
// her status words ⇄ the site's columns. The four basic ones map onto the template columns (or the first
// column of that kind); the others become columns of their own when a board needs them.
const TRK_STATUS = {'לא התחיל':'todo', 'בביצוע':'doing', 'ממתין לאישור':'wait', 'בוצע':'done', 'דורש תיקונים':'doing', 'לא בוצע':'done'};
const TRK_BASIC = new Set(['לא התחיל','בביצוע','ממתין לאישור','בוצע']);
const TRK_COL_OUT = {'לביצוע':'לא התחיל', 'בעבודה':'בביצוע', 'ממתין לאישור לקוח':'ממתין לאישור', 'הושלם':'בוצע', 'to do':'לא התחיל', 'in progress':'בביצוע', 'client approval':'ממתין לאישור', 'done':'בוצע', 'needs fixes':'דורש תיקונים', 'not done':'לא בוצע'};
// loose Hebrew match: ignore spaces/punctuation, the vowel letters ו/י and final forms (קניוני עופר = קניון עופר)
const FINALS = {'ך':'כ','ם':'מ','ן':'נ','ף':'פ','ץ':'צ'};
const loose = s=>norm(s).replace(/[\s\-–—'"׳״.,_()]/g,'').replace(/[ךםןףץ]/g, ch=>FINALS[ch]).replace(/[וי]/g,'');
const statusWord = c=>TRK_COL_OUT[norm(c.name)] || c.name;
function findTracker(sheets){
  if (sheets.some(s=>s.name === '_data')) return null; // one of our own backups
  for (const sh of sheets){
    const head = (sh.rows[0] || []).map(norm), hm = {};
    for (const [k, names] of Object.entries(TRK_HEAD)){ const ix = head.findIndex(h=>names.includes(h)); if (ix >= 0) hm[k] = ix; }
    if (hm.title != null && hm.status != null && hm.client != null && (hm.due != null || hm.start != null)) return {sheet:sh, hm};
  }
  return null;
}
function trkBoard(word){
  const w0 = String(word ?? '').trim(); if (!w0) return null;
  const byName = w=>S.boards.find(b=>sameText(b.name, w) || (b.alias && sameText(b.alias, w))) || S.boards.find(b=>loose(b.name) === loose(w) || (b.alias && loose(b.alias) === loose(w))) || null;
  return byName(w0) || byName(clientNow(w0));
}
function trkColumn(cols, word){
  const w = String(word ?? '').trim(); if (!w) return cols[0];
  const hit = cols.find(c=>sameText(c.name, w)) || cols.find(c=>sameText(statusWord(c), w)) || cols.find(c=>loose(c.name) === loose(w));
  if (hit) return hit;
  if (TRK_BASIC.has(w)) return cols.find(c=>c.cat === TRK_STATUS[w]) || null;
  return null;
}
function planTracker({sheet, hm}){
  const get = (r, k)=>hm[k] == null ? '' : r[hm[k]];
  const str = (r, k)=>String(get(r, k) ?? '').trim();
  const newBoards = new Map(), newCols = new Map(), used = new Set(), rows = [];
  const fallbackBoard = isAll() ? (boardById(P.lastBoard?.[wsId]) || S.boards[0]) : curBoard(), tplCols = COLS_PPC();
  sheet.rows.slice(1).forEach((r, k)=>{
    const title = str(r, 'title'); if (!title) return;
    const cw = str(r, 'client');
    let b = trkBoard(cw), bKey = b ? b.id : null;
    if (!b && cw){ bKey = 'new:' + loose(clientNow(cw)); if (!newBoards.has(bKey)) newBoards.set(bKey, clientNow(cw)); }
    if (!b && !cw){ b = fallbackBoard; bKey = b?.id; }
    if (!bKey) return;
    const cols = b ? b.columns : tplCols, sw = str(r, 'status');
    let col = sw ? trkColumn(cols, sw) : null;
    if (!col && sw){ const ck = bKey + '|' + loose(sw); if (!newCols.has(ck)) newCols.set(ck, {bKey, name:sw, cat:TRK_STATUS[sw] || 'doing'}); }
    const pw = str(r, 'prio'), prio = TRK_PRIO_IN[norm(pw)] || (pw ? 'medium' : null);
    const docCell = get(r, 'docs'), docText = String(docCell ?? '').trim(), docUrl = sheet.links?.[`${k+1},${hm.docs}`] || '';
    const links = []; const seen = new Set();
    const addL = (url, title)=>{ const u = normUrl(url); if (u && !seen.has(u)){ seen.add(u); links.push({url:u, title:title || ''}); } };
    if (docUrl) addL(docUrl, docText && !/^https?:\/\//i.test(docText) ? docText.split('\n')[0] : '');
    for (const m of docText.matchAll(/https?:\/\/[^\s<>"]+/g)) addL(m[0], '');
    const docNote = docText && !docUrl && !/https?:\/\//i.test(docText) ? docText : '';
    const row = {k, title, bKey, boardName:b ? b.name : cw, colId:b ? col?.id || null : null, statusWord:sw, prio, start:cellDate(get(r, 'start')), due:cellDate(get(r, 'due')),
      links, notes:[String(get(r, 'notes') ?? '').trim(), docNote ? `${TRACKER.head[6]}: ${docNote}` : ''].filter(Boolean).join('\n'), alias:b && cw && !isOldClient(cw) && !sameText(b.name, cw) && !(b.alias && sameText(b.alias, cw)) ? cw : null};
    // same client + same name = the same task
    const ex = b ? S.issues.find(i=>!used.has(i.id) && i.boardId === b.id && sameText(i.title, title)) : null;
    if (ex){ used.add(ex.id); row.id = ex.id; row.changed = trkChanged(ex, row);
      row.kept = (!row.statusWord ? 1 : 0) + (!row.prio ? 1 : 0) + (!row.due && ex.due ? 1 : 0) + (!row.notes && ex.desc.trim() ? 1 : 0); }
    rows.push(row);
  });
  return {rows, newBoards, newCols, sheetName:sheet.name};
}
// an empty cell never clears anything on the site: only filled cells update a task
function trkChanged(i, r){
  const col = colOf(i);
  return (r.prio && TRK_PRIO_OUT[i.priority] !== TRK_PRIO_OUT[r.prio]) || (!!r.statusWord && (r.colId ? r.colId !== i.status : statusWord(col) !== r.statusWord))
    || (r.start && r.start !== i.start) || (r.due && r.due !== i.due) || (r.notes && r.notes !== (i.desc || '').trim())
    || r.links.some(l=>!i.links.some(x=>x.url === l.url));
}
let trackerPending = null;
function openTrackerImport(plan){
  if (!plan.rows.length) return toast(t('trk.empty'));
  const add = plan.rows.filter(r=>!r.id).length, upd = plan.rows.filter(r=>r.id && r.changed).length, same = plan.rows.filter(r=>r.id && !r.changed).length;
  const nb = [...plan.newBoards.values()], nc = [...new Set([...plan.newCols.values()].map(c=>c.name))];
  const dlg = showModal(`${head(t('trk.title'))}<div class="c-body">
      <p class="confirm-p">${esc(t('trk.found', {sheet:plan.sheetName, n:plan.rows.length}))}</p>
      <ul class="trk-sum"><li><b>${add}</b> ${esc(t('trk.new'))}</li><li><b>${upd}</b> ${esc(t('trk.upd'))}</li>${same ? `<li><b>${same}</b> ${esc(t('trk.same'))}</li>` : ''}</ul>
      ${nb.length ? `<p class="confirm-p">${esc(t('trk.newBoards'))} <b data-bidi>${esc(nb.join(', '))}</b></p>` : ''}
      ${nc.length ? `<p class="confirm-p">${esc(t('trk.newCols'))} <b data-bidi>${esc(nc.join(', '))}</b></p>` : ''}
      ${(k=>k ? `<p class="confirm-p">${esc(tn('trk.kept', k))}</p>` : '')(plan.rows.reduce((n, r)=>n + (r.kept || 0), 0))}
      <p class="confirm-p muted-p">${esc(t('trk.keep'))}</p></div>
    <div class="m-foot"><span class="sp"></span><button class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button><button class="btn primary" data-act="trk-apply">${esc(t('trk.apply'))}</button></div>`, 'confirm', 'narrow');
  dlg.setAttribute('aria-label', t('trk.title'));
  trackerPending = plan;
}
function applyTracker(){
  const plan = trackerPending; trackerPending = null; if (!plan || !canEdit()) return;
  const boardOfKey = new Map(S.boards.map(b=>[b.id, b]));
  let order = (S.boards.reduce((m, b)=>Math.max(m, b.order ?? 0), 0)) + 1000;
  for (const [k, name] of plan.newBoards){
    let key = keyFromName(name) || 'C', n = 2; const base = key; while (S.boards.some(x=>x.key === key)) key = (base + n++).slice(0, 10);
    const b = {id:'b-'+uid(), name, key, color:PALETTE[S.boards.length % PALETTE.length][1], logo:null, seq:0, order, columns:COLS_PPC(), tc:TC}; order += 1000;
    S.boards.push(b); boardOfKey.set(k, b);
  }
  const colFor = (b, word)=>{
    let c = trkColumn(b.columns, word); if (c) return c;
    if (!word) return b.columns[0];
    c = {id:'c'+uid(), name:word, cat:TRK_STATUS[word] || 'doing'};
    // a new "done" kind goes last; anything else just before the first done column
    const di = b.columns.findIndex(x=>x.cat === 'done');
    if (c.cat === 'done' || di < 0) b.columns.push(c); else b.columns.splice(di, 0, c);
    return c;
  };
  const now = Date.now(); let made = 0, changed = 0;
  plan.rows.forEach((r, ix)=>{
    const b = boardOfKey.get(r.bKey); if (!b) return;
    if (r.alias && !b.alias) b.alias = r.alias;
    if (r.id){
      const i = get(r.id); if (!i) return;
      const col = r.statusWord ? colFor(b, r.statusWord) : null;
      if (!trkChanged(i, {...r, colId:col?.id || null})) return;
      if (r.prio && TRK_PRIO_OUT[i.priority] !== TRK_PRIO_OUT[r.prio]) i.priority = r.prio;
      if (col && col.id !== i.status){ const was = colOf(i); i.status = col.id; if (col.cat === 'done' && was?.cat !== 'done') i.doneAt = now; if (col.cat !== 'done') i.doneAt = null; }
      if (r.start) i.start = r.start;
      if (r.due) i.due = r.due;
      if (r.notes && r.notes !== i.desc.trim()) i.desc = r.notes;
      r.links.forEach(l=>{ if (!i.links.some(x=>x.url === l.url)) i.links.push({id:uid(), url:l.url, title:l.title}); });
      touch(i, t('trk.hist', {sheet:plan.sheetName})); changed++;
    } else {
      const col = colFor(b, r.statusWord);
      b.seq = Math.max(b.seq || 0, ...S.issues.filter(x=>x.boardId === b.id).map(x=>x.num)) + 1;
      const i = normIssue('i-'+uid(), {boardId:b.id, num:b.seq, type:S.types[0]?.id, title:r.title, status:col.id, priority:r.prio || 'medium',
        start:r.start || today(), due:r.due || '', desc:r.notes || '', labels:[normLabel(b.name)], links:r.links.map(l=>({id:uid(), ...l})),
        created:now + ix, updated:now, doneAt:col.cat === 'done' ? now : null, creator:me(), history:[{id:uid(), at:now, text:t('trk.histNew', {sheet:plan.sheetName}), by:user.name}]});
      i.rank = undefined; S.issues.push(i); made++;
    }
  });
  closeModal(); save(); render();
  toast(t('trk.done', {a:made, u:changed}), {ms:6000});
}
// export: every open (not archived) task of the workspace, in the order they were created
function exportTracker(){
  flush();
  const boards = S.boards.slice().sort((a, b)=>(a.order ?? 0) - (b.order ?? 0));
  const list = S.issues.filter(i=>!i.archived && boardOf(i)).sort((a, b)=>(a.created - b.created) || (a.num - b.num));
  const rows = list.map(i=>{
    const b = boardOf(i), c = colOf(i), l = i.links[0];
    const docs = !i.links.length ? null : i.links.length === 1 ? {text:l.title || l.url, url:l.url} : {text:i.links.map(x=>x.title ? `${x.title} – ${x.url}` : x.url).join('\n'), url:l.url};
    return {title:i.title, prio:TRK_PRIO_OUT[i.priority] || 'בינונית', client:b.alias || b.name, status:c ? statusWord(c) : '', start:i.start, due:i.due, docs, notes:i.desc};
  });
  const clients = boards.map(b=>b.alias || b.name);
  const statuses = [...new Set([...TRACKER.statuses, ...boards.flatMap(b=>b.columns.map(statusWord))])];
  const blob = writeTracker({rows, clients, statuses});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `task-tracker-${today()}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 1500);
  toast(tn('trk.exported', rows.length));
}

function openImportChoice(conv, raw){
  const dlg = showModal(`${head(t('imp.title'))}<p class="confirm-p">${esc(t('imp.text2', {b:tn('boards', conv.boards.length), i:tn('issues', conv.issues.length)}))}</p>
    ${isOwner() ? `<p class="confirm-p muted-p">${esc(t('imp.replaceHint'))}</p>` : ''}
    <div class="m-foot" style="border:0"><span class="sp"></span><button class="btn subtle" data-act="close-modal">${esc(t('btn.cancel'))}</button>
      ${isOwner() ? `<button class="btn danger" data-act="imp-replace">${esc(t('imp.replace'))}</button>` : ''}<button class="btn primary" data-act="imp-add">${esc(t('imp.add'))}</button></div>`, 'confirm', 'narrow');
  dlg.setAttribute('aria-label', t('imp.title'));
  importPending = {conv, raw};
}
let importPending = null;
function applyImport(mode){
  const p = importPending; importPending = null; if (!p) return;
  const conv = p.conv;
  if (mode === 'replace'){
    S.boards = conv.boards.map((b,ix)=>({...b, order:ix*1000}));
    S.types = conv.types.slice();
    S.issues = conv.issues.map(i=>({...i, rank:undefined}));
    if (p.raw?.archive) S.archive = normArchive(p.raw.archive);
    P.active[wsId] = ALL; savePrefs();
  } else {
    const startOrder = (S.boards.at(-1)?.order ?? 0) + 1000;
    conv.boards.forEach((b,ix)=>{ b.order = startOrder + ix*1000; let k = b.key, n = 2; while (S.boards.some(x=>x.key===k)) k = b.key + (n++); b.key = k; });
    const known = new Set(S.types.map(x=>x.id)), knownNames = new Map(S.types.map(x=>[norm(x.name), x.id]));
    const remap = new Map();
    conv.types.forEach(x=>{ if (known.has(x.id)) return; const same = knownNames.get(norm(x.name)); if (same) remap.set(x.id, same); else S.types.push(x); });
    S.boards.push(...conv.boards.map(b=>({...b, tc:TC}))); S.issues.push(...conv.issues.map(i=>({...i, type:remap.get(i.type) || i.type, rank:undefined})));
  }
  save(); closeModal(); render(); toast(t('toast.imported', {n:tn('issues', conv.issues.length)}));
}
$('#importFile').addEventListener('change', async e=>{
  const f = e.target.files[0]; e.target.value = ''; if (!f || !canEdit()) return;
  try {
    const raw = await readBackupFile(f);
    if (raw.tracker){ openTrackerImport(planTracker(raw.tracker)); return; }
    const conv = convertLocal(raw);
    if (!conv.boards.length) throw new Error('bad');
    openImportChoice(conv, raw);
  } catch(err){ console.error(err); toast(err?.code === 'no-inflate' ? t('toast.importOld') : t('toast.importBad')); }
});

/* ================= Theme, language, shell ================= */
const sysDark = matchMedia('(prefers-color-scheme: dark)');
const THEMES = ['light','dark','purple'];
const effectiveTheme = ()=>THEMES.includes(P.theme) ? P.theme : (sysDark.matches ? 'dark' : 'light');
const themeIcon = ()=>({light:I.sun, dark:I.moon, purple:I.palette})[effectiveTheme()];
const THEME_SW = {light:['#FFFFFF','linear-gradient(135deg,#F4F6FA,#E9F2FF)'], dark:['#2C333A','linear-gradient(135deg,#1D2125,#161A1D)'], purple:['linear-gradient(90deg,#4A2373,#7E4CB8)','linear-gradient(135deg,#FAF7FE,#EBDFFA)']};
function themePicker(){
  return `<div class="menu-h th">${esc(t('theme.title'))}</div><div class="theme-row" role="group" aria-label="${esc(t('theme.title'))}">${THEMES.map(th=>
    `<button type="button" class="theme-opt" data-act="set-theme" data-th="${th}" aria-pressed="${th===effectiveTheme()}"><span class="sw-t"><i style="background:${THEME_SW[th][0]}"></i><i style="background:${THEME_SW[th][1]}"></i></span>${esc(t('theme.'+th))}</button>`).join('')}</div>`;
}
function setTheme(th, animate){
  if (th){ P.theme = th; savePrefs(); }
  const root = document.documentElement, apply = ()=>{ root.dataset.theme = effectiveTheme(); };
  // Cross-fade the whole page on the GPU instead of transitioning every element's colors.
  if (animate && !reduce && document.startViewTransition && root.dataset.theme !== effectiveTheme()) document.startViewTransition(apply);
  else apply();
  const btn = $('#themeBtn'); if (btn){ btn.innerHTML = themeIcon(); btn.title = btn.ariaLabel = t('nav.theme'); btn.setAttribute('aria-haspopup', 'menu'); }
  const meta = document.querySelector('meta[name="theme-color"]'); if (meta) meta.content = {light:'#FFFFFF', dark:'#1D2125', purple:'#4A2373'}[effectiveTheme()];
}
sysDark.addEventListener?.('change', ()=>{ if (!P.theme) setTheme(null); });
function applyShell(){ shell.classList.toggle('side-closed', !P.side); }
matchMedia('(max-width:1024px)').addEventListener?.('change', ()=>shell.classList.remove('side-open'));
function staticText(){
  $$('[data-t]').forEach(el=>{ el.textContent = t(el.dataset.t); });
  $$('[data-tt]').forEach(el=>{ el.title = t(el.dataset.tt); el.setAttribute('aria-label', t(el.dataset.tt)); });
  $$('[data-tp]').forEach(el=>{ el.placeholder = t(el.dataset.tp); el.setAttribute('aria-label', t(el.dataset.tp)); });
}
function rerenderGate(){
  if (screen === 'signin') renderSignIn();
  else if (screen === 'welcome') renderWelcome();
}
function relocalize(){
  staticText(); setTheme(null); renderNavWs(); rerenderGate();
  if (modalKind){ const k = modalKind; closeModal(true); if (k === 'wsSettings') openWsSettings(); if (k === 'request') openRequestAccess(); }
  render();
}

/* ================= Init ================= */
$$('[data-i]').forEach(el=>el.insertAdjacentHTML('afterbegin', I[el.dataset.i]));
if (!['board','list','table','cal','dash','archive'].includes(P.view)) P.view = 'board';
setLang(getLang()); staticText(); setTheme(null); applyShell();
renderLoading();
setInterval(()=>{ if (screen === 'app' && !drag && !modalKind && !composer && !popAnchor) render(); }, 60000);
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(()=>{});
createStore().then(s=>{
  store = s;
  document.body.classList.toggle('demo', s.mode === 'demo');
  $('#demoBar').hidden = s.mode !== 'demo';
  if (s.mode === 'demo') window.__wb = { get S(){ return S; }, get ws(){ return ws; }, get wsList(){ return wsList; }, flush, render, P }; // test hook, demo mode only
  $('#sideNote').dataset.t = s.mode === 'demo' ? 'side.noteDemo' : 'side.note'; staticText();
  s.onAuth(onAuth, e=>{ console.error(e); toast(t('toast.signinFailed')); });
}).catch(err=>{ console.error(err); showGate(gateFrame(`<h1>${esc(t('gate.errTitle'))}</h1><p class="gate-sub">${esc(t('gate.errText'))}</p>`)); });

