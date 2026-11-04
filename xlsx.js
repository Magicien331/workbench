// Minimal .xlsx writer and reader, no dependencies.
// Writer: styled sheets (right-to-left, frozen header, filters, dropdown lists, colored "chips" via conditional formatting).
// Reader: plain cell values from files saved by this app, Excel, Google Sheets or LibreOffice.

const te = new TextEncoder(), td = new TextDecoder();

/* ---------- zip ---------- */
const CRC = (()=>{ const t = new Uint32Array(256); for (let n = 0; n < 256; n++){ let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(b){ let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
function zip(files){ // files: [{name, data:Uint8Array}] → stored (uncompressed) zip
  const parts = [], central = []; let off = 0;
  for (const f of files){
    const name = te.encode(f.name), crc = crc32(f.data), n = f.data.length;
    const h = new DataView(new ArrayBuffer(30));
    h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
    h.setUint16(10, 0, true); h.setUint16(12, 0x21, true); h.setUint32(14, crc, true); h.setUint32(18, n, true); h.setUint32(22, n, true);
    h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
    parts.push(new Uint8Array(h.buffer), name, f.data);
    const c = new DataView(new ArrayBuffer(46));
    c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true);
    c.setUint16(12, 0, true); c.setUint16(14, 0x21, true); c.setUint32(16, crc, true); c.setUint32(20, n, true); c.setUint32(24, n, true);
    c.setUint16(28, name.length, true); c.setUint32(42, off, true);
    central.push(new Uint8Array(c.buffer), name);
    off += 30 + name.length + n;
  }
  const cdSize = central.reduce((s, p)=>s + p.length, 0);
  const e = new DataView(new ArrayBuffer(22));
  e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
  return new Blob([...parts, ...central, new Uint8Array(e.buffer)], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
async function inflateRaw(bytes){
  if (typeof DecompressionStream === 'undefined') throw Object.assign(new Error('no-inflate'), {code:'no-inflate'});
  const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}
async function unzip(buf){
  const u8 = new Uint8Array(buf), dv = new DataView(buf), out = new Map();
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50){ eocd = i; break; }
  if (eocd < 0) throw new Error('not-zip');
  const count = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
  for (let k = 0; k < count; k++){
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('bad-zip');
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true),
          xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), loff = dv.getUint32(p + 42, true);
    const name = td.decode(u8.subarray(p + 46, p + 46 + nlen));
    const start = loff + 30 + dv.getUint16(loff + 26, true) + dv.getUint16(loff + 28, true);
    const raw = u8.subarray(start, start + csize);
    if (!name.endsWith('/')) out.set(name, method === 0 ? raw : method === 8 ? await inflateRaw(raw) : null);
    p += 46 + nlen + xlen + clen;
  }
  return out;
}

/* ---------- writer ---------- */
const wellFormed = s=>{ // drop lone surrogates (they would make the XML invalid)
  if (s.toWellFormed) return s.toWellFormed();
  let o = ''; for (let i = 0; i < s.length; i++){ const c = s.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF){ const d = s.charCodeAt(i + 1); if (d >= 0xDC00 && d <= 0xDFFF){ o += s[i] + s[i + 1]; i++; } continue; }
    if (c >= 0xDC00 && c <= 0xDFFF) continue; o += s[i]; }
  return o;
};
const xesc = s=>wellFormed(String(s)).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '')
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
export const colName = i=>{ let s = ''; i++; while (i){ const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
const argb = hex=>'FF' + String(hex || '#000000').replace('#', '').toUpperCase().padStart(6, '0').slice(-6);
const qs = name=>`'${String(name).replace(/'/g, "''")}'`;
export const MAX_CELL = 32000;

// Cell styles (index = s attribute). Kept in one table so sheets stay consistent.
export const ST = {plain:0, head:1, wrap:2, date:3, dtime:4, center:5, title:6, meta:7, num:8, key:9};
function stylesXml(accent, dxfs){
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy hh:mm"/></numFmts>
<fonts count="5"><font><sz val="11"/><color rgb="FF172B4D"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font><font><sz val="9"/><color rgb="FF626F86"/><name val="Arial"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FF172B4D"/><name val="Arial"/><family val="2"/></font><font><sz val="10"/><color rgb="FF44546F"/><name val="Consolas"/><family val="3"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="${argb(accent)}"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="thin"><color rgb="FFDCDFE4"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="10">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="165" fontId="2" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="${dxfs.length}">${dxfs.map(d=>`<dxf><font><b/><color rgb="${argb(d.fg)}"/></font><fill><patternFill patternType="solid"><bgColor rgb="${argb(d.bg)}"/></patternFill></fill></dxf>`).join('')}</dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>`;
}
function cellXml(ref, c){
  if (c == null || c === '') return '';
  if (typeof c !== 'object') c = {v:c};
  const s = c.s ? ` s="${c.s}"` : '';
  if (c.v == null || c.v === '') return c.s ? `<c r="${ref}"${s}/>` : '';
  if (typeof c.v === 'number' && Number.isFinite(c.v)) return `<c r="${ref}"${s}><v>${c.v}</v></c>`;
  const str = String(c.v).slice(0, MAX_CELL);
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xesc(str)}</t></is></c>`;
}
function sheetXml(sh, ix, dxfIndex, listsSheet){
  const rows = sh.rows || [], nCols = Math.max(1, ...rows.map(r=>r.length), (sh.cols||[]).length);
  const last = `${colName(nCols - 1)}${Math.max(1, rows.length)}`;
  const view = `<sheetViews><sheetView ${sh.rtl ? 'rightToLeft="1" ' : ''}${ix === 0 ? 'tabSelected="1" ' : ''}workbookViewId="0">${sh.freeze ? `<pane ySplit="${sh.freeze}" topLeftCell="A${sh.freeze + 1}" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A${sh.freeze + 1}" sqref="A${sh.freeze + 1}"/>` : ''}</sheetView></sheetViews>`;
  const cols = (sh.cols||[]).length ? `<cols>${sh.cols.map((c,i)=>`<col min="${i+1}" max="${i+1}" width="${c.width || 14}"${c.hidden ? ' hidden="1"' : ''} customWidth="1"/>`).join('')}</cols>` : '';
  const data = rows.map((r, ri)=>{
    const ht = ri === 0 && sh.headHeight ? ` ht="${sh.headHeight}" customHeight="1"` : '';
    return `<row r="${ri+1}"${ht}>${r.map((c, ci)=>cellXml(`${colName(ci)}${ri+1}`, c)).join('')}</row>`;
  }).join('');
  const af = sh.autoFilter && rows.length ? `<autoFilter ref="A1:${colName(nCols - 1)}${rows.length}"/>` : '';
  let pri = 1;
  const cf = (sh.chips||[]).filter(ch=>ch.values.length).map(ch=>{
    const ref = `${colName(ch.col)}2:${colName(ch.col)}${Math.max(2, rows.length + 500)}`;
    return `<conditionalFormatting sqref="${ref}">${ch.values.map(v=>`<cfRule type="cellIs" dxfId="${dxfIndex(v.bg, v.fg)}" priority="${pri++}" operator="equal"><formula>"${xesc(String(v.text).replace(/"/g, '""'))}"</formula></cfRule>`).join('')}</conditionalFormatting>`;
  }).join('');
  const dv = (sh.lists||[]).filter(l=>l.count > 0);
  const dvx = dv.length ? `<dataValidations count="${dv.length}">${dv.map(l=>`<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="0" sqref="${colName(l.col)}2:${colName(l.col)}${Math.max(2, rows.length + 500)}"><formula1>${xesc(`${qs(listsSheet)}!$${colName(l.listCol)}$1:$${colName(l.listCol)}$${l.count}`)}</formula1></dataValidation>`).join('')}</dataValidations>` : '';
  const tab = sh.hidden ? '' : `<sheetPr>${sh.tab ? `<tabColor rgb="${argb(sh.tab)}"/>` : ''}<pageSetUpPr fitToPage="1"/></sheetPr>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${tab}<dimension ref="A1:${last}"/>${view}<sheetFormatPr defaultRowHeight="18"/>${cols}<sheetData>${data}</sheetData>${af}${cf}${dvx}<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>${sh.hidden ? '' : '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>'}</worksheet>`;
}
// spec: {accent, sheets:[{name, rtl, hidden, freeze, headHeight, autoFilter, tab, cols:[{width,hidden}], rows:[[cell]], lists:[{col, values}], chips:[{col, values:[{text,bg,fg}]}]}]}
export function writeXlsx(spec){
  const LISTS = '_lists', sheets = spec.sheets.slice(), dxfs = [], dxfMap = new Map();
  const dxfIndex = (bg, fg)=>{ const k = bg + '|' + fg; if (!dxfMap.has(k)){ dxfMap.set(k, dxfs.length); dxfs.push({bg, fg}); } return dxfMap.get(k); };
  // dropdown sources live in one hidden sheet, one column per list
  const listCols = [];
  sheets.forEach(sh=>(sh.lists||[]).forEach(l=>{ l.values = [...new Set(l.values.filter(v=>v !== '' && v != null).map(String))]; l.listCol = listCols.length; l.count = l.values.length; listCols.push(l.values); }));
  if (listCols.length){
    const n = Math.max(...listCols.map(c=>c.length), 1), rows = [];
    for (let r = 0; r < n; r++) rows.push(listCols.map(c=>c[r] ?? ''));
    sheets.push({name:LISTS, hidden:true, rows});
  }
  const files = [], sheetEntries = [], rels = [], ct = [];
  sheets.forEach((sh, i)=>{
    files.push({name:`xl/worksheets/sheet${i+1}.xml`, data:te.encode(sheetXml(sh, i, dxfIndex, LISTS))});
    sheetEntries.push(`<sheet name="${xesc(sh.name)}" sheetId="${i+1}"${sh.hidden ? ' state="hidden"' : ''} r:id="rId${i+1}"/>`);
    rels.push(`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`);
    ct.push(`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  });
  const names = sheets.map((sh, i)=>sh.autoFilter && sh.rows.length ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${xesc(`${qs(sh.name)}!$A$1:$${colName(Math.max(...sh.rows.map(r=>r.length)) - 1)}$${sh.rows.length}`)}</definedName>` : '').join('');
  const nS = sheets.length;
  files.unshift(
    {name:'[Content_Types].xml', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${ct.join('')}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`)},
    {name:'_rels/.rels', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`)},
    {name:'docProps/core.xml', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${xesc(spec.title || 'Workbench')}</dc:title><dc:creator>Workbench</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0,19)}Z</dcterms:created></cp:coreProperties>`)},
    {name:'docProps/app.xml', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Workbench</Application></Properties>`)},
    {name:'xl/workbook.xml', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="16000" activeTab="0"/></bookViews><sheets>${sheetEntries.join('')}</sheets>${names ? `<definedNames>${names}</definedNames>` : ''}<calcPr calcId="191029"/></workbook>`)},
    {name:'xl/_rels/workbook.xml.rels', data:te.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}<Relationship Id="rId${nS+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`)}
  );
  // styles last: the dxf list is complete only after all sheets were built
  files.push({name:'xl/styles.xml', data:te.encode(stylesXml(spec.accent || '#5B2C83', dxfs))});
  return zip(files);
}

// raw package writer for files we lay out ourselves (e.g. the task-tracker sheet)
export const zipFiles = files=>zip(files.map(f=>({name:f.name, data:typeof f.data === 'string' ? te.encode(f.data) : f.data})));
export const xmlText = s=>xesc(s);

/* ---------- reader ---------- */
const parseXml = bytes=>{ const doc = new DOMParser().parseFromString(td.decode(bytes), 'application/xml'); if (doc.getElementsByTagName('parsererror').length) throw new Error('bad-xml'); return doc; };
const all = (el, tag)=>[...el.getElementsByTagNameNS('*', tag)];
const colIndex = ref=>{ const m = /^([A-Z]+)/.exec(ref || ''); if (!m) return -1; let n = 0; for (const ch of m[1]) n = n * 26 + ch.charCodeAt(0) - 64; return n - 1; };
const textOf = el=>all(el, 't').filter(t=>!t.closest || !t.parentElement || t.parentElement.localName !== 'rPh').map(t=>t.textContent).join('');
// returns {sheets:[{name, hidden, rows:[[value]]}]} — numbers stay numbers, everything else is text
export async function readXlsx(buf){
  const files = await unzip(buf);
  const get = n=>files.get(n) || files.get(n.replace(/^\//, ''));
  const wb = get('xl/workbook.xml'); if (!wb) throw new Error('not-xlsx');
  const wdoc = parseXml(wb), relsDoc = get('xl/_rels/workbook.xml.rels') ? parseXml(get('xl/_rels/workbook.xml.rels')) : null;
  const rel = new Map(relsDoc ? all(relsDoc, 'Relationship').map(r=>[r.getAttribute('Id'), r.getAttribute('Target')]) : []);
  const sst = get('xl/sharedStrings.xml') ? all(parseXml(get('xl/sharedStrings.xml')), 'si').map(textOf) : [];
  const sheets = [];
  for (const s of all(wdoc, 'sheet')){
    const rid = s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id');
    let target = rel.get(rid) || ''; target = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
    const bytes = get(target); if (!bytes) continue;
    const doc = parseXml(bytes), rows = [], links = {};
    // hyperlinks (e.g. Google Drive file chips in a "documents" column) → links["row,col"] = url
    const relPath = target.replace(/([^/]+)$/, '_rels/$1.rels'), srel = get(relPath) ? new Map(all(parseXml(get(relPath)), 'Relationship').map(r=>[r.getAttribute('Id'), r.getAttribute('Target')])) : new Map();
    for (const h of all(doc, 'hyperlink')){
      const ref = h.getAttribute('ref') || '', id = h.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || h.getAttribute('r:id');
      const url = id ? srel.get(id) : ''; const m = /^([A-Z]+)(\d+)/.exec(ref);
      if (url && m) links[`${Number(m[2]) - 1},${colIndex(m[1])}`] = url;
    }
    let rAuto = 0;
    for (const row of all(doc, 'row')){
      const ri = row.getAttribute('r') ? Number(row.getAttribute('r')) - 1 : rAuto; rAuto = ri + 1;
      const out = []; let cAuto = 0;
      for (const c of [...row.children].filter(x=>x.localName === 'c')){
        const ci = c.getAttribute('r') ? colIndex(c.getAttribute('r')) : cAuto; cAuto = ci + 1;
        const tp = c.getAttribute('t'), v = [...c.children].find(x=>x.localName === 'v');
        let val = null;
        if (tp === 's') val = sst[Number(v?.textContent)] ?? '';
        else if (tp === 'inlineStr'){ const is = [...c.children].find(x=>x.localName === 'is'); val = is ? textOf(is) : ''; }
        else if (tp === 'str' || tp === 'e') val = v?.textContent ?? '';
        else if (tp === 'b') val = v?.textContent === '1';
        else if (v) val = Number(v.textContent);
        const f = [...c.children].find(x=>x.localName === 'f'), hm = f && /HYPERLINK\(\s*"([^"]+)"/i.exec(f.textContent || '');
        if (hm) links[`${ri},${ci}`] = hm[1];
        if (val !== null && val !== '') out[ci] = val;
      }
      rows[ri] = out;
    }
    for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    sheets.push({name:s.getAttribute('name') || '', hidden:(s.getAttribute('state') || 'visible') !== 'visible', rows, links});
  }
  return {sheets};
}
// Excel date serials (days since 1899-12-30) ⇄ ISO dates
export const isoToSerial = iso=>{ const [y, m, d] = String(iso).split('-').map(Number); return Date.UTC(y, m - 1, d) / 86400000 + 25569; };
export const msToSerial = ms=>{ const d = new Date(ms); return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes()) / 86400000) + 25569; };
export const serialToIso = n=>{ const d = new Date(Math.round((n - 25569) * 86400000)); return d.toISOString().slice(0, 10); };
