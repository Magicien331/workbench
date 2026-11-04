// The "task tracker" sheet Elizabeth keeps in Google Sheets ("מעקב משימות"), rebuilt cell for cell:
// same columns, widths, fonts, table style, dropdowns, date check and colour rules, so the file the site
// exports opens exactly like hers and can be imported back.
import { zipFiles, xmlText, colName, isoToSerial } from './xlsx.js';

export const TRACKER = {
  sheet:'מעקב משימות', table:'מעקב_משימות',
  head:['משימה','עדיפות','לקוח','סטטוס','תאריך התחלה','דד ליין','מסמכים','הערות'],
  prios:['גבוהה','בינונית','נמוכה'],
  statuses:['לא התחיל','בביצוע','ממתין לאישור','בוצע','דורש תיקונים','לא בוצע'],
  minRows:205,
};
// her styles.xml, unchanged (fonts, dd/MM/yyyy, colour rules and the table style live here)
const STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><numFmts count="1"><numFmt numFmtId="164" formatCode="dd/MM/yyyy"/></numFmts><fonts count="2"><font><sz val="10.0"/><color rgb="FF000000"/><name val="Arial"/><scheme val="minor"/></font><font><color theme="1"/><name val="Assistant"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="lightGray"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" applyAlignment="1" applyFont="1"/></cellStyleXfs><cellXfs count="9"><xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment readingOrder="0" shrinkToFit="0" vertical="bottom" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="left" readingOrder="0" shrinkToFit="0" vertical="center" wrapText="1"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="left" readingOrder="2" shrinkToFit="0" vertical="center" wrapText="1"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="49" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment readingOrder="0" shrinkToFit="0" vertical="center" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="left" readingOrder="1" shrinkToFit="0" vertical="center" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="164" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="left" readingOrder="1" shrinkToFit="0" vertical="center" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="49" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="left" readingOrder="1" shrinkToFit="0" vertical="center" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="49" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment horizontal="right" readingOrder="2" shrinkToFit="0" vertical="center" wrapText="0"/></xf><xf borderId="0" fillId="0" fontId="1" numFmtId="0" xfId="0" applyAlignment="1" applyFont="1"><alignment horizontal="left" readingOrder="1" shrinkToFit="0" vertical="center" wrapText="0"/></xf></cellXfs><cellStyles count="1"><cellStyle xfId="0" name="Normal" builtinId="0"/></cellStyles><dxfs count="9"><dxf><font/><fill><patternFill patternType="none"/></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FF52405D"/><bgColor rgb="FF52405D"/></patternFill></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFF6F8F9"/><bgColor rgb="FFF6F8F9"/></patternFill></fill><border/></dxf><dxf><border><left style="thin"><color rgb="FF52405D"/></left><right style="thin"><color rgb="FF52405D"/></right><top style="thin"><color rgb="FF52405D"/></top><bottom style="thin"><color rgb="FF52405D"/></bottom></border></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFB7E1CD"/><bgColor rgb="FFB7E1CD"/></patternFill></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFEA9999"/><bgColor rgb="FFEA9999"/></patternFill></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFFCE5CD"/><bgColor rgb="FFFCE5CD"/></patternFill></fill><border/></dxf><dxf><font/><fill><patternFill patternType="solid"><fgColor rgb="FFF3F3F3"/><bgColor rgb="FFF3F3F3"/></patternFill></fill><border/></dxf></dxfs><tableStyles count="1"><tableStyle count="4" pivot="0" name="מעקב משימות-style"><tableStyleElement dxfId="1" type="headerRow"/><tableStyleElement dxfId="2" type="firstRowStripe"/><tableStyleElement dxfId="3" type="secondRowStripe"/><tableStyleElement dxfId="4" size="0" type="wholeTable"/></tableStyle></tableStyles></styleSheet>';
const W = [37.88, 15.13, 15.13, 15.13, 20.13, 20.13, 22.63, 21.75];
// cell styles from her file: header, filled row, empty row
const S_HEAD = [1,1,1,1,1,2,1,1], S_ROW = [3,4,4,4,5,5,6,3], S_EMPTY = [7,4,8,4,5,5,6,6];
const q = s=>xmlText(s).replace(/"/g, '&quot;');
const listF = items=>`"${items.map(x=>String(x).replace(/[",]/g, ' ').trim()).filter(Boolean).join(',')}"`;

// rows: [{title, prio, client, status, start, due, docs:{text, url}, notes}]
export function writeTracker({rows, clients, statuses}){
  const n = Math.max(TRACKER.minRows, rows.length + 20), last = n + 1, links = [];
  const cell = (r, c, v, s)=>{
    const ref = colName(c) + r;
    if (v == null || v === '') return `<c r="${ref}" s="${s}"/>`;
    if (typeof v === 'number') return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
    return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlText(v)}</t></is></c>`;
  };
  const out = [`<row r="1" ht="22.5" customHeight="1">${TRACKER.head.map((h, c)=>cell(1, c, h, S_HEAD[c])).join('')}</row>`];
  for (let k = 0; k < n; k++){
    const r = k + 2, x = rows[k];
    if (!x){ out.push(`<row r="${r}" ht="22.5" customHeight="1">${S_EMPTY.map((s, c)=>cell(r, c, '', s)).join('')}</row>`); continue; }
    if (x.docs?.url) links.push({ref:'G' + r, url:x.docs.url});
    const vals = [x.title, x.prio, x.client, x.status, x.start ? isoToSerial(x.start) : '', x.due ? isoToSerial(x.due) : '', x.docs?.text || '', x.notes];
    out.push(`<row r="${r}" ht="22.5" customHeight="1">${vals.map((v, c)=>cell(r, c, v, S_ROW[c])).join('')}</row>`);
  }
  const R = a=>`${a}2:${a}${last}`;
  const cf = (sq, dxf, pri, f, type='expression')=>`<conditionalFormatting sqref="${sq}"><cfRule type="${type}" dxfId="${dxf}" priority="${pri}"><formula>${q(f)}</formula></cfRule></conditionalFormatting>`;
  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr><dimension ref="A1:H${last}"/><sheetViews><sheetView rightToLeft="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection activeCell="A2" sqref="A2" pane="bottomLeft"/></sheetView></sheetViews><sheetFormatPr customHeight="1" defaultColWidth="12.63" defaultRowHeight="15.75"/><cols>${W.map((w, c)=>`<col customWidth="1" min="${c+1}" max="${c+1}" width="${w}"/>`).join('')}</cols><sheetData>${out.join('')}</sheetData>`
    + cf(`I6:I${last}`, 5, 1, 'LEN(TRIM(I6))>0', 'notContainsBlanks')
    // overdue (red), due today (orange), done (grey). Her sheet checked for "הושלם", a status it never offers; it checks "בוצע" here.
    + cf(`A2:H${last}`, 6, 2, `AND($F2<>"", $F2<TODAY(), $D2<>"בוצע", $D2<>"לא בוצע")`)
    + cf(`A2:H${last}`, 7, 3, `AND($F2=TODAY(), $D2<>"בוצע", $D2<>"לא בוצע")`)
    + cf(`A2:H${last}`, 8, 4, `$D2="בוצע"`)
    + `<dataValidations count="5"><dataValidation type="list" allowBlank="1" sqref="${R('B')}"><formula1>${q(listF(TRACKER.prios))}</formula1></dataValidation>`
    + `<dataValidation type="custom" allowBlank="1" showDropDown="1" sqref="E2:F${last}"><formula1>${q('OR(NOT(ISERROR(DATEVALUE(E2))), AND(ISNUMBER(E2), LEFT(CELL("format", E2))="D"))')}</formula1></dataValidation>`
    + `<dataValidation allowBlank="1" showDropDown="1" sqref="${R('A')} ${R('H')}"/>`
    + `<dataValidation type="list" allowBlank="1" sqref="${R('C')}"><formula1>${q(listF(clients))}</formula1></dataValidation>`
    + `<dataValidation type="list" allowBlank="1" sqref="${R('D')}"><formula1>${q(listF(statuses))}</formula1></dataValidation></dataValidations>`
    + (links.length ? `<hyperlinks>${links.map((l, k)=>`<hyperlink ref="${l.ref}" r:id="rIdL${k+1}"/>`).join('')}</hyperlinks>` : '')
    + `<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rIdT1"/></tableParts></worksheet>`;
  const sheetRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdT1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>${links.map((l, k)=>`<Relationship Id="rIdL${k+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${q(l.url)}" TargetMode="External"/>`).join('')}</Relationships>`;
  const table = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="${TRACKER.table}" displayName="${TRACKER.table}" ref="A1:H${last}"><tableColumns count="8">${TRACKER.head.map((h, c)=>`<tableColumn id="${c+1}" name="${q(h)}"/>`).join('')}</tableColumns><tableStyleInfo name="מעקב משימות-style" showFirstColumn="1" showLastColumn="1" showRowStripes="1" showColumnStripes="0"/></table>`;
  return zipFiles([
    {name:'[Content_Types].xml', data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`},
    {name:'_rels/.rels', data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`},
    {name:'xl/workbook.xml', data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr/><sheets><sheet state="visible" name="${TRACKER.sheet}" sheetId="1" r:id="rId1"/></sheets><definedNames/><calcPr/></workbook>`},
    {name:'xl/_rels/workbook.xml.rels', data:`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`},
    {name:'xl/styles.xml', data:STYLES},
    {name:'xl/worksheets/sheet1.xml', data:sheet},
    {name:'xl/worksheets/_rels/sheet1.xml.rels', data:sheetRels},
    {name:'xl/tables/table1.xml', data:table},
  ]);
}
