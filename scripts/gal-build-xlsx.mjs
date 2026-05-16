// Build final Excel with multiple sheets + chart using ExcelJS
import ExcelJS from 'exceljs';
import fs from 'node:fs';

const data = JSON.parse(fs.readFileSync('output/_gal_intermediate.json', 'utf8'));
const { clean, quality, branchRows, divisionRows, vehicleRows, charRows, faultRows, summary } = data;

const wb = new ExcelJS.Workbook();
wb.creator = 'גל — סוכן ניתוח Excel לצי רכב';
wb.created = new Date();
wb.views = [{ rightToLeft: true }];

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
const HEADER_FONT = { name: 'Tahoma', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
const BODY_FONT = { name: 'Tahoma', size: 11 };
const TITLE_FONT = { name: 'Tahoma', size: 16, bold: true, color: { argb: 'FF1F4E78' } };
const SUB_FONT = { name: 'Tahoma', size: 12, bold: true, color: { argb: 'FF1F4E78' } };

function styleHeader(row) {
  row.eachCell(c => { c.fill = HEADER_FILL; c.font = HEADER_FONT; c.alignment = { horizontal: 'center', vertical: 'middle', readingOrder: 'rtl' }; c.border = { bottom: { style: 'thin' } }; });
  row.height = 24;
}
function styleBody(row) {
  row.eachCell(c => { c.font = BODY_FONT; c.alignment = { vertical: 'middle', readingOrder: 'rtl' }; });
}
function autoWidth(ws, cols) {
  cols.forEach((c, i) => { ws.getColumn(i+1).width = c; });
}

// ============ 1) סיכום מנהלים ============
{
  const ws = wb.addWorksheet('סיכום מנהלים', { views: [{ rightToLeft: true }] });
  ws.getCell('A1').value = 'דוח תאונות צי רכב — רבעון 1, 2026';
  ws.getCell('A1').font = TITLE_FONT;
  ws.mergeCells('A1:F1');
  ws.getCell('A2').value = `קבוצת G1 | תאריך הפקה: ${new Date().toLocaleDateString('he-IL')}`;
  ws.getCell('A2').font = { name: 'Tahoma', size: 10, italic: true, color: { argb: 'FF555555' } };
  ws.mergeCells('A2:F2');

  ws.getCell('A4').value = 'מדדים מרכזיים — קבוצת G1';
  ws.getCell('A4').font = SUB_FONT;

  const kpis = [
    ['סך תאונות ברבעון', summary.totalAccidents],
    ['רכבים מעורבים (ייחודיים)', summary.distinctVehicles],
    ['סך הפסדים לחברה (₪)', summary.totalLoss],
    ['סך נזק לא מכוסה A (₪)', Math.round(summary.totalDamageA)],
    ['ממוצע הפסד לתאונה (₪)', summary.avgLossPerAccident],
    ['מספר חטיבות פעילות', summary.divisions],
    ['מספר סניפים מעורבים', summary.branches],
    ['שיעור תאונות לרכב מעורב', +(summary.totalAccidents/summary.distinctVehicles).toFixed(2)],
  ];
  let r = 5;
  for (const [label, value] of kpis) {
    ws.getCell(`A${r}`).value = label;
    ws.getCell(`A${r}`).font = { name: 'Tahoma', size: 11, bold: true };
    ws.getCell(`B${r}`).value = value;
    ws.getCell(`B${r}`).font = { name: 'Tahoma', size: 11 };
    if (typeof value === 'number' && value > 100) ws.getCell(`B${r}`).numFmt = '#,##0';
    r++;
  }

  r += 2;
  ws.getCell(`A${r}`).value = '3 ממצאים מרכזיים';
  ws.getCell(`A${r}`).font = SUB_FONT;
  r++;
  const findings = [
    `חטיבת "G-1 טכנולוגיות מיגון" אחראית ל-${divisionRows[0]['מס תאונות']} תאונות ול-₪${divisionRows[0]['סך הפסדים (₪)'].toLocaleString()} (≈${Math.round(divisionRows[0]['סך הפסדים (₪)']/summary.totalLoss*100)}% מסך ההפסדים), אף שהיא נמוכה במספר הסניפים.`,
    `סניף "ירושלים - מוקדים" הוא המוביל בהפסדים (₪${branchRows[0]['סך הפסדים (₪)'].toLocaleString()}) — עם ממוצע ₪${branchRows[0]['ממוצע הפסד לתאונה (₪)'].toLocaleString()} לתאונה, פי 3.3 מהממוצע הקבוצתי.`,
    `סניף "באר שבע - אבטחה" מציג ריכוז חריג של תאונות לרכב: 7 תאונות ב-3 רכבים בלבד (2.33 תאונות/רכב) — דורש בדיקת התנהגות נהיגה.`
  ];
  for (const f of findings) {
    const cell = ws.getCell(`A${r}`);
    cell.value = `• ${f}`;
    cell.font = BODY_FONT;
    cell.alignment = { wrapText: true, vertical: 'top', readingOrder: 'rtl' };
    ws.mergeCells(`A${r}:F${r}`);
    ws.getRow(r).height = 45;
    r++;
  }
  autoWidth(ws, [32, 22, 15, 15, 15, 15]);
}

// ============ 2) קבוצה ============
{
  const ws = wb.addWorksheet('קבוצה', { views: [{ rightToLeft: true }] });
  ws.getCell('A1').value = 'מדדים ברמת הקבוצה — G1 Group';
  ws.getCell('A1').font = TITLE_FONT;
  ws.mergeCells('A1:C1');
  const rows = [
    ['מדד', 'ערך', 'יחידה'],
    ['סך תאונות', summary.totalAccidents, 'מקרים'],
    ['רכבים מעורבים ייחודיים', summary.distinctVehicles, 'רכבים'],
    ['סך הפסדים', summary.totalLoss, '₪'],
    ['סך נזק לא מכוסה A', Math.round(summary.totalDamageA), '₪'],
    ['ממוצע הפסד לתאונה', summary.avgLossPerAccident, '₪'],
    ['תאונות לרכב מעורב', +(summary.totalAccidents/summary.distinctVehicles).toFixed(2), 'יחס'],
    ['חטיבות פעילות', summary.divisions, 'חטיבות'],
    ['סניפים מעורבים', summary.branches, 'סניפים'],
  ];
  const headerRow = ws.addRow(rows[0]);
  styleHeader(headerRow);
  for (let i = 1; i < rows.length; i++) {
    const r = ws.addRow(rows[i]);
    styleBody(r);
    if (typeof rows[i][1] === 'number' && rows[i][1] > 100) r.getCell(2).numFmt = '#,##0';
  }
  autoWidth(ws, [35, 18, 12]);
}

// ============ 3) חטיבות ============
{
  const ws = wb.addWorksheet('חטיבות', { views: [{ rightToLeft: true }] });
  const cols = ['חטיבה','מס סניפים','מס תאונות','רכבים מעורבים','סך הפסדים (₪)','נזק לא מכוסה A (₪)','ממוצע הפסד לתאונה (₪)','תאונות לרכב מעורב'];
  const h = ws.addRow(cols); styleHeader(h);
  for (const r of divisionRows) {
    const row = ws.addRow(cols.map(c => r[c]));
    styleBody(row);
    row.getCell(5).numFmt = '#,##0';
    row.getCell(6).numFmt = '#,##0';
    row.getCell(7).numFmt = '#,##0';
  }
  autoWidth(ws, [28, 12, 12, 14, 16, 18, 20, 18]);

  // Chart: bar chart of loss by division (ExcelJS chart support is limited;
  // fall back to embedding a sorted, formatted table that doubles as a chart-ready layout.)
  // Add an in-cell bar visualization using conditional formatting
  const maxLoss = Math.max(...divisionRows.map(r => r['סך הפסדים (₪)']));
  ws.addConditionalFormatting({
    ref: `E2:E${1 + divisionRows.length}`,
    rules: [{
      type: 'dataBar',
      cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: maxLoss }],
      color: { argb: 'FF1F4E78' },
      showValue: true
    }]
  });
}

// ============ 4) סניפים ============
{
  const ws = wb.addWorksheet('סניפים', { views: [{ rightToLeft: true }] });
  const cols = ['חטיבה','סניף','מס תאונות','רכבים מעורבים','סך הפסדים (₪)','נזק לא מכוסה A (₪)','ממוצע הפסד לתאונה (₪)','תאונות לרכב מעורב'];
  const h = ws.addRow(cols); styleHeader(h);
  // sort by division then loss desc
  const sorted = [...branchRows].sort((a,b)=> {
    if (a['חטיבה'] !== b['חטיבה']) return a['חטיבה'].localeCompare(b['חטיבה'],'he');
    return b['סך הפסדים (₪)'] - a['סך הפסדים (₪)'];
  });
  for (const r of sorted) {
    const row = ws.addRow(cols.map(c => r[c]));
    styleBody(row);
    row.getCell(5).numFmt = '#,##0';
    row.getCell(6).numFmt = '#,##0';
    row.getCell(7).numFmt = '#,##0';
  }
  autoWidth(ws, [28, 24, 12, 14, 16, 18, 20, 18]);

  const maxLoss = Math.max(...sorted.map(r => r['סך הפסדים (₪)']));
  ws.addConditionalFormatting({
    ref: `E2:E${1 + sorted.length}`,
    rules: [{
      type: 'dataBar',
      cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: maxLoss }],
      color: { argb: 'FFC00000' },
      showValue: true
    }]
  });
}

// ============ 5) רכבים ============
{
  const ws = wb.addWorksheet('רכבים', { views: [{ rightToLeft: true }] });
  const cols = ['מס רכב','נהג','סניף','חטיבה','מס תאונות','סך הפסדים (₪)'];
  const h = ws.addRow(cols); styleHeader(h);
  for (const r of vehicleRows) {
    const row = ws.addRow(cols.map(c => r[c]));
    styleBody(row);
    row.getCell(6).numFmt = '#,##0';
  }
  autoWidth(ws, [14, 26, 24, 28, 12, 16]);
}

// ============ 6) פילוח לפי איפיון ============
{
  const ws = wb.addWorksheet('פילוח לפי איפיון', { views: [{ rightToLeft: true }] });
  ws.getCell('A1').value = 'התפלגות תאונות לפי איפיון';
  ws.getCell('A1').font = TITLE_FONT;
  ws.mergeCells('A1:B1');
  const h = ws.addRow(['איפיון תאונה','מקרים']); styleHeader(h);
  const total = charRows.reduce((a,r)=>a+r['מקרים'],0);
  for (const r of charRows) {
    const row = ws.addRow([r['איפיון תאונה'], r['מקרים']]);
    styleBody(row);
  }
  // add percent column manually
  ws.getCell('C2').value = 'אחוז';
  ws.getCell('C2').fill = HEADER_FILL;
  ws.getCell('C2').font = HEADER_FONT;
  for (let i = 0; i < charRows.length; i++) {
    const cell = ws.getCell(`C${3+i}`);
    cell.value = charRows[i]['מקרים']/total;
    cell.numFmt = '0.0%';
    cell.font = BODY_FONT;
  }
  autoWidth(ws, [35, 12, 10]);
}

// ============ 7) נתונים מטוייבים ============
{
  const ws = wb.addWorksheet('נתונים מטוייבים', { views: [{ rightToLeft: true }] });
  const cols = ['מס תאונה','תאריך','מס רכב','שם נהג','סניף (מחלקה)','חטיבה (חברה)','קוד איפיון','איפיון','אשם/לא','תאור קצר','נזק לא מכוסה A (₪)','סוג ליסינג','בעלות','הפסדים (₪)','קובץ מקור','שורת מקור'];
  const h = ws.addRow(cols); styleHeader(h);
  for (const r of clean) {
    const row = ws.addRow([r.accid, r.date, r.vehicle_id, r.driver, r.branch, r.company, r.code, r.characterization, r.fault, r.description, r.damage_uncovered, r.leasing, r.owner, r.loss, r.source_file, r.source_row]);
    styleBody(row);
    row.getCell(11).numFmt = '#,##0';
    row.getCell(14).numFmt = '#,##0';
  }
  autoWidth(ws, [10, 12, 12, 22, 22, 26, 10, 26, 26, 30, 18, 14, 22, 14, 30, 10]);
  ws.views = [{ rightToLeft: true, state: 'frozen', ySplit: 1 }];
}

// ============ 8) מילון KPI ============
{
  const ws = wb.addWorksheet('מילון KPI', { views: [{ rightToLeft: true }] });
  const cols = ['שם KPI','הגדרה','נוסחה','יחידה','רמת חתך'];
  const h = ws.addRow(cols); styleHeader(h);
  const dict = [
    ['סך תאונות', 'מספר רשומות התאונה בתקופה', 'COUNT(records)', 'מקרים', 'רכב/סניף/חטיבה/קבוצה'],
    ['סך הפסדים', 'סכום עמודת "הפסדים" — עלות נטו לחברה', 'Σ(הפסדים)', '₪', 'רכב/סניף/חטיבה/קבוצה'],
    ['סך נזק לא מכוסה A', 'סכום עמודת "נזק לא מכוסה A" — חשיפה אקטוארית', 'Σ(נזק לא מכוסה A)', '₪', 'רכב/סניף/חטיבה/קבוצה'],
    ['ממוצע הפסד לתאונה', 'ההפסד הממוצע למקרה', 'Σ(הפסדים) / COUNT(records)', '₪', 'רכב/סניף/חטיבה/קבוצה'],
    ['רכבים מעורבים', 'מספר רכבים שונים שמופיעים ברשומות התקופה', 'COUNT(DISTINCT vehicle_id)', 'רכבים', 'סניף/חטיבה/קבוצה'],
    ['תאונות לרכב מעורב', 'מדד עומס תאונות יחסי על רכב מעורב', 'COUNT(records) / COUNT(DISTINCT vehicle_id)', 'יחס', 'סניף/חטיבה/קבוצה'],
    ['אחוז מתוך סך', 'משקל היחידה מסך התאונות/הפסדים בקבוצה', '(value_unit / value_group) × 100', '%', 'סניף/חטיבה'],
  ];
  for (const d of dict) {
    const row = ws.addRow(d);
    styleBody(row);
    row.eachCell(c => { c.alignment = { wrapText: true, vertical: 'top', readingOrder: 'rtl' }; });
  }
  autoWidth(ws, [22, 38, 38, 12, 25]);
}

// ============ 9) איכות נתונים ============
{
  const ws = wb.addWorksheet('איכות נתונים', { views: [{ rightToLeft: true }] });
  ws.getCell('A1').value = `סה"כ סוגיות שזוהו: ${quality.length}`;
  ws.getCell('A1').font = SUB_FONT;
  ws.mergeCells('A1:C1');
  ws.addRow([]);
  const cols = ['שורת מקור','סוג סוגיה','פירוט'];
  const h = ws.addRow(cols); styleHeader(h);
  for (const q of quality) {
    const row = ws.addRow([q.row, q.issue, q.detail]);
    styleBody(row);
  }
  autoWidth(ws, [14, 30, 60]);
}

// ============ 10) מקורות נתונים ============
{
  const ws = wb.addWorksheet('מקורות נתונים', { views: [{ rightToLeft: true }] });
  const cols = ['פרט','ערך'];
  const h = ws.addRow(cols); styleHeader(h);
  const meta = [
    ['קובץ מקור', 'input/תאונות  רבעון 1-2026.xls'],
    ['גיליון אנליטי', 'תאונות 1-2026'],
    ['פורמט קובץ', 'Excel 97-2003 (.xls legacy), Codepage 1255'],
    ['גיליונות בקובץ — סה"כ', '114 (היסטוריה רב-שנתית מ-2012 עד 2026)'],
    ['שורות נתונים שנקראו', String(clean.length)],
    ['תאריך קליטה', new Date().toLocaleString('he-IL')],
    ['תקופה מנותחת', '01/01/2026 – 31/03/2026 (רבעון 1, 2026)'],
    ['גרסת סקריפט', 'gal-build-xlsx.mjs / v1'],
    ['הערות', 'גיליונות "ריכוז השמירה" ו"ריכוז הט"מ" — תבניות ריכוז ריקות לרבעון 4-2026, לא רלוונטיות לתקופה זו. גיליון "נתוני מגמות" מכיל סדרת זמן רב-שנתית אך עמודת 1-26 כמעט ריקה.'],
  ];
  for (const m of meta) {
    const row = ws.addRow(m);
    styleBody(row);
    row.eachCell(c => { c.alignment = { wrapText: true, vertical: 'top', readingOrder: 'rtl' }; });
  }
  autoWidth(ws, [30, 80]);
}

// Reorder: put summary first
wb.worksheets.sort((a, b) => {
  const order = ['סיכום מנהלים','קבוצה','חטיבות','סניפים','רכבים','פילוח לפי איפיון','נתונים מטוייבים','מילון KPI','איכות נתונים','מקורות נתונים'];
  return order.indexOf(a.name) - order.indexOf(b.name);
});

// Filename
const now = new Date();
const pad = n => String(n).padStart(2,'0');
const stamp = `${now.getFullYear()}${pad(now.getMonth()+1)}${pad(now.getDate())}-${pad(now.getHours())}${pad(now.getMinutes())}`;
const outPath = `output/${stamp}-fleet-accidents-group-2026Q1.xlsx`;
await wb.xlsx.writeFile(outPath);
console.log('Wrote:', outPath);
