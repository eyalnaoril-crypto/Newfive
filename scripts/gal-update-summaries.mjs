// scripts/gal-update-summaries.mjs
// ----------------------------------
// Purpose: Update the two summary blocks "ריכוז השמירה" and "ריכוז הט"מ" for Q1-2026
// inside a COPY of the source workbook. The originals in input/ are never written to.
//
// Strategy:
//   1. Load source workbook from input/
//   2. For each summary sheet, locate the existing (empty) "רבעון 1-2026" block at the bottom
//   3. Compute per-branch aggregates from the raw "תאונות 1-2026" tab and "אירועים 1-2026" tab
//   4. Map each summary-sheet branch name to the corresponding source branch(es)
//   5. Write values into the existing block cells (preserve row positions to keep
//      any existing cross-sheet refs working in Excel later)
//   6. Update the totals row (סה"כ) with column sums
//   7. Save to output/<timestamp>-fleet-source-updated-2026Q1.xlsx
//
// Run: node scripts/gal-update-summaries.mjs

import * as XLSX from 'xlsx';
import fs from 'node:fs';
import path from 'node:path';

const SRC = 'input/תאונות  רבעון 1-2026.xls';
const ts = (() => {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
})();
const OUT = `output/${ts}-fleet-source-updated-2026Q1.xlsx`;

// ---------- load ----------
const wb = XLSX.read(fs.readFileSync(SRC), { type: 'buffer', cellDates: true, codepage: 1255 });

// helper: parse a numeric cell that might be a string like "1,000" or "₪ 280"
function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[₪,\s]/g, '');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

// ---------- aggregate raw accident data per source-branch ----------
const accAoa = XLSX.utils.sheet_to_json(wb.Sheets['תאונות 1-2026'], { header: 1, defval: null, raw: false });
// columns (0-indexed): 0=מס' תאונה 1=תאריך 2=מס' רכב 3=שם נהג 4=מחלקה(branch+div) 5=חברה(division)
// 6=קוד איפיון 7=איפיון תאונה 8=אשם/לא 9=תאור 10=נזק לא מכוסה A 11=סוג ליסינג 12=בעלות 13=הפסדים
const accidents = [];
for (let i = 1; i < accAoa.length; i++) {
  const r = accAoa[i] || [];
  if (!r[0]) continue;
  accidents.push({
    branch: String(r[4] || '').trim(),    // e.g. "אילת - אבטחה"
    company: String(r[5] || '').trim(),   // e.g. "G-1  פתרונות אבטחה"
    code: String(r[6] || '').trim(),
    type: String(r[7] || '').trim(),      // איפיון תאונה
    blame: String(r[8] || '').trim(),     // אשם/לא  – often a sub-category
    loss: num(r[13]),
    notCovered: num(r[10]),
  });
}
console.log(`Accidents parsed: ${accidents.length}`);

// classify each accident into one of the summary columns
// Categories (per the Q4-2025 column header):
//   F: נהג החברה אשם
//   G: צד ג' אשם + פריצה + חניה
//   H: נזקים + נזק מרכב תחתון
//   I: נזקי סיום עיסקה
//   J: טוטל-לוס
function classify(a) {
  const t = a.type + ' ' + a.blame + ' ' + a.code;
  if (/טוטל[ -]?לוס/i.test(t) || /טוטלוס/.test(t)) return 'J';
  if (/סיום\s*עסקה|בהחזרת רכב|ליסינג/.test(t) && /נזק/.test(t)) return 'I';
  if (/מרכב\s*תחתון|תחתון/.test(t)) return 'H';
  if (/צד\s*ג|פריצה|חניה|אומים|צמיגים|אביזר|פח"?ע|חפץ|בע"ח|גניב/.test(t)) return 'G';
  if (/נהג\s*החברה\s*אשם|עצמי|אופנוע/.test(t)) return 'F';
  // default
  return 'H';
}

// Build per-source-branch aggregate
const byBranch = new Map();
function ensure(b) {
  if (!byBranch.has(b)) byBranch.set(b, { F: 0, G: 0, H: 0, I: 0, J: 0, count: 0, totalCost: 0, eolCost: 0 });
  return byBranch.get(b);
}
for (const a of accidents) {
  const rec = ensure(a.branch);
  rec.count += 1;
  rec.totalCost += a.loss;
  const cat = classify(a);
  rec[cat] += 1;
  if (cat === 'I') rec.eolCost += a.loss;
}

// ---------- aggregate violations from אירועים 1-2026 ----------
const evAoa = XLSX.utils.sheet_to_json(wb.Sheets['אירועים 1-2026'], { header: 1, defval: null, raw: false });
// cols: 0=קוד 1=תאריך 2=רישוי 3=נהג 4=קוד קבוצה 5=קבוצת רכבים (=company) 6=מספר פנימי (=branch) 7=נושא 8=קוד 9=איפיון
const violationsByBranch = new Map();
for (let i = 1; i < evAoa.length; i++) {
  const r = evAoa[i] || [];
  if (!r[0]) continue;
  const branch = String(r[6] || '').trim();
  violationsByBranch.set(branch, (violationsByBranch.get(branch) || 0) + 1);
}
console.log(`Event-rows parsed: ${evAoa.length - 1}`);

// ---------- branch mapping ----------
// For each TEMPLATE branch name, list which source-branch keys flow into it.
// Source keys are in the form "<branch> - <division>" (mostly).
//
// SHEET: ריכוז השמירה (division = G-1 פתרונות אבטחה + G-1 פתרונות ניקיון)
// EXACT source-branch keys only — no substring matching to avoid double-counting.
const shemiraMap = {
  'אילת':            ['אילת - אבטחה'],
  'אשדוד':           ['אשדוד - אבטחה'],
  'באר שבע':         ['באר שבע - אבטחה'],
  'חדרה':            ['חדרה - אבטחה'],
  'חיפה +חדרה':      ['חיפה - אבטחה', 'חדרה - אבטחה'],
  'ירושלים':         ['ירושלים - אבטחה'],
  'מלמ':             ['מלמ - אבטחה'],
  'כפר סבא':         ['כפר סבא - אבטחה'],
  'מטה':             ['מטה - אבטחה'],
  'מפגשים':          ['מפגשים - אבטחה'],
  'מרכז':            ['מרכז - אבטחה'],
  'עכו':             ['עכו - אבטחה'],
  'עפולה':           ['עפולה - אבטחה'],
  'פ"ת':             ['פ"ת - אבטחה', 'פתח תקווה - אבטחה'],
  'ראש פינה':        ['ראש פינה - אבטחה'],
  'ראש פינה ניקיון': ['ראש פינה - ניקיון', 'ראש פינה ניקיון'],
  'רחובות':          ['רחובות - אבטחה'],
  'ת"א ניקיון':      ['ת"א ניקיון', 'תל אביב - ניקיון'],
  'ת"א ':            ['ת"א - אבטחה', 'תל אביב - אבטחה'],
  'עזריאלי ':        ['עזריאלי - אבטחה'],
  'אבטחה מדלגת':     ['אבטחה מדלגת'],
};

// SHEET: ריכוז הט"מ (G-1 טכנולוגיות מיגון + G1 מערכות ממוחשבות)
const tmMap = {
  'אילת':              ['אילת - מוקדים'],
  'חניונים':           ['חניונים - מוקדים'],
  'אשדוד':             ['אשדוד - מוקדים'],
  'ב"ש':               ['באר שבע - מוקדים'],
  'גוש דן ':           ['גוש דן - מוקדים'],
  'הוטלו':             ['הוטלו'],
  'הנהלה-מטה':         ['הנהלה'],
  'חדרה':              ['חדרה - מוקדים'],
  'חטיבת מוקדים ':     ['חטיבת מוקדים'],
  'חיפה':              ['חיפה - מוקדים'],
  'ירושלים':           ['ירושלים - מוקדים'],
  'כספים':             ['כספים'],
  'לוגיסטיקה':         ['מרכז לוגיסטי', 'לוגיסטיקה'],
  'עפולה':             ['עפולה - מוקדים'],
  'ערד':               ['ערד - מוקדים'],
  'פרויקט אזיקים':     ['פרויקט אזיקים'],
  'פרוייקטים +acvs':   ['טכנולוגיות'],
  'ציק פוינט':         ["צ'ק פוינט"],
  'משרד החינוך ':      ['משרד החינוך'],
  'כלבי אשמורת ':      ['גוש דן-כלבי אשמורת'],
};

// EXACT match only — no substring traps.
function aggregateForTemplate(srcKeys) {
  const r = { F: 0, G: 0, H: 0, I: 0, J: 0, count: 0, totalCost: 0, eolCost: 0, viol: 0 };
  for (const k of srcKeys) {
    const v = byBranch.get(k);
    if (v) {
      r.F += v.F; r.G += v.G; r.H += v.H; r.I += v.I; r.J += v.J;
      r.count += v.count; r.totalCost += v.totalCost; r.eolCost += v.eolCost;
    }
    const vio = violationsByBranch.get(k);
    if (vio) r.viol += vio;
  }
  return r;
}

// ---------- locate and fill the Q1-2026 block in each summary sheet ----------
function findBlockStart(aoa, label) {
  // label is e.g. "רבעון 1-2026" or "רבעון -1-2026"
  for (let i = aoa.length - 1; i >= 0; i--) {
    const v = (aoa[i] || [])[1];
    if (typeof v === 'string' && v.replace(/\s+/g, '').includes(label.replace(/\s+/g, ''))) return i;
  }
  return -1;
}

// Generic writer: writes value to ws cell at (row 0-indexed, col 0-indexed)
function setCell(ws, r, c, v) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (v === null || v === undefined || v === '') {
    if (ws[addr]) ws[addr] = { t: 's', v: '' };
    return;
  }
  if (typeof v === 'number') ws[addr] = { t: 'n', v };
  else ws[addr] = { t: 's', v: String(v) };
  // Extend !ref if needed
  if (!ws['!ref']) ws['!ref'] = `${addr}:${addr}`;
  const range = XLSX.utils.decode_range(ws['!ref']);
  if (r > range.e.r) range.e.r = r;
  if (c > range.e.c) range.e.c = c;
  ws['!ref'] = XLSX.utils.encode_range(range);
}

function fillBlock(sheetName, mapping, blockLabel) {
  console.log(`\n--- Filling ${sheetName} :: ${blockLabel} ---`);
  const ws = wb.Sheets[sheetName];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false, blankrows: true });
  const headerRow = findBlockStart(aoa, blockLabel);
  if (headerRow < 0) { console.log('  block not found'); return; }
  // The block layout (relative to header row):
  //   header row +0:  "רבעון 1-2026"        (col B / idx 1)
  //   +1:             blank
  //   +2:             column labels row    (col B onwards)
  //   +3 .. +N:       branch rows (one per template key, IN THE ORDER PRESENT IN THE BLOCK)
  //   then blank
  //   then סה"כ row
  console.log(`  block starts at row ${headerRow + 1}`);

  // Discover the branch rows by scanning until a blank row, then סה"כ
  const branchRows = []; // [{ rowIdx, name }]
  let i = headerRow + 3;
  while (i < aoa.length) {
    const v = (aoa[i] || [])[1];
    if (v === null || v === undefined || String(v).trim() === '') break;
    if (String(v).trim() === 'סה"כ') break;
    branchRows.push({ rowIdx: i, name: String(v) });
    i++;
  }
  // find סה"כ row
  let totalRow = -1;
  for (let j = i; j < Math.min(aoa.length, i + 6); j++) {
    const v = (aoa[j] || [])[1];
    if (typeof v === 'string' && v.trim() === 'סה"כ') { totalRow = j; break; }
  }
  console.log(`  template branches: ${branchRows.length}, totals row: ${totalRow + 1}`);

  // Column indices for the שמירה sheet block (from inspection):
  //  B=1: סניף name
  //  C=2: כמות כלי רכב בסניף
  //  D=3: כלי רכב סיור
  //  E=4: סך ק"מ לרבעון
  //  F=5: נהג החברה אשם
  //  G=6: צד ג' אשם + פריצה + חניה
  //  H=7: נזקים+נזק מרכב תחתון
  //  I=8: נזקי סיום עיסקה
  //  J=9: טוטל-לוס
  //  K=10: כמות תאונות ונזקים
  //  L=11: עלות כלל המקרים
  //  M=12: עלות נזקי סיום עסקה
  //  N=13: נוהל 6
  //  O=14: עבירות תנועה
  //  P=15: כמות תאונות ממוצעת לרכב
  //  Q=16: עלות נזקים ממוצעת לרכב
  //  R=17: עלות נזק ל-1000 ק"מ
  //  S=18: ציון סניף (שמירה only — in ט"מ this slot is "רבעון 1-25")
  // Track totals
  const totals = { F: 0, G: 0, H: 0, I: 0, J: 0, count: 0, totalCost: 0, eolCost: 0, viol: 0 };
  let filled = 0;
  for (const br of branchRows) {
    const keyName = br.name.trim();
    // find matching mapping entry, tolerant to extra spaces
    let srcKeys = null;
    for (const k of Object.keys(mapping)) {
      if (k.trim() === keyName) { srcKeys = mapping[k]; break; }
    }
    if (!srcKeys) {
      // try contains both ways
      for (const k of Object.keys(mapping)) {
        if (k.trim().replace(/\s+/g, '') === keyName.replace(/\s+/g, '')) { srcKeys = mapping[k]; break; }
      }
    }
    if (!srcKeys) {
      console.log(`  [skip] no mapping for branch "${keyName}"`);
      continue;
    }
    const agg = aggregateForTemplate(srcKeys);
    // Write F..J (5..9), K (10), L (11), M (12), O (14), N (13 — leave N/A)
    setCell(ws, br.rowIdx, 5, agg.F || 0);
    setCell(ws, br.rowIdx, 6, agg.G || 0);
    setCell(ws, br.rowIdx, 7, agg.H || 0);
    setCell(ws, br.rowIdx, 8, agg.I || 0);
    setCell(ws, br.rowIdx, 9, agg.J || 0);
    setCell(ws, br.rowIdx, 10, agg.count);
    setCell(ws, br.rowIdx, 11, Math.round(agg.totalCost));
    setCell(ws, br.rowIdx, 12, Math.round(agg.eolCost));
    // נוהל 6 = N/A (not in source) — leave blank
    setCell(ws, br.rowIdx, 14, agg.viol);
    // Leave C/D/E (fleet count, סיור, km) blank — N/A from source
    // Leave P/Q/R blank — depend on fleet count / km
    totals.F += agg.F; totals.G += agg.G; totals.H += agg.H;
    totals.I += agg.I; totals.J += agg.J;
    totals.count += agg.count; totals.totalCost += agg.totalCost;
    totals.eolCost += agg.eolCost; totals.viol += agg.viol;
    filled++;
  }
  // Totals row
  if (totalRow > 0) {
    setCell(ws, totalRow, 5, totals.F);
    setCell(ws, totalRow, 6, totals.G);
    setCell(ws, totalRow, 7, totals.H);
    setCell(ws, totalRow, 8, totals.I);
    setCell(ws, totalRow, 9, totals.J);
    setCell(ws, totalRow, 10, totals.count);
    setCell(ws, totalRow, 11, Math.round(totals.totalCost));
    setCell(ws, totalRow, 12, Math.round(totals.eolCost));
    setCell(ws, totalRow, 14, totals.viol);
  }
  console.log(`  filled ${filled} branches; totals: count=${totals.count}, totalCost=${Math.round(totals.totalCost).toLocaleString('he-IL')}, viol=${totals.viol}`);
  return { filled, totals, totalRow: totalRow + 1, headerRow: headerRow + 1 };
}

const r1 = fillBlock('ריכוז השמירה', shemiraMap, 'רבעון 1-2026');
const r2 = fillBlock('ריכוז הט"מ',  tmMap,      'רבעון -1-2026');

// ---------- save copy ----------
if (!fs.existsSync('output')) fs.mkdirSync('output', { recursive: true });
XLSX.writeFile(wb, OUT);
console.log(`\nSaved: ${OUT}`);
console.log(`Sheets in file: ${wb.SheetNames.length}`);

// ---------- write a small validation summary ----------
console.log('\n=== Validation summary ===');
console.log(`Total accidents read from source: ${accidents.length}`);
console.log(`Sum of accidents in שמירה totals: ${r1?.totals?.count}`);
console.log(`Sum of accidents in ט"מ totals:    ${r2?.totals?.count}`);
console.log(`Combined (should be <= ${accidents.length} since some branches may not map): ${(r1?.totals?.count || 0) + (r2?.totals?.count || 0)}`);

// list unmapped source branches (so we can flag them)
const mapped = new Set();
for (const k of Object.values(shemiraMap).flat()) mapped.add(k);
for (const k of Object.values(tmMap).flat()) mapped.add(k);
const unmapped = [];
for (const [b, v] of byBranch) {
  if (!mapped.has(b)) unmapped.push({ branch: b, count: v.count, cost: Math.round(v.totalCost) });
}
console.log('\nUnmapped source branches (data not propagated):');
for (const u of unmapped) console.log(`  - "${u.branch}"  (תאונות=${u.count}, הפסדים=${u.cost.toLocaleString('he-IL')})`);
