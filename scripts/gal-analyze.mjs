import * as XLSX from 'xlsx';
import fs from 'node:fs';

const file = 'input/תאונות  רבעון 1-2026.xls';
const buf = fs.readFileSync(file);
const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, codepage: 1255 });

const ws = wb.Sheets['תאונות 1-2026'];
const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });

// Helpers
const toNum = v => {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).replace(/[₪,\s]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
};
const norm = s => s == null ? '' : String(s).trim();

// Cleanse
const clean = [];
const quality = [];
for (let i = 0; i < rows.length; i++) {
  const r = rows[i];
  const accid = norm(r["מס' תאונה"]);
  const date = norm(r["תאריך תאונה"]);
  const veh = norm(r["מס' רכב"]);
  const driver = norm(r["שם נהג"]);
  const branch = norm(r["מחלקה"]);
  const company = norm(r["חברה"]);
  const codeChar = norm(r["קוד איפיון"]);
  const charac = norm(r["איפיון תאונה"]);
  const fault = norm(r["אשם/לא"]);
  const desc = norm(r["תאור תאונה קצר"]);
  const damageA = toNum(r["נזק לא מכוסה A"]);
  const leasing = norm(r["סוג ליסינג"]);
  const owner = norm(r["בעלות"]);
  const loss = toNum(r["הפסדים"]);

  if (!accid && !veh && !branch) {
    quality.push({ row: i+2, issue: 'שורה ריקה', detail: '' });
    continue;
  }
  if (!veh) quality.push({ row: i+2, issue: 'חסר מספר רכב', detail: `accid=${accid}` });
  if (!branch) quality.push({ row: i+2, issue: 'חסר מחלקה/סניף', detail: `accid=${accid}` });
  if (!company) quality.push({ row: i+2, issue: 'חסרה חברה', detail: `accid=${accid}` });
  if (damageA === 0 && loss === 0) quality.push({ row: i+2, issue: 'אין עלות מדווחת', detail: `accid=${accid} desc=${desc}` });
  if (damageA < 0 || loss < 0) quality.push({ row: i+2, issue: 'ערך שלילי בעלות', detail: `accid=${accid}` });

  clean.push({
    accid, date, vehicle_id: veh, driver, branch, company,
    code: codeChar, characterization: charac, fault, description: desc,
    damage_uncovered: damageA, leasing, owner, loss,
    source_file: file, source_row: i + 2
  });
}

// Duplicate detection
const seen = new Map();
for (const r of clean) {
  const k = r.accid;
  if (k && seen.has(k)) {
    quality.push({ row: r.source_row, issue: 'מס\' תאונה כפול', detail: `accid=${k}` });
  } else if (k) {
    seen.set(k, r.source_row);
  }
}

// Outliers (loss): mean ± 3σ
const losses = clean.map(r => r.loss).filter(v => v > 0);
const mean = losses.reduce((a,b)=>a+b,0)/losses.length;
const std = Math.sqrt(losses.reduce((a,b)=>a+(b-mean)**2,0)/losses.length);
const hi = mean + 3*std, lo = Math.max(0, mean - 3*std);
for (const r of clean) {
  if (r.loss > hi) quality.push({ row: r.source_row, issue: 'הפסד חריג סטטיסטית', detail: `loss=${r.loss} (> mean+3σ = ${hi.toFixed(0)})` });
}

console.log('Total rows in sheet:', rows.length);
console.log('Clean records:', clean.length);
console.log('Quality issues:', quality.length);
console.log('Loss stats: mean=', mean.toFixed(2), 'std=', std.toFixed(2), 'hi(3σ)=', hi.toFixed(2));

// Aggregate by branch, division (company), and group
const byBranch = new Map();
const byDivision = new Map();
const distinctVehicles = new Set();
let totalLoss = 0, totalDamageA = 0, totalAccidents = 0;
const faultByCounts = new Map();
const charByCounts = new Map();

for (const r of clean) {
  totalAccidents++;
  totalLoss += r.loss;
  totalDamageA += r.damage_uncovered;
  if (r.vehicle_id) distinctVehicles.add(r.vehicle_id);

  if (!byBranch.has(r.branch)) byBranch.set(r.branch, { branch: r.branch, division: r.company, accidents: 0, loss: 0, damage: 0, vehicles: new Set() });
  const b = byBranch.get(r.branch);
  b.accidents++; b.loss += r.loss; b.damage += r.damage_uncovered;
  if (r.vehicle_id) b.vehicles.add(r.vehicle_id);

  if (!byDivision.has(r.company)) byDivision.set(r.company, { division: r.company, accidents: 0, loss: 0, damage: 0, vehicles: new Set(), branches: new Set() });
  const d = byDivision.get(r.company);
  d.accidents++; d.loss += r.loss; d.damage += r.damage_uncovered;
  if (r.vehicle_id) d.vehicles.add(r.vehicle_id);
  d.branches.add(r.branch);

  faultByCounts.set(r.fault, (faultByCounts.get(r.fault)||0) + 1);
  charByCounts.set(r.characterization, (charByCounts.get(r.characterization)||0) + 1);
}

const branchRows = [...byBranch.values()].map(b => ({
  חטיבה: b.division,
  סניף: b.branch,
  'מס תאונות': b.accidents,
  'רכבים מעורבים': b.vehicles.size,
  'סך הפסדים (₪)': Math.round(b.loss),
  'נזק לא מכוסה A (₪)': Math.round(b.damage),
  'ממוצע הפסד לתאונה (₪)': b.accidents ? Math.round(b.loss/b.accidents) : 0,
  'תאונות לרכב מעורב': b.vehicles.size ? +(b.accidents/b.vehicles.size).toFixed(2) : 0
})).sort((a,b)=> b['סך הפסדים (₪)'] - a['סך הפסדים (₪)']);

const divisionRows = [...byDivision.values()].map(d => ({
  חטיבה: d.division,
  'מס סניפים': d.branches.size,
  'מס תאונות': d.accidents,
  'רכבים מעורבים': d.vehicles.size,
  'סך הפסדים (₪)': Math.round(d.loss),
  'נזק לא מכוסה A (₪)': Math.round(d.damage),
  'ממוצע הפסד לתאונה (₪)': d.accidents ? Math.round(d.loss/d.accidents) : 0,
  'תאונות לרכב מעורב': d.vehicles.size ? +(d.accidents/d.vehicles.size).toFixed(2) : 0
})).sort((a,b)=> b['סך הפסדים (₪)'] - a['סך הפסדים (₪)']);

// Top vehicles by loss
const byVehicle = new Map();
for (const r of clean) {
  if (!r.vehicle_id) continue;
  if (!byVehicle.has(r.vehicle_id)) byVehicle.set(r.vehicle_id, { vehicle_id: r.vehicle_id, driver: r.driver, branch: r.branch, division: r.company, accidents: 0, loss: 0 });
  const v = byVehicle.get(r.vehicle_id);
  v.accidents++; v.loss += r.loss;
}
const vehicleRows = [...byVehicle.values()].map(v => ({
  'מס רכב': v.vehicle_id, 'נהג': v.driver, 'סניף': v.branch, 'חטיבה': v.division,
  'מס תאונות': v.accidents, 'סך הפסדים (₪)': Math.round(v.loss)
})).sort((a,b)=> b['סך הפסדים (₪)'] - a['סך הפסדים (₪)']);

const charRows = [...charByCounts.entries()].map(([k,v])=>({ 'איפיון תאונה': k, 'מקרים': v })).sort((a,b)=> b['מקרים']-a['מקרים']);
const faultRows = [...faultByCounts.entries()].map(([k,v])=>({ 'אשמה / סטטוס': k, 'מקרים': v })).sort((a,b)=> b['מקרים']-a['מקרים']);

const summary = {
  totalAccidents, totalLoss, totalDamageA,
  distinctVehicles: distinctVehicles.size,
  avgLossPerAccident: totalAccidents ? Math.round(totalLoss/totalAccidents) : 0,
  divisions: byDivision.size,
  branches: byBranch.size
};

// Output JSON for next stage
fs.writeFileSync('output/_gal_intermediate.json', JSON.stringify({
  clean, quality, branchRows, divisionRows, vehicleRows, charRows, faultRows, summary
}, null, 2), 'utf8');

console.log('\n=== SUMMARY ===');
console.log(summary);
console.log('\nTOP 5 BRANCHES BY LOSS:');
console.table(branchRows.slice(0,5));
console.log('\nDIVISIONS:');
console.table(divisionRows);
console.log('\nTOP 5 VEHICLES BY LOSS:');
console.table(vehicleRows.slice(0,5));
console.log('\nCHARACTERIZATIONS:');
console.table(charRows.slice(0,10));
console.log('\nFAULT:');
console.table(faultRows);
