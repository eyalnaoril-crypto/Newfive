import * as XLSX from 'xlsx';
import fs from 'node:fs';

const file = 'input/תאונות  רבעון 1-2026.xls';
const buf = fs.readFileSync(file);
const wb = XLSX.read(buf, { type: 'buffer', cellDates: true, codepage: 1255 });

const targetSheets = ['תאונות 1-2026', 'אירועים 1-2026', 'ריכוז השמירה', 'ריכוז הט"מ', 'נתוני מגמות'];

for (const name of targetSheets) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log(`MISSING: ${name}`); continue; }
  // Read both as array-of-arrays (to see headers/structure) and as JSON
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });
  console.log(`\n========== SHEET: ${name} ==========`);
  console.log('Total rows (incl headers):', aoa.length);
  console.log('First 8 rows (raw):');
  for (let i = 0; i < Math.min(8, aoa.length); i++) {
    console.log(`R${i}:`, JSON.stringify(aoa[i]));
  }
  if (aoa.length > 8) {
    console.log('Last 3 rows (raw):');
    for (let i = Math.max(0, aoa.length - 3); i < aoa.length; i++) {
      console.log(`R${i}:`, JSON.stringify(aoa[i]));
    }
  }
}
