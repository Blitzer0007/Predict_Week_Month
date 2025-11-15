// separate_predict_year.js
// Usage: node Prediction_Scripts\separate_predict_year.js Test_Data\weekly.xlsx Test_Data\monthly.xlsx
// Outputs: ./separate_predict_year/weekly_year_predictions.xlsx and ./separate_predict_year/monthly_year_predictions.xlsx

const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const fs = require('fs');
const path = require('path');

if (process.argv.length < 4) {
  console.error('Usage: node separate_predict_year.js weekly.xlsx monthly.xlsx');
  process.exit(1);
}

const weeklyFile = process.argv[2];
const monthlyFile = process.argv[3];

// ---------- create output folder based on script name ----------
const scriptBase = path.basename(process.argv[1], path.extname(process.argv[1]));
const outputFolder = path.join(process.cwd(), scriptBase);
if (!fs.existsSync(outputFolder)) fs.mkdirSync(outputFolder, { recursive: true });
console.log('Output folder:', outputFolder);

// ---------- helpers ----------
const is3Digits = s => {
  if (s === null || s === undefined) return false;
  const str = String(s).trim();
  const m = str.match(/(\d{1,3})$/);
  if (!m) return false;
  return /^[0-9]{3}$/.test(m[1].padStart(3,'0'));
};
const normalize3 = s => {
  const str = String(s||'').trim();
  const m = str.match(/(\d{1,3})$/);
  if (!m) return null;
  return m[1].padStart(3,'0');
};
function mkCounter(){ return { pos0:{}, pos1:{}, pos2:{}, triplets:{}, total:0 }; }
function addToCounter(counter, trip){
  counter.total = (counter.total||0) + 1;
  counter.pos0[trip[0]] = (counter.pos0[trip[0]]||0) + 1;
  counter.pos1[trip[1]] = (counter.pos1[trip[1]]||0) + 1;
  counter.pos2[trip[2]] = (counter.pos2[trip[2]]||0) + 1;
  counter.triplets[trip] = (counter.triplets[trip]||0) + 1;
}

// ---------- read sheet into matrix ----------
async function sheetToMatrix(pathArg, preferredNames=[]){
  if (!fs.existsSync(pathArg)) throw new Error('File not found: ' + pathArg);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(pathArg);
  let sheet = null;
  for (const n of preferredNames) {
    const candidate = wb.getWorksheet(n);
    if (candidate) { sheet = candidate; break; }
  }
  if (!sheet) sheet = wb.worksheets[0];
  const matrix = [];
  sheet.eachRow((row) => {
    const vals = row.values ? row.values.slice(1).map(v => (v === null || v === undefined) ? '' : String(v)) : [];
    matrix.push(vals);
  });
  return matrix;
}

// ---------- parse weekly matrix ----------
function parseWeeklyMatrix(mat){
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const dayNames = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
  let headerIdx = -1;
  for (let i=0;i<Math.min(mat.length,6);i++){
    const row = mat[i].map(c=>up(c));
    let matches = 0;
    for (const dn of dayNames) if (row.some(cell => cell.includes(dn))) matches++;
    if (matches >= 3) { headerIdx = i; break; }
  }
  let start = 0; let colMap = null;
  if (headerIdx >= 0) {
    start = headerIdx + 1;
    colMap = {};
    const header = mat[headerIdx].map(c=>up(c));
    header.forEach((cell, idx) => {
      if (!cell) return;
      if (cell.includes('SUN')) colMap[idx] = 0;
      else if (cell.includes('MON')) colMap[idx] = 1;
      else if (cell.includes('TUE')) colMap[idx] = 2;
      else if (cell.includes('WED')) colMap[idx] = 3;
      else if (cell.includes('THU')) colMap[idx] = 4;
      else if (cell.includes('FRI')) colMap[idx] = 5;
      else if (cell.includes('SAT')) colMap[idx] = 6;
    });
  } else {
    const maxCols = Math.max(...mat.map(r=>r.length));
    if (maxCols >= 7) {
      colMap = {}; for (let c=0;c<7;c++) colMap[c]=c; start = 0;
    } else {
      return [];
    }
  }

  const entries = [];
  for (let r=start; r<mat.length; r++){
    const row = mat[r];
    if (!row) continue;
    for (const [col, weekday] of Object.entries(colMap)){
      const cidx = Number(col);
      const cell = row[cidx];
      if (!cell) continue;
      if (is3Digits(cell)) entries.push({ trip: normalize3(cell), weekday: Number(weekday) });
    }
  }
  return entries;
}

// ---------- parse monthly matrix ----------
function parseMonthlyMatrix(mat){
  if (!mat || mat.length === 0) return [];
  const up = s => String(s||'').toUpperCase();
  const headerIdx = mat.findIndex(row => row && row.some(c => up(c).includes('DATE')));
  const headerRow = mat[ headerIdx >=0 ? headerIdx : 0 ].map(c=>up(c));
  const months = { 'JAN':1,'FEB':2,'MAR':3,'APR':4,'MAY':5,'JUN':6,'JUL':7,'AUG':8,'SEP':9,'OCT':10,'NOV':11,'DEC':12 };
  const colToMonth = {};
  headerRow.forEach((cell, idx) => {
    if (!cell) return;
    for (const key of Object.keys(months)) if (cell.includes(key)) { colToMonth[idx] = months[key]; break; }
    const short = cell.slice(0,3);
    if (!colToMonth[idx] && months[short]) colToMonth[idx] = months[short];
  });
  let dateCol = headerRow.findIndex(c=>up(c).includes('DATE'));
  if (dateCol < 0) dateCol = 0;
  const start = (headerIdx >=0) ? headerIdx + 1 : 1;
  const entries = [];
  for (let r=start; r<mat.length; r++){
    const row = mat[r];
    if (!row) continue;
    const dateVal = row[dateCol];
    const dayNum = parseInt(String(dateVal||'').trim(), 10);
    if (!dayNum || dayNum < 1 || dayNum > 31) continue;
    for (const [cStr, mon] of Object.entries(colToMonth)){
      const cidx = Number(cStr);
      const cell = row[cidx];
      if (!cell) continue;
      if (is3Digits(cell)) entries.push({ trip: normalize3(cell), month: Number(mon), day: dayNum });
    }
  }
  return entries;
}

// ---------- build counters ----------
async function buildWeeklyCounters(weeklyPath){
  const mat = await sheetToMatrix(weeklyPath, ['weekly','Week']);
  const parsed = parseWeeklyMatrix(mat);
  const weekCounters = {}; // 0..6
  for (const p of parsed){
    if (!weekCounters[p.weekday]) weekCounters[p.weekday] = mkCounter();
    addToCounter(weekCounters[p.weekday], p.trip);
  }
  // also build overall weekly aggregate if useful
  const overall = mkCounter();
  for (let w=0; w<=6; w++){
    const c = weekCounters[w] || mkCounter();
    // add counts to overall
    for (const trip of Object.keys(c.triplets)) {
      const cnt = c.triplets[trip];
      for (let i=0;i<cnt;i++) addToCounter(overall, trip);
    }
  }
  return { weekCounters, overall };
}

async function buildMonthlyCounters(monthlyPath){
  const mat = await sheetToMatrix(monthlyPath, ['monthly','Month']);
  const parsed = parseMonthlyMatrix(mat);
  const monthCounters = {}; // 1..12
  const monthDayTrip = {}; // key `${m}-${d}` -> counter of triplets (only triplet counts)
  const overall = mkCounter();
  for (const p of parsed){
    if (!monthCounters[p.month]) monthCounters[p.month] = mkCounter();
    addToCounter(monthCounters[p.month], p.trip);
    const key = `${p.month}-${p.day}`;
    if (!monthDayTrip[key]) monthDayTrip[key] = {};
    monthDayTrip[key][p.trip] = (monthDayTrip[key][p.trip]||0) + 1;
    addToCounter(overall, p.trip);
  }
  // convert monthDayTrip objects to counters
  const monthDayCounters = {};
  for (const [key, map] of Object.entries(monthDayTrip)){
    const cnt = mkCounter();
    for (const [trip, c] of Object.entries(map)){
      cnt.triplets[trip] = c;
      cnt.total += c;
      // update positional counts
      for (let i=0;i<c;i++){
        cnt.pos0[trip[0]] = (cnt.pos0[trip[0]]||0) + 1;
        cnt.pos1[trip[1]] = (cnt.pos1[trip[1]]||0) + 1;
        cnt.pos2[trip[2]] = (cnt.pos2[trip[2]]||0) + 1;
      }
    }
    monthDayCounters[key] = cnt;
  }
  return { monthCounters, monthDayCounters, overall };
}

// ---------- probability & ranking ----------
function buildProbMaps(counter){
  const total = counter.total || 1;
  const p0 = {}, p1 = {}, p2 = {};
  for (let d=0; d<=9; d++){
    const k = String(d);
    p0[k] = (counter.pos0[k] || 0) / total;
    p1[k] = (counter.pos1[k] || 0) / total;
    p2[k] = (counter.pos2[k] || 0) / total;
  }
  return { p0,p1,p2, triplets: counter.triplets || {}, total: counter.total || 0 };
}

function rankCandidatesFromCounter(counter, topKPerPos=6, topN=30){
  // Use triplet evidence if exists (we'll include all triplets)
  const combos = new Map();
  // triplet evidence
  for (const [t, cnt] of Object.entries(counter.triplets || {})){
    const score = (counter.total ? cnt / counter.total : 0);
    combos.set(t, Math.max(combos.get(t) || 0, score));
  }
  // per-position supplement
  const pm = buildProbMaps(counter);
  const sortDigits = pmap => Object.keys(pmap).sort((a,b)=> pmap[b]-pmap[a]);
  const top0 = sortDigits(pm.p0).slice(0, topKPerPos);
  const top1 = sortDigits(pm.p1).slice(0, topKPerPos);
  const top2 = sortDigits(pm.p2).slice(0, topKPerPos);
  for (const a of top0) for (const b of top1) for (const c of top2){
    const n = `${a}${b}${c}`;
    const score = (pm.p0[a]||0)*(pm.p1[b]||0)*(pm.p2[c]||0);
    combos.set(n, Math.max(combos.get(n)||0, score));
  }
  const arr = Array.from(combos.entries()).map(([num,score]) => ({ num, score }));
  arr.sort((x,y) => y.score - x.score);
  return arr.slice(0, topN).map(x => ({ num: x.num, score: x.score, count: (counter.triplets && counter.triplets[x.num]) || 0 }));
}

// ---------- generate date list for next 365 days (starting tomorrow) ----------
function nextNDates(n){
  const out = [];
  let day = dayjs().startOf('day').add(1,'day');
  for (let i=0;i<n;i++){
    out.push(day.format('YYYY-MM-DD'));
    day = day.add(1,'day');
  }
  return out;
}

// ---------- create output workbook for weekly ----------
async function produceWeeklyWorkbook(weeklyPath, outFile){
  const { weekCounters, overall } = await buildWeeklyCounters(weeklyPath);

  const wb = new ExcelJS.Workbook();

  // positional_overall (aggregate of all weekdays)
  const posAllRows = [['Digit','Count_pos0','Prob_pos0','Count_pos1','Prob_pos1','Count_pos2','Prob_pos2']];
  const totalAll = overall.total || 0;
  for (let d=0; d<=9; d++){
    const k = String(d);
    const c0 = overall.pos0[k]||0, c1 = overall.pos1[k]||0, c2 = overall.pos2[k]||0;
    posAllRows.push([k, c0, totalAll? c0/totalAll:0, c1, totalAll? c1/totalAll:0, c2, totalAll? c2/totalAll:0]);
  }
  posAllRows.push(['TOTAL', totalAll]);

  const wsPosAll = wb.addWorksheet('positional_overall');
  posAllRows.forEach(r => wsPosAll.addRow(r));

  // positional_weekday (block per weekday)
  const wsPosWeek = wb.addWorksheet('positional_weekday');
  for (let w=0; w<=6; w++){
    wsPosWeek.addRow([`WEEKDAY_${w}`]);
    const cnt = weekCounters[w] || mkCounter();
    const block = [['Digit','Count_pos0','Prob_pos0','Count_pos1','Prob_pos1','Count_pos2','Prob_pos2']];
    const tot = cnt.total || 0;
    for (let d=0; d<=9; d++){
      const k = String(d);
      const c0 = cnt.pos0[k]||0, c1 = cnt.pos1[k]||0, c2 = cnt.pos2[k]||0;
      block.push([k, c0, tot? c0/tot:0, c1, tot? c1/tot:0, c2, tot? c2/tot:0]);
    }
    block.push(['TOTAL', tot]);
    block.forEach(r => wsPosWeek.addRow(r));
    wsPosWeek.addRow([]);
  }

  // triplets_weekday
  const wsTripWeek = wb.addWorksheet('triplets_weekday');
  wsTripWeek.addRow(['Weekday','Triplet','Count','Probability']);
  for (let w=0; w<=6; w++){
    const cnt = weekCounters[w] || mkCounter();
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c]) => ({t,c, p: tot? c/tot:0}));
    items.sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripWeek.addRow([w, it.t, it.c, it.p]);
  }

  // predictions for next 365 days (weekly-only)
  const dates = nextNDates(365);
  const wsPred = wb.addWorksheet('predictions');
  wsPred.addRow(['Date','Weekday','ObservationsUsed','TopCandidates']);
  for (const d of dates){
    const dt = dayjs(d);
    const wd = dt.day();
    const cnt = weekCounters[wd] || mkCounter();
    const candidates = rankCandidatesFromCounter(cnt, 6, 30);
    const obs = cnt.total || 0;
    const topText = candidates.map(c=> `${c.num} (score:${(c.score*100).toFixed(4)}%, count:${c.count})`).join(' | ');
    wsPred.addRow([d, wd, obs, topText]);
  }

  const outPath = path.join(outputFolder, outFile);
  await wb.xlsx.writeFile(outPath);
  console.log('Weekly workbook written:', outPath);
}

// ---------- create output workbook for monthly ----------
async function produceMonthlyWorkbook(monthlyPath, outFile){
  const { monthCounters, monthDayCounters, overall } = await buildMonthlyCounters(monthlyPath);

  const wb = new ExcelJS.Workbook();

  // positional_month (per month block)
  const wsPosMonth = wb.addWorksheet('positional_month');
  for (let m=1; m<=12; m++){
    wsPosMonth.addRow([`MONTH_${m}`]);
    const cnt = monthCounters[m] || mkCounter();
    const tot = cnt.total || 0;
    wsPosMonth.addRow(['Digit','Count_pos0','Prob_pos0','Count_pos1','Prob_pos1','Count_pos2','Prob_pos2']);
    for (let d=0; d<=9; d++){
      const k = String(d);
      const c0 = cnt.pos0[k]||0, c1 = cnt.pos1[k]||0, c2 = cnt.pos2[k]||0;
      wsPosMonth.addRow([k, c0, tot? c0/tot:0, c1, tot? c1/tot:0, c2, tot? c2/tot:0]);
    }
    wsPosMonth.addRow(['TOTAL', tot]);
    wsPosMonth.addRow([]);
  }

  // triplets_month
  const wsTripMonth = wb.addWorksheet('triplets_month');
  wsTripMonth.addRow(['Month','Triplet','Count','Probability']);
  for (let m=1; m<=12; m++){
    const cnt = monthCounters[m] || mkCounter();
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c])=>({t,c,p: tot? c/tot:0}));
    items.sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripMonth.addRow([m, it.t, it.c, it.p]);
  }

  // triplets_month_day (for exact day-of-month evidence)
  const wsTripMonthDay = wb.addWorksheet('triplets_month_day');
  wsTripMonthDay.addRow(['Month','Day','Triplet','Count','Probability']);
  for (const key of Object.keys(monthDayCounters)) {
    const [m,d] = key.split('-').map(Number);
    const cnt = monthDayCounters[key];
    const tot = cnt.total || 0;
    const items = Object.entries(cnt.triplets || {}).map(([t,c]) => ({t,c,p: tot? c/tot:0}));
    items.sort((a,b)=> b.c - a.c);
    for (const it of items) wsTripMonthDay.addRow([m,d,it.t,it.c,it.p]);
  }

  // predictions for next 365 days (monthly-only)
  const dates = nextNDates(365);
  const wsPred = wb.addWorksheet('predictions');
  wsPred.addRow(['Date','Month','Day','ObservationsUsed','SourceUsed','TopCandidates']);

  for (const d of dates){
    const dt = dayjs(d);
    const m = dt.month() + 1;
    const dayNum = dt.date();
    const key = `${m}-${dayNum}`;
    let source = 'none';
    let cnt = null;
    // prefer exact month-day counter
    if (monthDayCounters[key] && (monthDayCounters[key].total || 0) > 0) {
      cnt = monthDayCounters[key];
      source = 'month-day';
    } else if (monthCounters[m] && (monthCounters[m].total || 0) > 0) {
      cnt = monthCounters[m];
      source = 'month';
    } else {
      cnt = overall;
      source = 'overall';
    }
    const candidates = rankCandidatesFromCounter(cnt, 6, 30);
    const obs = cnt.total || 0;
    const topText = candidates.map(c=> `${c.num} (score:${(c.score*100).toFixed(4)}%, count:${c.count})`).join(' | ');
    wsPred.addRow([d, m, dayNum, obs, source, topText]);
  }

  const outPath = path.join(outputFolder, outFile);
  await wb.xlsx.writeFile(outPath);
  console.log('Monthly workbook written:', outPath);
}

// ---------- run ----------
(async () => {
  try {
    await produceWeeklyWorkbook(weeklyFile, 'weekly_year_predictions.xlsx');
  } catch (err) {
    console.error('Weekly generation failed:', err.message);
  }
  try {
    await produceMonthlyWorkbook(monthlyFile, 'monthly_year_predictions.xlsx');
  } catch (err) {
    console.error('Monthly generation failed:', err.message);
  }
})();
