/* ---------------- State ---------------- */
let TRADES = [];        // {date: Date, dateStr, symbol, type, volume, profit}
let FILTERED_TRADES = []; // TRADES narrowed by the active date-range filter
let capital = 1100;
let goal = 800;
let calYear, calMonth;  // 0-indexed month
let dateFilter = 'month'; // 'month' | '3m' | '6m' | 'all'
let symbolChart, trendChart, dayChart, winLossChart;
let selectedDay = null;
let currentView = 'resumen';
let profitMode = 'net'; // 'net' includes commission+swap (matches MT5's own Beneficio), 'gross' excludes them
function pv(t){ return profitMode === 'gross' ? t.grossProfit : t.profit; }
let sortKey = 'date', sortDir = -1;
let page = 0;
const PAGE_SIZE = 50;

const STORAGE_KEY = 'oro-trading-dashboard-v1';

/* ---------------- Utils ---------------- */
function norm(s){
  return String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function fmtMoney(n){
  const sign = n < 0 ? '-' : '';
  return sign + '$' + Math.abs(n).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2});
}
function fmtPct(n){ return (n>=0?'+':'') + n.toFixed(2) + '%'; }
function setStatus(msg, cls){
  const el = document.getElementById('status');
  el.textContent = msg; el.className = cls || '';
}
function parseNum(val){
  if (typeof val === 'number') return val;
  if (val == null) return NaN;
  let s = String(val).trim();
  if (s === '') return NaN;
  s = s.replace(/\s/g,'');
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/,/g,'');
  } else if (s.includes(',') && !s.includes('.')) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return n;
}
function parseMT5Date(val){
  if (val instanceof Date && !isNaN(val)) return val;
  if (typeof val === 'number') {
    try {
      const d = XLSX.SSF.parse_date_code(val);
      if (d) return new Date(d.y, d.m-1, d.d, d.H||0, d.M||0, d.S||0);
    } catch(e){}
  }
  if (typeof val === 'string') {
    const s = val.trim();
    const m = s.match(/(\d{4})[.\-\/](\d{2})[.\-\/](\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/);
    if (m) return new Date(+m[1], +m[2]-1, +m[3], +(m[4]||0), +(m[5]||0), +(m[6]||0));
    const d2 = new Date(s);
    if (!isNaN(d2)) return d2;
  }
  return null;
}
function toDateStr(d){
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}

/* ---------------- Parsing ---------------- */
function findHeaderRow(aoa){
  const symRe = /simbolo|symbol/;
  const profitRe = /beneficio|profit|ganancia neta|ganancia/;
  for (let i=0; i<Math.min(aoa.length, 500); i++){
    const row = aoa[i];
    if (!row) continue;
    const cells = row.map(norm);
    const hasSym = cells.some(c => symRe.test(c));
    const hasProfit = cells.some(c => profitRe.test(c));
    if (hasSym && hasProfit) return i;
  }
  return -1;
}

function guessColumns(headers){
  const n = headers.map(norm);
  const find = (re) => n.findIndex(h => re.test(h));
  const findLast = (re) => { for(let i=n.length-1;i>=0;i--) if(re.test(n[i])) return i; return -1; };

  const profitCol = findLast(/beneficio|profit|ganancia/);
  const symbolCol = find(/simbolo|symbol/);
  const volumeCol = find(/volumen|volume|lotes|lots?/);
  const typeCol = find(/^tipo$|^type$/);
  const commissionCol = find(/comision|commission/);
  const swapCol = find(/^swap$/);

  const dateCandidates = [];
  n.forEach((h,i) => { if (/fecha|hora|time|date/.test(h)) dateCandidates.push(i); });
  let dateCol = -1;
  if (dateCandidates.length){
    const before = dateCandidates.filter(i => profitCol === -1 || i < profitCol);
    dateCol = before.length ? before[before.length-1] : dateCandidates[dateCandidates.length-1];
  }
  return { dateCol, symbolCol, profitCol, volumeCol, typeCol, commissionCol, swapCol };
}

function isSectionMarker(row){
  if (!row || row[0] === undefined || row[0] === null) return false;
  const first = String(row[0]).trim();
  if (first === '') return false;
  for (let i=1; i<row.length; i++){
    const c = row[i];
    if (c !== undefined && c !== null && String(c).trim() !== '') return false;
  }
  return true;
}

function extractTrades(aoa, headerRowIndex, cols){
  const stopRe = /^(total|balance|deposito|dep[oó]sito|retiro|withdrawal|resultado|profit factor|resumen|summary)/;
  const trades = [];
  for (let i = headerRowIndex+1; i < aoa.length; i++){
    const row = aoa[i];
    if (!row || row.every(c => c === undefined || c === null || String(c).trim()==='')) {
      if (trades.length > 0) break; else continue;
    }
    if (isSectionMarker(row)) {
      if (trades.length > 0) break; else continue;
    }
    const firstCell = norm(row[0]);
    if (stopRe.test(firstCell)) break;

    const symbolRaw = cols.symbolCol > -1 ? row[cols.symbolCol] : null;
    const profitRaw = cols.profitCol > -1 ? row[cols.profitCol] : null;
    const dateRaw = cols.dateCol > -1 ? row[cols.dateCol] : null;
    const volumeRaw = cols.volumeCol > -1 ? row[cols.volumeCol] : null;
    const typeRaw = cols.typeCol > -1 ? row[cols.typeCol] : '';
    const commissionRaw = cols.commissionCol > -1 ? row[cols.commissionCol] : null;
    const swapRaw = cols.swapCol > -1 ? row[cols.swapCol] : null;

    const grossProfit = parseNum(profitRaw);
    if (!isFinite(grossProfit)) continue; // not a data row (header/section/blank)
    const symbol = symbolRaw ? String(symbolRaw).trim() : '';
    if (!symbol) continue; // deposits/credits and other non-trade rows have no symbol
    const date = parseMT5Date(dateRaw);
    if (!date) continue;
    const volume = parseNum(volumeRaw);
    const commission = isFinite(parseNum(commissionRaw)) ? parseNum(commissionRaw) : 0;
    const swap = isFinite(parseNum(swapRaw)) ? parseNum(swapRaw) : 0;
    const profit = grossProfit + commission + swap; // net result, matches MT5's own "Beneficio" total

    trades.push({
      date, dateStr: toDateStr(date), symbol,
      type: typeRaw ? String(typeRaw).trim() : '',
      volume: isFinite(volume) ? volume : 0,
      grossProfit, commission, swap, profit
    });
  }
  return trades;
}

function buildWorkbookAOA(file, cb){
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      let wb;
      if (file.name.toLowerCase().endsWith('.csv')) {
        wb = XLSX.read(e.target.result, { type: 'string', cellDates:true });
      } else {
        wb = XLSX.read(e.target.result, { type: 'array', cellDates:true });
      }
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header:1, raw:true, defval:'' });
      cb(null, aoa);
    } catch(err){ cb(err); }
  };
  reader.onerror = () => cb(new Error('No se pudo leer el archivo.'));
  if (file.name.toLowerCase().endsWith('.csv')) reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
}

/* ---------------- Aggregation ---------------- */
function aggregate(trades){
  const byDate = {};
  const bySymbol = {};
  let totalProfit = 0, wins = 0;
  trades.forEach(t => {
    const p = pv(t);
    totalProfit += p;
    if (p > 0) wins++;
    if (!byDate[t.dateStr]) byDate[t.dateStr] = { profit:0, count:0, wins:0 };
    byDate[t.dateStr].profit += p;
    byDate[t.dateStr].count += 1;
    if (p > 0) byDate[t.dateStr].wins += 1;

    if (!bySymbol[t.symbol]) bySymbol[t.symbol] = { count:0, volume:0, profit:0, wins:0, losses:0 };
    bySymbol[t.symbol].count += 1;
    bySymbol[t.symbol].volume += t.volume;
    bySymbol[t.symbol].profit += p;
    if (p > 0) bySymbol[t.symbol].wins += 1;
    else if (p < 0) bySymbol[t.symbol].losses += 1;
  });
  return { byDate, bySymbol, totalProfit, wins, total: trades.length };
}

/* ---------------- Chart data labels ---------------- */
function makeLabelPlugin(type, formatter){
  return {
    id: 'dl_' + type + '_' + Math.random().toString(36).slice(2),
    afterDatasetsDraw(chart){
      const { ctx } = chart;
      ctx.save();
      ctx.font = '600 11px ui-monospace, SFMono-Regular, Menlo, monospace';
      ctx.fillStyle = '#EDEDED';
      chart.data.datasets.forEach((dataset, dsIndex) => {
        const meta = chart.getDatasetMeta(dsIndex);
        if (meta.hidden) return;
        meta.data.forEach((el, i) => {
          const raw = dataset.data[i];
          const text = formatter(raw, i, dataset, chart);
          if (text === null || text === undefined || text === '') return;
          if (type === 'hbar'){
            const isNeg = raw < 0;
            const areaLeft = chart.chartArea.left, areaRight = chart.chartArea.right;
            const textWidth = ctx.measureText(text).width;
            ctx.textBaseline = 'middle';
            if (isNeg){
              const outsideX = el.x - 6;
              if (outsideX - textWidth < areaLeft){
                // not enough room to the left of the bar -> draw inside the bar's end, white text
                ctx.textAlign = 'left';
                ctx.fillText(text, el.x + 6, el.y);
              } else {
                ctx.textAlign = 'right';
                ctx.fillText(text, outsideX, el.y);
              }
            } else {
              const outsideX = el.x + 6;
              if (outsideX + textWidth > areaRight){
                ctx.textAlign = 'right';
                ctx.fillText(text, el.x - 6, el.y);
              } else {
                ctx.textAlign = 'left';
                ctx.fillText(text, outsideX, el.y);
              }
            }
          } else if (type === 'line'){
            ctx.textAlign = 'center';
            ctx.textBaseline = raw >= 0 ? 'bottom' : 'top';
            ctx.fillText(text, el.x, el.y + (raw >= 0 ? -8 : 8));
          } else if (type === 'doughnut'){
            const mid = (el.startAngle + el.endAngle) / 2;
            const r = (el.innerRadius + el.outerRadius) / 2;
            const lx = el.x + Math.cos(mid) * r;
            const ly = el.y + Math.sin(mid) * r;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            const prevFill = ctx.fillStyle, prevFont = ctx.font;
            ctx.fillStyle = '#0B0D10';
            ctx.font = '700 12px ui-monospace, SFMono-Regular, Menlo, monospace';
            const lines = String(text).split('\n');
            lines.forEach((line, li) => ctx.fillText(line, lx, ly + (li - (lines.length-1)/2) * 14));
            ctx.fillStyle = prevFill; ctx.font = prevFont;
          }
        });
      });
      ctx.restore();
    }
  };
}

/* ---------------- Rendering ---------------- */
function getFilteredTrades(){
  if (!TRADES.length || dateFilter === 'all') return TRADES;
  let monthsBack = 0;
  if (dateFilter === '3m') monthsBack = 2;
  else if (dateFilter === '6m') monthsBack = 5;
  const cutoff = new Date(calYear, calMonth-monthsBack, 1);
  const upper = new Date(calYear, calMonth+1, 1);
  return TRADES.filter(t => t.date >= cutoff && t.date < upper);
}

function renderAll(){
  document.getElementById('emptyState').style.display = TRADES.length ? 'none' : 'block';
  document.getElementById('dashboard').style.display = TRADES.length ? 'block' : 'none';
  if (!TRADES.length) return;

  FILTERED_TRADES = getFilteredTrades();

  const periodLabels = { month: MONTHS[calMonth]+' '+calYear, '3m':'últimos 3 meses (hasta '+MONTHS[calMonth]+')', '6m':'últimos 6 meses (hasta '+MONTHS[calMonth]+')', all:'histórico completo' };
  document.getElementById('resumenPeriodLabel').textContent = periodLabels[dateFilter];

  const agg = aggregate(FILTERED_TRADES);
  const winRate = agg.total ? (agg.wins/agg.total*100) : 0;
  const roi = capital > 0 ? (agg.totalProfit/capital*100) : 0;

  const profitEl = document.getElementById('kpiProfit');
  profitEl.textContent = fmtMoney(agg.totalProfit);
  profitEl.className = 'kpi-value ' + (agg.totalProfit>=0?'pos':'neg');
  document.getElementById('kpiProfitFoot').textContent = agg.total + ' operaciones · ' + (profitMode==='net' ? 'incluye comisión y swap' : 'sin comisión ni swap');

  document.getElementById('kpiWinRate').textContent = winRate.toFixed(1) + '%';
  document.getElementById('kpiWinRateFoot').textContent = agg.wins + ' ganadoras / ' + (agg.total-agg.wins) + ' perdedoras';

  let totalWon = 0, totalLost = 0;
  FILTERED_TRADES.forEach(t => { const p = pv(t); if (p>0) totalWon += p; else if (p<0) totalLost += Math.abs(p); });
  document.getElementById('kpiWon').textContent = fmtMoney(totalWon);
  document.getElementById('kpiWonFoot').textContent = agg.wins + ' operaciones ganadoras';
  document.getElementById('kpiLost').textContent = fmtMoney(-totalLost);
  document.getElementById('kpiLostFoot').textContent = (agg.total-agg.wins) + ' operaciones perdedoras';

  const roiEl = document.getElementById('kpiRoi');
  roiEl.textContent = fmtPct(roi);
  roiEl.className = 'kpi-value ' + (roi>=0?'pos':'neg');
  document.getElementById('kpiRoiFoot').textContent = 'sobre capital de ' + fmtMoney(capital);

  document.getElementById('kpiTrades').textContent = agg.total;
  const dates = Object.keys(agg.byDate);
  document.getElementById('kpiTradesFoot').textContent = dates.length ? (dates.length + ' días con actividad') : '';

  renderCalendar(aggregate(TRADES));
  renderSymbolWinLoss(agg);
  renderLotCard(FILTERED_TRADES);
  renderTrendChart(agg);
  renderTable();
  renderWinLossBar(totalWon, totalLost);
  renderForecast();
  if (selectedDay) renderDayDetail(selectedDay);
}

/* --- Win vs Loss pie --- */
function centerTextPlugin(bigText, smallText, color){
  return {
    id: 'centerText_' + Math.random().toString(36).slice(2),
    afterDraw(chart){
      const { ctx, chartArea } = chart;
      const cx = (chartArea.left + chartArea.right) / 2;
      const cy = (chartArea.top + chartArea.bottom) / 2;
      ctx.save();
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillStyle = color;
      ctx.font = '700 28px ui-monospace, SFMono-Regular, Menlo, monospace';
      ctx.fillText(bigText, cx, cy - 9);
      ctx.fillStyle = '#8A8F98';
      ctx.font = '600 10.5px -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif';
      ctx.fillText(smallText, cx, cy + 15);
      ctx.restore();
    }
  };
}

function renderWinLossBar(totalWon, totalLost){
  const canvas = document.getElementById('winLossBar');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  if (winLossChart) winLossChart.destroy();
  const net = totalWon - totalLost;
  const netColor = net>=0 ? '#C9A227' : '#E5484D';

  const labels = ['Ganado', 'Perdido', 'Neto'];
  const data = [totalWon, totalLost, Math.abs(net)];
  const colors = ['#3FB27F', '#E5484D', netColor];

  winLossChart = new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[
      { data, backgroundColor:colors, barThickness:70, borderRadius:8, borderSkipped:false }
    ]},
    options:{
      responsive:true, maintainAspectRatio:false, animation:{duration:250},
      layout:{ padding:{ top:10 } },
      plugins:{
        legend:{display:false},
        tooltip:{ callbacks:{ label:(c)=> c.dataIndex===1 ? fmtMoney(-totalLost) : (c.dataIndex===2 ? fmtMoney(net) : fmtMoney(c.raw)) } }
      },
      scales:{
        x:{ grid:{display:false}, ticks:{color:'#EDEDED', font:{size:13, weight:600}} },
        y:{ grid:{color:'#262B33'}, ticks:{color:'#8A8F98', font:{family:'monospace', size:11}} }
      }
    }
  });

  document.getElementById('winLossLegend').innerHTML = `
    <div class="pie-legend-item">
      <span class="pie-dot" style="background:#3FB27F"></span>
      <span class="pie-legend-label">Ganado</span>
      <span class="pie-legend-value pos">${fmtMoney(totalWon)}</span>
    </div>
    <div class="pie-legend-item">
      <span class="pie-dot" style="background:#E5484D"></span>
      <span class="pie-legend-label">Perdido</span>
      <span class="pie-legend-value neg">${fmtMoney(-totalLost)}</span>
    </div>
    <div class="pie-legend-item">
      <span class="pie-dot" style="background:${net>=0?'#C9A227':'#E5484D'}"></span>
      <span class="pie-legend-label">Neto</span>
      <span class="pie-legend-value ${net>=0?'pos':'neg'}">${fmtMoney(net)}</span>
    </div>
  `;
}

/* --- Forecast --- */
function renderForecast(){
  const card = document.getElementById('forecastCard');
  if (!TRADES.length){ card.style.display='none'; return; }

  const last = TRADES.reduce((a,b)=> a.date>b.date?a:b);
  const fy = last.date.getFullYear(), fm = last.date.getMonth();
  const prefix = fy + '-' + String(fm+1).padStart(2,'0');

  const byDay = {};
  TRADES.forEach(t => {
    if (!t.dateStr.startsWith(prefix)) return;
    const day = +t.dateStr.split('-')[2];
    byDay[day] = (byDay[day]||0) + pv(t);
  });
  const dayNumbers = Object.keys(byDay).map(Number);
  if (!dayNumbers.length){ card.style.display='none'; return; }

  const daysElapsed = Math.max(...dayNumbers);
  const monthProfitToDate = Object.values(byDay).reduce((a,b)=>a+b,0);
  const totalDaysInMonth = new Date(fy, fm+1, 0).getDate();
  const dailyAvg = monthProfitToDate / daysElapsed;
  const forecast = dailyAvg * totalDaysInMonth;

  card.style.display = 'flex';
  const valueEl = document.getElementById('forecastValue');
  valueEl.textContent = fmtMoney(forecast);
  valueEl.className = 'forecast-value ' + (forecast>=0?'pos':'neg');
  document.getElementById('forecastSub').textContent =
    `si mantienes tu ritmo actual en ${MONTHS[fm]} ${fy}` + (capital>0 ? ' · ' + fmtPct(forecast/capital*100) + ' sobre capital' : '');

  let note = `Cálculo: ${fmtMoney(monthProfitToDate)} acumulado ÷ ${daysElapsed} días transcurridos de ${MONTHS[fm]} = promedio de ${fmtMoney(dailyAvg)}/día. Ese promedio × ${totalDaysInMonth} días del mes = el pronóstico. Es una proyección lineal simple sobre tu historial, no una garantía — tu resultado real puede variar.`;
  if (daysElapsed < 5) note += ` Con solo ${daysElapsed} día(s) de datos este mes, tómalo como referencia aproximada, no como algo preciso.`;
  document.getElementById('forecastNote').textContent = note;

  const goalStatusEl = document.getElementById('goalStatus');
  if (goal > 0){
    const diffVsGoal = forecast - goal;
    if (diffVsGoal >= 0){
      goalStatusEl.textContent = `▲ Tu pronóstico supera tu meta por ${fmtMoney(diffVsGoal)}`;
      goalStatusEl.className = 'goal-status pos';
    } else {
      goalStatusEl.textContent = `▼ Tu pronóstico queda ${fmtMoney(Math.abs(diffVsGoal))} por debajo de tu meta`;
      goalStatusEl.className = 'goal-status neg';
    }
  } else {
    goalStatusEl.textContent = '';
  }

  const daysPct = Math.min(100, daysElapsed/totalDaysInMonth*100);
  const goalPct = goal > 0 ? Math.max(0, Math.min(100, monthProfitToDate/goal*100)) : 0;
  document.getElementById('progDaysPct').textContent = daysPct.toFixed(0) + '%';
  document.getElementById('progDaysBar').style.width = daysPct + '%';
  document.getElementById('progGoalPct').textContent = goalPct.toFixed(0) + '%';
  const goalBar = document.getElementById('progGoalBar');
  goalBar.style.width = goalPct + '%';
  goalBar.className = 'progress-fill ' + (monthProfitToDate>=0 ? 'progress-fill-pos' : 'progress-fill-neg');

  const paceEl = document.getElementById('progPace');
  if (goal > 0){
    const diff = goalPct - daysPct;
    if (Math.abs(diff) < 3) paceEl.textContent = '⏺ Vas justo al ritmo de tu pronóstico.';
    else if (diff > 0) paceEl.textContent = `▲ Vas ${diff.toFixed(0)} puntos adelante de tu ritmo esperado.`;
    else paceEl.textContent = `▼ Vas ${Math.abs(diff).toFixed(0)} puntos atrás de tu ritmo esperado.`;
  } else {
    paceEl.textContent = '';
  }
}

/* --- Calendar --- */
const DOW = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
const MONTHS = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

function renderCalendar(agg){
  document.getElementById('monthLabel').textContent = MONTHS[calMonth] + ' ' + calYear;
  const grid = document.getElementById('calGrid');
  grid.innerHTML = '';
  DOW.forEach(d => grid.insertAdjacentHTML('beforeend', `<div class="cal-dow">${d}</div>`));

  const first = new Date(calYear, calMonth, 1);
  const startOffset = first.getDay();
  const daysInMonth = new Date(calYear, calMonth+1, 0).getDate();

  for (let i=0;i<startOffset;i++) grid.insertAdjacentHTML('beforeend', '<div class="cal-cell empty-day"></div>');

  let monthProfit=0, winDays=0, lossDays=0, tradedDays=0;
  for (let day=1; day<=daysInMonth; day++){
    const ds = calYear + '-' + String(calMonth+1).padStart(2,'0') + '-' + String(day).padStart(2,'0');
    const rec = agg.byDate[ds];
    let cls = '';
    let inner = `<span class="cal-daynum">${day}</span>`;
    if (rec){
      tradedDays++;
      monthProfit += rec.profit;
      cls = rec.profit >= 0 ? 'pos' : 'neg';
      if (rec.profit > 0) winDays++; else if (rec.profit < 0) lossDays++;
      const dayWr = rec.count ? Math.round(rec.wins/rec.count*100) : 0;
      inner += `<span class="wick"></span><span class="cal-amt">${fmtMoney(rec.profit)}</span><span class="cal-wr">${rec.count} op · ${dayWr}% WR</span>`;
    } else {
      inner += `<span class="cal-wr">&nbsp;</span>`;
    }
    grid.insertAdjacentHTML('beforeend', `<div class="cal-cell ${cls}" data-date="${ds}">${inner}</div>`);
  }
  grid.querySelectorAll('.cal-cell[data-date]').forEach(cell => {
    cell.addEventListener('click', () => selectDay(cell.dataset.date));
  });
  document.getElementById('calMonthProfit').textContent = fmtMoney(monthProfit);
  document.getElementById('calMonthProfit').style.color = monthProfit>=0 ? 'var(--pos)' : 'var(--neg)';
  document.getElementById('calWinDays').textContent = winDays;
  document.getElementById('calLossDays').textContent = lossDays;
  document.getElementById('calMonthWinRate').textContent = tradedDays ? Math.round(winDays/tradedDays*100)+'%' : '0%';
}

/* --- Symbol bar chart --- */
function renderSymbolWinLoss(agg){
  const rows = Object.entries(agg.bySymbol).map(([symbol,v]) => ({symbol, ...v}));
  rows.sort((a,b) => b.count - a.count);

  document.getElementById('symbolHint').textContent = rows.length === 1
    ? `De tus ${rows[0].count} operaciones en ${rows[0].symbol}, cuántas cerraron ganadoras y cuántas perdedoras.`
    : 'De tus operaciones por símbolo, cuántas cerraron ganadoras y cuántas perdedoras (número de operaciones, no montos).';

  const labels = rows.map(r=>r.symbol);
  const winsData = rows.map(r=>r.wins);
  const lossData = rows.map(r=>r.losses);

  const ctx = document.getElementById('symbolChart').getContext('2d');
  if (symbolChart) symbolChart.destroy();
  document.getElementById('symbolChartWrap').style.height = Math.max(90, rows.length*50 + 60) + 'px';
  symbolChart = new Chart(ctx, {
    type:'bar',
    data:{ labels, datasets:[
      { label:'Ganadas', data:winsData, backgroundColor:'#3FB27F', stack:'s',
        borderRadius:{ topLeft:8, bottomLeft:8, topRight:0, bottomRight:0 }, borderSkipped:false },
      { label:'Perdidas', data:lossData, backgroundColor:'#E5484D', stack:'s',
        borderRadius:{ topRight:8, bottomRight:8, topLeft:0, bottomLeft:0 }, borderSkipped:false }
    ]},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false, animation:{duration:250},
      plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label:(c)=> `${c.dataset.label}: ${c.raw}` } } },
      scales:{
        x:{ stacked:true, grid:{color:'#262B33'}, ticks:{color:'#8A8F98', font:{family:'monospace', size:11}, precision:0} },
        y:{ stacked:true, grid:{display:false}, ticks:{color:'#EDEDED', font:{size:12}} }
      }
    }
  });

  document.getElementById('symbolLegend').innerHTML = rows.map(r => `
    <div class="pie-legend-item">
      <span class="pie-dot" style="background:#3FB27F"></span>
      <span class="pie-legend-label">${r.symbol}</span>
      <span class="pie-legend-value pos">${r.wins} ganadas</span>
      <span class="pie-legend-value neg" style="margin-left:10px;">${r.losses} perdidas</span>
    </div>`).join('');
}

function renderLotCard(trades){
  if (!trades.length){
    document.getElementById('lotAvgValue').textContent = '0.00';
    document.getElementById('lotBreakdown').innerHTML = '';
    return;
  }
  const totalVolume = trades.reduce((sum,t)=>sum+t.volume, 0);
  const avg = totalVolume / trades.length;
  document.getElementById('lotAvgValue').textContent = avg.toFixed(2);

  const bySymbol = {};
  trades.forEach(t => {
    if (!bySymbol[t.symbol]) bySymbol[t.symbol] = { volume:0, count:0 };
    bySymbol[t.symbol].volume += t.volume;
    bySymbol[t.symbol].count += 1;
  });
  const rows = Object.entries(bySymbol);
  document.getElementById('lotBreakdown').innerHTML = rows.length > 1
    ? rows.map(([symbol,v]) => `<div class="lot-breakdown-item"><span>${symbol}</span><b>${(v.volume/v.count).toFixed(2)} lotes</b></div>`).join('')
    : `<div class="lot-breakdown-item"><span>Volumen total operado</span><b>${totalVolume.toFixed(2)} lotes</b></div>`;
}

/* --- Trend line --- */
function renderTrendChart(agg){
  const dates = Object.keys(agg.byDate).sort();
  const labels = dates;
  const data = dates.map(d => capital>0 ? (agg.byDate[d].profit/capital*100) : 0);
  const colors = data.map(v => v>=0 ? '#3FB27F' : '#E5484D');

  const ctx = document.getElementById('trendChart').getContext('2d');
  if (trendChart) trendChart.destroy();
  trendChart = new Chart(ctx, {
    type:'line',
    data:{ labels, datasets:[{
      label:'Rentabilidad diaria (%)', data,
      borderColor:'#C9A227', backgroundColor:'rgba(201,162,39,.12)', fill:true, tension:.25,
      pointRadius:3, pointBackgroundColor: colors, pointHoverRadius:5, borderWidth:2
    }]},
    options:{
      responsive:true, maintainAspectRatio:false, animation:{duration:250},
      layout:{ padding:{ top:22, bottom:14 } },
      plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label:(c)=> fmtPct(c.raw) } } },
      scales:{
        x:{ grid:{display:false}, ticks:{color:'#8A8F98', font:{family:'monospace', size:10}, maxRotation:60, minRotation:0} },
        y:{ grid:{color:'#262B33'}, ticks:{color:'#8A8F98', callback:(v)=>v+'%'} }
      }
    },
    plugins:[ makeLabelPlugin('line', (raw)=> fmtPct(raw)) ]
  });
}

/* --- Table --- */
function renderTable(){
  const sorted = [...FILTERED_TRADES].sort((a,b) => {
    let av=a[sortKey], bv=b[sortKey];
    if (sortKey==='date'){ av=a.date.getTime(); bv=b.date.getTime(); }
    if (sortKey==='profit'){ av=pv(a); bv=pv(b); }
    if (typeof av === 'string') return av.localeCompare(bv) * sortDir;
    return (av-bv) * sortDir;
  });
  const start = page*PAGE_SIZE;
  const pageRows = sorted.slice(start, start+PAGE_SIZE);
  const body = document.getElementById('tradesBody');
  body.innerHTML = pageRows.map(t => `
    <tr>
      <td>${t.date.toLocaleString('es-GT',{year:'numeric',month:'2-digit',day:'2-digit',hour:'2-digit',minute:'2-digit'})}</td>
      <td>${t.symbol}</td>
      <td>${t.type}</td>
      <td>${t.volume.toFixed(2)}</td>
      <td class="${pv(t)>=0?'pos':'neg'}">${fmtMoney(pv(t))}</td>
    </tr>`).join('');
  document.getElementById('pagerInfo').textContent =
    `Mostrando ${sorted.length? start+1:0}–${Math.min(start+PAGE_SIZE, sorted.length)} de ${sorted.length}`;
}

/* --- View switching --- */
function switchView(view){
  currentView = view;
  document.querySelectorAll('#viewTabs button').forEach(b => b.classList.toggle('active', b.dataset.view===view));
  document.getElementById('view-resumen').style.display = view==='resumen' ? '' : 'none';
  document.getElementById('view-detalles').style.display = view==='detalles' ? '' : 'none';
  document.getElementById('view-dia').style.display = view==='dia' ? '' : 'none';
  // Chart.js needs a nudge after its canvas becomes visible again
  requestAnimationFrame(() => {
    if (view==='resumen'){ symbolChart && symbolChart.resize(); trendChart && trendChart.resize(); }
    if (view==='dia'){ dayChart && dayChart.resize(); }
  });
}

function selectDay(dateStr){
  selectedDay = dateStr;
  switchView('dia');
  renderDayDetail(dateStr);
}

/* --- Day detail --- */
const DOW_LONG = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
function renderDayDetail(dateStr){
  const dayTrades = TRADES.filter(t => t.dateStr === dateStr).sort((a,b)=>a.date-b.date);
  document.getElementById('dayEmptyState').style.display = dayTrades.length ? 'none' : 'block';
  document.getElementById('dayContent').style.display = dayTrades.length ? 'block' : 'none';
  if (!dayTrades.length) return;

  const [y,m,d] = dateStr.split('-').map(Number);
  const dObj = new Date(y, m-1, d);
  document.getElementById('dayLabel').textContent = `${DOW_LONG[dObj.getDay()]} ${d} de ${MONTHS[m-1]} ${y}`;

  let winsCount=0, lossesCount=0, winsSum=0, lossesSum=0;
  dayTrades.forEach(t => {
    const p = pv(t);
    if (p > 0){ winsCount++; winsSum += p; }
    else if (p < 0){ lossesCount++; lossesSum += Math.abs(p); }
  });
  const net = winsSum - lossesSum;
  const wr = (winsCount+lossesCount) ? (winsCount/(winsCount+lossesCount)*100) : 0;

  document.getElementById('dayTotal').textContent = winsCount + lossesCount;
  document.getElementById('dayWins').textContent = winsCount;
  document.getElementById('dayLosses').textContent = lossesCount;
  document.getElementById('dayWinRate').textContent = wr.toFixed(1) + '%';
  const netEl = document.getElementById('dayNet');
  netEl.textContent = fmtMoney(net);
  netEl.className = 'kpi-value ' + (net>=0?'pos':'neg');
  document.getElementById('dayNetFoot').textContent =
    `${fmtMoney(winsSum)} ganado − ${fmtMoney(lossesSum)} perdido · ${capital>0 ? fmtPct(net/capital*100)+' sobre capital' : ''}`;

  const ctx = document.getElementById('dayChart').getContext('2d');
  if (dayChart) dayChart.destroy();
  const barLabels = ['Ganado','Perdido'];
  const barData = [winsSum, lossesSum];
  const barColors = ['#3FB27F','#E5484D'];
  const dayLabelFmt = (raw) => fmtMoney(raw);
  dayChart = new Chart(ctx, {
    type:'bar',
    data:{ labels:barLabels, datasets:[{ data:barData, backgroundColor:barColors, borderRadius:6, barThickness:34 }]},
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false, animation:{duration:250},
      layout:{ padding:{ right:60 } },
      plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label:(c)=>fmtMoney(c.raw) } } },
      scales:{
        x:{ grid:{color:'#262B33'}, ticks:{color:'#8A8F98', font:{family:'monospace', size:11}} },
        y:{ grid:{display:false}, ticks:{color:'#EDEDED', font:{size:13, weight:600}} }
      }
    },
    plugins:[ makeLabelPlugin('hbar', dayLabelFmt) ]
  });

  const body = document.getElementById('dayTradesBody');
  body.innerHTML = dayTrades.map(t => `
    <tr>
      <td>${t.date.toLocaleTimeString('es-GT',{hour:'2-digit',minute:'2-digit',second:'2-digit'})}</td>
      <td>${t.symbol}</td>
      <td>${t.type}</td>
      <td>${t.volume.toFixed(2)}</td>
      <td class="${pv(t)>=0?'pos':'neg'}">${fmtMoney(pv(t))}</td>
    </tr>`).join('');
}

/* ---------------- Storage ---------------- */
async function saveToStorage(){
  try {
    await window.storage.set(STORAGE_KEY, JSON.stringify({
      trades: TRADES.map(t=>({...t, date:t.date.toISOString()})),
      capital, goal
    }), false);
  } catch(e){ console.error('storage set failed', e); }
}
async function loadFromStorage(){
  try {
    const res = await window.storage.get(STORAGE_KEY, false);
    if (res && res.value){
      const parsed = JSON.parse(res.value);
      TRADES = parsed.trades.map(t => ({...t, date:new Date(t.date)}));
      capital = parsed.capital || capital;
      goal = (parsed.goal !== undefined && parsed.goal !== null) ? parsed.goal : goal;
      document.getElementById('capitalInput').value = capital;
      document.getElementById('goalInput').value = goal;
      const last = TRADES.length ? TRADES.reduce((a,b)=> a.date>b.date?a:b) : null;
      calYear = last ? last.date.getFullYear() : new Date().getFullYear();
      calMonth = last ? last.date.getMonth() : new Date().getMonth();
      renderAll();
      setStatus('Datos de tu última carga restaurados.', 'ok');
    }
  } catch(e){ /* no data saved yet */ }
}

/* ---------------- Events ---------------- */
document.querySelectorAll('#viewTabs button').forEach(btn => {
  btn.addEventListener('click', () => switchView(btn.dataset.view));
});

document.querySelectorAll('#dateFilterTabs button').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('#dateFilterTabs button').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    dateFilter = btn.dataset.range;
    page = 0;
    if (TRADES.length) renderAll();
  });
});

document.getElementById('fileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  setStatus('Leyendo ' + file.name + '…');
  buildWorkbookAOA(file, (err, aoa) => {
    if (err){ setStatus('Error al leer el archivo: ' + err.message, 'err'); return; }
    const headerRowIndex = findHeaderRow(aoa);
    if (headerRowIndex === -1){
      setStatus('No se encontraron columnas de Símbolo/Beneficio en este archivo. Verifica que sea un informe de historial exportado desde MT5.', 'err');
      return;
    }
    const headers = aoa[headerRowIndex];
    const cols = guessColumns(headers);

    const trades = extractTrades(aoa, headerRowIndex, cols);
    if (!trades.length){
      setStatus('Se encontraron columnas pero ninguna operación con resultado distinto de cero.', 'err');
      return;
    }
    TRADES = trades;
    const last = TRADES.reduce((a,b)=> a.date>b.date?a:b);
    calYear = last.date.getFullYear(); calMonth = last.date.getMonth();
    renderAll();
    saveToStorage();
    setStatus(`${trades.length} operaciones cargadas correctamente.`, 'ok');
  });
});

document.getElementById('capitalInput').addEventListener('input', (e) => {
  const v = parseFloat(e.target.value);
  capital = isFinite(v) ? v : 0;
  if (TRADES.length) renderAll();
  saveToStorage();
});

document.getElementById('goalInput').addEventListener('input', (e) => {
  const v = parseFloat(e.target.value);
  goal = isFinite(v) ? v : 0;
  if (TRADES.length) renderAll();
  saveToStorage();
});

document.getElementById('clearBtn').addEventListener('click', async () => {
  try { await window.storage.delete(STORAGE_KEY, false); } catch(e){}
  TRADES = [];
  renderAll();
  setStatus('Datos borrados.', '');
});

document.getElementById('prevMonth').addEventListener('click', () => {
  calMonth--; if (calMonth<0){calMonth=11; calYear--;}
  renderAll();
});
document.getElementById('nextMonth').addEventListener('click', () => {
  calMonth++; if (calMonth>11){calMonth=0; calYear++;}
  renderAll();
});
document.getElementById('todayMonth').addEventListener('click', () => {
  const n = new Date(); calYear = n.getFullYear(); calMonth = n.getMonth();
  renderAll();
});

document.querySelectorAll('table.trades thead th[data-key]').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.key;
    if (sortKey===key) sortDir *= -1; else { sortKey=key; sortDir=-1; }
    page=0; renderTable();
  });
});
document.getElementById('pagerPrev').addEventListener('click', () => { if(page>0){page--; renderTable();} });
document.getElementById('pagerNext').addEventListener('click', () => {
  if ((page+1)*PAGE_SIZE < FILTERED_TRADES.length){ page++; renderTable(); }
});

/* ---------------- Init ---------------- */
(function init(){
  const now = new Date();
  calYear = now.getFullYear(); calMonth = now.getMonth();
  document.getElementById('capitalInput').value = capital;
  document.getElementById('goalInput').value = goal;
  loadFromStorage();
})();
