import { IOpsDataset, IOpsRow, Direction } from './OpsDataService';
import { IOpsNote, noteKey } from './OpsNotesService';

// Chart.js is attached to window by the web part (bundled). Reference as global.
/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Chart: any;

export const OPS_DASHBOARD_VERSION: string = '1.0.0 · 2026-07-23';

export interface IOpsDashboardOptions {
  fetchData: () => Promise<IOpsDataset>;
  fetchNotes: () => Promise<{ [key: string]: IOpsNote }>;
  saveNote: (branch: string, metric: string, month: string, note: string, existingId: number | null) => Promise<number | null>;
}

export interface IOpsController { destroy: () => void; }

const FY_MONTH_ORDER: string[] = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
const QUARTERS: { [q: string]: string[] } = { Q1: ['Apr', 'May', 'Jun'], Q2: ['Jul', 'Aug', 'Sep'], Q3: ['Oct', 'Nov', 'Dec'], Q4: ['Jan', 'Feb', 'Mar'] };

export function initOpsDashboard(root: HTMLElement, opts: IOpsDashboardOptions): IOpsController {
  const el = (name: string): HTMLElement => root.querySelector(`[data-el="${name}"]`) as HTMLElement;

  let dataset: IOpsDataset = { rows: [], branches: [], metrics: [], units: {}, directions: {}, monthsPresent: [] };
  let notesMap: { [key: string]: IOpsNote } = {};
  const notesLocal: { [key: string]: string } = {};
  const notesDirty: { [key: string]: boolean } = {};
  const charts: { [id: string]: any } = {};

  let selectedMonth: string | null = null;
  let selectedQuarter: string | null = null;
  let selectedBranchMonth: string | null = null;

  const verEl = root.querySelector('[data-el="appVersion"]');
  if (verEl) { verEl.textContent = 'v' + OPS_DASHBOARD_VERSION; }

  // ---- helpers -------------------------------------------------------------
  function escapeHtml(s: unknown): string {
    return ('' + s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
  }
  function unitOf(metric: string): string { return dataset.units[metric] || ''; }
  function dirOf(metric: string): Direction { return dataset.directions[metric] || 'higher'; }

  function fmtV(v: number, metric: string): string {
    const u: string = unitOf(metric);
    if (u === '$') {
      const abs: number = Math.abs(v), sign: string = v < 0 ? '-' : '';
      if (abs >= 1e6) { return sign + '$' + (abs / 1e6).toFixed(1) + 'M'; }
      if (abs >= 1000) { return sign + '$' + Math.round(abs / 1000) + 'K'; }
      return sign + '$' + Math.round(abs);
    }
    if (u === 'rate') { return v.toFixed(3); }
    if (u === '%') { return v.toFixed(1) + '%'; }
    return Math.round(v).toLocaleString();
  }

  function completedMonths(): string[] {
    return FY_MONTH_ORDER.filter((m) =>
      dataset.rows.some((r) => r.month === m && r.metric === 'Revenue' && r.actual > 0));
  }
  function quarterForMonth(m: string): string | null {
    const found = Object.keys(QUARTERS).filter((q) => QUARTERS[q].indexOf(m) >= 0);
    return found.length ? found[0] : null;
  }
  function currentQuarter(comp: string[]): string {
    if (!comp.length) { return 'Q1'; }
    return quarterForMonth(comp[comp.length - 1]) || 'Q1';
  }

  function getRows(branch: string | null, metric: string, months: string[] | null): IOpsRow[] {
    return dataset.rows.filter((r) => {
      if (r.metric !== metric) { return false; }
      if (branch && branch !== 'ALL' && r.branch !== branch) { return false; }
      if (months && months.indexOf(r.month) < 0) { return false; }
      return true;
    });
  }
  function sumRows(rows: IOpsRow[]): { actual: number; target: number } {
    return { actual: rows.reduce((s, r) => s + r.actual, 0), target: rows.reduce((s, r) => s + r.target, 0) };
  }
  function isPass(actual: number, target: number, metric: string): boolean {
    if (dirOf(metric) === 'higher') { return actual >= target; }
    return (actual === 0 && target === 0) || actual <= target;
  }
  function monthActual(branch: string, metric: string, month: string): number {
    return sumRows(getRows(branch === 'ALL' ? null : branch, metric, [month])).actual;
  }
  function momChange(metric: string, branch: string, fromMonth: string, toMonth: string): number | null {
    const lSum: number = monthActual(branch, metric, toMonth);
    const pSum: number = monthActual(branch, metric, fromMonth);
    if (pSum === 0) { return null; }
    return ((lSum - pSum) / Math.abs(pSum)) * 100;
  }

  function bFilter(): string { return (el('branchSel') as HTMLSelectElement).value; }
  function mFilter(): string { return (el('metricSel') as HTMLSelectElement).value; }
  function legendHtml(): string {
    return `<div class="legend"><span><span class="leg-sq" style="background:#185FA5"></span>Actual</span><span><span class="leg-sq" style="background:rgba(24,95,165,0.2);border:1px solid #185FA5"></span>Budget</span></div>`;
  }

  // ---- charts --------------------------------------------------------------
  function destroyChart(id: string): void { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }
  function destroyAllCharts(): void { Object.keys(charts).forEach(destroyChart); }

  function buildHorizBar(id: string, data: { b: string; actual: number; target: number }[]): void {
    if (typeof Chart === 'undefined') { return; }
    destroyChart(id);
    const canvas = root.querySelector('#' + id) as HTMLCanvasElement;
    if (!canvas) { return; }
    charts[id] = new Chart(canvas, {
      type: 'bar',
      data: {
        labels: data.map((d) => d.b), datasets: [
          { label: 'Actual', data: data.map((d) => d.actual), backgroundColor: '#185FA5', borderRadius: 3, barPercentage: 0.55 },
          { label: 'Budget', data: data.map((d) => d.target), backgroundColor: 'rgba(24,95,165,0.13)', borderColor: '#185FA5', borderWidth: 1, borderRadius: 3, barPercentage: 0.55 }
        ]
      },
      options: {
        indexAxis: 'y', responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx: any) => ` ${ctx.dataset.label}: ${fmtV(ctx.raw, 'Revenue')}` } } },
        scales: { x: { grid: { color: 'rgba(0,0,0,.05)' }, ticks: { callback: (v: any) => fmtV(v, 'Revenue'), font: { size: 11 } } }, y: { grid: { display: false }, ticks: { font: { size: 12 } } } }
      }
    });
  }
  function buildGroupedBar(id: string, labels: any[], datasets: any[]): void {
    if (typeof Chart === 'undefined') { return; }
    destroyChart(id);
    const canvas = root.querySelector('#' + id) as HTMLCanvasElement;
    if (!canvas) { return; }
    charts[id] = new Chart(canvas, {
      type: 'bar', data: { labels, datasets },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { callbacks: { label: (ctx: any) => ` ${ctx.dataset.label}: ${fmtV(ctx.raw, 'Revenue')}` } } },
        scales: { x: { grid: { display: false }, ticks: { font: { size: 12 } } }, y: { grid: { color: 'rgba(0,0,0,.05)' }, ticks: { callback: (v: any) => fmtV(v, 'Revenue'), font: { size: 11 } } } }
      }
    });
  }

  // ---- shared blocks -------------------------------------------------------
  function kpiBlock(months: string[]): HTMLElement {
    const branch: string = bFilter();
    const { actual: revA, target: revT } = sumRows(getRows(branch === 'ALL' ? null : branch, 'Revenue', months));
    const variance: number = revA - revT;
    const pct: number = revT > 0 ? ((revA / revT - 1) * 100) : 0;
    const g = document.createElement('div'); g.className = 'kpi-grid';
    [{ label: 'Revenue actual', val: fmtV(revA, 'Revenue'), sub: `${months.length} month${months.length !== 1 ? 's' : ''}`, cls: '' },
     { label: 'Revenue budget', val: fmtV(revT, 'Revenue'), sub: 'prorated to period', cls: '' },
     { label: 'Variance', val: (variance >= 0 ? '+' : '') + fmtV(variance, 'Revenue'), sub: 'actual minus budget', cls: variance >= 0 ? 'pos' : 'neg' },
     { label: 'vs budget', val: (pct >= 0 ? '+' : '') + pct.toFixed(1) + '%', sub: variance >= 0 ? 'ahead of pace' : 'behind pace', cls: pct >= 0 ? 'pos' : 'neg' }
    ].forEach((k) => {
      const d = document.createElement('div'); d.className = 'kpi-card';
      d.innerHTML = `<div class="kpi-label">${k.label}</div><div class="kpi-value ${k.cls}">${k.val}</div><div class="kpi-sub">${k.sub}</div>`;
      g.appendChild(d);
    });
    return g;
  }

  function horizBarPanel(canvasId: string, months: string[], titleTxt: string): HTMLElement {
    const branch: string = bFilter();
    const data = dataset.branches.filter((b) => branch === 'ALL' ? true : b === branch)
      .map((b) => { const r = getRows(b, 'Revenue', months); return { b, ...sumRows(r) }; });
    const h: number = Math.max(180, data.length * 44 + 60);
    const panel = document.createElement('div'); panel.className = 'panel';
    panel.innerHTML = `<div class="panel-title">${escapeHtml(titleTxt || 'Revenue vs budget by branch')}</div>` + legendHtml() +
      `<div class="chart-wrap" style="height:${h}px"><canvas id="${canvasId}"></canvas></div>`;
    setTimeout(() => buildHorizBar(canvasId, data), 0);
    return panel;
  }

  function perfMetricPanel(months: string[]): HTMLElement {
    const branch: string = bFilter();
    const comp: string[] = completedMonths();
    const availMonthsInSet: string[] = FY_MONTH_ORDER.filter((m) => months.indexOf(m) >= 0 && comp.indexOf(m) >= 0);
    const prevM: string | null = availMonthsInSet.length >= 2 ? availMonthsInSet[availMonthsInSet.length - 2] : null;
    const lastM: string | null = availMonthsInSet.length >= 1 ? availMonthsInSet[availMonthsInSet.length - 1] : null;
    const momLabel: string = prevM && lastM ? `${prevM} → ${lastM}` : 'vs prior month';

    const stats = dataset.metrics.map((m) => {
      const rows = getRows(branch === 'ALL' ? null : branch, m, months);
      const actualSum = rows.reduce((s, r) => s + r.actual, 0);
      const budgetSum = rows.reduce((s, r) => s + r.target, 0);
      const pct = budgetSum !== 0 ? Math.round(actualSum / budgetSum * 100) : (actualSum === 0 ? 100 : 0);
      const mom = (prevM && lastM) ? momChange(m, branch, prevM, lastM) : null;
      return { m, pct, dir: dirOf(m), mom };
    });

    let html = `<table><thead><tr><th>Metric</th><th>Direction</th><th>${escapeHtml(momLabel)}</th><th>% on Budget</th></tr></thead><tbody>`;
    stats.forEach((s) => {
      const isGoodPct = s.dir === 'lower' ? s.pct <= 100 : s.pct >= 100;
      const barColor = isGoodPct ? '#1D9E75' : '#E24B4A';
      const barWidth = Math.max(0, Math.min(100, s.pct));
      const bar = `<div class="pbar-row"><div class="pbar-track"><div class="pbar-fill" style="width:${barWidth}%;background:${barColor}"></div></div><span class="pbar-label">${s.pct}%</span></div>`;
      let momCell = '<span class="mom-pill flat">—</span>';
      if (s.mom !== null) {
        const arrow = s.mom > 0 ? '↑' : '↓';
        const isGood = s.dir === 'lower' ? s.mom <= 0 : s.mom >= 0;
        const cls = s.mom === 0 ? 'flat' : (isGood ? 'up' : 'down');
        momCell = `<span class="mom-pill ${cls}">${arrow} ${s.mom > 0 ? '+' : ''}${s.mom.toFixed(1)}%</span>`;
      }
      html += `<tr><td class="fw">${escapeHtml(s.m)}</td><td class="muted">${s.dir === 'higher' ? '↑ higher' : '↓ lower'}</td><td>${momCell}</td><td style="min-width:200px">${bar}</td></tr>`;
    });
    html += `</tbody></table>`;
    const panel = document.createElement('div'); panel.className = 'panel';
    panel.innerHTML = `<div class="panel-title">Performance metrics</div><div class="tbl-wrap">${html}</div>`;
    return panel;
  }

  function titleRow(title: string, sub: string, badgeTxt: string, rightExtras?: HTMLElement): HTMLElement {
    const div = document.createElement('div'); div.className = 'page-title-row';
    const left = document.createElement('div');
    left.innerHTML = `<div class="page-title">${escapeHtml(title)}</div><div class="page-sub">${escapeHtml(sub)}</div>`;
    const right = document.createElement('div'); right.className = 'page-title-right';
    right.innerHTML = `<span class="period-badge">${escapeHtml(badgeTxt)}</span>`;
    if (rightExtras) { right.appendChild(rightExtras); }
    div.appendChild(left); div.appendChild(right);
    return div;
  }

  function branchMetricDetailPanel(months: string[], metric: string, titleTxt: string): HTMLElement {
    const branch: string = bFilter();
    let displayBranches: string[] = branch === 'ALL' ? dataset.branches : [branch];
    const variances: { [b: string]: number } = {};
    displayBranches.forEach((b) => {
      let tA = 0, tT = 0;
      months.forEach((mo) => { getRows(b, metric, [mo]).forEach((r) => { tA += r.actual; tT += r.target; }); });
      variances[b] = tA - tT;
    });
    displayBranches = [...displayBranches].sort((a, b) => variances[b] - variances[a]);

    const tblPanel = document.createElement('div'); tblPanel.className = 'panel';
    tblPanel.innerHTML = `<div class="panel-title">${escapeHtml(titleTxt)}</div>`;
    const tblWrap = document.createElement('div'); tblWrap.className = 'tbl-wrap';
    let html = `<table><thead><tr><th>Branch</th>`;
    months.forEach((m) => { html += `<th>${escapeHtml(m)} actual</th><th>${escapeHtml(m)} vs budget</th>`; });
    html += `<th>Total actual</th><th>Total budget</th><th>Variance</th></tr></thead><tbody>`;
    displayBranches.forEach((b) => {
      html += `<tr><td class="fw">${escapeHtml(b)}</td>`;
      let tActual = 0, tTarget = 0;
      months.forEach((mo) => {
        const rows = getRows(b, metric, [mo]);
        if (!rows.length) { html += `<td class="muted">—</td><td class="muted">—</td>`; return; }
        const { actual, target } = sumRows(rows);
        tActual += actual; tTarget += target;
        const pass = isPass(actual, target, metric);
        const pct = target > 0 ? ((actual / target - 1) * 100) : 0;
        html += `<td>${fmtV(actual, metric)}</td><td class="${pass ? 'pos' : 'neg'}">${(pct >= 0 ? '+' : '') + pct.toFixed(1)}%</td>`;
      });
      const tVar = tActual - tTarget; const tPct = tTarget > 0 ? ((tActual / tTarget - 1) * 100) : 0;
      const tPass = isPass(tActual, tTarget, metric);
      html += `<td class="fw">${fmtV(tActual, metric)}</td><td class="muted">${fmtV(tTarget, metric)}</td><td class="${tPass ? 'pos' : 'neg'} fw">${tVar >= 0 ? '+' : ''}${fmtV(tVar, metric)} (${(tPct >= 0 ? '+' : '') + tPct.toFixed(1)}%)</td></tr>`;
    });
    html += `</tbody></table>`;
    tblWrap.innerHTML = html; tblPanel.appendChild(tblWrap);
    return tblPanel;
  }

  // ---- views ---------------------------------------------------------------
  function renderOverview(): void {
    const comp: string[] = completedMonths();
    const c = el('view-overview'); c.innerHTML = '';
    c.appendChild(titleRow('Overview', `Full FY27 · all ${comp.length} completed months`, `FY27 · ${comp.length} months complete`));
    if (!comp.length) { c.innerHTML += '<div class="panel"><p class="muted">No completed months yet (no Revenue actuals &gt; 0).</p></div>'; return; }
    c.appendChild(kpiBlock(comp));
    c.appendChild(horizBarPanel('chart-ov-bar', comp, `Revenue vs budget · ${comp.join(', ')}`));
    c.appendChild(perfMetricPanel(comp));
    dataset.metrics.forEach((m) => {
      c.appendChild(branchMetricDetailPanel(comp, m, `FY27 detail by branch — ${m} (${comp.join(', ')})`));
    });
  }

  function renderMonthly(skipSelReset?: boolean): void {
    const comp: string[] = completedMonths();
    if (!skipSelReset) { selectedMonth = comp.length ? comp[comp.length - 1] : null; }
    const c = el('view-monthly'); c.innerHTML = '';
    if (!comp.length) { c.innerHTML = '<div class="panel"><p class="muted">No completed months yet.</p></div>'; return; }
    if (!selectedMonth || comp.indexOf(selectedMonth) < 0) { selectedMonth = comp[comp.length - 1]; }

    const sel = document.createElement('select');
    sel.className = 'period-select';
    comp.forEach((m) => { const o = document.createElement('option'); o.value = m; o.text = m; if (m === selectedMonth) { o.selected = true; } sel.appendChild(o); });
    sel.addEventListener('change', () => { selectedMonth = sel.value; renderMonthly(true); });
    const selWrap = document.createElement('div'); selWrap.className = 'period-selector';
    selWrap.innerHTML = '<label>Viewing month:</label>'; selWrap.appendChild(sel);

    const mo: string = selectedMonth as string;
    const prevIdx: number = comp.indexOf(mo) - 1;
    const prevM: string | null = prevIdx >= 0 ? comp[prevIdx] : null;

    c.appendChild(titleRow('Monthly', 'Single month snapshot', `${mo} FY27`, selWrap));
    c.appendChild(kpiBlock([mo]));
    c.appendChild(horizBarPanel('chart-mo-bar', [mo], `Revenue vs budget · ${mo}`));
    c.appendChild(perfMetricPanel([mo]));

    const branch: string = bFilter(); const metric: string = mFilter();
    const displayBranches: string[] = branch === 'ALL' ? dataset.branches : [branch];
    const displayMetrics: string[] = metric === 'ALL' ? dataset.metrics : [metric];

    const panel = document.createElement('div'); panel.className = 'panel';
    panel.innerHTML = `<div class="panel-title">${escapeHtml(mo)} detail — all metrics by branch</div>`;
    const tblWrap = document.createElement('div'); tblWrap.className = 'tbl-wrap';
    let html = `<table><thead><tr><th>Branch</th><th>Metric</th><th>Actual</th><th>Budget</th><th>Variance</th><th>vs budget</th>`;
    if (prevM) { html += `<th>vs ${escapeHtml(prevM)}</th>`; }
    html += `</tr></thead><tbody>`;
    displayBranches.forEach((b) => {
      displayMetrics.forEach((dm) => {
        const rows = getRows(b, dm, [mo]);
        if (!rows.length) { return; }
        const { actual, target } = sumRows(rows);
        const variance = actual - target;
        const pass = isPass(actual, target, dm);
        const pct = target > 0 ? ((actual / target - 1) * 100) : 0;
        html += `<tr><td class="fw">${escapeHtml(b)}</td><td class="muted">${escapeHtml(dm)}</td><td>${fmtV(actual, dm)}</td><td class="muted">${fmtV(target, dm)}</td><td class="${variance >= 0 ? 'pos' : 'neg'}">${variance >= 0 ? '+' : ''}${fmtV(variance, dm)}</td><td class="${pass ? 'pos' : 'neg'} fw">${(pct >= 0 ? '+' : '') + pct.toFixed(1)}%</td>`;
        if (prevM) {
          const prevActual = monthActual(b, dm, prevM);
          const mom = prevActual !== 0 ? ((actual - prevActual) / Math.abs(prevActual) * 100) : null;
          html += `<td class="${mom === null ? 'muted' : mom >= 0 ? 'pos' : 'neg'}">${mom === null ? '—' : (mom >= 0 ? '+' : '') + mom.toFixed(1) + '%'}</td>`;
        }
        html += `</tr>`;
      });
    });
    html += `</tbody></table>`;
    tblWrap.innerHTML = html; panel.appendChild(tblWrap); c.appendChild(panel);
  }

  function renderQuarterly(skipSelReset?: boolean): void {
    const comp: string[] = completedMonths();
    if (!skipSelReset) { selectedQuarter = currentQuarter(comp); }
    const c = el('view-quarterly'); c.innerHTML = '';
    const availableQuarters: string[] = Object.keys(QUARTERS).filter((q) => QUARTERS[q].some((m) => comp.indexOf(m) >= 0));
    if (!availableQuarters.length) { c.innerHTML = '<div class="panel"><p class="muted">No completed quarters yet.</p></div>'; return; }
    if (!selectedQuarter || availableQuarters.indexOf(selectedQuarter) < 0) { selectedQuarter = availableQuarters[availableQuarters.length - 1]; }

    const sel = document.createElement('select'); sel.className = 'period-select';
    availableQuarters.forEach((q) => { const o = document.createElement('option'); o.value = q; o.text = q; if (q === selectedQuarter) { o.selected = true; } sel.appendChild(o); });
    sel.addEventListener('change', () => { selectedQuarter = sel.value; renderQuarterly(true); });
    const selWrap = document.createElement('div'); selWrap.className = 'period-selector';
    selWrap.innerHTML = '<label>Viewing quarter:</label>'; selWrap.appendChild(sel);

    const branch: string = bFilter();
    const qMonths: string[] = QUARTERS[selectedQuarter as string];
    const completedInQ: string[] = qMonths.filter((m) => comp.indexOf(m) >= 0);

    c.appendChild(titleRow('Quarterly', `${selectedQuarter} snapshot · ${completedInQ.join(', ')}`, `${selectedQuarter} FY27 · ${completedInQ.length}/${qMonths.length} months`, selWrap));
    c.appendChild(kpiBlock(completedInQ));
    c.appendChild(horizBarPanel('chart-qu-bar', completedInQ, `Revenue vs budget · ${selectedQuarter} (${completedInQ.join(', ')})`));
    c.appendChild(perfMetricPanel(completedInQ));

    const allQData = Object.keys(QUARTERS).map((q) => {
      const cInQ = QUARTERS[q].filter((m) => comp.indexOf(m) >= 0);
      if (!cInQ.length) { return { q, actual: 0, target: 0, hasData: false, completedInQ: cInQ }; }
      const rows = getRows(branch === 'ALL' ? null : branch, 'Revenue', cInQ);
      return { q, ...sumRows(rows), hasData: true, completedInQ: cInQ };
    }).filter((d) => d.hasData);

    const chartPanel = document.createElement('div'); chartPanel.className = 'panel';
    chartPanel.innerHTML = `<div class="panel-title">Revenue by quarter — all completed quarters</div>` + legendHtml();
    const wrap = document.createElement('div'); wrap.className = 'chart-wrap'; wrap.style.height = '240px';
    wrap.innerHTML = `<canvas id="chart-qu-grp"></canvas>`;
    chartPanel.appendChild(wrap); c.appendChild(chartPanel);
    setTimeout(() => {
      buildGroupedBar('chart-qu-grp',
        allQData.map((d) => `${d.q} (${d.completedInQ.join(',')})`),
        [{ label: 'Actual', data: allQData.map((d) => d.actual), backgroundColor: allQData.map((d) => d.q === selectedQuarter ? '#185FA5' : 'rgba(24,95,165,0.35)'), borderRadius: 3, barPercentage: 0.6 },
         { label: 'Budget', data: allQData.map((d) => d.target), backgroundColor: 'rgba(24,95,165,0.1)', borderColor: '#185FA5', borderWidth: 1, borderRadius: 3, barPercentage: 0.6 }]);
    }, 0);

    c.appendChild(branchMetricDetailPanel(completedInQ, 'Revenue', `${selectedQuarter} detail by branch — revenue`));
  }

  function showToast(msg: string): void {
    let t = root.querySelector('[data-el="opsToast"]') as HTMLElement;
    if (!t) {
      t = document.createElement('div'); t.setAttribute('data-el', 'opsToast');
      t.style.cssText = 'position:fixed;bottom:28px;right:28px;background:#1a2740;color:#fff;font-size:13px;font-weight:500;padding:10px 18px;border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,0.25);z-index:9999;transition:opacity 0.4s;opacity:0;pointer-events:none;';
      document.body.appendChild(t);
    }
    t.textContent = msg; t.style.opacity = '1';
    clearTimeout((t as any)._timer);
    (t as any)._timer = setTimeout(() => { t.style.opacity = '0'; }, 2500);
  }

  async function saveDirtyNotes(btn: HTMLButtonElement): Promise<void> {
    const keys = Object.keys(notesDirty).filter((k) => notesDirty[k]);
    if (!keys.length) { showToast('No note changes to save'); return; }
    btn.disabled = true; const orig = btn.textContent; btn.textContent = 'Saving…';
    try {
      for (const key of keys) {
        const parts = key.split('|'); const branch = parts[0], metric = parts[1], month = parts[2];
        const existing = notesMap[key] ? notesMap[key].id : null;
        const newId = await opts.saveNote(branch, metric, month, notesLocal[key] || '', existing);
        notesMap[key] = { id: newId, note: notesLocal[key] || '' };
        delete notesDirty[key];
      }
      showToast('Notes saved successfully');
    } catch (e) {
      showToast('Could not save notes — check list permissions');
      /* eslint-disable-next-line no-console */
      console.error(e);
    } finally {
      btn.disabled = false; btn.textContent = orig;
    }
  }

  function renderBranch(skipMonthReset?: boolean): void {
    const comp: string[] = completedMonths();
    if (!skipMonthReset) { selectedBranchMonth = comp.length ? comp[comp.length - 1] : null; }
    const c = el('view-branch'); c.innerHTML = '';
    const branch: string = bFilter(); const metric: string = mFilter();
    if (!comp.length) { c.appendChild(titleRow('Branch detail', 'Actuals, budgets & notes per branch', 'FY27 · 0 months complete')); c.innerHTML += '<div class="panel"><p class="muted">No completed months yet.</p></div>'; return; }
    if (!selectedBranchMonth || comp.indexOf(selectedBranchMonth) < 0) { selectedBranchMonth = comp[comp.length - 1]; }

    const monthSel = document.createElement('select'); monthSel.className = 'period-select';
    comp.forEach((m) => { const o = document.createElement('option'); o.value = m; o.text = m; if (m === selectedBranchMonth) { o.selected = true; } monthSel.appendChild(o); });
    monthSel.addEventListener('change', () => { selectedBranchMonth = monthSel.value; renderBranch(true); });
    const monthWrap = document.createElement('div'); monthWrap.className = 'period-selector';
    monthWrap.innerHTML = '<label>Viewing month:</label>'; monthWrap.appendChild(monthSel);

    const saveBtn = document.createElement('button'); saveBtn.className = 'btn'; saveBtn.textContent = 'Save notes';
    saveBtn.addEventListener('click', () => { saveDirtyNotes(saveBtn).catch(() => undefined); });

    const rightExtras = document.createElement('div'); rightExtras.style.cssText = 'display:flex;align-items:center;gap:10px;flex-wrap:wrap;';
    rightExtras.appendChild(monthWrap); rightExtras.appendChild(saveBtn);

    const mo: string = selectedBranchMonth as string;
    c.appendChild(titleRow('Branch detail', `Month: ${mo} · actuals, budgets & notes`, `FY27 · ${comp.length} months complete`, rightExtras));

    const displayBranches: string[] = branch === 'ALL' ? dataset.branches : [branch];
    const displayMetrics: string[] = metric === 'ALL' ? dataset.metrics : [metric];

    displayBranches.forEach((b) => {
      const section = document.createElement('div'); section.className = 'branch-section';
      const heading = document.createElement('div'); heading.className = 'branch-heading'; heading.textContent = b;
      section.appendChild(heading);
      const tblWrap = document.createElement('div'); tblWrap.className = 'branch-table-wrap';
      const innerWrap = document.createElement('div'); innerWrap.className = 'tbl-wrap';
      let html = `<table><thead><tr><th>Metric</th><th>Actual</th><th>Budget</th><th>Variance</th><th>vs budget</th><th style="min-width:240px">Notes — ${escapeHtml(mo)}</th></tr></thead><tbody>`;
      displayMetrics.forEach((dm) => {
        const rows = getRows(b, dm, [mo]);
        if (!rows.length) { return; }
        const { actual, target } = sumRows(rows); const variance = actual - target;
        const pass = isPass(actual, target, dm); const pct = target > 0 ? ((actual / target - 1) * 100) : 0;
        const nk = noteKey(b, dm, mo);
        const noteVal = (nk in notesLocal) ? notesLocal[nk] : (notesMap[nk] ? notesMap[nk].note : '');
        html += `<tr><td class="fw">${escapeHtml(dm)}</td><td>${fmtV(actual, dm)}</td><td class="muted">${fmtV(target, dm)}</td><td class="${variance >= 0 ? 'pos' : 'neg'}">${variance >= 0 ? '+' : ''}${fmtV(variance, dm)}</td><td class="${pass ? 'pos' : 'neg'} fw">${(pct >= 0 ? '+' : '') + pct.toFixed(1)}%</td><td><textarea class="note" data-key="${escapeHtml(nk)}" placeholder="Add note for ${escapeHtml(mo)}…">${escapeHtml(noteVal)}</textarea></td></tr>`;
      });
      html += `</tbody></table>`;
      innerWrap.innerHTML = html;
      innerWrap.querySelectorAll('textarea.note').forEach((ta) => {
        ta.addEventListener('input', (e) => {
          const t = e.target as HTMLTextAreaElement; const key = t.getAttribute('data-key') as string;
          notesLocal[key] = t.value; notesDirty[key] = true;
        });
      });
      tblWrap.appendChild(innerWrap); section.appendChild(tblWrap); c.appendChild(section);
    });
  }

  // ---- nav / boot ----------------------------------------------------------
  function renderAll(): void {
    destroyAllCharts();
    renderOverview();
    renderMonthly(true);
    renderQuarterly(true);
    renderBranch(true);
  }
  function switchTab(tab: string): void {
    root.querySelectorAll('.view').forEach((v) => v.classList.remove('active'));
    const view = el('view-' + tab); if (view) { view.classList.add('active'); }
    const main = el('main'); if (main) { main.scrollTop = 0; }
  }

  function populateFilters(): void {
    const bs = el('branchSel') as HTMLSelectElement;
    const ms = el('metricSel') as HTMLSelectElement;
    while (bs.children.length > 1) { bs.removeChild(bs.lastChild as Node); }
    while (ms.children.length > 1) { ms.removeChild(ms.lastChild as Node); }
    dataset.branches.forEach((b) => { const o = document.createElement('option'); o.value = b; o.text = b; bs.appendChild(o); });
    dataset.metrics.forEach((m) => { const o = document.createElement('option'); o.value = m; o.text = m; ms.appendChild(o); });

    const nav = el('branchNav'); nav.innerHTML = '';
    const allBtn = document.createElement('button'); allBtn.className = 'nav-item'; allBtn.setAttribute('data-branch', 'ALL'); allBtn.textContent = 'All branches';
    nav.appendChild(allBtn);
    dataset.branches.forEach((b) => {
      const btn = document.createElement('button'); btn.className = 'nav-item'; btn.setAttribute('data-branch', b); btn.textContent = b;
      nav.appendChild(btn);
    });
  }

  function wireNav(): void {
    // Delegate: branch nav buttons are created later (populateFilters), so a
    // single delegated handler covers both the static Views buttons and the
    // dynamic branch buttons.
    root.addEventListener('click', (e) => {
      const b = (e.target as HTMLElement).closest('.nav-item') as HTMLElement;
      if (!b || !root.contains(b)) { return; }
      const branch = b.getAttribute('data-branch');
      root.querySelectorAll('.nav-item').forEach((n) => n.classList.remove('active'));
      b.classList.add('active');
      if (branch) { (el('branchSel') as HTMLSelectElement).value = branch; renderAll(); switchTab('overview'); }
      else { switchTab(b.getAttribute('data-tab') as string); }
    });
    (el('branchSel') as HTMLSelectElement).addEventListener('change', renderAll);
    (el('metricSel') as HTMLSelectElement).addEventListener('change', renderAll);
  }

  function stateMsg(title: string, sub: string, detail?: string): void {
    const c = el('view-overview'); c.innerHTML = `<div class="panel"><div class="state-screen"><div class="state-title">${escapeHtml(title)}</div><div class="state-sub">${escapeHtml(sub)}</div>${detail ? `<div class="state-detail">${escapeHtml(detail)}</div>` : ''}</div></div>`;
    switchTab('overview');
  }

  async function load(): Promise<void> {
    stateMsg('Loading…', 'Reading the FY27 workbook from SharePoint.');
    try {
      const [ds, nm] = await Promise.all([opts.fetchData(), opts.fetchNotes()]);
      dataset = ds; notesMap = nm;
      populateFilters();
      renderAll();
      switchTab('overview');
    } catch (err) {
      const msg = (err && (err as Error).message) ? (err as Error).message : ('' + err);
      stateMsg('Could not load the dashboard', 'Check the workbook path and that the file is the FY27 metrics workbook, then reload.', msg);
    }
  }

  // static Views nav is already in the DOM; wire it once the branch nav exists.
  wireNav();
  load().catch(() => undefined);

  return { destroy(): void { destroyAllCharts(); } };
}
