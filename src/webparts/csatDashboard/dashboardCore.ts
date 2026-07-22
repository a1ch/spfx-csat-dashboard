import { ICsatItem } from './CsatDataService';

// Chart.js and ExcelJS are loaded from CDN by the web part and attached to
// window, so we reference them as ambient globals here.
/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Chart: any;
declare const ExcelJS: any;

export interface IDashboardOptions {
  fetchItems: () => Promise<ICsatItem[]>;
  autoRefreshSeconds: number;
}

export interface IDashboardController {
  destroy: () => void;
}

// SharePoint may call the web part's render() rapidly; never let auto-refresh
// hammer the list faster than this, whatever the configured value.
const MIN_AUTO_REFRESH_SECONDS: number = 30;

const starQs: { col: keyof ICsatItem; label: string }[] = [
  { col: 'r_overall', label: 'Overall satisfaction' },
  { col: 'r_technical', label: 'Technical competency' },
  { col: 'r_timeliness', label: 'Timeliness' },
  { col: 'r_communication', label: 'Communication' },
  { col: 'r_quality', label: 'Workmanship quality' },
  { col: 'r_professionalism', label: 'Professionalism' },
  { col: 'r_cleanliness', label: 'Site cleanliness' }
];

const EXPORT_FIELDS: string[] = [
  'id', 'serviceDate', 'branch', 'technician', 'company', 'contactName', 'contactTitle',
  'contactInfo', 'rigWellName', 'location', 'workOrder', 'serviceType',
  'r_overall', 'r_technical', 'r_timeliness', 'r_communication', 'r_quality',
  'r_professionalism', 'r_cleanliness', 'nps', 'safety', 'improvementAreas',
  'exemplary', 'improveSuggestion', 'comments', 'serviceNotes', 'Created'
];

// Free-text columns that hold long verbatim answers — fixed readable width + wrap.
const WRAP_FIELDS: string[] = ['improvementAreas', 'exemplary', 'improveSuggestion', 'comments', 'serviceNotes'];
const WRAP_WIDTH: number = 40;
const XL_DARK: string = 'FF1A3A5C';
const XL_LIGHT: string = 'FFE6F1FB';
const AXIS_5: any = { min: 0, max: 5, ticks: { stepSize: 1 } };

/**
 * Boots the dashboard inside `root`. All DOM lookups are scoped to `root`
 * (via data-el attributes). Data comes from opts.fetchItems().
 */
export function initDashboard(root: HTMLElement, opts: IDashboardOptions): IDashboardController {
  const el = (name: string): HTMLElement => root.querySelector(`[data-el="${name}"]`) as HTMLElement;

  let allItems: ICsatItem[] = [];
  let trendChartInst: any = null;
  let branchChartInst: any = null;
  let autoTimer: any = null;

  // ---- small helpers -------------------------------------------------------
  function setLive(state: string, text: string): void {
    const dot = el('liveDot');
    dot.className = 'live-dot' + (state === 'error' ? ' error' : state === 'loading' ? ' loading' : '');
    el('liveText').textContent = text;
  }
  function showState(html: string): void {
    el('stateScreen').innerHTML = html;
    el('stateScreen').style.display = 'block';
    el('dashboardContent').style.display = 'none';
  }
  function hideState(): void {
    el('stateScreen').style.display = 'none';
    el('dashboardContent').style.display = 'block';
  }
  function avg(arr: (number | null)[]): number | null {
    const a = arr.filter((v) => v !== null && v !== undefined && !isNaN(v as number)) as number[];
    return a.length ? a.reduce((x, y) => x + y, 0) / a.length : null;
  }
  function fmt(v: number | null, d: number = 1): string {
    return v === null || v === undefined || isNaN(v) ? '—' : v.toFixed(d);
  }
  function round1(v: number | null): number | null {
    return v === null || v === undefined || isNaN(v) ? null : Math.round(v * 10) / 10;
  }
  function escapeHtml(s: unknown): string {
    return ('' + s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
  }
  function itemDate(d: ICsatItem): Date | null {
    const raw = d.serviceDate || d.Created;
    return raw ? new Date(raw) : null;
  }
  function commenterOf(d: ICsatItem): string {
    return d.contactName ? (d.company ? `${d.contactName} (${d.company})` : d.contactName) : (d.company || 'Customer');
  }
  function commentCard(d: ICsatItem, text: string, quote: boolean): string {
    const body = quote ? `"${text}"` : text;
    return `<div class="comment-card"><div class="comment-text">${escapeHtml(body)}</div><div class="comment-meta">${escapeHtml(commenterOf(d))} · ${escapeHtml(d.branch || '—')} · ${((d.serviceDate || '') + '').split('T')[0]}${d.serviceType ? ' · ' + escapeHtml(d.serviceType) : ''}</div></div>`;
  }
  function npsScoreOf(items: ICsatItem[]): number | null {
    const vals = items.map((d) => d.nps).filter((v) => v !== null && v !== undefined) as number[];
    if (!vals.length) { return null; }
    const det = vals.filter((v) => v <= 6).length;
    const pro = vals.filter((v) => v >= 9).length;
    return Math.round(((pro - det) / vals.length) * 100);
  }
  function setDelta(name: string, cur: number | null, prev: number | null, digits: number): void {
    const node = el(name);
    if (cur === null || prev === null || prev === undefined) { node.textContent = ''; return; }
    const diff = cur - prev;
    const arrow = diff > 0.0001 ? '▲' : diff < -0.0001 ? '▼' : '▬';
    const cls = diff > 0.0001 ? 'up' : diff < -0.0001 ? 'down' : 'flat';
    node.className = 'delta ' + cls;
    node.textContent = ` ${arrow} ${Math.abs(diff).toFixed(digits)}`;
    node.title = 'vs previous period';
  }

  // ---- filtering -----------------------------------------------------------
  interface IFilter { branch?: string; technician?: string; days?: number; from?: Date; to?: Date; }

  function currentFilters(): IFilter {
    return {
      branch: (el('branchFilter') as HTMLSelectElement).value,
      technician: (el('techFilter') as HTMLSelectElement).value,
      days: parseInt((el('rangeFilter') as HTMLSelectElement).value, 10)
    };
  }
  function filterItems(f: IFilter): ICsatItem[] {
    return allItems.filter((d) => {
      if (f.branch && d.branch !== f.branch) { return false; }
      if (f.technician && d.technician !== f.technician) { return false; }
      if (f.days !== null && f.days !== undefined && f.days < 9999) {
        const sd = itemDate(d);
        if (sd) { const cutoff = new Date(); cutoff.setDate(cutoff.getDate() - f.days); if (sd < cutoff) { return false; } }
      }
      if (f.from || f.to) {
        const sd = itemDate(d);
        if (sd) { if (f.from && sd < f.from) { return false; } if (f.to && sd >= f.to) { return false; } }
      }
      return true;
    });
  }
  function filtered(): ICsatItem[] { return filterItems(currentFilters()); }
  function previousPeriod(): ICsatItem[] | null {
    const f = currentFilters();
    if (!f.days || f.days >= 9999) { return null; }
    const now = new Date();
    const start = new Date(); start.setDate(now.getDate() - f.days);
    const prevStart = new Date(); prevStart.setDate(now.getDate() - f.days * 2);
    return filterItems({ branch: f.branch, technician: f.technician, from: prevStart, to: start });
  }

  function initFilters(): void {
    const bf = el('branchFilter') as HTMLSelectElement;
    const tf = el('techFilter') as HTMLSelectElement;
    const prevB = bf.value, prevT = tf.value;
    [bf, tf].forEach((sel) => { while (sel.children.length > 1) { sel.removeChild(sel.lastChild as Node); } });
    Array.from(new Set(allItems.map((d) => d.branch).filter(Boolean))).sort().forEach((v: string) => { const o = document.createElement('option'); o.value = v; o.textContent = v; bf.appendChild(o); });
    Array.from(new Set(allItems.map((d) => d.technician).filter(Boolean))).sort().forEach((v: string) => { const o = document.createElement('option'); o.value = v; o.textContent = v; tf.appendChild(o); });
    bf.value = prevB; tf.value = prevT;
  }

  // ---- render --------------------------------------------------------------
  function render(): void {
    const data = filtered();
    const prev = previousPeriod();

    el('kTotal').textContent = '' + data.length;
    el('kTotalSub').textContent = prev ? `prev period: ${prev.length}` : '';

    const csat = avg(data.map((d) => d.r_overall));
    el('kCsat').textContent = fmt(csat);
    setDelta('kCsatDelta', csat, prev ? avg(prev.map((d) => d.r_overall)) : null, 2);

    const npsVals = data.map((d) => d.nps).filter((v) => v !== null && v !== undefined) as number[];
    const det = npsVals.filter((v) => v <= 6).length;
    const pas = npsVals.filter((v) => v >= 7 && v <= 8).length;
    const pro = npsVals.filter((v) => v >= 9).length;
    const npsScore = npsScoreOf(data);
    const npsTxt = npsScore !== null ? (npsScore > 0 ? '+' : '') + npsScore : '—';
    el('kNps').textContent = npsTxt;
    el('npsScoreVal').textContent = npsTxt;
    el('detCount').textContent = '' + det;
    el('pasCount').textContent = '' + pas;
    el('proCount').textContent = '' + pro;
    setDelta('kNpsDelta', npsScore, prev ? npsScoreOf(prev) : null, 0);

    const safety = avg(data.map((d) => d.safety));
    el('kSafety').textContent = fmt(safety);
    setDelta('kSafetyDelta', safety, prev ? avg(prev.map((d) => d.safety)) : null, 2);

    renderAttention(data);

    const sb = el('starBreakdown'); sb.innerHTML = '';
    starQs.forEach((q) => {
      const a = avg(data.map((d) => d[q.col] as number | null));
      const rounded = a ? Math.round(a) : 0;
      const stars = '★'.repeat(rounded) + '☆'.repeat(5 - rounded);
      sb.innerHTML += `<div class="star-row"><span class="star-name">${q.label}</span><span class="stars-mini">${stars}</span><span class="star-score">${fmt(a)}</span></div>`;
    });

    const sorted = allItems.length ? [...data].sort((a, b) => ((itemDate(a) as any) || 0) - ((itemDate(b) as any) || 0)) : [];
    const labels = sorted.map((d) => ((d.serviceDate || d.Created || '') + '').split('T')[0]);
    const scores = sorted.map((d) => (d.r_overall === null || d.r_overall === undefined ? null : d.r_overall));
    if (typeof Chart !== 'undefined') {
      if (trendChartInst) { trendChartInst.destroy(); }
      trendChartInst = new Chart(el('trendChart'), {
        type: 'line',
        data: { labels, datasets: [{ label: 'Overall rating', data: scores, borderColor: '#378add', backgroundColor: 'rgba(55,138,221,0.12)', pointRadius: 0, borderWidth: 2, tension: 0.35, fill: true, spanGaps: true }] },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 1, max: 5, grid: { color: 'rgba(0,0,0,0.06)' } }, x: { ticks: { maxTicksLimit: 10, maxRotation: 30 }, grid: { display: false } } }, plugins: { legend: { display: false } } }
      });

      const branchMap: { [k: string]: number[] } = {};
      data.forEach((d) => { if (!d.branch) { return; } if (!branchMap[d.branch]) { branchMap[d.branch] = []; } if (d.r_overall !== null && d.r_overall !== undefined) { branchMap[d.branch].push(d.r_overall); } });
      const branchEntries = Object.keys(branchMap).map((k) => [k, branchMap[k]] as [string, number[]]);
      if (branchChartInst) { branchChartInst.destroy(); }
      branchChartInst = new Chart(el('branchChart'), {
        type: 'bar',
        data: { labels: branchEntries.map((e) => e[0]), datasets: [{ label: 'Avg CSAT', data: branchEntries.map((e) => avg(e[1])), backgroundColor: '#378add', borderRadius: 4 }] },
        options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 0, max: 5, grid: { color: 'rgba(0,0,0,0.06)' } }, x: { grid: { display: false }, ticks: { autoSkip: false, maxRotation: 20 } } }, plugins: { legend: { display: false } } }
      });
    }

    const imprMap: { [k: string]: number } = {};
    data.forEach((d) => { (d.improvementAreas || '').split(';').map((s) => s.trim()).filter(Boolean).forEach((a) => { imprMap[a] = (imprMap[a] || 0) + 1; }); });
    const imprSorted = Object.keys(imprMap).map((k) => [k, imprMap[k]] as [string, number]).sort((a, b) => b[1] - a[1]);
    const ib = el('improvBars'); ib.innerHTML = '';
    if (!imprSorted.length) { ib.innerHTML = '<div class="empty-note">No improvement areas cited in this period.</div>'; }
    else { imprSorted.forEach(([label, count]) => { const pct = Math.round(count / data.length * 100); ib.innerHTML += `<div class="improv-bar"><div class="improv-label"><span>${escapeHtml(label)}</span><span>${count} (${pct}%)</span></div><div class="improv-track"><div class="improv-fill" style="width:${pct}%"></div></div></div>`; }); }

    const techMap: { [k: string]: { jobs: number; ratings: number[] } } = {};
    data.forEach((d) => { if (!d.technician) { return; } if (!techMap[d.technician]) { techMap[d.technician] = { jobs: 0, ratings: [] }; } techMap[d.technician].jobs++; if (d.r_overall !== null && d.r_overall !== undefined) { techMap[d.technician].ratings.push(d.r_overall); } });
    const techEntries = Object.keys(techMap).map((k) => [k, techMap[k]] as [string, { jobs: number; ratings: number[] }]).sort((a, b) => b[1].jobs - a[1].jobs);
    const tbody = el('techTable'); tbody.innerHTML = '';
    if (!techEntries.length) { tbody.innerHTML = '<tr><td colspan="4" class="empty-note">No data in this period.</td></tr>'; }
    else {
      techEntries.forEach(([name, d]) => {
        const ra = avg(d.ratings);
        const badge = (ra as number) >= 4.5 ? 'badge-good' : (ra as number) >= 3.5 ? 'badge-warn' : 'badge-bad';
        const status = (ra as number) >= 4.5 ? 'Excellent' : (ra as number) >= 3.5 ? 'Good' : 'Needs review';
        tbody.innerHTML += `<tr><td>${escapeHtml(name)}</td><td>${d.jobs}</td><td><span class="badge ${badge}">${fmt(ra)}</span></td><td><span class="badge ${badge}">${status}</span></td></tr>`;
      });
    }

    const exList = el('exemplaryList'); exList.innerHTML = '';
    const exItems = data.filter((d) => d.exemplary && ('' + d.exemplary).trim().length > 10);
    if (!exItems.length) { exList.innerHTML = '<div class="empty-note">No exemplary comments recorded in this period.</div>'; }
    else { exItems.slice(-8).reverse().forEach((d) => { exList.innerHTML += commentCard(d, ('' + d.exemplary).trim(), true); }); }

    const imprList = el('improveList'); imprList.innerHTML = '';
    const imprItems = data.filter((d) => d.improveSuggestion && ('' + d.improveSuggestion).trim().length > 5);
    if (!imprItems.length) { imprList.innerHTML = '<div class="empty-note">No improvement suggestions recorded in this period.</div>'; }
    else { imprItems.slice(-8).reverse().forEach((d) => { imprList.innerHTML += commentCard(d, ('' + d.improveSuggestion).trim(), false); }); }
  }

  function renderAttention(data: ICsatItem[]): void {
    const card = el('attentionCard');
    const list = el('attentionList');
    const flagged = data.filter((d) =>
      (d.r_overall !== null && d.r_overall !== undefined && d.r_overall <= 3) ||
      (d.safety !== null && d.safety !== undefined && d.safety <= 3) ||
      (d.nps !== null && d.nps !== undefined && d.nps <= 6)
    ).sort((a, b) => ((a.r_overall === null ? 5 : a.r_overall) - (b.r_overall === null ? 5 : b.r_overall)));

    el('attnCount').textContent = '' + flagged.length;
    if (!flagged.length) { card.style.display = 'none'; return; }
    card.style.display = 'block';
    list.innerHTML = '';
    flagged.slice(0, 12).forEach((d) => {
      const reasons: string[] = [];
      if (d.r_overall !== null && d.r_overall !== undefined && d.r_overall <= 3) { reasons.push(`CSAT ${d.r_overall}/5`); }
      if (d.safety !== null && d.safety !== undefined && d.safety <= 3) { reasons.push(`Safety ${d.safety}/5`); }
      if (d.nps !== null && d.nps !== undefined && d.nps <= 6) { reasons.push(`NPS ${d.nps}`); }
      const note = (d.improveSuggestion || d.comments || '').toString().trim();
      list.innerHTML += `<div class="comment-card attention"><div class="comment-text"><strong>${escapeHtml(reasons.join(' · '))}</strong>${note ? ' — ' + escapeHtml(note) : ''}</div><div class="comment-meta">${escapeHtml(d.technician || '—')} · ${escapeHtml(commenterOf(d))} · ${escapeHtml(d.branch || '—')} · ${((d.serviceDate || '') + '').split('T')[0]}</div></div>`;
    });
  }

  // ---- CSV export (dependency-free) ----------------------------------------
  function exportRows(): { [k: string]: unknown }[] {
    return filtered().map((d) => {
      const row: { [k: string]: unknown } = {};
      EXPORT_FIELDS.forEach((f) => { const v = (d as any)[f]; row[f] = (v === null || v === undefined) ? '' : v; });
      return row;
    });
  }
  function exportStamp(): string {
    const n = new Date();
    const p = (x: number): string => ('' + x).padStart(2, '0');
    return `${n.getFullYear()}${p(n.getMonth() + 1)}${p(n.getDate())}_${p(n.getHours())}${p(n.getMinutes())}`;
  }
  function triggerDownload(blob: Blob, filename: string): void {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  }
  function exportCSV(): void {
    closeExportMenu();
    const rows = exportRows();
    const esc = (v: unknown): string => { const s = ('' + v).replace(/"/g, '""'); return /[",\n]/.test(s) ? `"${s}"` : s; };
    const lines = [EXPORT_FIELDS.join(',')];
    rows.forEach((r) => lines.push(EXPORT_FIELDS.map((f) => esc(r[f])).join(',')));
    const blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    triggerDownload(blob, `csat_responses_${exportStamp()}.csv`);
  }

  // ---- Excel export (ExcelJS: Summary + per-branch + All Responses) --------
  // Charts are rendered from Chart.js to PNG and embedded as images (the free
  // Excel libraries cannot author native chart objects).
  const whiteBgPlugin: any = {
    id: 'whiteBg',
    beforeDraw(chart: any): void {
      const ctx = chart.canvas.getContext('2d');
      ctx.save();
      ctx.globalCompositeOperation = 'destination-over';
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, chart.width, chart.height);
      ctx.restore();
    }
  };

  // Renders a chart config off-screen and returns a base64 PNG (2x for crispness).
  async function chartPNG(config: any, width: number, height: number): Promise<string> {
    const holder = el('exportChartHolder');
    const canvas = document.createElement('canvas');
    canvas.width = width; canvas.height = height;
    canvas.style.width = width + 'px'; canvas.style.height = height + 'px';
    holder.appendChild(canvas);

    const cfg = JSON.parse(JSON.stringify(config)); // configs are plain data
    cfg.options = Object.assign({}, cfg.options, {
      responsive: false, maintainAspectRatio: false,
      animation: false, devicePixelRatio: 2
    });
    cfg.plugins = [whiteBgPlugin];

    const chart = new Chart(canvas, cfg);
    // With animation disabled Chart.js paints synchronously on construction.
    // Yield via setTimeout (rAF is suspended while the tab is hidden).
    await new Promise((r) => setTimeout(r, 0));
    const url = canvas.toDataURL('image/png');
    chart.destroy();
    holder.removeChild(canvas);
    return url;
  }

  // ---- chart configs used by the export ------------------------------------
  function cfgBar(labels: any[], values: any[], axis?: any): any {
    return {
      type: 'bar',
      data: { labels, datasets: [{ data: values, backgroundColor: '#378add', borderRadius: 3 }] },
      options: { scales: { y: axis || AXIS_5, x: { ticks: { autoSkip: false, maxRotation: 30, minRotation: 0 } } }, plugins: { legend: { display: false } } }
    };
  }
  function cfgLine(labels: any[], values: any[]): any {
    return {
      type: 'line',
      data: { labels, datasets: [{ data: values, borderColor: '#378add', backgroundColor: 'rgba(55,138,221,0.12)', borderWidth: 2, tension: 0.35, fill: true, pointRadius: 2, spanGaps: true }] },
      options: { scales: { y: { min: 1, max: 5 }, x: { ticks: { maxTicksLimit: 12, maxRotation: 40 } } }, plugins: { legend: { display: false } } }
    };
  }
  function cfgDoughnut(labels: any[], values: any[]): any {
    return {
      type: 'doughnut',
      data: { labels, datasets: [{ data: values, backgroundColor: ['#a32d2d', '#d9a441', '#0f6e56'] }] },
      options: { plugins: { legend: { position: 'right' } } }
    };
  }

  // ---- aggregation helpers -------------------------------------------------
  function summaryOf(rows: ICsatItem[]): any {
    const npsVals = rows.map((d) => d.nps).filter((v) => v !== null && v !== undefined) as number[];
    return {
      count: rows.length,
      csat: avg(rows.map((d) => d.r_overall)),
      safety: avg(rows.map((d) => d.safety)),
      nps: npsScoreOf(rows),
      det: npsVals.filter((v) => v <= 6).length,
      pas: npsVals.filter((v) => v >= 7 && v <= 8).length,
      pro: npsVals.filter((v) => v >= 9).length,
      categories: starQs.map((q) => ({ label: q.label, value: avg(rows.map((d) => d[q.col] as number | null)) }))
    };
  }
  function trendSeries(rows: ICsatItem[]): { labels: string[]; values: (number | null)[] } {
    const byDate: { [k: string]: number[] } = {};
    rows.forEach((d) => {
      const key = ((d.serviceDate || d.Created || '') + '').split('T')[0];
      if (!key || d.r_overall === null || d.r_overall === undefined) { return; }
      (byDate[key] = byDate[key] || []).push(d.r_overall);
    });
    const labels = Object.keys(byDate).sort();
    return { labels, values: labels.map((k) => avg(byDate[k])) };
  }
  function improvementCounts(rows: ICsatItem[]): [string, number][] {
    const map: { [k: string]: number } = {};
    rows.forEach((d) => (d.improvementAreas || '').split(';').map((s) => s.trim()).filter(Boolean)
      .forEach((a) => { map[a] = (map[a] || 0) + 1; }));
    return Object.keys(map).map((k) => [k, map[k]] as [string, number]).sort((a, b) => b[1] - a[1]);
  }

  // ---- worksheet helpers ---------------------------------------------------
  function sanitizeSheetName(name: string, used: Set<string>): string {
    let s = ('' + (name || 'Unknown')).replace(/[\\/?*[\]:]/g, '-').trim().slice(0, 31) || 'Sheet';
    const base = s; let i = 2;
    while (used.has(s.toLowerCase())) { const sfx = ' (' + i + ')'; s = base.slice(0, 31 - sfx.length) + sfx; i++; }
    used.add(s.toLowerCase());
    return s;
  }
  function titleBlock(ws: any, title: string, subtitle: string, span: number): void {
    ws.mergeCells(1, 1, 1, span);
    const t = ws.getCell(1, 1);
    t.value = title;
    t.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
    t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_DARK } };
    t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    ws.getRow(1).height = 26;
    ws.mergeCells(2, 1, 2, span);
    const s = ws.getCell(2, 1);
    s.value = subtitle;
    s.font = { size: 9, color: { argb: 'FF666666' } };
  }
  function writeTable(ws: any, startRow: number, headers: any[], rows: any[][]): number {
    const hr = ws.getRow(startRow);
    headers.forEach((h, i) => {
      const c = hr.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_DARK } };
    });
    hr.height = 17;
    rows.forEach((r, ri) => {
      const row = ws.getRow(startRow + 1 + ri);
      r.forEach((v, i) => { row.getCell(i + 1).value = (v === undefined ? null : v); });
      if (ri % 2 === 1) {
        r.forEach((_, i) => {
          row.getCell(i + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: XL_LIGHT } };
        });
      }
    });
    return startRow + rows.length + 2;
  }
  // Sizes every column to its widest actual cell value. Merged cells (title
  // banner) are skipped or they'd blow out column A.
  function autoFitColumns(ws: any, opts?: any): void {
    const o = Object.assign({ min: 9, max: 55, padding: 2 }, opts || {});
    const longest: number[] = [];
    ws.eachRow({ includeEmpty: false }, (row: any) => {
      row.eachCell({ includeEmpty: false }, (cell: any, colNumber: number) => {
        if (cell.isMerged) { return; }
        let v = cell.value;
        if (v === null || v === undefined) { return; }
        if (typeof v === 'object') {
          if (v.richText) { v = v.richText.map((t: any) => t.text).join(''); }
          else if (v.text !== undefined) { v = v.text; }
          else if (v.result !== undefined) { v = v.result; }
          else { return; }
        }
        ('' + v).split('\n').forEach((line) => {
          if (line.length > (longest[colNumber] || 0)) { longest[colNumber] = line.length; }
        });
      });
    });
    longest.forEach((len, colNumber) => {
      if (!colNumber) { return; }
      ws.getColumn(colNumber).width = Math.min(o.max, Math.max(o.min, len) + o.padding);
    });
  }
  function wrapColumns(ws: any, colNumbers: number[], width?: number): void {
    const targets = new Set(colNumbers.filter(Boolean));
    if (!targets.size) { return; }
    targets.forEach((cn) => { ws.getColumn(cn).width = width || WRAP_WIDTH; });
    ws.eachRow({ includeEmpty: false }, (row: any) => {
      row.eachCell({ includeEmpty: false }, (cell: any, cn: number) => {
        if (!targets.has(cn) || cell.isMerged) { return; }
        cell.alignment = Object.assign({}, cell.alignment || {}, { wrapText: true, vertical: 'top' });
      });
    });
  }
  // Returns the first column index at/after the given pixel offset — used to
  // place a second chart beside the first without overlapping it.
  function colAtPixel(ws: any, px: number): number {
    let acc = 0, col = 0;
    while (acc < px && col < 80) { acc += (ws.getColumn(col + 1).width || 9) * 7; col++; }
    return col;
  }
  async function addChart(wb: any, ws: any, config: any, anchorCol: number, anchorRow: number, w: number, h: number): Promise<void> {
    const png = await chartPNG(config, w, h);
    const id = wb.addImage({ base64: png.split(',')[1], extension: 'png' });
    ws.addImage(id, { tl: { col: anchorCol, row: anchorRow }, ext: { width: w, height: h } });
  }

  function filterDescription(): string {
    const f = currentFilters();
    const parts: string[] = [];
    parts.push(f.branch ? `Branch: ${f.branch}` : 'Branch: All');
    parts.push((f.days as number) >= 9999 ? 'Range: All time' : `Range: Last ${f.days} days`);
    parts.push(f.technician ? `Technician: ${f.technician}` : 'Technician: All');
    return parts.join('  ·  ');
  }

  // ---- sheet builders ------------------------------------------------------
  async function buildSummarySheet(wb: any, rows: ICsatItem[]): Promise<any> {
    const ws = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
    titleBlock(ws, 'Stream-Flo — Field Service CSAT Summary',
      `Generated ${new Date().toLocaleString()}   ·   ${filterDescription()}`, 12);

    const s = summaryOf(rows);
    let r = writeTable(ws, 4,
      ['Metric', 'Value'],
      [['Total responses', s.count],
       ['Avg CSAT (out of 5)', round1(s.csat)],
       ['NPS score', s.nps],
       ['Safety score (out of 5)', round1(s.safety)],
       ['Promoters (9–10)', s.pro],
       ['Passives (7–8)', s.pas],
       ['Detractors (0–6)', s.det]]);

    const branches = Array.from(new Set(rows.map((d) => d.branch).filter(Boolean))).sort();
    const branchRows = branches.map((b) => {
      const bs = summaryOf(rows.filter((d) => d.branch === b));
      return [b, bs.count, round1(bs.csat), bs.nps, round1(bs.safety)];
    });
    r = writeTable(ws, r, ['Branch', 'Responses', 'Avg CSAT', 'NPS', 'Safety'], branchRows);

    r = writeTable(ws, r, ['Category', 'Avg rating'],
      s.categories.map((c: any) => [c.label, round1(c.value)]));

    autoFitColumns(ws);

    const trend = trendSeries(rows);
    let chartRow = r + 1;
    const rightCol = colAtPixel(ws, 640);
    await addChart(wb, ws, cfgBar(branches, branchRows.map((x) => x[2])), 0, chartRow, 620, 300);
    await addChart(wb, ws, cfgDoughnut(['Detractors', 'Passives', 'Promoters'], [s.det, s.pas, s.pro]), rightCol, chartRow, 400, 300);
    chartRow += 16;
    await addChart(wb, ws, cfgLine(trend.labels, trend.values.map(round1)), 0, chartRow, 620, 300);
    await addChart(wb, ws, cfgBar(s.categories.map((c: any) => c.label), s.categories.map((c: any) => round1(c.value))), rightCol, chartRow, 480, 300);
    chartRow += 16;

    const impr = improvementCounts(rows);
    if (impr.length) {
      await addChart(wb, ws, cfgBar(impr.map((x) => x[0]), impr.map((x) => x[1]),
        { beginAtZero: true, ticks: { precision: 0 } }), 0, chartRow, 620, 300);
    }
    return ws;
  }

  async function buildBranchSheet(wb: any, wsName: string, branch: string, rows: ICsatItem[]): Promise<any> {
    const ws = wb.addWorksheet(wsName, { views: [{ showGridLines: false }] });
    titleBlock(ws, `${branch} — CSAT Detail`,
      `${rows.length} response(s)   ·   Generated ${new Date().toLocaleString()}`, 12);

    const s = summaryOf(rows);
    let r = writeTable(ws, 4, ['Metric', 'Value'],
      [['Responses', s.count],
       ['Avg CSAT', round1(s.csat)],
       ['NPS score', s.nps],
       ['Safety score', round1(s.safety)]]);

    const techMap: { [k: string]: (number | null)[] } = {};
    rows.forEach((d) => {
      if (!d.technician) { return; }
      (techMap[d.technician] = techMap[d.technician] || []).push(d.r_overall);
    });
    const techRows = Object.keys(techMap)
      .map((n) => [n, techMap[n].length, round1(avg(techMap[n]))])
      .sort((a, b) => (b[1] as number) - (a[1] as number));
    if (techRows.length) { r = writeTable(ws, r, ['Technician', 'Jobs', 'Avg rating'], techRows); }

    const trend = trendSeries(rows);
    const chartRow = r + 1;

    const headers = ['Service date', 'Technician', 'Company', 'Contact', 'Service type',
                     'Overall', 'Technical', 'Timeliness', 'Comms', 'Quality',
                     'Professional', 'Cleanliness', 'NPS', 'Safety',
                     'Improvement areas', 'Exemplary', 'Suggestion'];
    const detail = rows.map((d) => [
      ((d.serviceDate || '') + '').split('T')[0], d.technician || '', d.company || '',
      d.contactName || '', d.serviceType || '',
      d.r_overall, d.r_technical, d.r_timeliness, d.r_communication, d.r_quality,
      d.r_professionalism, d.r_cleanliness, d.nps, d.safety,
      d.improvementAreas || '', d.exemplary || '', d.improveSuggestion || ''
    ]);
    const tableStart = chartRow + 16 + 1;
    writeTable(ws, tableStart, headers, detail);
    ws.autoFilter = { from: { row: tableStart, column: 1 }, to: { row: tableStart, column: headers.length } };

    autoFitColumns(ws);
    wrapColumns(ws, ['Improvement areas', 'Exemplary', 'Suggestion'].map((h) => headers.indexOf(h) + 1));
    await addChart(wb, ws, cfgBar(s.categories.map((c: any) => c.label), s.categories.map((c: any) => round1(c.value))), 0, chartRow, 560, 290);
    if (trend.labels.length > 1) {
      await addChart(wb, ws, cfgLine(trend.labels, trend.values.map(round1)), colAtPixel(ws, 580), chartRow, 500, 290);
    }
    return ws;
  }

  function buildAllResponsesSheet(wb: any, rows: ICsatItem[]): any {
    const ws = wb.addWorksheet('All Responses');
    const headers = EXPORT_FIELDS;
    writeTable(ws, 1, headers, rows.map((d) => headers.map((f) =>
      ((d as any)[f] === null || (d as any)[f] === undefined) ? '' : (d as any)[f])));
    autoFitColumns(ws);
    wrapColumns(ws, WRAP_FIELDS.map((f) => headers.indexOf(f) + 1));
    ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headers.length } };
    ws.views = [{ state: 'frozen', ySplit: 1 }];
    return ws;
  }

  async function exportXLSX(): Promise<void> {
    closeExportMenu();
    if (typeof ExcelJS === 'undefined') {
      alert('Excel export library did not load (offline?). Use CSV instead.');
      return;
    }
    const btn = root.querySelector('[data-action="exportToggle"]') as HTMLButtonElement;
    const original = btn ? btn.textContent : '';
    if (btn) { btn.textContent = '⏳ Building…'; btn.disabled = true; }

    try {
      const rows = filtered();
      if (!rows.length) { alert('No responses in the current filter to export.'); return; }

      const wb = new ExcelJS.Workbook();
      wb.creator = 'Stream-Flo CSAT Dashboard';
      wb.created = new Date();

      await buildSummarySheet(wb, rows);

      const used = new Set(['summary', 'all responses']);
      const branches = Array.from(new Set(rows.map((d) => d.branch).filter(Boolean))).sort();
      for (const b of branches) {
        const name = sanitizeSheetName(b, used);
        await buildBranchSheet(wb, name, b, rows.filter((d) => d.branch === b));
      }

      buildAllResponsesSheet(wb, rows);

      const buf = await wb.xlsx.writeBuffer();
      triggerDownload(
        new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        `csat_dashboard_${exportStamp()}.xlsx`);
    } catch (err) {
      /* eslint-disable-next-line no-console */
      console.error(err);
      const msg = (err && (err as Error).message) ? (err as Error).message : ('' + err);
      alert('Could not build the Excel file: ' + msg);
    } finally {
      if (btn) { btn.textContent = original; btn.disabled = false; }
    }
  }

  function toggleExportMenu(e: Event): void { e.stopPropagation(); el('exportMenu').classList.toggle('open'); }
  function closeExportMenu(): void { const m = root.querySelector('[data-el="exportMenu"]'); if (m) { m.classList.remove('open'); } }

  // ---- auto refresh --------------------------------------------------------
  // start/stop are idempotent: startAuto never creates a second timer, so
  // repeated calls (e.g. from repeated web part render()) cannot stack up and
  // fire loadData in rapid succession.
  function autoIntervalMs(): number {
    const secs = opts.autoRefreshSeconds > 0 ? opts.autoRefreshSeconds : 120;
    return Math.max(secs, MIN_AUTO_REFRESH_SECONDS) * 1000;
  }
  function startAuto(): void {
    if (autoTimer) { return; }
    const btn = root.querySelector('[data-action="auto"]') as HTMLElement;
    autoTimer = setInterval(() => { loadData().catch(() => undefined); }, autoIntervalMs());
    if (btn) { btn.classList.add('active'); btn.textContent = '◷ Auto on'; }
  }
  function stopAuto(): void {
    if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
    const btn = root.querySelector('[data-action="auto"]') as HTMLElement;
    if (btn) { btn.classList.remove('active'); btn.textContent = '◷ Auto'; }
  }
  function toggleAuto(): void {
    if (autoTimer) { stopAuto(); } else { startAuto(); }
  }

  // ---- load ----------------------------------------------------------------
  async function loadData(): Promise<void> {
    setLive('loading', 'Syncing…');
    el('connSub').textContent = 'Connecting…';
    const refreshBtn = root.querySelector('[data-action="refresh"]') as HTMLButtonElement;
    if (refreshBtn) { refreshBtn.disabled = true; }

    try {
      allItems = await opts.fetchItems();

      if (!allItems.length) {
        setLive('ok', 'Connected · 0 responses');
        el('connSub').textContent = 'CSAT Responses · 0 total';
        showState('<div class="state-screen"><div class="state-icon">📋</div><div class="state-title">Connected — no responses yet</div><div class="state-sub">The list is reachable but has no responses yet. Results will appear here as branches submit the survey.</div></div>');
        return;
      }

      setLive('ok', `Connected · ${allItems.length} responses`);
      el('connSub').textContent = `CSAT Responses · ${allItems.length} total`;
      el('lastSync').textContent = 'Last synced: ' + new Date().toLocaleString();
      hideState();
      initFilters();
      render();
    } catch (err) {
      const msg = (err && (err as Error).message) ? (err as Error).message : ('' + err);
      setLive('error', 'Connection failed');
      el('connSub').textContent = 'Could not connect';
      showState(`<div class="state-screen"><div class="state-icon">⚠️</div><div class="state-title">Could not load dashboard data</div><div class="state-sub">The web part could not read the CSAT list. Check that the list name and site URL are correct in the web part properties, and that you have access to that list.</div><div class="state-detail">Error: ${escapeHtml(msg)}</div></div>`);
    } finally {
      if (refreshBtn) { refreshBtn.disabled = false; }
    }
  }

  // ---- wire events ---------------------------------------------------------
  ['branchFilter', 'rangeFilter', 'techFilter'].forEach((n) => {
    const node = el(n);
    if (node) { node.addEventListener('change', render); }
  });
  root.querySelectorAll('[data-action]').forEach((node) => {
    const action = (node as HTMLElement).getAttribute('data-action');
    node.addEventListener('click', (e) => {
      if (action === 'refresh') { loadData().catch(() => undefined); }
      else if (action === 'auto') { toggleAuto(); }
      else if (action === 'exportToggle') { toggleExportMenu(e); }
      else if (action === 'exportCsv') { exportCSV(); }
      else if (action === 'exportXlsx') { exportXLSX().catch(() => undefined); }
    });
  });
  // clicking anywhere else closes the export menu
  root.addEventListener('click', (e) => {
    const t = e.target as HTMLElement;
    if (!t.closest || !t.closest('.dropdown')) { closeExportMenu(); }
  });

  // initial load, then arm auto-refresh once (guarded — never stacks)
  loadData().catch(() => undefined);
  if (opts.autoRefreshSeconds > 0) { startAuto(); }

  return {
    destroy(): void {
      stopAuto();
      if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
    }
  };
}
