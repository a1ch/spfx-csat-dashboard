import { ICsatItem } from './CsatDataService';

// Chart.js and SheetJS are loaded from CDN by the web part and attached to
// window, so we reference them as ambient globals here.
/* eslint-disable @typescript-eslint/no-explicit-any */
declare const Chart: any;
declare const XLSX: any;

export interface IDashboardOptions {
  fetchItems: () => Promise<ICsatItem[]>;
  autoRefreshSeconds: number;
}

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

/**
 * Boots the dashboard inside `root`. All DOM lookups are scoped to `root`
 * (via data-el attributes) so multiple web parts / the SharePoint page around
 * it are never touched. Data comes from opts.fetchItems().
 */
export function initDashboard(root: HTMLElement, opts: IDashboardOptions): void {
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

  // ---- export --------------------------------------------------------------
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
  function exportXLSX(): void {
    closeExportMenu();
    if (typeof XLSX === 'undefined') { alert('Excel export library did not load. Use CSV instead.'); return; }
    const ws = XLSX.utils.json_to_sheet(exportRows(), { header: EXPORT_FIELDS });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'CSAT Responses');
    XLSX.writeFile(wb, `csat_responses_${exportStamp()}.xlsx`);
  }
  function toggleExportMenu(e: Event): void { e.stopPropagation(); el('exportMenu').classList.toggle('open'); }
  function closeExportMenu(): void { const m = root.querySelector('[data-el="exportMenu"]'); if (m) { m.classList.remove('open'); } }

  // ---- auto refresh --------------------------------------------------------
  function toggleAuto(): void {
    const btn = root.querySelector('[data-action="auto"]') as HTMLElement;
    if (autoTimer) {
      clearInterval(autoTimer); autoTimer = null;
      btn.classList.remove('active'); btn.textContent = '◷ Auto';
    } else {
      const secs = opts.autoRefreshSeconds > 0 ? opts.autoRefreshSeconds : 120;
      autoTimer = setInterval(loadData, secs * 1000);
      btn.classList.add('active'); btn.textContent = '◷ Auto on';
    }
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
      else if (action === 'exportXlsx') { exportXLSX(); }
    });
  });
  // clicking anywhere else closes the export menu
  root.addEventListener('click', (e) => {
    const t = e.target as HTMLElement;
    if (!t.closest || !t.closest('.dropdown')) { closeExportMenu(); }
  });

  // initial load
  loadData().catch(() => undefined);
}
