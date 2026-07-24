import { IAssetDataset, IAssetRow, IAssetForm } from './AssetFormsService';
import { IGlEntry, glKey } from './AssetGlService';

/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ExcelJS: any = require('exceljs/dist/exceljs.min.js');

export const ASSET_DASHBOARD_VERSION: string = '1.0.4 · 2026-07-24';

export interface IAssetDashboardOptions {
  fetchData: () => Promise<IAssetDataset>;
  fetchGl: () => Promise<{ [key: string]: IGlEntry }>;
  saveGl: (emp: string, item: string, serial: string, gl: string, existingId: number | null) => Promise<number | null>;
  uploadUrl?: string;
}
export interface IAssetController { destroy: () => void; }

const NAVY: string = 'FF1F3864';
const BLUE: string = 'FF2E5496';
const LIGHT: string = 'FFD9E1F2';

export function initAssetDashboard(root: HTMLElement, opts: IAssetDashboardOptions): IAssetController {
  const el = (n: string): HTMLElement => root.querySelector(`[data-el="${n}"]`) as HTMLElement;

  let data: IAssetDataset = { forms: [], rows: [], totalBeforeTax: 0, skipped: [], superseded: [] };
  let gl: { [key: string]: IGlEntry } = {};

  const verEl = root.querySelector('[data-el="appVersion"]');
  if (verEl) { verEl.textContent = 'v' + ASSET_DASHBOARD_VERSION; }

  // Classic SharePoint pages wrap the web part in an ASP.NET <form>. Pressing
  // Enter in a text field would submit it and reload the page, so swallow Enter
  // on our inputs (change/blur handlers still fire).
  root.addEventListener('keydown', (e) => {
    const t = e.target as HTMLElement;
    if ((e as KeyboardEvent).key === 'Enter' && t && (t.tagName === 'INPUT' || t.tagName === 'SELECT')) {
      e.preventDefault();
      (t as HTMLElement).blur();
    }
  });

  // ---- helpers -------------------------------------------------------------
  function esc(s: unknown): string {
    return ('' + s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string));
  }
  function money(v: number): string {
    const neg: boolean = v < 0; const a: number = Math.abs(v);
    return (neg ? '-$' : '$') + a.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }
  function uniq(vals: string[]): string[] {
    return Array.from(new Set(vals.filter((v) => v !== null && v !== undefined && v !== ''))).sort();
  }
  function glFor(r: IAssetRow): string {
    const k: string = glKey(r.emp, r.item, r.serial);
    return gl[k] ? gl[k].gl : (r.cc || '');
  }

  // ---- filtering -----------------------------------------------------------
  function filters(): { emp: string; dept: string; loc: string; cls: string; year: string; q: string } {
    return {
      emp: (el('fEmp') as HTMLSelectElement).value,
      dept: (el('fDept') as HTMLSelectElement).value,
      loc: (el('fLoc') as HTMLSelectElement).value,
      cls: (el('fCls') as HTMLSelectElement).value,
      year: (el('fYear') as HTMLSelectElement).value,
      q: ((el('fQ') as HTMLInputElement).value || '').trim().toLowerCase()
    };
  }
  function rows(): IAssetRow[] {
    const f = filters();
    return data.rows.filter((r) => {
      if (f.emp && r.emp !== f.emp) { return false; }
      if (f.dept && r.dept !== f.dept) { return false; }
      if (f.loc && r.loc !== f.loc) { return false; }
      if (f.cls && r.cls !== f.cls) { return false; }
      if (f.year && r.date.slice(0, 4) !== f.year) { return false; }
      if (f.q) {
        const hay: string = [r.item, r.serial, r.po, r.ticket, r.emp, r.dept, r.loc, r.notes, glFor(r)].join(' ').toLowerCase();
        if (hay.indexOf(f.q) < 0) { return false; }
      }
      return true;
    });
  }
  function formsFiltered(): IAssetForm[] {
    const keep: { [fid: number]: boolean } = {};
    rows().forEach((r) => { keep[r.fid] = true; });
    return data.forms.filter((_, i) => keep[i]);
  }

  // ---- rollups -------------------------------------------------------------
  function rollup(rs: IAssetRow[], key: (r: IAssetRow) => string): { k: string; n: number; v: number }[] {
    const m: { [k: string]: { n: number; v: number } } = {};
    rs.forEach((r) => {
      const k: string = key(r) || '(blank)';
      if (!m[k]) { m[k] = { n: 0, v: 0 }; }
      m[k].n += r.qty; m[k].v += r.total;
    });
    return Object.keys(m).map((k) => ({ k, n: m[k].n, v: Math.round(m[k].v * 100) / 100 }))
      .sort((a, b) => b.v - a.v);
  }
  function barTable(title: string, list: { k: string; n: number; v: number }[]): string {
    const max: number = list.reduce((s, x) => Math.max(s, x.v), 0) || 1;
    let h = `<div class="panel"><div class="panel-title">${esc(title)}</div><div class="tbl-wrap"><table><thead><tr><th>Name</th><th class="num">Items</th><th class="num">Value</th><th style="width:120px"></th></tr></thead><tbody>`;
    list.forEach((x) => {
      h += `<tr><td class="fw">${esc(x.k)}</td><td class="num">${x.n}</td><td class="num">${money(x.v)}</td><td><div class="bar-track"><div class="bar-fill" style="width:${Math.round(x.v / max * 100)}%"></div></div></td></tr>`;
    });
    return h + `</tbody></table></div></div>`;
  }

  // ---- views ---------------------------------------------------------------
  function renderDash(): void {
    const rs: IAssetRow[] = rows();
    const fs: IAssetForm[] = formsFiltered();
    const total: number = Math.round(rs.reduce((s, r) => s + r.total, 0) * 100) / 100;
    const items: number = rs.reduce((s, r) => s + r.qty, 0);
    const emps: number = uniq(rs.map((r) => r.emp)).length;

    let h = `<div class="kpi-grid">
      <div class="kpi"><div class="kpi-label">Total value</div><div class="kpi-value">${money(total)}</div><div class="kpi-sub">before tax</div></div>
      <div class="kpi"><div class="kpi-label">Line items</div><div class="kpi-value">${rs.length}</div><div class="kpi-sub">${items} unit(s)</div></div>
      <div class="kpi"><div class="kpi-label">Employees</div><div class="kpi-value">${emps}</div><div class="kpi-sub">with equipment</div></div>
      <div class="kpi"><div class="kpi-label">Forms</div><div class="kpi-value">${fs.length}</div><div class="kpi-sub">allocation forms</div></div>
    </div>`;

    if (data.skipped.length || data.superseded.length) {
      const bits: string[] = [];
      if (data.superseded.length) { bits.push(`${data.superseded.length} form(s) superseded by a newer upload`); }
      if (data.skipped.length) { bits.push(`${data.skipped.length} file(s) skipped: ` + data.skipped.map((s) => `${s.file} (${s.reason})`).join('; ')); }
      h += `<div class="warn">${esc(bits.join(' · '))}</div>`;
    }

    h += `<div class="grid2">${barTable('By department', rollup(rs, (r) => r.dept))}${barTable('By location', rollup(rs, (r) => r.loc))}</div>`;
    h += `<div class="grid2">${barTable('By class', rollup(rs, (r) => r.cls))}${barTable('By request type', rollup(rs, (r) => r.type))}</div>`;
    el('view-dash').innerHTML = h;
  }

  function renderItems(): void {
    const rs: IAssetRow[] = rows();
    let h = `<div class="panel"><div class="panel-title">Line items (${rs.length})</div><div class="tbl-wrap"><table><thead><tr>
      <th>Date</th><th>Employee</th><th>Department</th><th>Location</th><th>Item</th><th>Serial</th>
      <th>Class</th><th>PO #</th><th class="num">Unit</th><th class="num">Qty</th><th class="num">Total</th><th>GL / Cost Center</th></tr></thead><tbody>`;
    rs.forEach((r) => {
      const k: string = glKey(r.emp, r.item, r.serial);
      h += `<tr><td class="muted">${esc(r.date)}</td><td class="fw">${esc(r.emp)}</td><td class="muted">${esc(r.dept)}</td><td class="muted">${esc(r.loc)}</td>
        <td>${esc(r.item)}</td><td class="muted">${esc(r.serial)}</td><td><span class="pill">${esc(r.cls)}</span></td><td class="muted">${esc(r.po)}</td>
        <td class="num">${money(r.price)}</td><td class="num">${r.qty}</td><td class="num fw">${money(r.total)}</td>
        <td><input class="gl" data-gl="${esc(k)}" data-emp="${esc(r.emp)}" data-item="${esc(r.item)}" data-serial="${esc(r.serial)}" value="${esc(glFor(r))}" placeholder="add code"></td></tr>`;
    });
    h += `</tbody></table></div></div>`;
    el('view-items').innerHTML = h;
    wireGlInputs();
  }

  function renderEmps(): void {
    const rs: IAssetRow[] = rows();
    const list = rollup(rs, (r) => r.emp);
    let h = `<div class="panel"><div class="panel-title">Employees (${list.length})</div><div class="tbl-wrap"><table><thead><tr>
      <th>Employee</th><th>Job title</th><th>Department</th><th>Location</th><th class="num">Items</th><th class="num">Value</th></tr></thead><tbody>`;
    list.forEach((x) => {
      const first = rs.filter((r) => r.emp === x.k)[0];
      h += `<tr><td class="fw">${esc(x.k)}</td><td class="muted">${esc(first ? first.title : '')}</td><td class="muted">${esc(first ? first.dept : '')}</td><td class="muted">${esc(first ? first.loc : '')}</td><td class="num">${x.n}</td><td class="num fw">${money(x.v)}</td></tr>`;
    });
    h += `</tbody></table></div></div>`;
    el('view-emps').innerHTML = h;
  }

  function renderForms(): void {
    const rs: IAssetRow[] = rows();
    const fs: IAssetForm[] = formsFiltered();
    let grand: number = 0;
    let h = `<div class="panel"><div class="panel-title">Allocation forms (${fs.length})</div><div class="tbl-wrap"><table><thead><tr>
      <th>Employee</th><th>Ticket #</th><th>Date</th><th>Type</th><th>PO #(s)</th><th class="num">Subtotal</th><th class="num">Tax</th><th class="num">Grand total</th><th>Source file</th></tr></thead><tbody>`;
    data.forms.forEach((f, i) => {
      if (fs.indexOf(f) < 0) { return; }
      const sub: number = Math.round(rs.filter((r) => r.fid === i).reduce((s, r) => s + r.total, 0) * 100) / 100;
      const tax: number = Math.round(sub * f.rate * 100) / 100;
      grand += sub + tax;
      const pos: string = uniq(f.items.map((it) => it.po)).join(', ');
      h += `<tr><td class="fw">${esc(f.emp)}</td><td class="muted">${esc(f.ticket)}</td><td class="muted">${esc(f.date.toISOString().slice(0, 10))}</td>
        <td class="muted">${esc(f.type)}</td><td class="muted">${esc(pos)}</td><td class="num">${money(sub)}</td>
        <td class="num muted">${f.rate ? money(tax) + ' (' + (f.rate * 100).toFixed(2) + '%)' : '—'}</td>
        <td class="num fw">${money(sub + tax)}</td><td class="muted">${esc(f.file)}</td></tr>`;
    });
    h += `</tbody><tfoot><tr><td colspan="7" class="fw">TOTAL charged (with tax)</td><td class="num fw">${money(Math.round(grand * 100) / 100)}</td><td></td></tr></tfoot></table></div></div>`;
    el('view-forms').innerHTML = h;
  }

  function renderAll(): void {
    const rs = rows();
    el('fCount').textContent = `${rs.length} of ${data.rows.length} items`;
    renderDash(); renderItems(); renderEmps(); renderForms();
  }

  // ---- GL editing ----------------------------------------------------------
  function wireGlInputs(): void {
    root.querySelectorAll('input.gl').forEach((node) => {
      node.addEventListener('change', (e) => {
        const t = e.target as HTMLInputElement;
        const k: string = t.getAttribute('data-gl') as string;
        const emp: string = t.getAttribute('data-emp') as string;
        const item: string = t.getAttribute('data-item') as string;
        const serial: string = t.getAttribute('data-serial') as string;
        const val: string = t.value.trim();
        const existing: number | null = gl[k] ? gl[k].id : null;
        t.disabled = true;
        t.style.borderColor = '';
        opts.saveGl(emp, item, serial, val, existing)
          .then((id) => { gl[k] = { id: id, gl: val }; t.disabled = false; t.style.borderColor = '#0F6E56'; })
          .catch((err) => {
            t.disabled = false;
            const msg: string = (err && (err as Error).message) ? (err as Error).message : ('' + err);
            /* eslint-disable-next-line no-console */
            console.error('Asset dashboard: GL save failed —', msg);
            t.style.borderColor = '#A32D2D';
            t.title = msg;   // hover the red field to see the exact error
            if (!(window as any).__astGlAlerted) {
              (window as any).__astGlAlerted = true;
              alert('Could not save GL / Cost Center:\n\n' + msg +
                '\n\nCheck that the GL list exists on this site with columns Employee, Item, Serial, GL, and that you have edit access.');
            }
          });
      });
    });
  }

  // ---- Excel export --------------------------------------------------------
  function hdr(ws: any, rowIdx: number, names: string[]): void {
    const r = ws.getRow(rowIdx);
    names.forEach((n, i) => {
      const c = r.getCell(i + 1);
      c.value = n;
      c.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    r.height = 26;
  }
  async function exportExcel(btn: HTMLButtonElement): Promise<void> {
    const orig = btn.textContent; btn.disabled = true; btn.textContent = 'Building…';
    try {
      const rs: IAssetRow[] = rows();
      const wb: any = new ExcelJS.Workbook();
      wb.creator = 'Stream-Flo IT Equipment Dashboard';
      wb.created = new Date();

      // Assets
      const ws: any = wb.addWorksheet('Assets', { views: [{ showGridLines: false, state: 'frozen', ySplit: 3 }] });
      ws.mergeCells(1, 1, 1, 18);
      const t = ws.getCell(1, 1);
      t.value = 'Stream-Flo Industries — IT Equipment Asset Log  (prices before tax)';
      t.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
      t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      t.alignment = { vertical: 'middle', indent: 1 };
      ws.getRow(1).height = 26;
      const cols = ['Date Issued', 'Employee', 'Job Title', 'Department', 'Location', 'Ticket #', 'PO #',
        'Request Type', 'Item / Description', 'Serial / Part #', 'Class', 'Unit Price', 'Qty',
        'Line Total', 'GL / Cost Center', 'Requester', 'Approved By', 'Notes'];
      hdr(ws, 3, cols);
      rs.forEach((r, i) => {
        const row = ws.getRow(4 + i);
        [r.date, r.emp, r.title, r.dept, r.loc, r.ticket, r.po, r.type, r.item, r.serial, r.cls,
         r.price, r.qty, r.total, glFor(r), r.req, r.appr, r.notes]
          .forEach((v, ci) => { row.getCell(ci + 1).value = (v === undefined ? '' : v as any); });
        row.getCell(12).numFmt = '$#,##0.00';
        row.getCell(14).numFmt = '$#,##0.00';
      });
      [11, 18, 22, 20, 12, 14, 12, 12, 32, 13, 11, 11, 6, 11, 13, 16, 15, 46]
        .forEach((w, i) => { ws.getColumn(i + 1).width = w; });
      if (rs.length) { ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3, column: cols.length } }; }

      // Forms
      const wf: any = wb.addWorksheet('Forms', { views: [{ showGridLines: false }] });
      wf.mergeCells(1, 1, 1, 9);
      const ft = wf.getCell(1, 1);
      ft.value = 'Submitted Forms — one row per allocation form (what Accounting is charged)';
      ft.font = { bold: true, size: 13, color: { argb: 'FFFFFFFF' } };
      ft.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      ft.alignment = { vertical: 'middle', indent: 1 };
      wf.getRow(1).height = 24;
      hdr(wf, 3, ['Employee', 'Ticket #', 'Date', 'Request Type', 'PO #(s)', 'Subtotal (before tax)', 'Tax Rate', 'Sales Tax', 'Grand Total']);
      let fr: number = 3;
      data.forms.forEach((f, i) => {
        const sub: number = Math.round(rs.filter((r) => r.fid === i).reduce((s, r) => s + r.total, 0) * 100) / 100;
        if (!sub && rs.filter((r) => r.fid === i).length === 0) { return; }
        fr++;
        const row = wf.getRow(fr);
        const tax: number = Math.round(sub * f.rate * 100) / 100;
        [f.emp, f.ticket, f.date, f.type, uniq(f.items.map((it) => it.po)).join(', '), sub, f.rate, tax, sub + tax]
          .forEach((v, ci) => { row.getCell(ci + 1).value = v as any; });
        row.getCell(3).numFmt = 'mm/dd/yyyy';
        row.getCell(6).numFmt = '$#,##0.00';
        row.getCell(7).numFmt = '0.00%';
        row.getCell(8).numFmt = '$#,##0.00';
        row.getCell(9).numFmt = '$#,##0.00';
      });
      const tot = wf.getRow(fr + 1);
      tot.getCell(1).value = 'TOTAL'; tot.getCell(1).font = { bold: true };
      [6, 8, 9].forEach((c) => {
        const cell = tot.getCell(c);
        cell.value = { formula: `SUM(${String.fromCharCode(64 + c)}4:${String.fromCharCode(64 + c)}${fr})` };
        cell.numFmt = '$#,##0.00'; cell.font = { bold: true };
      });
      [20, 16, 12, 14, 26, 18, 10, 13, 14].forEach((w, i) => { wf.getColumn(i + 1).width = w; });

      // Summary
      const wsum: any = wb.addWorksheet('Summary', { views: [{ showGridLines: false }] });
      wsum.mergeCells(1, 1, 1, 3);
      const st = wsum.getCell(1, 1);
      st.value = 'Equipment Charge Summary — Stream-Flo IT';
      st.font = { bold: true, size: 14, color: { argb: 'FFFFFFFF' } };
      st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      st.alignment = { vertical: 'middle', indent: 1 };
      wsum.getRow(1).height = 28;
      let sr: number = 3;
      const sections: { title: string; list: { k: string; n: number; v: number }[] }[] = [
        { title: 'By Department', list: rollup(rs, (r) => r.dept) },
        { title: 'By Location', list: rollup(rs, (r) => r.loc) },
        { title: 'By Class', list: rollup(rs, (r) => r.cls) },
        { title: 'By Employee', list: rollup(rs, (r) => r.emp) }
      ];
      sections.forEach((sec) => {
        wsum.mergeCells(sr, 1, sr, 3);
        const h2 = wsum.getCell(sr, 1);
        h2.value = sec.title; h2.font = { bold: true, size: 11, color: { argb: 'FFFFFFFF' } };
        h2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
        h2.alignment = { vertical: 'middle', indent: 1 };
        sr++;
        ['Name', '# Items', 'Value before tax'].forEach((n, i) => {
          const c = wsum.getCell(sr, i + 1);
          c.value = n; c.font = { bold: true };
          c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } };
        });
        sr++;
        sec.list.forEach((x) => {
          wsum.getCell(sr, 1).value = x.k;
          wsum.getCell(sr, 2).value = x.n;
          const cv = wsum.getCell(sr, 3); cv.value = x.v; cv.numFmt = '$#,##0.00';
          sr++;
        });
        const tr = wsum.getRow(sr);
        tr.getCell(1).value = 'Total'; tr.getCell(1).font = { bold: true };
        tr.getCell(2).value = sec.list.reduce((s, x) => s + x.n, 0); tr.getCell(2).font = { bold: true };
        const tv = tr.getCell(3);
        tv.value = Math.round(sec.list.reduce((s, x) => s + x.v, 0) * 100) / 100;
        tv.numFmt = '$#,##0.00'; tv.font = { bold: true };
        sr += 2;
      });
      wsum.getColumn(1).width = 30; wsum.getColumn(2).width = 12; wsum.getColumn(3).width = 20;

      const buf: any = await wb.xlsx.writeBuffer();
      const n = new Date(); const p = (x: number): string => ('' + x).padStart(2, '0');
      const name = `Equipment-Asset-Log_${n.getFullYear()}${p(n.getMonth() + 1)}${p(n.getDate())}.xlsx`;
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = name;
      document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    } catch (e) {
      /* eslint-disable-next-line no-console */
      console.error(e);
      alert('Could not build the Excel log: ' + ((e as Error).message || e));
    } finally {
      btn.disabled = false; btn.textContent = orig;
    }
  }

  // ---- filters / tabs / boot ----------------------------------------------
  function populateFilters(): void {
    const set = (name: string, vals: string[], all: string): void => {
      const s = el(name) as HTMLSelectElement;
      const prev = s.value;
      s.innerHTML = `<option value="">${all}</option>`;
      vals.forEach((v) => { const o = document.createElement('option'); o.value = v; o.text = v; s.appendChild(o); });
      s.value = prev;
    };
    set('fEmp', uniq(data.rows.map((r) => r.emp)), 'All employees');
    set('fDept', uniq(data.rows.map((r) => r.dept)), 'All departments');
    set('fLoc', uniq(data.rows.map((r) => r.loc)), 'All locations');
    set('fCls', uniq(data.rows.map((r) => r.cls)), 'All classes');
    set('fYear', uniq(data.rows.map((r) => r.date.slice(0, 4))), 'All years');
  }

  function state(title: string, sub: string, detail?: string): void {
    el('view-dash').innerHTML = `<div class="panel"><div class="state"><div class="state-title">${esc(title)}</div><div>${esc(sub)}</div>${detail ? `<div class="state-detail">${esc(detail)}</div>` : ''}</div></div>`;
  }

  async function load(): Promise<void> {
    state('Loading…', 'Reading the allocation forms from SharePoint.');
    try {
      const [ds, glMap] = await Promise.all([opts.fetchData(), opts.fetchGl()]);
      data = ds; gl = glMap;
      el('headerSub').textContent =
        `${data.forms.length} forms · ${data.rows.length} line items · ${money(data.totalBeforeTax)} before tax`;
      populateFilters();
      renderAll();
    } catch (err) {
      const msg = (err && (err as Error).message) ? (err as Error).message : ('' + err);
      state('Could not load the asset data', 'Check the forms folder URL in the web part properties, then Refresh.', msg);
    }
  }

  ['fEmp', 'fDept', 'fLoc', 'fCls', 'fYear'].forEach((n) => {
    const node = el(n); if (node) { node.addEventListener('change', renderAll); }
  });
  const q = el('fQ'); if (q) { q.addEventListener('input', renderAll); }
  const reset = el('fReset');
  if (reset) {
    reset.addEventListener('click', () => {
      ['fEmp', 'fDept', 'fLoc', 'fCls', 'fYear'].forEach((n) => { (el(n) as HTMLSelectElement).value = ''; });
      (el('fQ') as HTMLInputElement).value = '';
      renderAll();
    });
  }
  root.querySelectorAll('.tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      root.querySelectorAll('.tab').forEach((t) => t.classList.remove('active'));
      btn.classList.add('active');
      root.querySelectorAll('.view').forEach((v) => v.classList.remove('active'));
      const v = el('view-' + btn.getAttribute('data-tab'));
      if (v) { v.classList.add('active'); }
    });
  });
  const up = root.querySelector('[data-el="uploadBtn"]') as HTMLAnchorElement;
  if (up) { if (opts.uploadUrl) { up.href = opts.uploadUrl; } else { up.style.display = 'none'; } }
  const rb = el('refreshBtn'); if (rb) { rb.addEventListener('click', () => { load().catch(() => undefined); }); }
  const eb = el('exportBtn');
  if (eb) { eb.addEventListener('click', () => { exportExcel(eb as HTMLButtonElement).catch(() => undefined); }); }

  load().catch(() => undefined);
  return { destroy(): void { /* no timers held */ } };
}
