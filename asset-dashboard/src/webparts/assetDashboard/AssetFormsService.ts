import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// ExcelJS is bundled (self-contained browser build) and used to read every
// allocation form. Port of update_dashboard.py's parse_form()/collect().
/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ExcelJS: any = require('exceljs/dist/exceljs.min.js');

export interface IAssetItem {
  item: string; po: string; serial: string; cls: string; price: number; qty: number;
}
export interface IAssetForm {
  file: string; mtime: number; emp: string; title: string; dept: string; loc: string;
  ticket: string; type: string; date: Date; rate: number; req: string; appr: string;
  cc: string; notes: string; items: IAssetItem[]; poPrimary: string; warns: string[];
}
export interface IAssetRow {
  date: string; dateObj: Date; emp: string; title: string; dept: string; loc: string;
  ticket: string; po: string; type: string; item: string; serial: string; cls: string;
  price: number; qty: number; total: number; rate: number; cc: string; req: string;
  appr: string; notes: string; fid: number; src: string;
}
export interface IAssetDataset {
  forms: IAssetForm[];
  rows: IAssetRow[];
  totalBeforeTax: number;
  skipped: { file: string; reason: string }[];
  superseded: { older: string; newer: string }[];
}

// ---- config mirrored from the Python updater -------------------------------
const DEPT_NORMALIZE: { [k: string]: string } = { 'service': 'SFUSA Service' };
const TITLE_NORMALIZE: { [k: string]: string } = { 'service cordinator': 'Service Coordinator' };

const LABELS: { [k: string]: string } = {
  'employee name': 'emp', 'job title': 'title', 'department': 'dept',
  'location / office': 'loc', 'location/office': 'loc', 'location': 'loc',
  'ticket #': 'ticket', 'ticket': 'ticket',
  'requester': 'req', 'approved by': 'appr',
  'approval date': 'adate', 'issue date': 'idate', 'purchase date': 'pdate',
  'request type': 'type', 'po #': 'po', 'purchase order #': 'po',
  'gl / cost center': 'cc', 'gl/cost center': 'cc', 'gl / cost center (required)': 'cc'
};
const TOTAL_PAT: RegExp = /subtotal|grand total|sales tax|total \(before|^total$|reason/i;

function isDate(v: any): boolean { return v instanceof Date && !isNaN(v.getTime()); }

// ExcelJS cell values: formulas come back as {formula,result}, rich text as
// {richText:[...]}, hyperlinks as {text,hyperlink}. Reduce to a plain value.
function cellVal(ws: any, r: number, c: number): any {
  const cell: any = ws.getCell(r, c);
  let v: any = cell ? cell.value : null;
  if (v === null || v === undefined) { return null; }
  if (typeof v === 'object') {
    if (v instanceof Date) { return v; }
    if (v.result !== undefined) { v = v.result; }
    else if (v.richText) { v = v.richText.map((t: any) => t.text).join(''); }
    else if (v.text !== undefined) { v = v.text; }
    else if (v.error !== undefined) { return null; }
    else { return null; }
  }
  if (typeof v === 'string' && v.trim() === '') { return null; }
  return v;
}

function clean(v: any): any {
  if (v === null || v === undefined) { return ''; }
  if (typeof v === 'number' && v === Math.floor(v)) { return String(v); }
  if (isDate(v)) { return v; }
  return String(v).trim();
}

function parseDateVal(v: any): Date | null {
  if (isDate(v)) { return new Date(v.getFullYear(), v.getMonth(), v.getDate()); }
  const s: string = String(v || '').trim();
  const m1 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m1) {
    let y: number = parseInt(m1[3], 10); if (y < 100) { y += 2000; }
    return new Date(y, parseInt(m1[1], 10) - 1, parseInt(m1[2], 10));
  }
  const m2 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m2) { return new Date(parseInt(m2[1], 10), parseInt(m2[2], 10) - 1, parseInt(m2[3], 10)); }
  return null;
}

function normLabel(s: any): string {
  return String(s).replace(/\s+/g, ' ').trim().replace(/:+$/, '').trim().toLowerCase();
}

function classify(item: string): string {
  const t: string = item.toLowerCase();
  const has = (words: string[]): boolean => words.some((w) => t.indexOf(w) >= 0);
  if (has(['warranty', 'support', 'premier', 'onsite'])) { return 'Warranty'; }
  if (has(['software', 'license', 'subscription'])) { return 'Software'; }
  if (has(['case', 'mouse', 'keyboard', 'webcam', 'headset', 'cable', 'adapter',
           'bag', 'stand', 'sleeve', 'm650', 'm605', 'mx keys'])) { return 'Peripheral'; }
  return 'Capital';
}

function inferType(notes: string): string {
  const t: string = (notes || '').toLowerCase();
  if (['new intern', 'new employee', 'new hire', 'new user'].some((w) => t.indexOf(w) >= 0)) { return 'New Hire'; }
  if (t.indexOf('replac') >= 0) { return 'Replacement'; }
  return '';
}

// Pick the worksheet that holds the equipment table (skip Instructions tabs).
function pickSheet(wb: any): any {
  let found: any = null;
  wb.eachSheet((ws: any) => {
    if (found) { return; }
    if (String(ws.name || '').toLowerCase().indexOf('instruction') >= 0) { return; }
    const maxR: number = Math.min(ws.rowCount || 0, 60);
    const maxC: number = Math.min(ws.columnCount || 0, 12);
    for (let r = 1; r <= maxR; r++) {
      for (let c = 1; c <= maxC; c++) {
        const v: any = cellVal(ws, r, c);
        if (typeof v === 'string' && v.toLowerCase().indexOf('item name / description') >= 0) { found = ws; return; }
      }
    }
  });
  return found;
}

/** Parses one allocation form. Returns null + reason when it cannot be read. */
export function parseForm(wb: any, fileName: string, mtime: number): { form: IAssetForm | null; reason: string } {
  const ws: any = pickSheet(wb);
  if (!ws) { return { form: null, reason: 'no equipment table found' }; }

  const maxR: number = Math.min(ws.rowCount || 0, 60);
  const maxC: number = ws.columnCount || 0;

  // header row that starts the item table
  let hdrRow: number = 0;
  for (let r = 1; r <= maxR && !hdrRow; r++) {
    for (let c = 1; c <= maxC; c++) {
      const v: any = cellVal(ws, r, c);
      if (typeof v === 'string' && v.toLowerCase().indexOf('item name / description') >= 0) { hdrRow = r; break; }
    }
  }
  if (!hdrRow) { return { form: null, reason: 'no item header row' }; }

  // form-level fields live ABOVE the item table; scanning below would mistake
  // item column headers (e.g. "PO #") for form fields.
  const fields: { [k: string]: any } = {};
  for (let r = 1; r < hdrRow; r++) {
    for (let c = 1; c <= maxC; c++) {
      const v: any = cellVal(ws, r, c);
      if (typeof v !== 'string') { continue; }
      const key: string = LABELS[normLabel(v)];
      if (!key || fields[key] !== undefined) { continue; }
      for (let dc = 1; dc <= 4; dc++) {
        const nv: any = cellVal(ws, r, c + dc);
        if (nv !== null && String(nv).trim() !== '') { fields[key] = clean(nv); break; }
      }
    }
  }

  // item table columns
  const colmap: { [k: string]: number } = {};
  for (let c = 1; c <= maxC; c++) {
    const h: any = cellVal(ws, hdrRow, c);
    if (typeof h !== 'string') { continue; }
    const hn: string = normLabel(h);
    if (hn.indexOf('item name') >= 0) { colmap.item = c; }
    else if (hn === 'po #') { colmap.po = c; }
    else if (hn.indexOf('serial') >= 0) { colmap.serial = c; }
    else if (hn === 'class') { colmap.cls = c; }
    else if (hn.indexOf('unit price') >= 0) { colmap.price = c; }
    else if (hn === 'qty') { colmap.qty = c; }
    else if (hn.indexOf('line total') >= 0) { colmap.total = c; }
  }
  if (!colmap.item || !colmap.price) {
    return { form: null, reason: 'item table missing Item/Unit Price columns' };
  }

  const items: IAssetItem[] = [];
  const warns: string[] = [];
  for (let r = hdrRow + 1; r < hdrRow + 30; r++) {
    // stop at the subtotal / tax / reason block
    let stop: boolean = false;
    for (let c = 1; c <= maxC; c++) {
      const v: any = cellVal(ws, r, c);
      if (typeof v === 'string' && TOTAL_PAT.test(v)) { stop = true; break; }
    }
    if (stop) { break; }

    const itemRaw: any = clean(cellVal(ws, r, colmap.item));
    if (!itemRaw || isDate(itemRaw)) { continue; }
    const item: string = String(itemRaw);
    const price: any = cellVal(ws, r, colmap.price);
    if (typeof price !== 'number') {
      warns.push(`row ${r}: item '${item}' has no unit price, skipped`);
      continue;
    }
    const qtyRaw: any = colmap.qty ? cellVal(ws, r, colmap.qty) : null;
    const qty: number = (typeof qtyRaw === 'number' && qtyRaw > 0) ? Math.floor(qtyRaw) : 1;
    const cached: any = colmap.total ? cellVal(ws, r, colmap.total) : null;
    if (typeof cached === 'number' && Math.abs(cached - price * qty) > 0.01) {
      warns.push(`row ${r}: '${item}' cached line total ${cached} != ${price}*${qty}`);
    }
    let serial: any = colmap.serial ? clean(cellVal(ws, r, colmap.serial)) : '';
    if (isDate(serial)) { serial = ''; }
    const clsCell: any = colmap.cls ? clean(cellVal(ws, r, colmap.cls)) : '';
    items.push({
      item,
      po: String(colmap.po ? clean(cellVal(ws, r, colmap.po)) : '') || String(fields.po || ''),
      serial: String(serial).trim(),
      cls: String(clsCell || '') || classify(item),
      price: Math.round(Number(price) * 100) / 100,
      qty
    });
  }
  if (!items.length) { return { form: null, reason: 'no line items found' }; }

  // sales tax rate: a 0<x<0.5 number on the Sales Tax row, else the % in its label
  let rate: number = 0;
  for (let r = 1; r <= maxR && !rate; r++) {
    for (let c = 1; c <= maxC; c++) {
      const v: any = cellVal(ws, r, c);
      if (typeof v === 'string' && v.toLowerCase().indexOf('sales tax') >= 0) {
        for (let c2 = 1; c2 <= maxC; c2++) {
          const rv: any = cellVal(ws, r, c2);
          if (typeof rv === 'number' && rv > 0 && rv < 0.5) { rate = rv; break; }
        }
        if (!rate) {
          const m = v.match(/([\d.]+)\s*%/);
          if (m) { rate = parseFloat(m[1]) / 100; }
        }
        break;
      }
    }
  }

  // reason / justification note
  let notes: string = '';
  for (let r = 1; r <= maxR && !notes; r++) {
    for (let c = 1; c <= maxC; c++) {
      const v: any = cellVal(ws, r, c);
      if (typeof v === 'string' && /REASON\s*\/\s*JUSTIFICATION/i.test(v)) {
        for (let r2 = r + 1; r2 <= r + 3 && !notes; r2++) {
          for (let c2 = 1; c2 <= maxC; c2++) {
            const nv: any = cellVal(ws, r2, c2);
            if (typeof nv === 'string' && nv.trim()) { notes = nv.trim(); break; }
          }
        }
        break;
      }
    }
  }

  const emp: string = String(fields.emp || '').trim();
  if (!emp) { return { form: null, reason: 'no employee name' }; }

  let dept: string = String(fields.dept || '').trim();
  dept = DEPT_NORMALIZE[dept.toLowerCase()] || dept;

  let dt: Date | null = null;
  ['idate', 'adate', 'pdate'].forEach((k) => {
    if (!dt && fields[k]) { dt = parseDateVal(fields[k]); }
  });
  if (!dt) { dt = new Date(mtime); warns.push('no date field found, used file modified date'); }

  const ftype: string = String(fields.type || '').trim() || inferType(notes);

  const poCounts: { [po: string]: number } = {};
  items.forEach((it) => { if (it.po) { poCounts[it.po] = (poCounts[it.po] || 0) + 1; } });
  let poPrimary: string = String(fields.po || '').trim();
  if (!poPrimary) {
    const keys = Object.keys(poCounts);
    poPrimary = keys.length ? keys.reduce((a, b) => poCounts[a] >= poCounts[b] ? a : b) : '';
  }

  const rawTitle: string = String(fields.title || '').trim();
  return {
    form: {
      file: fileName, mtime, emp,
      title: TITLE_NORMALIZE[rawTitle.toLowerCase()] || rawTitle,
      dept, loc: String(fields.loc || '').trim(),
      ticket: String(fields.ticket || '').trim(),
      type: ftype, date: dt as Date, rate,
      req: String(fields.req || '').trim(), appr: String(fields.appr || '').trim(),
      cc: String(fields.cc || '').trim(), notes,
      items, poPrimary, warns
    },
    reason: ''
  };
}

// A PO typed into the Ticket field does not identify a request.
function ticketId(f: IAssetForm): string {
  const t: string = f.ticket.trim();
  return /^45\d{8}$/.test(t) ? '' : t;
}
function sameRequest(a: IAssetForm, b: IAssetForm): boolean {
  if (a.emp.toLowerCase() !== b.emp.toLowerCase()) { return false; }
  const ta: string = ticketId(a), tb: string = ticketId(b);
  if (ta && tb) { return ta === tb; }
  return !!a.poPrimary && a.poPrimary === b.poPrimary;
}

function pad(n: number): string { return (n < 10 ? '0' : '') + n; }
function ymd(d: Date): string { return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; }

function serverRelative(u0: string): string {
  const u: string = (u0 || '').trim().replace(/\/$/, '');
  if (!u) { return ''; }
  if (/^https?:\/\//i.test(u)) {
    const idx: number = u.indexOf('/', u.indexOf('://') + 3);
    return idx >= 0 ? u.slice(idx) : u;
  }
  return u.charAt(0) === '/' ? u : '/' + u;
}

/**
 * Reads every .xlsx allocation form in a SharePoint folder, parses each one,
 * applies the "newer file supersedes the same request" rule, and flattens to
 * one row per line item. Mirrors update_dashboard.py's collect().
 */
export async function fetchAssetData(
  spHttpClient: SPHttpClient, siteUrl: string, folderUrl: string
): Promise<IAssetDataset> {
  const site: string = siteUrl.replace(/\/$/, '');
  const rel: string = serverRelative(folderUrl);
  if (!rel) { throw new Error('No forms folder configured. Set the forms folder URL in the web part properties.'); }

  const listUrl: string =
    `${site}/_api/web/getfolderbyserverrelativeurl('${rel.replace(/'/g, "''")}')/Files` +
    `?$select=Name,ServerRelativeUrl,TimeLastModified&$top=500`;
  const listRes: SPHttpClientResponse = await spHttpClient.get(listUrl, SPHttpClient.configurations.v1, {
    headers: { 'Accept': 'application/json;odata=nometadata' }
  });
  if (!listRes.ok) {
    throw new Error(`Could not read the forms folder (${listRes.status}). Check the folder path and your access.`);
  }
  const listData: any = await listRes.json();
  const files: any[] = (Array.isArray(listData.value) ? listData.value : [])
    .filter((f: any) => /\.xlsx$/i.test(f.Name) && String(f.Name).indexOf('~$') !== 0)
    .sort((a: any, b: any) => String(a.Name).localeCompare(String(b.Name)));
  if (!files.length) { throw new Error('No .xlsx forms found in that folder.'); }

  const parsed: IAssetForm[] = [];
  const skipped: { file: string; reason: string }[] = [];

  for (const f of files) {
    try {
      const res: SPHttpClientResponse = await spHttpClient.get(
        `${site}/_api/web/getfilebyserverrelativeurl('${String(f.ServerRelativeUrl).replace(/'/g, "''")}')/$value`,
        SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/octet-stream' } });
      if (!res.ok) { skipped.push({ file: f.Name, reason: `download failed (${res.status})` }); continue; }
      const buf: ArrayBuffer = await res.arrayBuffer();
      const wb: any = new ExcelJS.Workbook();
      await wb.xlsx.load(buf);
      const mtime: number = new Date(f.TimeLastModified).getTime();
      const out = parseForm(wb, f.Name, mtime);
      if (!out.form) { skipped.push({ file: f.Name, reason: out.reason }); }
      else { parsed.push(out.form); }
    } catch (e) {
      skipped.push({ file: f.Name, reason: 'error: ' + ((e as Error).message || e) });
    }
  }

  // newer file wins for the same employee+request
  const kept: IAssetForm[] = [];
  const superseded: { older: string; newer: string }[] = [];
  parsed.slice().sort((a, b) => a.mtime - b.mtime).forEach((form) => {
    for (let i = 0; i < kept.length; i++) {
      if (sameRequest(kept[i], form)) {
        superseded.push({ older: kept[i].file, newer: form.file });
        kept[i] = form;
        return;
      }
    }
    kept.push(form);
  });
  const forms: IAssetForm[] = kept.sort((a, b) =>
    (a.date.getTime() - b.date.getTime()) || a.emp.localeCompare(b.emp));

  const rows: IAssetRow[] = [];
  forms.forEach((f, fi) => {
    f.items.forEach((it) => {
      rows.push({
        date: ymd(f.date), dateObj: f.date, emp: f.emp, title: f.title, dept: f.dept,
        loc: f.loc, ticket: f.ticket, po: it.po, type: f.type, item: it.item,
        serial: it.serial, cls: it.cls, price: it.price, qty: it.qty,
        total: Math.round(it.price * it.qty * 100) / 100,
        rate: f.rate, cc: f.cc, req: f.req, appr: f.appr, notes: f.notes,
        fid: fi, src: f.file
      });
    });
  });

  const totalBeforeTax: number = Math.round(rows.reduce((s, r) => s + r.total, 0) * 100) / 100;
  return { forms, rows, totalBeforeTax, skipped, superseded };
}
