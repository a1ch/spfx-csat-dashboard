import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// ExcelJS is bundled (self-contained browser build) — used here to READ the
// uploaded workbook. Bundling avoids any runtime CDN/AMD load problems.
/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ExcelJS: any = require('exceljs/dist/exceljs.min.js');

export type Direction = 'higher' | 'lower';

// One normalized fact: a single Branch × Metric × Month actual/target.
export interface IOpsRow {
  month: string;        // FY month name, e.g. 'Apr'
  monthIndex: number;   // 1..12 in fiscal order (Apr = 1)
  branch: string;
  metric: string;
  unit: string;         // 'rate' | '$' | 'count' | '%' (from the workbook)
  actual: number;
  target: number;
  direction: Direction;
}

export interface IOpsDataset {
  rows: IOpsRow[];
  branches: string[];               // distinct, sorted
  metrics: string[];                // canonical order, extras appended
  units: { [metric: string]: string };
  directions: { [metric: string]: Direction };
  monthsPresent: string[];          // months that appear in the data, FY order
}

const FACT_SHEET: string = 'Fact_MonthlyInputs';
// Fiscal year starts in April. Index 1..12.
const FY_MONTH_ORDER: string[] = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Jan', 'Feb', 'Mar'];
const METRIC_ORDER: string[] = ['TRIR', 'COPQ', 'CSTAT', 'Revenue', 'Operational Cost', 'Operational Profit'];

const CAL_MONTH_SHORT: string[] = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// A calendar month number (1=Jan) mapped to fiscal index (Apr=1 … Mar=12).
function fyIndexFromCalMonth(calMonth1to12: number): number {
  return ((calMonth1to12 - 4 + 12) % 12) + 1;
}

// ExcelJS cell values are messy: formula cells come back as {formula,result},
// dates as Date, rich text as {richText:[...]}. Reduce to a plain value.
function cellValue(cell: any): any {
  if (cell === null || cell === undefined) { return null; }
  const v: any = cell.value;
  if (v === null || v === undefined) { return null; }
  if (typeof v === 'object') {
    if (v instanceof Date) { return v; }
    if (v.result !== undefined) { return v.result; }              // formula → cached result
    if (v.richText) { return v.richText.map((t: any) => t.text).join(''); }
    if (v.text !== undefined) { return v.text; }
    if (v.error !== undefined) { return null; }
    return null;
  }
  return v;
}

function toNumber(v: any): number {
  if (v === null || v === undefined || v === '') { return 0; }
  if (typeof v === 'number') { return isNaN(v) ? 0 : v; }
  const n: number = Number(String(v).replace(/[$,%\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toDir(v: any): Direction {
  return String(v || '').toLowerCase().indexOf('lower') >= 0 ? 'lower' : 'higher';
}

// Derive an { Apr, 4 } style month from a PeriodStart cell (a Date, an Excel
// serial number, or a 'PeriodLabel' fallback like '01  Apr FY27').
function monthFromRow(periodStart: any, periodLabel: any): { name: string; index: number } | null {
  let d: Date | null = null;
  if (periodStart instanceof Date) {
    d = periodStart;
  } else if (typeof periodStart === 'number' && periodStart > 0) {
    // Excel serial date → JS date (Excel epoch 1899-12-30).
    d = new Date(Math.round((periodStart - 25569) * 86400 * 1000));
  }
  if (d && !isNaN(d.getTime())) {
    const name: string = CAL_MONTH_SHORT[d.getUTCMonth()];
    return { name, index: fyIndexFromCalMonth(d.getUTCMonth() + 1) };
  }
  // Fallback: parse a label such as '01  Apr FY27'
  const label: string = String(periodLabel || '');
  const m: RegExpMatchArray | null = label.match(/([A-Za-z]{3})/);
  if (m && FY_MONTH_ORDER.indexOf(m[1]) >= 0) {
    return { name: m[1], index: FY_MONTH_ORDER.indexOf(m[1]) + 1 };
  }
  return null;
}

function serverRelative(fileUrl: string): string {
  const u: string = (fileUrl || '').trim();
  if (!u) { return ''; }
  if (/^https?:\/\//i.test(u)) {
    const idx: number = u.indexOf('/', u.indexOf('://') + 3);
    return idx >= 0 ? u.slice(idx) : u;
  }
  return u.charAt(0) === '/' ? u : '/' + u;
}

/**
 * Downloads the workbook from a SharePoint document library (under the signed-in
 * user's permissions) and parses the Fact_MonthlyInputs sheet into normalized
 * rows. Columns are located by header NAME so the template and the filled export
 * (which has extra columns) both work.
 */
export async function fetchOpsData(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  fileUrl: string
): Promise<IOpsDataset> {
  const site: string = siteUrl.replace(/\/$/, '');
  const rel: string = serverRelative(fileUrl);
  if (!rel) { throw new Error('No Excel file path configured. Set the workbook URL in the web part properties.'); }

  const api: string =
    `${site}/_api/web/getfilebyserverrelativeurl('${rel.replace(/'/g, "''")}')/$value`;

  const res: SPHttpClientResponse = await spHttpClient.get(api, SPHttpClient.configurations.v1, {
    headers: { 'Accept': 'application/octet-stream' }
  });
  if (!res.ok) {
    const body: string = await res.text().catch(() => '');
    throw new Error(`Could not read the workbook (${res.status} ${res.statusText}). Check the file path and your access. ${body.slice(0, 200)}`);
  }

  const buf: ArrayBuffer = await res.arrayBuffer();
  const wb: any = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const ws: any = wb.getWorksheet(FACT_SHEET);
  if (!ws) { throw new Error(`The workbook has no "${FACT_SHEET}" sheet. Is this the FY27 metrics workbook?`); }

  // Header row → column index (1-based), matched case-insensitively.
  const headerRow: any = ws.getRow(1);
  const col: { [k: string]: number } = {};
  headerRow.eachCell((cell: any, c: number) => {
    const name: string = String(cellValue(cell) || '').trim().toLowerCase();
    if (name) { col[name] = c; }
  });
  const need: (name: string) => number = (name) => col[name.toLowerCase()] || 0;
  const cPeriod: number = need('PeriodStart');
  const cBranch: number = need('Branch');
  const cMetric: number = need('Metric');
  const cUnit: number = need('Unit');
  const cActual: number = need('ActualValue');
  const cTarget: number = need('TargetValue');
  const cDir: number = need('Direction');
  const cLabel: number = need('PeriodLabel');
  if (!cBranch || !cMetric || !cActual) {
    throw new Error('The Fact_MonthlyInputs sheet is missing expected columns (Branch / Metric / ActualValue).');
  }

  const rows: IOpsRow[] = [];
  const unitByMetric: { [m: string]: string } = {};
  const dirByMetric: { [m: string]: Direction } = {};

  const last: number = ws.rowCount;
  for (let r = 2; r <= last; r++) {
    const row: any = ws.getRow(r);
    const branch: string = String(cellValue(row.getCell(cBranch)) || '').trim();
    const metric: string = String(cellValue(row.getCell(cMetric)) || '').trim();
    if (!branch || !metric) { continue; }

    const mo = monthFromRow(cellValue(row.getCell(cPeriod)), cLabel ? cellValue(row.getCell(cLabel)) : null);
    if (!mo) { continue; }

    const unit: string = cUnit ? String(cellValue(row.getCell(cUnit)) || '').trim() : '';
    const direction: Direction = cDir ? toDir(cellValue(row.getCell(cDir))) : 'higher';

    rows.push({
      month: mo.name,
      monthIndex: mo.index,
      branch,
      metric,
      unit,
      actual: toNumber(cellValue(row.getCell(cActual))),
      target: cTarget ? toNumber(cellValue(row.getCell(cTarget))) : 0,
      direction
    });
    if (unit && !unitByMetric[metric]) { unitByMetric[metric] = unit; }
    if (!dirByMetric[metric]) { dirByMetric[metric] = direction; }
  }

  if (!rows.length) {
    throw new Error(`The "${FACT_SHEET}" sheet has no data rows yet.`);
  }

  const branches: string[] = Array.from(new Set(rows.map((r) => r.branch))).sort();
  const found: string[] = Array.from(new Set(rows.map((r) => r.metric)));
  const metrics: string[] = METRIC_ORDER.filter((m) => found.indexOf(m) >= 0)
    .concat(found.filter((m) => METRIC_ORDER.indexOf(m) < 0));
  const monthsPresent: string[] = FY_MONTH_ORDER.filter((m) => rows.some((r) => r.month === m));

  return { rows, branches, metrics, units: unitByMetric, directions: dirByMetric, monthsPresent };
}
