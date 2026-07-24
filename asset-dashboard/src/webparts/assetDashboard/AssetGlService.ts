import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/* eslint-disable @typescript-eslint/no-explicit-any */

// A GL / Cost Center value Accounting typed in the dashboard. Keyed by
// employee + item + serial so it survives forms being re-uploaded — this
// replaces the Python updater's "carry values over from the old xlsx" trick.
export interface IGlEntry { id: number | null; gl: string; }

export function glKey(emp: string, item: string, serial: string): string {
  return `${(emp || '').toLowerCase()}|${(item || '').toLowerCase()}|${(serial || '').trim()}`;
}

function str(v: any): string { return (v === null || v === undefined) ? '' : String(v); }

interface IFieldMap { Employee: string; Item: string; Serial: string; GL: string; }

// SharePoint column internal names can differ from their display names (renames,
// spaces, etc.), which makes writes fail with "property 'X' does not exist".
// Resolve the real internal names from the list by matching display name, so the
// dashboard works regardless of how the columns were created. Cached per list.
const fieldCache: { [k: string]: IFieldMap } = {};

async function resolveFields(
  spHttpClient: SPHttpClient, site: string, listName: string
): Promise<IFieldMap> {
  const cacheKey: string = site + '|' + listName;
  if (fieldCache[cacheKey]) { return fieldCache[cacheKey]; }
  const fallback: IFieldMap = { Employee: 'Employee', Item: 'Item', Serial: 'Serial', GL: 'GL' };
  try {
    const url: string =
      `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/fields` +
      `?$select=Title,InternalName,Hidden,ReadOnlyField&$top=500`;
    const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });
    if (!res.ok) { return fallback; }
    const data: any = await res.json();
    const byDisplay: { [display: string]: string } = {};
    (Array.isArray(data.value) ? data.value : []).forEach((f: any) => {
      if (f && !f.Hidden && !f.ReadOnlyField && f.Title && f.InternalName) {
        byDisplay[String(f.Title).trim().toLowerCase()] = f.InternalName;
      }
    });
    const pick = (names: string[], def: string): string => {
      for (const n of names) { if (byDisplay[n.toLowerCase()]) { return byDisplay[n.toLowerCase()]; } }
      return def;
    };
    const map: IFieldMap = {
      Employee: pick(['Employee', 'Employee Name'], 'Employee'),
      Item: pick(['Item', 'Item / Description', 'Item Name'], 'Item'),
      Serial: pick(['Serial', 'Serial / Part #', 'Serial Number'], 'Serial'),
      GL: pick(['GL', 'GL / Cost Center', 'GL/Cost Center', 'Cost Center'], 'GL')
    };
    fieldCache[cacheKey] = map;
    return map;
  } catch (e) {
    return fallback;
  }
}

async function describeError(res: SPHttpClientResponse): Promise<Error> {
  const t: string = await res.text().catch(() => '');
  let detail: string = '';
  try {
    const j: any = JSON.parse(t);
    const m: any = (j && (j['odata.error'] || j.error) || {}).message;
    detail = (m && typeof m === 'object' && m.value) ? String(m.value) : (m ? String(m) : '');
  } catch (e) { detail = t.slice(0, 200); }
  return new Error(`Save failed (${res.status}) ${detail || ''}`.trim());
}

/** Loads every GL override into a { key -> {id,gl} } map. Never throws. */
export async function fetchGlOverrides(
  spHttpClient: SPHttpClient, siteUrl: string, listName: string
): Promise<{ [key: string]: IGlEntry }> {
  const site: string = siteUrl.replace(/\/$/, '');
  const map: { [key: string]: IGlEntry } = {};
  if (!listName) { return map; }

  try {
    const fields: IFieldMap = await resolveFields(spHttpClient, site, listName);
    const sel: string = ['Id', fields.Employee, fields.Item, fields.Serial, fields.GL]
      .filter((v, i, a) => a.indexOf(v) === i).join(',');
    let url: string =
      `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items?$select=${sel}&$top=2000`;
    while (url) {
      const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' }
      });
      if (!res.ok) { return map; }
      const data: any = await res.json();
      (Array.isArray(data.value) ? data.value : []).forEach((it: any) => {
        const key: string = glKey(str(it[fields.Employee]), str(it[fields.Item]), str(it[fields.Serial]));
        map[key] = { id: Number(it.Id), gl: str(it[fields.GL]) };
      });
      url = data['odata.nextLink'] || data['@odata.nextLink'] || '';
    }
  } catch (e) {
    /* eslint-disable-next-line no-console */
    console.error('Asset dashboard: could not load GL overrides', e);
  }
  return map;
}

/** Creates or updates one GL override, writing to the list's real internal
 *  field names. Returns the item id. */
export async function saveGlOverride(
  spHttpClient: SPHttpClient, siteUrl: string, listName: string,
  emp: string, itemName: string, serial: string, gl: string, existingId: number | null
): Promise<number | null> {
  const site: string = siteUrl.replace(/\/$/, '');
  const fields: IFieldMap = await resolveFields(spHttpClient, site, listName);
  const base: string = `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items`;

  const payload: any = { Title: `${emp} — ${itemName}`.slice(0, 255) };
  payload[fields.Employee] = emp;
  payload[fields.Item] = itemName;
  payload[fields.Serial] = serial;
  payload[fields.GL] = gl;
  const body: string = JSON.stringify(payload);

  const headers: any = {
    'Accept': 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata'
  };
  if (existingId) {
    const res: SPHttpClientResponse = await spHttpClient.post(
      `${base}(${existingId})`, SPHttpClient.configurations.v1,
      { headers: { ...headers, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }, body });
    if (!res.ok) { throw await describeError(res); }
    return existingId;
  }
  const res: SPHttpClientResponse = await spHttpClient.post(base, SPHttpClient.configurations.v1, { headers, body });
  if (!res.ok) { throw await describeError(res); }
  const data: any = await res.json().catch(() => ({}));
  return data && data.Id ? Number(data.Id) : null;
}
