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

/** Loads every GL override into a { key -> {id,gl} } map. Never throws: a
 *  missing/unreadable list just means no overrides. */
export async function fetchGlOverrides(
  spHttpClient: SPHttpClient, siteUrl: string, listName: string
): Promise<{ [key: string]: IGlEntry }> {
  const site: string = siteUrl.replace(/\/$/, '');
  const map: { [key: string]: IGlEntry } = {};
  if (!listName) { return map; }

  let url: string =
    `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items` +
    `?$select=Id,Employee,Item,Serial,GL&$top=2000`;
  try {
    while (url) {
      const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' }
      });
      if (!res.ok) { return map; }
      const data: any = await res.json();
      (Array.isArray(data.value) ? data.value : []).forEach((it: any) => {
        map[glKey(str(it.Employee), str(it.Item), str(it.Serial))] = { id: Number(it.Id), gl: str(it.GL) };
      });
      url = data['odata.nextLink'] || data['@odata.nextLink'] || '';
    }
  } catch (e) {
    /* eslint-disable-next-line no-console */
    console.error('Asset dashboard: could not load GL overrides', e);
  }
  return map;
}

/** Creates or updates one GL override. Returns the item id. */
export async function saveGlOverride(
  spHttpClient: SPHttpClient, siteUrl: string, listName: string,
  emp: string, item: string, serial: string, gl: string, existingId: number | null
): Promise<number | null> {
  const site: string = siteUrl.replace(/\/$/, '');
  const base: string = `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items`;
  const body: string = JSON.stringify({ Employee: emp, Item: item, Serial: serial, GL: gl });

  const headers: any = {
    'Accept': 'application/json;odata=nometadata',
    'Content-Type': 'application/json;odata=nometadata'
  };
  if (existingId) {
    const res: SPHttpClientResponse = await spHttpClient.post(
      `${base}(${existingId})`, SPHttpClient.configurations.v1,
      { headers: { ...headers, 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }, body });
    if (!res.ok) { throw new Error(`Save failed (${res.status}).`); }
    return existingId;
  }
  const res: SPHttpClientResponse = await spHttpClient.post(base, SPHttpClient.configurations.v1, { headers, body });
  if (!res.ok) { throw new Error(`Save failed (${res.status}).`); }
  const data: any = await res.json().catch(() => ({}));
  return data && data.Id ? Number(data.Id) : null;
}
