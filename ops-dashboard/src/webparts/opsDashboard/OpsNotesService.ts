import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

// A note keyed by Branch | Metric | Month. `id` is the SharePoint list item id
// (present once the note has been saved at least once).
export interface IOpsNote {
  id: number | null;
  note: string;
}

export function noteKey(branch: string, metric: string, month: string): string {
  return `${branch}|${metric}|${month}`;
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function str(v: any): string { return (v === null || v === undefined) ? '' : String(v); }

/**
 * Loads all notes from the shared notes list into a { key -> {id,note} } map.
 * Runs under the signed-in user's context. Returns an empty map if the list is
 * missing/unreadable so the dashboard still renders (notes just won't persist).
 */
export async function fetchNotes(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listName: string
): Promise<{ [key: string]: IOpsNote }> {
  const site: string = siteUrl.replace(/\/$/, '');
  const map: { [key: string]: IOpsNote } = {};
  if (!listName) { return map; }

  let url: string =
    `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items` +
    `?$select=Id,Branch,Metric,Month,Note&$top=2000`;

  try {
    while (url) {
      const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
        headers: { 'Accept': 'application/json;odata=nometadata' }
      });
      if (!res.ok) { return map; }
      const data: any = await res.json();
      const items: any[] = Array.isArray(data.value) ? data.value : [];
      items.forEach((it) => {
        const key: string = noteKey(str(it.Branch), str(it.Metric), str(it.Month));
        map[key] = { id: Number(it.Id), note: str(it.Note) };
      });
      url = data['odata.nextLink'] || data['@odata.nextLink'] || '';
    }
  } catch (e) {
    /* eslint-disable-next-line no-console */
    console.error('Ops dashboard: could not load notes', e);
  }
  return map;
}

/**
 * Creates or updates one note in the list. Returns the (possibly new) item id.
 * Uses SPHttpClient, which supplies the request digest automatically.
 */
export async function saveNote(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listName: string,
  branch: string,
  metric: string,
  month: string,
  note: string,
  existingId: number | null
): Promise<number | null> {
  const site: string = siteUrl.replace(/\/$/, '');
  const base: string = `${site}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items`;
  const body: string = JSON.stringify({ Branch: branch, Metric: metric, Month: month, Note: note });

  if (existingId) {
    const res: SPHttpClientResponse = await spHttpClient.post(
      `${base}(${existingId})`, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body
      });
    if (!res.ok) { throw new Error(`Save failed (${res.status}).`); }
    return existingId;
  }

  const res: SPHttpClientResponse = await spHttpClient.post(
    base, SPHttpClient.configurations.v1, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata'
      },
      body
    });
  if (!res.ok) { throw new Error(`Save failed (${res.status}).`); }
  const data: any = await res.json().catch(() => ({}));
  return data && data.Id ? Number(data.Id) : null;
}
