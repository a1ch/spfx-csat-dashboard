import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

/**
 * One normalized CSAT response row, using the simplified field names the
 * dashboard rendering logic expects. This is intentionally the same shape the
 * old Power Automate "read" flow used to return, so the dashboard code did not
 * have to change when we swapped the data source to a direct list read.
 */
export interface ICsatItem {
  id: number;
  Created: string;
  timestamp: string;
  serviceDate: string;
  branch: string;
  technician: string;
  company: string;
  contactName: string;
  contactTitle: string;
  contactInfo: string;
  rigWellName: string;
  location: string;
  workOrder: string;
  serviceType: string;
  serviceNotes: string;
  r_overall: number | null;
  r_technical: number | null;
  r_timeliness: number | null;
  r_communication: number | null;
  r_quality: number | null;
  r_professionalism: number | null;
  r_cleanliness: number | null;
  nps: number | null;
  safety: number | null;
  improvementAreas: string;
  exemplary: string;
  improveSuggestion: string;
  comments: string;
  avgRating: number | null;
}

// SharePoint internal column names we pull. Note "serviceNptes" is the actual
// (typo'd) internal name of the notes column in the CSAT RESPONSES list.
const SELECT_FIELDS: string[] = [
  'Id', 'Created', 'timestamp', 'serviceDate', 'branch', 'technician',
  'company', 'contactName', 'contactTitle', 'contactInfo', 'rigWellName',
  'location', 'workOrder', 'serviceType', 'serviceNptes',
  'r_overall', 'r_technical', 'r_timeliness', 'r_communication', 'r_quality',
  'r_professionalism', 'r_cleanliness', 'nps', 'safety',
  'improvementAreas', 'exemplary', 'improveSuggestion', 'comments', 'avgRating'
];

function num(v: unknown): number | null {
  if (v === null || v === undefined || v === '') { return null; }
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function str(v: unknown): string {
  return (v === null || v === undefined) ? '' : String(v);
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function mapRow(f: any): ICsatItem {
  return {
    id: num(f.Id) || 0,
    Created: str(f.Created),
    timestamp: str(f.timestamp),
    serviceDate: str(f.serviceDate),
    branch: str(f.branch),
    technician: str(f.technician),
    company: str(f.company),
    contactName: str(f.contactName),
    contactTitle: str(f.contactTitle),
    contactInfo: str(f.contactInfo),
    rigWellName: str(f.rigWellName),
    location: str(f.location),
    workOrder: str(f.workOrder),
    serviceType: str(f.serviceType),
    serviceNotes: str(f.serviceNptes), // typo'd column name -> clean name
    r_overall: num(f.r_overall),
    r_technical: num(f.r_technical),
    r_timeliness: num(f.r_timeliness),
    r_communication: num(f.r_communication),
    r_quality: num(f.r_quality),
    r_professionalism: num(f.r_professionalism),
    r_cleanliness: num(f.r_cleanliness),
    nps: num(f.nps),
    safety: num(f.safety),
    improvementAreas: str(f.improvementAreas),
    exemplary: str(f.exemplary),
    improveSuggestion: str(f.improveSuggestion),
    comments: str(f.comments),
    avgRating: num(f.avgRating)
  };
}
/* eslint-enable @typescript-eslint/no-explicit-any */

/**
 * Reads all items from the CSAT RESPONSES list via SharePoint REST, following
 * @odata.nextLink paging so it is not capped at a single page. Runs under the
 * signed-in user's context — no anonymous flow, no stored credentials.
 */
export async function fetchCsatItems(
  spHttpClient: SPHttpClient,
  siteUrl: string,
  listName: string
): Promise<ICsatItem[]> {
  const select: string = SELECT_FIELDS.join(',');
  const cleanSite: string = siteUrl.replace(/\/$/, '');
  let url: string =
    `${cleanSite}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')/items` +
    `?$select=${select}&$top=2000&$orderby=Created desc`;

  const items: ICsatItem[] = [];

  while (url) {
    const res: SPHttpClientResponse = await spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { 'Accept': 'application/json;odata=nometadata' }
    });

    if (!res.ok) {
      const body: string = await res.text().catch(() => '');
      throw new Error(`SharePoint returned ${res.status} ${res.statusText}. ${body.slice(0, 300)}`);
    }

    /* eslint-disable-next-line @typescript-eslint/no-explicit-any */
    const data: any = await res.json();
    const rows: unknown[] = Array.isArray(data.value) ? data.value : [];
    rows.forEach((r) => items.push(mapRow(r)));

    url = data['odata.nextLink'] || data['@odata.nextLink'] || '';
  }

  return items;
}
