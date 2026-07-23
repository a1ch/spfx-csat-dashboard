import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

// Chart.js is bundled (chart.js/auto auto-registers all controllers) and
// attached to window so the core can use it as a global — no runtime CDN load.
// Namespace import + .default fallback keeps this robust to esModuleInterop.
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as ChartAuto from 'chart.js/auto';

import * as strings from 'OpsDashboardWebPartStrings';
import { OPS_DASHBOARD_HTML } from './opsDashboardTemplate';
import { initOpsDashboard, IOpsController } from './opsDashboardCore';
import { fetchOpsData } from './OpsDataService';
import { fetchNotes, saveNote } from './OpsNotesService';

export interface IOpsDashboardWebPartProps {
  listSiteUrl: string;
  workbookUrl: string;
  notesListName: string;
}

export default class OpsDashboardWebPart extends BaseClientSideWebPart<IOpsDashboardWebPartProps> {

  private _dashboard: IOpsController | undefined;
  private _builtKey: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    (window as any).Chart = (ChartAuto as any).default || ChartAuto;
  }

  public render(): void {
    const siteUrl: string = (this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl).trim();
    const workbookUrl: string = (this.properties.workbookUrl || '').trim();
    const notesListName: string = (this.properties.notesListName || 'FY27 Ops Notes').trim();

    // SharePoint calls render() often; only rebuild when config changed.
    const key: string = `${siteUrl}|${workbookUrl}|${notesListName}`;
    const hasRoot: boolean = !!this.domElement.querySelector('.opsRoot');
    if (hasRoot && this._dashboard && key === this._builtKey) { return; }
    if (this._dashboard) { this._dashboard.destroy(); this._dashboard = undefined; }
    this._builtKey = key;

    this.domElement.innerHTML = OPS_DASHBOARD_HTML;
    const root: HTMLElement = this.domElement.querySelector('.opsRoot') as HTMLElement;
    if (!root) { return; }

    this._dashboard = initOpsDashboard(root, {
      fetchData: () => fetchOpsData(this.context.spHttpClient, siteUrl, workbookUrl),
      fetchNotes: () => fetchNotes(this.context.spHttpClient, siteUrl, notesListName),
      saveNote: (b, m, mo, note, id) => saveNote(this.context.spHttpClient, siteUrl, notesListName, b, m, mo, note, id)
    });
  }

  protected onDispose(): void {
    if (this._dashboard) { this._dashboard.destroy(); this._dashboard = undefined; }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('workbookUrl', {
                  label: strings.WorkbookUrlFieldLabel,
                  description: 'Full or server-relative URL of the FY27 metrics .xlsx in a document library (e.g. https://…/sites/…/Shared Documents/FY27_Master_Metrics.xlsx). Replace this file to update the dashboard.',
                  multiline: true
                }),
                PropertyPaneTextField('listSiteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  description: 'Site that holds the workbook library and the notes list. Leave blank to use the current site.'
                }),
                PropertyPaneTextField('notesListName', {
                  label: strings.NotesListFieldLabel,
                  description: 'Shared list where per-branch notes are saved (columns: Branch, Metric, Month, Note).'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
