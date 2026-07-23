import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as strings from 'CsatDashboardWebPartStrings';
import { DASHBOARD_HTML } from './dashboardTemplate';
import { initDashboard, IDashboardController } from './dashboardCore';
import { fetchCsatItems, ICsatItem } from './CsatDataService';

export interface ICsatDashboardWebPartProps {
  listSiteUrl: string;
  listName: string;
  autoRefreshSeconds: number;
  chartJsUrl: string;
  excelJsUrl: string;
}

// Public CDN defaults. If a tenant blocks the CDN (common for the larger
// ExcelJS file), the URLs can be overridden in the property pane to point at a
// same-tenant copy uploaded to e.g. Site Assets.
const CHARTJS_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js';
const EXCELJS_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';

export default class CsatDashboardWebPart extends BaseClientSideWebPart<ICsatDashboardWebPartProps> {

  private _dashboard: IDashboardController | undefined;
  private _builtKey: string = '';

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Load Chart.js and ExcelJS once. Both are UMD bundles, and the SharePoint
    // page already has an AMD loader (define/require) present. Loaded as a plain
    // script, the UMD wrapper registers as an anonymous AMD module instead of
    // assigning window.Chart / window.ExcelJS — so the globals the dashboard
    // checks for stay undefined (blank charts, failed Excel export). Passing
    // globalExportsName makes SPComponentLoader suppress AMD during the load and
    // capture the real browser global.
    //
    // URLs default to the public CDN but can be overridden in the property pane
    // with a same-tenant copy — needed when the tenant blocks the CDN (the
    // 925 KB ExcelJS file is the usual casualty). Each loads independently so
    // one failing still leaves the other working.
    const chartUrl: string = (this.properties.chartJsUrl || '').trim() || CHARTJS_URL;
    const excelUrl: string = (this.properties.excelJsUrl || '').trim() || EXCELJS_URL;
    await Promise.all([
      SPComponentLoader.loadScript(chartUrl, { globalExportsName: 'Chart' })
        .catch((e: unknown) => {
          /* eslint-disable-next-line no-console */
          console.error('CSAT dashboard: Chart.js failed to load from ' + chartUrl, e);
        }),
      SPComponentLoader.loadScript(excelUrl, { globalExportsName: 'ExcelJS' })
        .catch((e: unknown) => {
          /* eslint-disable-next-line no-console */
          console.error('CSAT dashboard: ExcelJS failed to load from ' + excelUrl, e);
        })
    ]);
  }

  public render(): void {
    const siteUrl: string = (this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl).trim();
    const listName: string = (this.properties.listName || 'CSAT RESPONSES').trim();
    const autoRefreshSeconds: number = this.properties.autoRefreshSeconds || 0;

    // SharePoint calls render() often (selection, resize, layout, property
    // edits). Rebuilding the DOM + re-initialising the dashboard every time
    // caused visible flicker and stacked auto-refresh timers. Only rebuild when
    // the config that actually affects the dashboard changed; otherwise leave
    // the running instance (and its single refresh timer) untouched.
    const key: string = `${siteUrl}|${listName}|${autoRefreshSeconds}`;
    const hasRoot: boolean = !!this.domElement.querySelector('.sfCsatRoot');
    if (hasRoot && this._dashboard && key === this._builtKey) { return; }

    if (this._dashboard) { this._dashboard.destroy(); this._dashboard = undefined; }
    this._builtKey = key;

    this.domElement.innerHTML = DASHBOARD_HTML;

    const root: HTMLElement = this.domElement.querySelector('.sfCsatRoot') as HTMLElement;
    if (!root) { return; }

    // initDashboard arms auto-refresh itself (a single, non-stacking timer) when
    // autoRefreshSeconds > 0 — no programmatic button click needed.
    this._dashboard = initDashboard(root, {
      autoRefreshSeconds,
      fetchItems: (): Promise<ICsatItem[]> => fetchCsatItems(this.context.spHttpClient, siteUrl, listName)
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
                PropertyPaneTextField('listSiteUrl', {
                  label: strings.ListSiteUrlFieldLabel,
                  description: 'Full URL of the site that holds the list (e.g. https://…/sites/SUSTeam-SUSBranchOperations). Leave blank to use the current site.'
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                }),
                PropertyPaneSlider('autoRefreshSeconds', {
                  label: strings.AutoRefreshFieldLabel,
                  min: 0,
                  max: 600,
                  step: 30,
                  showValue: true
                })
              ]
            },
            {
              groupName: 'Script sources (advanced)',
              groupFields: [
                PropertyPaneTextField('excelJsUrl', {
                  label: 'ExcelJS library URL',
                  description: 'Used by Export → Excel. If your tenant blocks the public CDN, upload exceljs.min.js to a library on this site (e.g. Site Assets) and paste its full URL here. Leave blank to use the public CDN.'
                }),
                PropertyPaneTextField('chartJsUrl', {
                  label: 'Chart.js library URL',
                  description: 'Used by the dashboard charts. Same idea: upload chart.umd.js to this site and paste its URL if the CDN is blocked. Leave blank to use the public CDN.'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
