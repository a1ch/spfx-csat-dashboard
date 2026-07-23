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

// ExcelJS is bundled straight into this package (its self-contained browser
// build) instead of being fetched at runtime. Loading it from a CDN or Site
// Assets proved unreliable on this tenant — the 925 KB UMD kept failing to
// attach window.ExcelJS (CDN block / MIME / AMD-loader interaction), leaving
// Excel export broken while the smaller Chart.js worked. Bundling removes every
// one of those failure modes: the global is guaranteed to exist.
/* eslint-disable @typescript-eslint/no-explicit-any */
// eslint-disable-next-line @typescript-eslint/no-var-requires
const ExcelJSBundled: any = require('exceljs/dist/exceljs.min.js');

export interface ICsatDashboardWebPartProps {
  listSiteUrl: string;
  listName: string;
  autoRefreshSeconds: number;
  chartJsUrl: string;
}

// Chart.js still loads at runtime (small, reliable). CDN default; overridable in
// the property pane with a same-tenant copy if a tenant blocks the CDN.
const CHARTJS_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js';

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
    // ExcelJS is bundled — just expose it as the browser global the dashboard's
    // export code expects. Guaranteed present; nothing to download or race.
    (window as any).ExcelJS = ExcelJSBundled && ExcelJSBundled.default ? ExcelJSBundled.default : ExcelJSBundled;

    // Chart.js is loaded at runtime (works reliably; CDN or self-hosted URL).
    const chartUrl: string = (this.properties.chartJsUrl || '').trim() || CHARTJS_URL;
    try {
      await SPComponentLoader.loadScript(chartUrl, { globalExportsName: 'Chart' });
    } catch (e) {
      /* eslint-disable-next-line no-console */
      console.error('CSAT dashboard: Chart.js failed to load from ' + chartUrl, e);
    }
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
                PropertyPaneTextField('chartJsUrl', {
                  label: 'Chart.js library URL',
                  description: 'Used by the dashboard charts. If your tenant blocks the public CDN, upload chart.umd.js to a library on this site (e.g. Site Assets) and paste its full URL here. Leave blank to use the public CDN. (ExcelJS is now built into the web part — no URL needed.)'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
