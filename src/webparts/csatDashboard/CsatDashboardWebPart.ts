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
import { initDashboard } from './dashboardCore';
import { fetchCsatItems, ICsatItem } from './CsatDataService';

export interface ICsatDashboardWebPartProps {
  listSiteUrl: string;
  listName: string;
  autoRefreshSeconds: number;
}

const CHARTJS_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js';
const EXCELJS_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';

export default class CsatDashboardWebPart extends BaseClientSideWebPart<ICsatDashboardWebPartProps> {

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Load Chart.js and ExcelJS once, from CDN. Both are UMD bundles, and the
    // SharePoint page already has an AMD loader (define/require) present. Loaded
    // as a plain script, the UMD wrapper registers as an anonymous AMD module
    // instead of assigning window.Chart / window.ExcelJS — so the globals the
    // dashboard checks for stay undefined (blank charts, failed Excel export).
    // Passing globalExportsName makes SPComponentLoader suppress AMD during the
    // load and capture the real browser global. Each loads independently so one
    // failing (or a tenant blocking the CDN) still leaves the other working.
    await Promise.all([
      SPComponentLoader.loadScript(CHARTJS_URL, { globalExportsName: 'Chart' })
        .catch((e: unknown) => {
          /* eslint-disable-next-line no-console */
          console.error('CSAT dashboard: Chart.js failed to load', e);
        }),
      SPComponentLoader.loadScript(EXCELJS_URL, { globalExportsName: 'ExcelJS' })
        .catch((e: unknown) => {
          /* eslint-disable-next-line no-console */
          console.error('CSAT dashboard: ExcelJS failed to load', e);
        })
    ]);
  }

  public render(): void {
    const siteUrl: string = (this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl).trim();
    const listName: string = (this.properties.listName || 'CSAT RESPONSES').trim();
    const autoRefreshSeconds: number = this.properties.autoRefreshSeconds || 0;

    this.domElement.innerHTML = DASHBOARD_HTML;

    const root: HTMLElement = this.domElement.querySelector('.sfCsatRoot') as HTMLElement;
    if (!root) { return; }

    initDashboard(root, {
      autoRefreshSeconds,
      fetchItems: (): Promise<ICsatItem[]> => fetchCsatItems(this.context.spHttpClient, siteUrl, listName)
    });

    // If auto-refresh is configured, start it after first render by clicking
    // the Auto toggle programmatically once the dashboard has wired up.
    if (autoRefreshSeconds > 0) {
      const autoBtn = root.querySelector('[data-action="auto"]') as HTMLElement;
      if (autoBtn) { window.setTimeout(() => autoBtn.click(), 3000); }
    }
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
            }
          ]
        }
      ]
    };
  }
}
