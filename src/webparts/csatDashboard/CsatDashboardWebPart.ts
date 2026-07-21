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
const XLSX_URL: string = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

export default class CsatDashboardWebPart extends BaseClientSideWebPart<ICsatDashboardWebPartProps> {

  private _librariesLoaded: boolean = false;

  protected async onInit(): Promise<void> {
    await super.onInit();
    // Load Chart.js and SheetJS (xlsx) once, from CDN. If a tenant blocks the
    // CDN the dashboard still renders — charts/export just degrade gracefully.
    try {
      await SPComponentLoader.loadScript(CHARTJS_URL);
      await SPComponentLoader.loadScript(XLSX_URL);
      this._librariesLoaded = true;
    } catch (e) {
      this._librariesLoaded = false;
    }
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
