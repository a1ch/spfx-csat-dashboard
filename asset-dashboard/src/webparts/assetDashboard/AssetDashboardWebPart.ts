import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

/* eslint-disable @typescript-eslint/no-explicit-any */
import * as strings from 'AssetDashboardWebPartStrings';
import { ASSET_DASHBOARD_HTML } from './assetDashboardTemplate';
import { initAssetDashboard, IAssetController } from './assetDashboardCore';
import { fetchAssetData } from './AssetFormsService';
import { fetchGlOverrides, saveGlOverride } from './AssetGlService';

export interface IAssetDashboardWebPartProps {
  listSiteUrl: string;
  formsFolderUrl: string;
  glListName: string;
}

export default class AssetDashboardWebPart extends BaseClientSideWebPart<IAssetDashboardWebPartProps> {

  private _dashboard: IAssetController | undefined;
  private _builtKey: string = '';

  public render(): void {
    const siteUrl: string = (this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl).trim();
    const formsFolderUrl: string = (this.properties.formsFolderUrl || '').trim();
    const glListName: string = (this.properties.glListName || 'Asset GL Overrides').trim();

    // SharePoint calls render() often; only rebuild when config changed.
    const key: string = `${siteUrl}|${formsFolderUrl}|${glListName}`;
    const hasRoot: boolean = !!this.domElement.querySelector('.astRoot');
    if (hasRoot && this._dashboard && key === this._builtKey) { return; }
    if (this._dashboard) { this._dashboard.destroy(); this._dashboard = undefined; }
    this._builtKey = key;

    this.domElement.innerHTML = ASSET_DASHBOARD_HTML;
    const root: HTMLElement = this.domElement.querySelector('.astRoot') as HTMLElement;
    if (!root) { return; }

    this._dashboard = initAssetDashboard(root, {
      uploadUrl: formsFolderUrl.replace(/\/$/, ''),
      fetchData: () => fetchAssetData(this.context.spHttpClient, siteUrl, formsFolderUrl),
      fetchGl: () => fetchGlOverrides(this.context.spHttpClient, siteUrl, glListName),
      saveGl: (emp, item, serial, gl, id) =>
        saveGlOverride(this.context.spHttpClient, siteUrl, glListName, emp, item, serial, gl, id)
    });
  }

  protected onDispose(): void {
    if (this._dashboard) { this._dashboard.destroy(); this._dashboard = undefined; }
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyPaneTextField('formsFolderUrl', {
                  label: strings.FormsFolderFieldLabel,
                  description: 'URL of the folder holding the Equipment Asset Allocation forms (e.g. https://…/SiteAssets/AssetForms). Every .xlsx in it is parsed. The "Upload form" button opens this folder.',
                  multiline: true
                }),
                PropertyPaneTextField('listSiteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  description: 'Site that holds the forms folder and the GL list. Leave blank to use the current site.'
                }),
                PropertyPaneTextField('glListName', {
                  label: strings.GlListFieldLabel,
                  description: 'List where GL / Cost Center values typed in the dashboard are saved (columns: Employee, Item, Serial, GL).'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
