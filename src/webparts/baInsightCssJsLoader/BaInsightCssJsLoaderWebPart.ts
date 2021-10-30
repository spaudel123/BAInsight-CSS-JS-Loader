import { Version,Log,DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './BaInsightCssJsLoaderWebPart.module.scss';
import * as strings from 'BaInsightCssJsLoaderWebPartStrings';

import {SPComponentLoader} from '@microsoft/sp-loader';


export interface IBaInsightCssJsLoaderWebPartProps {
  description: string;
  cssfullpath: string;
  jsfullpath: string;
}

export default class BaInsightCssJsLoaderWebPart extends BaseClientSideWebPart<IBaInsightCssJsLoaderWebPartProps> {

  public render(): void {
    if(this.displayMode == DisplayMode.Edit){
      //Modern SharePoint in Edit Mode
      this.domElement.innerHTML = `
      <div class="${ styles.baInsightCssJsLoader }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Load External CSS/JS</span>
              <p class="${ styles.subTitle }">Edit webpart and Specify css/js path to load</p>
            
            </div>
          </div>
        </div>
      </div>`;
      }else{
      //Modern SharePoint in Read Mode
      this.domElement.innerHTML = '';
      }
  
     
  }

  protected onInit():Promise<void>{
    var cpath = escape(this.properties.cssfullpath);
    var jpath = escape(this.properties.jsfullpath);
    Log.info("BAI Custom css js loader","Custom css path:" + cpath);
    Log.info("BAI Custom css js loader","Custom js path:" + jpath);
    SPComponentLoader.loadCss(cpath);
    SPComponentLoader.loadScript(jpath);
    return super.onInit();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('cssfullpath', {
                  label: strings.CSSFullPathLabel
                }),
                PropertyPaneTextField('jsfullpath', {
                  label: strings.JSFullPathLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
