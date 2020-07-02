import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PnPDemoWebPart.module.scss';
import * as strings from 'PnPDemoWebPartStrings';

import { sp } from "@pnp/sp/presets/all";
import { ICourse } from "../../common/ICourse";

export interface IPnPDemoWebPartProps {
  description: string;
}

export default class PnPDemoWebPart extends BaseClientSideWebPart<IPnPDemoWebPartProps> {
  // initialse SP object for PNP,
  protected onInit(): Promise<void> {
    sp.setup({
      spfxContext: this.context
    });
    return Promise.resolve();
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnPDemo}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description}" id="output">
              
              </p>

            </div>
          </div>
        </div>
      </div>`;
    sp.web.lists.getByTitle('Courses').items.get()
      .then((courses: ICourse[]) => {
        let html = "";
        for (let c of courses) {

          html += `<div> 
          ${ c.Title} : $ { c.Category}
          
          <div>`
        }
        this.domElement.querySelector("#output").innerHTML = html;
      });
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
