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

import { CourseProvider } from "../../services/CourseProvider";

export interface IPnPDemoWebPartProps {
  description: string;
}

export default class PnPDemoWebPart extends BaseClientSideWebPart<IPnPDemoWebPartProps> {
  private provider: CourseProvider;

  protected onInit(): Promise<void> {
    this.provider = new CourseProvider("Courses", this.context);

    return Promise.resolve();
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.pnPDemo}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">PnP Demo!</span>
              <p class="${ styles.subTitle}">Using Pnp/sp Library</p>
              <div class="${ styles.description}" id="output">
                Loading...
              </div>
            </div>
          </div>
        </div>
      </div>`;

    // Add Items
    console.log("Adding new Item...");

    this.provider.addItem({
      CourseID: 8001,
      Title: 'SQL Server Programming',
      Description: 'SQL Server Triggers and Procedures',
      Category: 'Web Development',
      Duration: 40,
      Price: 99.25,
      Technology: 'Databases'
    } as ICourse).then(item => {
      console.log("Added item successfully!");
      console.log(`Item ID: ${item['ID']} and ETag: ${item["etag"]}`);
    }).catch(err => {
      console.log("Error adding item : " + err);
    });

    this.provider.getCategories().then(output => {
      console.log(JSON.stringify(output));
    });

    this.provider.getItems()
      .then((courses: ICourse[]) => {
        let html = "";

        for (let c of courses) {
          html += `<div class="${styles.course}>
                      ${ c.Title} <br/>
                      ${ c.Category} <br/>
                      ${ c.CourseID} <br/>
                      ${ c.CourseID} hrs</br>
                      ${ c.Price}
                    <div>`;
        }

        this.domElement.querySelector("#output").innerHTML = html;
      }).catch(err => {
        this.domElement.querySelector("#output").innerHTML = `<div>
            Error getting Items: ${ err}
          </div>`;
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
