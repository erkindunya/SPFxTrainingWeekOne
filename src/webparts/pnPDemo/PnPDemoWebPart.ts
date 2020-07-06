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

    /*this.provider.addItem({
      CourseID: 8002,
      Title: 'Entity Framework',
      Description: 'Entity Framework with SQL Server',
      Category: 'Web Development',
      Duration: 40,
      Price: 99.25,
      Technology: 'Databases'
    } as ICourse).then(item => {
      console.log("Added item successfully!");
      console.log(`Item ID: ${ item['ID'] } and ETag: ${ item["odata.etag"]}`);
    }).catch(err=> {
      console.log("Error adding item : " + err);
    });*/

    // console.log("Updating item...");
    // this.provider.updateItem(7, {
    //   CourseID: 1007,
    //   Title: 'Swift Programming for iOS',
    //   Description: 'Mobile App Dev with Swift',
    //   Category: 'Mobile Development',
    //   Technology: 'Swift',
    //   Duration: 40,
    //   Price: 200
    // } as ICourse).then(flag => {
    //   if (flag) {
    //     console.log("Item Updated successfully!");
    //   } else {
    //     console.log("Item update failed!");
    //   }
    // });

    //test Delete
    // this.provider.deleteItem(15)
    //   .then(_ => {
    //     console.log("Delete successful!");
    //   })
    //   .catch(err => {
    //     console.log("Delete failed - " + err);
    //   });

    // this.provider.getCategories().then(output => {
    //   console.log(JSON.stringify(output));
    // });

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
