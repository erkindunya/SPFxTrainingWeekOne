import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CustomersWebPart.module.scss';
import * as strings from 'CustomersWebPartStrings';
import  {ICustomer} from '../../common/ICustomer';

const data : ICustomer[] = [
    {
      id: 1,
      name: "ABC Ltd",
      phone: "43434343",
      balance: 3000.25
    },
    {
      id: 2,
      name: "PQR Ltd",
      phone: "24343343",
      balance: 8541.25
    },
    {
      id: 3,
      name: "MNB Ltd",
      phone: "87867676",
      balance: 8965.21
    },
    {
      id: 4,
      name: "YUI Ltd",
      phone: "53434344",
      balance: 9652.40
    }

  ];

export interface ICustomersWebPartProps {
  description: string;
}

export default class CustomersWebPart extends BaseClientSideWebPart <ICustomersWebPartProps> {
private customers : ICustomer[];
private getData() : ICustomer[]{
    return data;

}
  public onInit() : Promise<void> {
    console.log("** OnInit Fired!");
  // Get the Customers and Store
      this.customers = this.getData();
    return Promise.resolve();
  }

  public onDispose() {
    console.log("** OnDispose Fired!");
  }

  public render(): void {
    console.log("** Render Fired!");

    this.domElement.innerHTML = `
      <div class="${ styles.customers }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">CUSTOMERS!</span>
              <p class="${ styles.subTitle }">List of Customers</p>
              <div id="output">
                Loading...
              </div>
            </div>
          </div>
        </div>
      </div>`;
      this.domElement.querySelector("#output").innerHTML = this.getHTML(this.customers);
  }
private getHTML(items: ICustomer[]) : string {

let html="";

for (let c of items) {
html += `<div>
${c.id} </br>
${c.name} </br>
${c.phone} </br>
${c.balance} </br>
</div>`;

}
return html;
}
  protected get dataVersion(): Version {
  return Version.parse('1.0');
}

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log("** getPropertyPaneConfig fired!");

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
