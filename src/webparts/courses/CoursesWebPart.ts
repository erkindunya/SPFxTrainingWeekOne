import { Version } from "@microsoft/sp-core-library";
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from "./CoursesWebPart.module.scss";
import * as strings from "CoursesWebPartStrings";

export interface ICoursesWebPartProps {
    description: string;
}

interface ICourse {
    CourseID: number;
    Category: string;
    Title: string;
    Description: string;
    Technology: string;
    Duration: number;
    Price: number;
}
const url =
    "https://selchuk.sharepoint.com/_api/lists/getbytitle('courses')/items";
export default class CoursesWebPart extends BaseClientSideWebPart<
    ICoursesWebPartProps
> {
    public render(): void {
        this.domElement.innerHTML = `
      <div class="${styles.courses}">
        <div class="${styles.container}">
           <div class="${styles.row}">
              <div class="${styles.column}">
                <span class="${styles.title}">Courses!</span>
                <p class="${styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
                <div id="output">Loading...</div>
            </div>
          </div>
        </div>
      </div>`;
        this.getData(url).then((courses: ICourse[]) => {
            this.domElement.querySelector("#output").innerHTML = this.getHTML(
                courses
            );
        });
    }
    private getHTML(courses: ICourse[]): string {
        let html = "";

        for (let c of courses) {
            html += `
      <div class="${styles.coursebox}">
          ${c.CourseID} <br/>
          ${c.Title} <br/>
          ${c.Description} <br/>
          ${c.Technology} <br/>
          ${c.Price} <br/>
          ${c.Duration}
        </div>
      `;
        }

        return html;
    }
    // private getData (url : string): Promise<ICourse[]>{
    //   return this.context.spHttpClient.get(url,SPHttpClient.configurations.v1)
    // .then (resp :SPHttpClientResponse) => {
    //   return resp.json();
    //   }).then(data => {
    //     return data.value as ICourse[];
    //   });
    //   }

    private getData(url: string): Promise<ICourse[]> {
        return this.context.spHttpClient
            .get(url, SPHttpClient.configurations.v1)
            .then((resp: SPHttpClientResponse) => {
                return resp.json();
            })
            .then((data) => {
                return data.value as ICourse[];
            })
            .catch((err) => {
                console.log("getData()-> Error in REST Call : " + err);
                return [];
            });
    }

    private async getData2(url: string): Promise<ICourse[]> {
        let resp = await this.context.spHttpClient.get(
            url,
            SPHttpClient.configurations.v1
        );

        let data = await resp.json();

        let courses = data.value as ICourse[];

        return Promise.resolve(courses);
    }

    protected get dataVersion(): Version {
        return Version.parse("1.0");
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField("description", {
                                    label: strings.DescriptionFieldLabel,
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
