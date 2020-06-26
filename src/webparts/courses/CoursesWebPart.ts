import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown,
    IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import styles from './CoursesWebPart.module.scss';
import * as strings from 'CoursesWebPartStrings';

export interface ICoursesWebPartProps {
    count: number;
    category: string;
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

const svcUrl = "https://selchuk.sharepoint.com/_api/Lists/GetByTitle('Courses')/Items";

export default class CoursesWebPart extends BaseClientSideWebPart<ICoursesWebPartProps> {
    private catValues: IPropertyPaneDropdownOption[] = [];
    private courses: ICourse[] = [];

    protected onInit(): Promise<void> {

        // One call for Category choic values for the
        // Prop pane dropdown
        this.getCategories()
            .then((data: string[]) => {
                this.catValues = data.map((item => {
                    return {
                        key: item,
                        text: item
                    } as IPropertyPaneDropdownOption;
                }));

                console.log("Prop Pane Options : " + JSON.stringify(this.catValues));
            });

        return Promise.resolve();
    }

    public render(): void {
        this.domElement.innerHTML = `
        <div class="${ styles.courses}">
          <div class="${ styles.container}">
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <span class="${ styles.title}">Courses!</span>
                <p class="${ styles.subTitle}">Course data from SP List.</p>
                <div id="output">Loading...</div>
              </div>
            </div>
          </div>
        </div>`;

        // Get the Courses
        this.getData(svcUrl, this.properties.count, this.properties.category)
            .then((courses: ICourse[]) => {
                this.domElement.querySelector("#output").innerHTML = this.getHTML(courses);
            });
    }

    private getHTML(courses: ICourse[]): string {
        let html = "";

        for (let c of courses) {
            html += `
          <div class="${ styles.coursebox}">
            ID: ${ c.CourseID} <br/>
            NAME: ${ c.Title} <br/>
            DESC: ${ c.Description} <br/>
            TECH: ${ c.Technology} <br/>
            PRICE: ${ c.Price} <br/>
            HOURS: ${ c.Duration}
          </div>
        `;
        }

        return html;
    }

    private getData(url: string, count: number, category?: string): Promise<ICourse[]> {
        url += "?$top=" + count;

        if (category) {
            url += `&$filter=Category eq '${category}'`;
        }

        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((resp: SPHttpClientResponse) => {
                return resp.json();
            }).then(data => {
                return data.value as ICourse[];
            }).catch(err => {
                console.log("getData()-> Error in REST Call : " + err);
                return [];
            });
    }

    private async getData2(url: string): Promise<ICourse[]> {
        try {
            let resp = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);

            let data = await resp.json();

            let courses = data.value as ICourse[];

            return Promise.resolve(courses);

        } catch (err) {
            console.log(err);
        }

        return Promise.resolve([]);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    private getCategories(): Promise<string[]> {

        return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
            + "/_api/web/lists/GetByTitle('courses')/fields?$filter=EntityPropertyName eq 'Category'",
            SPHttpClient.configurations.v1
        ).then(resp => {
            return resp.json();
        }).then(data => {
            console.log(JSON.stringify(data));

            return data.value[0].Choices as string[];
        });
    }

    protected getPropertyPaneConfiguration = (): IPropertyPaneConfiguration => {
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
                                PropertyPaneTextField('count', {
                                    label: "Count",
                                    onGetErrorMessage: (value: string) => {
                                        let c: number = parseInt(value);

                                        if (isNaN(c)) {
                                            return "Invalid number";
                                        }

                                        if (c <= 0) {
                                            return "Count must be > 0";
                                        }

                                        return "";
                                    },
                                    deferredValidationTime: 300
                                }),
                                PropertyPaneDropdown('category', {
                                    label: 'Category',
                                    options: this.catValues
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}
