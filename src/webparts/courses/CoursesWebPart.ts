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

import { ICourse } from "../../common/ICourse";

import { CourseService } from "../../services/CourseService";

export interface ICoursesWebPartProps {
    count: number;
    category: string;
}

export default class CoursesWebPart extends BaseClientSideWebPart<ICoursesWebPartProps> {
    private provider: CourseService;
    private catValues: IPropertyPaneDropdownOption[] = [];

    protected onInit(): Promise<void> {
        //Create Course Service
        this.provider = new CourseService(`${this.context.pageContext.web.absoluteUrl}/_api/Lists/GetByTitle('Courses')/Items`,
            this.context);

        // One call for Category choic values for the
        // Prop pane dropdown
        this.provider.getCategories()
            .then((data: string[]) => {
                this.catValues = data.map((item => {
                    return {
                        key: item,
                        text: item
                    } as IPropertyPaneDropdownOption
                }));

                console.log("Prop Pane Options : " + JSON.stringify(this.catValues));
            });

        this.provider.updateCourse(1, {
            CourseID: 9001,
            Category: "Web Development",
            Title: "Test ",
            Description: "test Programming",
            Duration: 60,
            Price: 199,
            Technology: "Test"
        }).then(status => {
            console.log("Item updated : " + status);
        });

        this.provider.deleteCourse(10).then(status => console.log("Item deleted : " + status));

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
        this.provider.getData(this.properties.count, this.properties.category)
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

    protected get dataVersion(): Version {
        return Version.parse('1.0');
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