import { Version } from '@microsoft/sp-core-library';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
    PropertyPaneDropdown,
    IPropertyPaneDropdownOption,
    PropertyPaneDropdownOptionType
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CoursesWebPart.module.scss';
import * as strings from 'CoursesWebPartStrings';

import * as $ from "jquery";

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

                //console.log("Prop Pane Options : " + JSON.stringify(this.catValues));
            });

        return Promise.resolve();
    }

    public render(): void {
        $(this.domElement).html(`
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
      </div>`);

        // Get the Courses
        this.provider.getData(this.properties.count, this.properties.category == "All" ? undefined : this.properties.category)
            .then((courses: ICourse[]) => {
                $("#output", this.domElement).html(this.getHTML(courses));
            });
    }

    private getHTML(courses: ICourse[]): string {
        let html = "<table>";

        for (let c of courses) {
            html += `
        <tr>
          <td>${ c.CourseID} </td>
          <td>${ c.Title} </td>
          <td>${ c.Description} </td>
          <td> ${ c.Technology} </td>
          <td>${ c.Price} </td>
          <td>${ c.Duration} </td>
        </tr>
      `;
        }

        return html + "</table>";
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
                                    options: [{
                                        key: "All",
                                        text: "Show All"
                                    },
                                    {
                                        key: "div1",
                                        text: "-",
                                        type: PropertyPaneDropdownOptionType.Divider
                                    }, ...this.catValues]
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}