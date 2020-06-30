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
    private currentID: number = 0;

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

                $("#category", this.domElement).html(this.getCatSelectOptions(this.catValues));

                //console.log("Prop Pane Options : " + JSON.stringify(this.catValues));
            });

        return Promise.resolve();
    }

    private getCatSelectOptions(items: IPropertyPaneDropdownOption[]): string {
        let html = "";

        items.forEach(i => {
            html += `<option value='${i.key}'>${i.text}</option>`;
        });

        return html;
    }

    public render(): void {
        $(this.domElement).html(`
      <div class="${ styles.courses}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Courses!</span>
              <p class="${ styles.description}">List of Courses</p>
              <div>
                <input type="button" id="btnnew" value="New..." />
              </div>
              <div id="output">Loading...</div>
              <div id="addform">
                <h2>Add Course</h2>
                Course ID: <input type="text" id="courseid" /><br/>
                Name: <input type="text" id="coursename" /><br/>
                Details: <br/><textarea id="coursedesc" cols="40" rows="6"></textarea><br/>
                Category: <select id="category">
                  ${
            this.getCatSelectOptions(this.catValues)
            }
                </select><br/>
                Technology: <input type="text" id="technology" /><br/>
                Duration: <input type="text" id="duration" /><br/>
                Price: <input type="text" id="price" /><br/><br/>
                <input type="button" value="Save" id="btnaddsave" />&nbsp;
                <input type="button" value="Cancel" id="btnaddcancel" />
              </div>
              <div id="editform">
                <h2>Edit Course</h2>
                Course ID: <input type="text" id="ecourseid" /><br/>
                Name: <input type="text" id="ecoursename" /><br/>
                Details: <br/><textarea id="ecoursedesc" cols="40" rows="6"></textarea><br/>
                Category: <select id="ecategory">
                  ${
            this.getCatSelectOptions(this.catValues)
            }
                </select><br/>
                Technology: <input type="text" id="etechnology" /><br/>
                Duration: <input type="text" id="eduration" /><br/>
                Price: <input type="text" id="eprice" /><br/><br/>
                <input type="button" value="Save" id="btneditsave" />&nbsp;
                <input type="button" value="Cancel" id="btneditcancel" />
              </div>
            </div>
          </div>
        </div>
      </div>`);

        // New Button Event Handlers
        $("#btnnew", this.domElement).on('click', () => {
            $("#output", this.domElement).hide();
            $("#addform", this.domElement).show();
            $("#btnnew", this.domElement).hide();
        });

        $("#btnaddcancel", this.domElement).on('click', () => {
            $("#output", this.domElement).show();
            $("#addform", this.domElement).hide();
            $("#btnnew", this.domElement).show();
        });

        $("#btnaddsave", this.domElement).on('click', () => {
            let item: ICourse = {
                CourseID: $("#courseid", this.domElement).val() as number,
                Title: $("#coursename", this.domElement).val() as string,
                Description: $("#coursedesc", this.domElement).val() as string,
                Category: $("#category", this.domElement).val() as string,
                Duration: $("#duration", this.domElement).val() as number,
                Price: parseFloat($("#price", this.domElement).val() as string),
                Technology: $("#technology", this.domElement).val() as string
            };

            console.log("New Item : " + JSON.stringify(item));

            this.provider.addCourse(item).then(newItem => {
                console.log("Add success!");
                alert("Added Item!");
                $("#output", this.domElement).show();
                $("#addform", this.domElement).hide();
                $("#btnnew", this.domElement).show();

                this.render();
            }).catch(err => {
                alert("Error adding Item!");
                $("#output", this.domElement).show();
                $("#addform", this.domElement).hide();
                $("#btnnew", this.domElement).show();
            });

        });

        $("#btneditcancel", this.domElement).on('click', () => {
            $("#output", this.domElement).show();
            $("#editform", this.domElement).hide();
            $("#btnnew", this.domElement).show();
        });

        $("#btneditsave", this.domElement).on('click', () => {

            if (!confirm("Do you want to save changed?")) {
                $("#output", this.domElement).show();
                $("#editform", this.domElement).hide();
                $("#btnnew", this.domElement).show();
                return;
            }

            let item: ICourse = {
                CourseID: $("#ecourseid", this.domElement).val() as number,
                Title: $("#ecoursename", this.domElement).val() as string,
                Description: $("#ecoursedesc", this.domElement).val() as string,
                Category: $("#ecategory", this.domElement).val() as string,
                Duration: $("#eduration", this.domElement).val() as number,
                Price: parseFloat($("#eprice", this.domElement).val() as string),
                Technology: $("#etechnology", this.domElement).val() as string
            };

            this.provider.updateCourse(this.currentID, item).then(status => {
                if (status) {
                    alert("Item Updated!");
                }

                $("#output", this.domElement).show();
                $("#editform", this.domElement).hide();
                $("#btnnew", this.domElement).show();

                this.render();
            }).catch(err => {
                alert("Error Updating Item!");
                $("#output", this.domElement).show();
                $("#editform", this.domElement).hide();
                $("#btnnew", this.domElement).show();
            });

        });

        $("#addform", this.domElement).hide();
        $("#editform", this.domElement).hide();


        // Get the Courses
        this.provider.getData(this.properties.count, this.properties.category == "All" ? undefined : this.properties.category)
            .then((courses: ICourse[]) => {
                $("#output", this.domElement).html(this.getHTML(courses));

                // Register the Edit/Del Link handlers
                this.registerDelHandlers();

                this.registerEditHandlers();
            });
    }

    private registerEditHandlers() {
        $('a[id^="edt"]', this.domElement).on('click', (event) => {
            event.preventDefault();

            let itemID: number = parseInt($(event.currentTarget).attr("href"));

            this.currentID = itemID;

            this.provider.getItemById(itemID).then((course: ICourse) => {
                $("#output", this.domElement).hide();
                $("#editform", this.domElement).show();
                $("#btnnew", this.domElement).hide();

                // Populate the edit form
                $("#ecourseid", this.domElement).val(course.CourseID.toString());
                $("#ecoursename", this.domElement).val(course.Title);
                $("#ecoursedesc", this.domElement).val(course.Description);
                $("#ecategory", this.domElement).val(course.Category);
                $("#eduration", this.domElement).val(course.Duration.toString());
                $("#eprice", this.domElement).val(course.Price.toString());
                $("#etechnology", this.domElement).val(course.Technology);
                $("#ecourseid", this.domElement).focus();
            }).catch(err => {
                alert("Unable to fetch details!");
                $("#output", this.domElement).show();
                $("#editform", this.domElement).hide();
                $("#btnnew", this.domElement).show();
            });
        });
    }

    private registerDelHandlers() {
        $('a[id^="del"]', this.domElement).on('click', (event) => {
            event.preventDefault();

            let itemID: number = parseInt($(event.currentTarget).attr("href"));

            if (confirm("Delete this Course?")) {
                this.provider.deleteCourse(itemID).then(success => {
                    if (success) {
                        alert("Course Deleted!");
                        this.render();
                    }
                }).catch(err => {
                    alert("Error deleting course! : " + err);
                });
            }
        });
    }

    private getHTML(courses: ICourse[]): string {
        let html = `<table>
                <tr>
                  <th>&nbsp;</th>
                  <th>ID</th>
                  <th>Name</th>
                  <th>Category</th>
                  <th>Details</th>
                  <th>Technology</th>
                  <th>Price<th>
                  <th>Hours</th>
                  <th>&nbsp;</th>
                </tr>`;

        for (let c of courses) {
            html += `
        <tr class="${ styles.courserow}">
          <td>
            <a href="${ c["ID"]}" id="edt">Edit</a>&nbsp;
          </td>
          <td>${ c.CourseID} </td>
          <td>${ c.Title} </td>
          <td>${ c.Category} <td>
          <td>${ c.Description} </td>
          <td> ${ c.Technology} </td>
          <td>${ c.Price} </td>
          <td>${ c.Duration} </td>
          <td>
            <a href="${ c["ID"]}" id="del">Del</a>
          </td>
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