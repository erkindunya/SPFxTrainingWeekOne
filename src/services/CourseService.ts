import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICourse } from "../common/ICourse";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";

export class CourseService {

  constructor(private url: string, private context: WebPartContext) {

  }

  public addCourse(newItem: ICourse): Promise<ICourse> {
    return this.context.spHttpClient.post(this.url, SPHttpClient.configurations.v1, {
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      body: JSON.stringify(newItem)
    }).then(resp => {
      return resp.json();
    }).then(data => {
      return data as ICourse;
    });
  }

  // this build URL and it is dynamicly build as web part runs so it captures here  (this.url + `(${id})then it goes to  CourseWebPart.ts line 34, Update and Delete addes extra Rest filter and 
  // https:/.../_api/Lists/GetByTitle('Courses')/Items(10) 

  public updateCourse(id: number, item: ICourse, etag: string): Promise<boolean> {
    return this.context.spHttpClient.post(this.url + `(${id})`, SPHttpClient.configurations.v1, {
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "X-Http-Method": "PATCH",
        "IF-Match": etag
      },
      body: JSON.stringify(item)
    }).then(resp => {
      console.log("Status Text: " + resp.statusText);
      return resp.ok;
    }).catch(err => {
      console.log("Error updating Course : " + err);
      return false;
    });
  }

  public deleteCourse(id: number, etag: string): Promise<boolean> {
    return this.context.spHttpClient.post(this.url + `(${id})`, SPHttpClient.configurations.v1, {
      headers: {
        "Accept": "application/json",
        "X-Http-Method": "DELETE",
        "IF-Match": etag
      }
    }).then(resp => {
      return resp.ok;
    }).catch(err => {
      console.log("Error deleting Course");
      return false;
    });
  }

  public getData(count: number = 100, category?: string): Promise<ICourse[]> {
    let url = this.url;
    url += "?$top=" + count;

    if (category) {
      url += `&$filter=Category eq '${category}'`;
    }

    console.log("getData() url : " + url);

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

  public getItemById(id: number): Promise<ICourse> {
    let url = this.url;

    url += `?$filter=ID eq ${id}`;

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((resp: SPHttpClientResponse) => {
        return resp.json();
      }).then(data => {
        return data.value[0] as ICourse;
      }).catch(err => {
        console.log("getData()-> Error in REST Call : " + err);
        return null;
      });
  }

  public getCategories(): Promise<string[]> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl
      + "/_api/web/lists/GetByTitle('courses')/fields?$filter=EntityPropertyName eq 'Category'",
      SPHttpClient.configurations.v1
    ).then(resp => {
      return resp.json();
    }).then(data => {
      console.log(JSON.stringify(data));

      return data.value[0].Choices as string[]
    });
  }

}