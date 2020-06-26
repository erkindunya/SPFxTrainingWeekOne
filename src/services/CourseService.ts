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

  public getData(count: number = 100, category?: string): Promise<ICourse[]> {
    let url = this.url;
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

  public getItemById(id: number): Promise<ICourse> {
    let url = this.url;

    url += `&$filter=ID eq ${id}`;

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