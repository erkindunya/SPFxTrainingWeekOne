import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ICourse } from "../common/ICourse";

import { sp, IItemAddResult, IItemUpdateResult } from "@pnp/sp/presets/all";

export class CourseProvider {

  constructor(private listName: string, private context: WebPartContext) {
    sp.setup({
      spfxContext: this.context
    });
  }

  public addItem(newItem: ICourse): Promise<ICourse> {
    return sp.web.lists.getByTitle(this.listName).items.add(newItem)
      .then((result: IItemAddResult) => {
        console.log("New Item : " + JSON.stringify(result.data));

        return result.data as ICourse;
      });
  }

  public updateItem(id: number, item: ICourse, eTag: string): Promise<boolean> {
    return sp.web.lists.getByTitle(this.listName).items.getById(id)
      .update(item, eTag)
      .then((result: IItemUpdateResult) => {
        return true;
      })
      .catch(err => {
        console.log("Updated Failed: " + err);
        return false;
      });
  }

  public deleteItem(id: number, eTag: string): Promise<any> {
    return sp.web.lists.getByTitle(this.listName).items.getById(id)
      .delete(eTag);
  }

  public getItemsByCategory(count: number = 100, category: string): Promise<ICourse[]> {
    if (!category) {
      return sp.web.lists.getByTitle(this.listName).items
        .top(count)
        .get<ICourse[]>();
    }

    return sp.web.lists.getByTitle(this.listName).items
      .top(count)
      .filter(`Category eq '${category}'`)
      .get<ICourse[]>();
  }

  public getItemsByTitle(title: string): Promise<ICourse> {
    let query: string = `<View>
        <Where>
          <Eq>
            <FieldRef Name="Title" />
            <Value Type="Text">${ title}</Value>
          </Eq>
        </Where>
      </View>`;

    return sp.web.lists.getByTitle(this.listName).getItemsByCAMLQuery({
      ViewXml: query
    }).then((courses: ICourse[]) => {
      return courses[0] as ICourse;
    });
  }

  public getItems(count: number = 100): Promise<ICourse[]> {
    return sp.web.lists.getByTitle(this.listName).items
      .top(count)
      .get<ICourse[]>();
  }

  public getItemById(id: number): Promise<ICourse> {
    return sp.web.lists.getByTitle(this.listName).items
      .getById(id)
      .get<ICourse>();
  }

  public getCategories(): Promise<string[]> {
    return sp.web.lists.getByTitle(this.listName).fields
      .filter("EntityPropertyName eq 'Category'")
      .get<any[]>()
      .then(data => {
        return data[0].Choices as string[];
      });
  }

}