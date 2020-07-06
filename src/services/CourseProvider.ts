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

  public updateItem(id: number, item: ICourse): Promise<boolean> {
    return sp.web.lists.getByTitle(this.listName).items.getById(id)
      .update(item)
      .then((result: IItemUpdateResult) => {
        return true;
      })
      .catch(err => {
        console.log("Updated Failed: " + err);
        return err;
      });
  }

  public deleteItem(id: number): Promise<any> {
    return sp.web.lists.getByTitle(this.listName).items.getById(id)
      .delete();
  }

  public getItemsByCategory(count: number = 100, category: string): Promise<ICourse[]> {
    return sp.web.lists.getByTitle(this.listName).items
      .top(count)
      .filter(`Category eq ${category}`)
      .get<ICourse[]>();
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