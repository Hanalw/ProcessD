import { SPFI } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import { IViewInfo } from "@pnp/sp/views";

export const fetchListViews = async (sp: SPFI, listId: string) => {
  if (!sp || !listId) return [];
  try {
    const views = await sp.web.lists.getById(listId).views
      .select("Id", "Title", "Hidden", "DefaultView", "ServerRelativeUrl")();
    return views;
  } catch (err) {
    console.error("fetchListViews error:", err);
    return [];
  }
}

export const findView = (views: IViewInfo[], title: string) => {
  return views.find(view => view.Title === title);
}