import { SPFI } from "@pnp/sp";
import { IColumn } from "@fluentui/react";

// fetched fields to display
export const fetchLibraryViewFields = async (sp: SPFI, libraryId: string, viewId: string): Promise<{ columns: IColumn[]; viewQuery: string}> => {
  const list = sp.web.lists.getById(libraryId);
  const fields = await list.fields();
  const view = await list.views.getById(viewId)();
  const viewQuery = view.ViewQuery || '';

  const viewFieldsResp = await list.views.getById(viewId).fields();
  const viewFieldInternalNames: string[] = viewFieldsResp.Items as string[];

  const visibleFields = fields.filter(f => !f.Hidden && viewFieldInternalNames.includes(f.InternalName));
  const ordered = viewFieldInternalNames
    .map(name => visibleFields.find(f => f.InternalName === name));

  const columns = ordered.map(f => ({
    key: f!.InternalName,
    name: f!.Title,
    fieldName: f!.InternalName,
    minWidth: 100,
    maxWidth: 150,
    isResizable: true,
    data: f!.TypeAsString,
    isSorted: false,
    isSortedDescending: false,
    isRowHeader: true,
    iconName: f!.InternalName === 'DocIcon' ? 'Page' : '',
    isIconOnly: f!.InternalName === 'DocIcon' ? true : false
  }));
  return {columns, viewQuery};
}

// fetches all fields in the library
export const fetchLibraryFields = async (sp: SPFI, libraryId: string): Promise<IColumn[]> => {
  const fields = await sp.web.lists.getById(libraryId).fields();
  return fields.filter(f => !f.Hidden).map(f => ({
    key: f.InternalName,
    name: f.Title,
    fieldName: f.InternalName,
    minWidth: 100,
    maxWidth: 150,
    isResizable: true,
    data: f.TypeAsString,
    isSorted: false,
    isSortedDescending: false,
    isRowHeader: true,
    iconName: f.InternalName === 'DocIcon' ? 'Page' : '',
    isIconOnly: f.InternalName === 'DocIcon' ? true : false,
    TypeAsString: f.TypeAsString
  }));
}