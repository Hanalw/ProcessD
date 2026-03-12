import type { SPFI } from "@pnp/sp";
import "@pnp/sp/batching";
import * as _ from "lodash";
import { IViewColumn, IFieldInfo, IColumn } from '../models'

/**
 * Helpers for underscore-prefixed SharePoint internal fields
 * e.g., _UIVersionString must be requested as OData__UIVersionString
 */
const isUnderscoreField = (name?: string): boolean =>
  !!name && name.startsWith("_");

const toODataName = (internalName: string): string =>
  isUnderscoreField(internalName)
    ? `OData__${internalName.substring(1)}`
    : internalName;

/**
 * Mirror OData__ props back to underscore equivalents on an object
 * so the rest of the UI can continue to use the original internal names.
 */
const mirrorUnderscoreProps = (obj: any): void => {
  if (!obj || typeof obj !== "object") return;
  for (const key of Object.keys(obj)) {
    if (key.startsWith("OData__")) {
      const underscore = `_${key.substring("OData__".length)}`;
      if (obj[underscore] == null) {
        obj[underscore] = obj[key];
      }
    }
  }
};

export function fieldInfoToColumn(f: IFieldInfo): IColumn {
  return {
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
    iconName: f.InternalName === "DocIcon" ? "Page" : "",
    isIconOnly: f.InternalName === "DocIcon",
  };
}

// Figure out which selects/expands are needed based on the view columns + field types
export const buildSelectExpand = (
  viewColumns: IViewColumn[],
  fieldInfos: Map<string, IFieldInfo>
): { select: string[]; expand: string[] } => {
  // Always include FieldValuesAsText to be able to project text for any field
  const select = new Set<string>(["Id", "FileLeafRef", "FileRef", "FieldValuesAsText"]);
  const expand = new Set<string>(["FieldValuesAsText"]);

  for (const col of viewColumns) {
    const internal = col.fieldName;
    if (!internal) continue;

    if (internal === "DocIcon") continue;

    if (internal === "LinkFilename" || internal === "LinkFilenameNoMenu") {
      select.add("FileLeafRef");
      select.add(`FieldValuesAsText/FileLeafRef`);
      continue;
    }

    // Allow underscore pseudo-fields even if they don't exist in fieldInfos
    const includeCol = fieldInfos.has(internal) || isUnderscoreField(internal);
    if (!includeCol) continue;

    // Map underscore fields to their OData names for REST
    const odata = toODataName(internal);

    // Always include the text projection for the field
    select.add(`FieldValuesAsText/${odata}`);

    const info = fieldInfos.get(internal);
    const t = info?.TypeAsString;

    if (t === "User" || t === "UserMulti") {
      // User fields use the original field name path for expand/select of subprops
      expand.add(internal);
      select.add(`${internal}/Title`);
      select.add(`${internal}/Id`);
      select.add(`${internal}/EMail`);
      continue;
    }

    if (t === "Lookup" || t === "LookupMulti") {
      expand.add(internal);
      const lookupDisplay = info?.LookupField || "Title";
      select.add(`${internal}/Id`);
      select.add(`${internal}/${lookupDisplay}`);
      continue;
    }

    if (t === "TaxonomyFieldType" || t === "TaxonomyFieldTypeMulti") {
      // Taxonomy: select both raw and text (raw is the non-OData internal, not underscore)
      select.add(internal);
      select.add(`FieldValuesAsText/${internal}`);
      continue;
    }

    // Other simple types (including underscore pseudo-fields)
    select.add(odata);
  }

  return { select: Array.from(select), expand: Array.from(expand) };
};

// Generic parsers/formatters

export const parseTaxonomyLabel = (val?: string): string => {
  if (!val) return "";
  // Common format: "Label|GUID;#Label2|GUID2"
  if (val.includes(";#") || val.includes("|")) {
    const parts = val.split(";#").map(p => p.split("|")[0]);
    return parts.filter(Boolean).join("; ");
  }
  return val;
};

export const getFileNameFromLeaf = (fileLeafRef?: string): string => {
  return fileLeafRef ?? "";
};

export const getFileExtensionFromLeaf = (fileLeafRef?: string): string => {
  if (!fileLeafRef) return "";
  const i = fileLeafRef.lastIndexOf(".");
  return i > -1 ? fileLeafRef.substring(i + 1).toLowerCase() : "";
};

// Batches list item fetches for the given IDs with the provided select/expand
export const batchGetItemsByIds = async (
  sp: SPFI,
  listId: string,
  ids: number[],
  select: string[],
  expand: string[]
): Promise<any[]> => {
  if (!ids.length) return [];

  // Use the batched SPFI to queue all requests in a single batch
  const [batched, exec] = sp.batched();
  const list = batched.web.lists.getById(listId);
  const promises: Promise<any>[] = [];

  for (const id of ids) {
    promises.push(
      list.items.getById(id).select(...select).expand(...expand)()
    );
  }

  await exec();
  const results = await Promise.all(promises);
  return results;
};

function chunk<T>(array: T[], size: number): T[][] {
  return Array.from({length: Math.ceil(array.length / size)}, (_, i) =>
    array.slice(i * size, i * size + size)
  );
}

export const getItemsById = async (
  sp: SPFI,
  listId: string,
  ids: number[],
  select: string[],
  expand: string[]
): Promise<any[]> => {
  if (!ids.length) return [];
  const filterString = ids.map(id => `Id eq ${id}`).join(" or ")
  const items = await sp.web.lists.getById(listId).items
    .filter(filterString)
    .select(...select)
    .expand(...expand)();
  return items;
}

// Hydrate search results (by ListItemId) with actual list data for the current view
export const hydrateSearchResultsForView = async (
  sp: SPFI,
  listId: string,
  viewColumns: IViewColumn[],
  searchRows: any[],
  fieldInfos: Map<string, IFieldInfo>
): Promise<{ items: any[]; fieldInfos: Map<string, IFieldInfo> }> => {
  // Keep fields that exist in fieldInfos OR underscore pseudo-fields (e.g., _UIVersionString)
  const filteredViewColumns = viewColumns.filter(col =>
    fieldInfos.has(col.fieldName) || isUnderscoreField(col.fieldName)
  );

  const { select, expand } = buildSelectExpand(filteredViewColumns, fieldInfos);

  const ids = searchRows.map(r => Number(r.ListItemId)).filter(n => !Number.isNaN(n));
  const MAX_BATCH = 1000;
  const idChunks = chunk(ids, MAX_BATCH);
  let listItems: any[] = [];
  for (const idChunk of idChunks) {
    const result = await getItemsById(sp, listId, idChunk, select, expand);
    listItems.push(...result);
  }

  // Mirror OData__ props back to underscore names on raw item and FieldValuesAsText
  for (const li of listItems) {
    mirrorUnderscoreProps(li);
    if (li.FieldValuesAsText) mirrorUnderscoreProps(li.FieldValuesAsText);
  }

  // Map by ID for quick merge
  const byId = new Map<number, any>();
  for (const li of listItems) byId.set(Number(li.Id), li);

  // Merge: prefer hydrated list fields for columns, but keep useful search props like Path/FileType
  const items = searchRows.map(sr => {
    const li = byId.get(Number(sr.ListItemId)) || {};
    return {
      ...sr,
      ...li,
      // Provide computed helpers
      __DocIcon: getFileExtensionFromLeaf(li.FileLeafRef || (sr.FileName ? `${sr.FileName}.${sr.FileExtension ?? ""}` : "")),
      __LinkFilename: getFileNameFromLeaf(li.FileLeafRef),
    };
  });

  return { items, fieldInfos };
};