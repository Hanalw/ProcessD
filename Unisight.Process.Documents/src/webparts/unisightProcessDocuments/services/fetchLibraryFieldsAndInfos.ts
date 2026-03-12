import { SPFI } from '@pnp/sp';

import { IFieldInfo, IColumn } from '../models'
import {fieldInfoToColumn} from '../helpers/spViewHydrator';

// Cache field metadata per list to avoid repeated calls
const fieldInfoCache = new Map<string, Map<string, IFieldInfo>>();

export const fetchLibraryFieldsAndInfos = async (
  sp: SPFI,
  listId: string,
  opts?: { onlyViewFieldNames?: string[]}
): Promise<{ columns: IColumn[], fieldInfos: Map<string, IFieldInfo> }> => {
  if (fieldInfoCache.has(listId)) {
    // Columns can't be cached because each view might ask for different ones.
    return {
      columns: opts?.onlyViewFieldNames
        ? Array.from(fieldInfoCache.get(listId)!.values())
            .filter(f => opts.onlyViewFieldNames!.includes(f.InternalName))
            .map(f => fieldInfoToColumn(f))
        : Array.from(fieldInfoCache.get(listId)!.values()).map(f => fieldInfoToColumn(f)),
      fieldInfos: fieldInfoCache.get(listId)!
    };
  }

  // Fetch all fields for this list
  const rawFields: any[] = await sp.web.lists.getById(listId).fields
    .select("InternalName", "TypeAsString", "LookupField", "AllowMultipleValues", "Title", "Hidden")();

  // Only use visible fields
  let visibleFields = rawFields.filter(f => !f.Hidden);
  if (opts?.onlyViewFieldNames) {
    visibleFields = visibleFields.filter(f => opts.onlyViewFieldNames!.includes(f.InternalName));
  }

  // Build the fieldInfos map for downstream logic (hydration, taxonomy, etc.)
  const fieldInfos = new Map<string, IFieldInfo>();
  for (const f of visibleFields) {
    fieldInfos.set(f.InternalName, {
      InternalName: f.InternalName,
      TypeAsString: f.TypeAsString,
      LookupField: f.LookupField,
      AllowMultipleValues: f.AllowMultipleValues,
      Title: f.Title,
    });
  }
  // Cache for later
  fieldInfoCache.set(listId, fieldInfos);

  // Build the columns for UI table
  const columns: IColumn[] = visibleFields.map(fieldInfoToColumn);

  return { columns, fieldInfos };
};