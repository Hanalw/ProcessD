import { IFieldInfo } from "../models";

export type TaxoOption = { label: string; value: string }; // value = TermGuid

export const resolveKind = (fieldInfo?: IFieldInfo): 'text' | 'date' => {
  console.log(fieldInfo)
  if (!fieldInfo) return 'text';
  if (fieldInfo.TypeAsString === 'DateTime') return 'date';
  return 'text';
}

export const fieldType = (fieldInfo?: IFieldInfo): string => {
  if (!fieldInfo) return '';
  return fieldInfo.TypeAsString;
}

  export const getDistinctTaxonomyValues = (items: any[], fieldName: string): TaxoOption[] => {
  const seen = new Set<string>();
  const results: TaxoOption[] = [];

  for (const item of items) {
    // Get TermGuid from actual field value
    const taxoValue = item[fieldName];
    // Get label from FieldValuesAsText
    const label = item.FieldValuesAsText?.[fieldName];

    if (taxoValue) {
      // Support both single and multi taxonomy
      if (Array.isArray(taxoValue)) {
        taxoValue.forEach((taxoEntry: any) => {
          if (taxoEntry.TermGuid && label) {
            if (!seen.has(taxoEntry.TermGuid)) {
              results.push({ label, value: taxoEntry.TermGuid });
              seen.add(taxoEntry.TermGuid);
            }
          }
        });
      } else if (taxoValue.TermGuid && label) {
        if (!seen.has(taxoValue.TermGuid)) {
          results.push({ label, value: taxoValue.TermGuid });
          seen.add(taxoValue.TermGuid);
        }
      }
    }
  }
  return results;
}

export const getTaxonomyLabel = (fieldName: string, termGuid: string, docs: any[]): string => {
  for (const doc of docs) {
    // Taxonomy field can be single or array
    const taxoVal = doc[fieldName];
    if (Array.isArray(taxoVal)) {
      for (const entry of taxoVal) {
        if (entry.TermGuid === termGuid) {
          return doc.FieldValuesAsText?.[fieldName] || entry.Label || termGuid;
        }
      }
    } else if (taxoVal && taxoVal.TermGuid === termGuid) {
      return doc.FieldValuesAsText?.[fieldName] || taxoVal.Label || termGuid;
    }
  }
  return termGuid; // fallback
}

export const getUserIdAndName = (fieldName: string, selectedId: string, docs: any[]): { id: string, name: string } | string => {
  for (const doc of docs) {
    const userObj = doc[fieldName];
    if (Array.isArray(userObj)) {
      for (const entry of userObj) {
        if (String(entry.Id) === selectedId) {
          return { id: String(entry.Id), name: entry.Title || entry.Email || entry.Id };
        }
      }
    } else if (userObj && String(userObj.Id) === selectedId) {
      return { id: String(userObj.Id), name: userObj.Title || userObj.Email || userObj.Id };
    }
  }
  return '';
}
