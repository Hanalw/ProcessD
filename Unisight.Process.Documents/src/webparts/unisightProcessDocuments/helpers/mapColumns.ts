import {parseTaxonomyLabel} from './spViewHydrator'

// Build filename from FileName/FileExtension or Path
export const getLeafFromPath = (path: string, fileName?: string, fileExt?: string) => {
  if (fileName && fileExt) return `${fileName}.${fileExt}`;
  try {
    const url = new URL(path);
    const last = decodeURIComponent(url.pathname.split("/").pop() || "");
    return last;
  } catch {
    const noQuery = path.split(/[?#]/)[0];
    return decodeURIComponent(noQuery.split("/").pop() || "");
  }
};

// Taxonomy in search often comes as "Label|GUID;#Label|GUID"
// or sometimes just "Label". Return a readable string.
// export const parseTaxonomyLabel = (val?: string) => {
//   if (!val) return "";
//   // Multi-value: "Label|GUID;#Label2|GUID2"
//   if (val.includes(";#") || val.includes("|")) {
//     console.log(val)
//     const parts = val.split(";#").map(p => p.split("|")[0]);
//     return parts.filter(Boolean).join("; ");
//   }
//   return val; // already a simple label
// };

// Map a single search result row -> object with your view's field names
export const mapSearchRowToViewItem = (r: any, taxonomyColumn: string) => {
  const leaf = getLeafFromPath(r.Path, r.FileName, r.FileExtension);

  // Try a few candidates for the taxonomy label
  const taxonomyRaw =
    r[taxonomyColumn] ??            // if you have a custom managed property mapped to the label
    r[`ows${taxonomyColumn}`] ?? // sometimes ows<Field> carries label|guid
    r[`owstaxId${taxonomyColumn}`] ??    // often present, may be ids or "Label|Guid"
    "";

  return {
    ...r,
    // Columns your DetailsList expects:
    DocIcon: r.FileType || r.FileExtension || "",     // used for icon
    LinkFilename: leaf,                                // "Namn"
    Title: r.Title ?? "",
    [taxonomyColumn]: parseTaxonomyLabel(taxonomyRaw),       // "Process"
    Modified: r.LastModifiedTime ?? r.Write ?? "",     // "Ändrat" (date)
    Editor: r.Editor ?? "",
    Author: r.Author ?? "",
  };
};