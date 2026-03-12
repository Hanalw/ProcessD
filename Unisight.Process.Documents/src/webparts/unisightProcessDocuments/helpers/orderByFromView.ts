export const NAME_FIELDS = ["FileLeafRef", "LinkFilename", "LinkFilenameNoMenu", "FileName"];
export const NAME_FIELDS_LOWER = new Set(NAME_FIELDS.map((n: string) => n.toLowerCase()));

export interface ParsedByOrder {
  field: string;
  ascending: boolean;
}

export const parseOrderBy = (viewQuery: string): ParsedByOrder[] => {
  if (!viewQuery) return [];
  const xml = `<View>${viewQuery}</View>`
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const orderBy = doc.getElementsByTagName('OrderBy')[0];
  if (!orderBy) return [];
  const refs = Array.from(orderBy.getElementsByTagName('FieldRef'));
  return refs
    .map(r => {
      const name = r.getAttribute('Name') || '';
      const ascAttr = r.getAttribute('Ascending');
      const ascending = ascAttr ? ascAttr.toUpperCase() !== 'FALSE' : true;
      return {field: name, ascending};
    })
    .filter(r => r.field);
};

const getLastSegment = (url: string) => {
  try {
    const clean = url.split("?")[0].split("#")[0];
    const parts = clean.split("/");
    return parts[parts.length - 1] || "";
  } catch {
    return "";
  }
}

export const normalizeNameFields = (item: any) => {
  const fileLeafRef = item?.FileLeafRef ?? item?.fileLeafRef;
  const linkFilename = item?.LinkFilename ?? item?.linkFilename;
  const linkFilenameNoMenu = item?.LinkFilenameNoMenu ?? item?.linkFilenameNoMenu;
  const fileName = item?.FileName ?? item?.fileName;
  const path = item?.Path ?? item?.path;

  let name = fileLeafRef ?? linkFilename ?? linkFilenameNoMenu ?? fileName;
  if (!name && path) name = getLastSegment(path);

  if (name) item.FileLeafRef = name;
  if (!linkFilename) item.LinkFilename = item.FileLeafRef;
  if (!linkFilenameNoMenu) item.LinkFilenameNoMenu = item.FileLeafRef;
  return item;
}

export interface SortDef {
  key?: string;
  originalKey?: string;
  ascending: boolean;
  originalWasName: boolean;
}

export const getSortDefFromView = (viewQuery: string): SortDef => {
  const parts = parseOrderBy(viewQuery);
  if (!parts.length) return {key: undefined, originalKey: undefined, ascending: true, originalWasName: false};

  const primary = parts[0];
  const orginalLower = primary.field.toLowerCase();
  const originalWasName = NAME_FIELDS_LOWER.has(orginalLower);

  return {
    key: originalWasName ? "FileLeafRef" : primary.field,
    originalKey: primary.field,
    ascending: primary.ascending,
    originalWasName
  }
}

export const buildComparator = (key: string, ascending: boolean) => {
  const dir = ascending ? 1 : -1;
  return (a: any, b: any) => {
    const av = a?.[key];
    const bv = b?.[key];

    if (av === null && bv === null) return 0;
    if (av === null) return 1;
    if (bv === null) return -1;

    const an = typeof av === 'number' ? av : Number(av);
    const bn = typeof bv === 'number' ? bv : Number(bv);
    const anIsNum = !isNaN(an);
    const bnIsNum = !isNaN(bn);

    if (anIsNum && bnIsNum) return (an - bn) * dir;

    const ad = new Date(av);
    const bd = new Date(bv);
    const adIsDate = !isNaN(ad.getTime());
    const bdIsDate = !isNaN(bd.getTime());
    if (adIsDate && bdIsDate) return (ad.getTime() - bd.getTime()) * dir;

    const as = String(av).toLowerCase();
    const bs = String(bv).toLowerCase();
    return as < bs ? -1 * dir : as > bs ? 1 * dir : 0;
  }
}