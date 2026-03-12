import type { SPFI } from "@pnp/sp";
import type { IViewColumn, IFieldInfo } from "../models";
import { getTaxonomyText } from "./taxonomyDisplay";

export const fillMissingTaxonomyLabelsFromHiddenList = async (
  sp: SPFI,
  items: any[],
  viewColumns: IViewColumn[],
  fieldInfos: Map<string, IFieldInfo>
): Promise<any[]> => {
  const taxoFields = viewColumns
    .map(vc => ({ name: vc.fieldName, info: fieldInfos.get(vc.fieldName) }))
    .filter(x => x.info && (x.info!.TypeAsString === "TaxonomyFieldType" || x.info!.TypeAsString === "TaxonomyFieldTypeMulti"));

  if (!taxoFields.length || !items.length) return items;

  // Collect WssIds that still need resolution
  const wssToFetch = new Set<number>();
  for (const it of items) {
    for (const f of taxoFields) {
      const display = getTaxonomyText(it, f.name);
      if (display) continue;
      const raw = it?.[f.name];
      const add = (w: any) => {
        const n = Number(w);
        if (!Number.isNaN(n) && n > 0) wssToFetch.add(n);
      };
      if (Array.isArray(raw)) {
        raw.forEach(v => add(v?.WssId));
      } else if (raw && typeof raw === "object") {
        add((raw as any).WssId);
      }
    }
  }

  if (!wssToFetch.size) return items;

  const wssIdArray = Array.from(wssToFetch);
  const MAX_BATCH = 100;
  const idToTitle = new Map<number, string>();
  for (let i = 0; i < wssIdArray.length; i += MAX_BATCH) {
    const chunkIds = wssIdArray.slice(i, i + MAX_BATCH);
    const filter = chunkIds.map(id => `Id eq ${id}`).join(" or ");
    const thlItems = await sp.web.lists.getByTitle("TaxonomyHiddenList").items
      .filter(filter)
      .select("Id","Title")();

    for (const item of thlItems) {
      if (item.Title) idToTitle.set(Number(item.Id), item.Title as string)
    }
  }

  // Populate __TaxoResolved per item/field
  for (const it of items) {
    it.__TaxoResolved = it.__TaxoResolved || {};
    for (const f of taxoFields) {
      const already = getTaxonomyText(it, f.name);
      if (already) continue;
      const raw = it?.[f.name];
      let labels: string[] = [];
      if (Array.isArray(raw)) {
        labels = raw
          .map(v => idToTitle.get(Number(v?.WssId)))
          .filter((x): x is string => !!x);
      } else if (raw && typeof raw === "object") {
        const t = idToTitle.get(Number((raw as any).WssId));
        if (t) labels = [t];
      }
      if (labels.length) {
        it.__TaxoResolved[f.name] = labels.join("; ");
        // Optionally also set FieldValuesAsText to simplify renderers
        it.FieldValuesAsText = { ...(it.FieldValuesAsText ?? {}), [f.name]: it.__TaxoResolved[f.name] };
      }
    }
  }

  return items;
};