// More defensive parsing + helpers for taxonomy
const coerceToString = (val: unknown): string => {
  if (val == null) return "";
  if (typeof val === "string") return val;
  if (typeof val === "number" || typeof val === "boolean") return String(val);
  if (Array.isArray(val)) return val.map(coerceToString).join(";#");
  if (typeof val === "object") {
    const anyVal: any = val as any;
    if (Array.isArray(anyVal.results)) return anyVal.results.map(coerceToString).join(";#");
    try {
      const s = String(val);
      return s === "[object Object]" ? "" : s;
    } catch {
      return "";
    }
  }
  return "";
};

const isDigitsOnly = (s?: string) => !!s && /^\d+$/.test(s.trim());

export const parseHiddenNote = (note?: unknown): string[] => {
  const text = coerceToString(note);
  if (!text || isDigitsOnly(text)) return [];
  const parts = text.split(";#").filter(Boolean);
  return parts
    .map(p => {
      if (p.includes("|")) return p.split("|")[0];
      if (isDigitsOnly(p)) return "";
      return p;
    })
    .filter(Boolean);
};

// Resolve a displayable string from item data; no async calls here.
// A separate resolver can populate item.__TaxoResolved[internal] if needed.
export const getTaxonomyText = (item: any, internalName: string): string => {
  // 1) Prefer FieldValuesAsText if not just an ID
  const fvatStr = coerceToString(item?.FieldValuesAsText?.[internalName]).trim();
  if (fvatStr && !isDigitsOnly(fvatStr)) return fvatStr;

  // 2) Hidden note
  const hiddenRaw = item?.[`${internalName}_0`] ?? item?.FieldValuesAsText?.[`${internalName}_0`];
  const labelsFromNote = parseHiddenNote(hiddenRaw);
  if (labelsFromNote.length) return labelsFromNote.join("; ");

  // 3) Raw field
  const raw = item?.[internalName];
  if (Array.isArray(raw)) {
    const labels = raw
      .map(v => {
        if (!v) return "";
        if (typeof v === "object") {
          const lbl = (v as any).Label;
          const wss = (v as any).WssId;
          if (lbl != null && String(lbl) !== String(wss ?? "")) return String(lbl);
          return "";
        }
        if (typeof v === "string") return v.includes("|") ? v.split("|")[0] : v;
        return "";
      })
      .filter(Boolean);
    if (labels.length) return labels.join("; ");
  } else if (raw && typeof raw === "object") {
    const lbl = (raw as any).Label;
    const wss = (raw as any).WssId;
    if (lbl != null && String(lbl) !== String(wss ?? "")) return String(lbl);
  } else if (typeof raw === "string") {
    return raw.includes("|") ? raw.split("|")[0] : raw;
  }

  // A later pass may fill item.__TaxoResolved[internalName]
  return "";
};