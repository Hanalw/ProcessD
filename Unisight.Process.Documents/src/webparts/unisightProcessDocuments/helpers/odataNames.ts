export const isUnderscoreField = (name?: string): boolean =>
  !!name && name.startsWith("_");

export const toODataName = (internalName: string): string =>
  isUnderscoreField(internalName)
    ? `OData__${internalName.substring(1)}`
    : internalName;

export const mirrorUnderscoreProps = (obj: any): void => {
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