export interface IFieldInfo {
  InternalName: string;
  TypeAsString: string;        // e.g., "Text", "Note", "User", "UserMulti", "Lookup", "LookupMulti", "TaxonomyFieldType", "TaxonomyFieldTypeMulti", "DateTime", "Number", "Boolean", "Computed", ...
  LookupField?: string;        // for Lookup fields, which display field to use (defaults to Title)
  AllowMultipleValues?: boolean;
  Title?: string;
}