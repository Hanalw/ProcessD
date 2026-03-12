export interface IViewColumn {
  key: string;
  name: string;
  fieldName: string; // internal name from the view
  minWidth?: number;
  maxWidth?: number;
  data: string;
}