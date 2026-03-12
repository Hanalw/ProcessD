// import LocalizedStrings, { LocalizedStringsMethods } from "react-localization";
// import {svStrings} from './sv-se';
// import {enStrings} from './en-us';

declare interface IUnisightProcessDocumentsWebPartStrings {
  ConfigureWebPart: string;
  EmptyConfigure: string;
  EmptyData: string;
  EmptyDescription: string;
  WebPartTitle: string;
  PropertyPaneGroupTerms: string;
  PropertyPaneGroupConnectors: string;
  PropertyPaneGroupTabs: string;
  PropertyPaneGroupColumns: string;
  PropertyPaneLabelConnectors: string;
  Yes: string;
  No: string;
  loading: string;
  PropertyPaneGroupList: string;
  FieldLabelProcessObjectList: string;
  FieldLabelTermPicker: string;
  FieldLabelTermColumn: string;
  PanelTitleTermPicker: string;
  ButtonLabelManageTabs: string;
  PanelTitleManageTabs: string;
  TabTitleOrder: string;
  TabTitleName: string;
  TabTitleIcon: string;
  TabTitleShowIcon: string;
  TabTitleSource: string;
  TabTitleType: string;
  TabSourceEmpty: string;
  TypeTextDescription: string;
  TypeTextDocument: string;
  TypeTextList: string;
  TypeTextPages: string;
  TypeTextLinks: string;
  connectedEmpty: string;
  notConnectedEmpty: string;
  noData: string;
  sortdesc: string;
  sortasc: string;
  oldToNew: string;
  newToOld: string;
  filter: string;
  ColumnNameTitle: string;
  ColumnNameDescription: string;
}

// export const strings = new LocalizedStrings({
//   en: enStrings,
//   sv: svStrings
// })

declare module 'UnisightProcessDocumentsWebPartStrings' {
  const strings: IUnisightProcessDocumentsWebPartStrings;
  export = strings;
}