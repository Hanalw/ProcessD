import { DynamicProperty } from "@microsoft/sp-component-base";
import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { DisplayMode } from '@microsoft/sp-core-library';
import { IColumn } from "@fluentui/react";
import { IFieldInfo } from "../../models";

export type TextFilterOp = 'contains' | 'equals';
export type DateFilterOp = 'onOrAfter' | 'onOrBefore';

export interface BaseFilter {
  fieldName: string;
  kind: 'text' | 'date';
}
export interface TextFilter extends BaseFilter {
  kind: 'text';
  op: TextFilterOp;
  value: string;
}
export interface DateFilter extends BaseFilter {
  kind: 'date';
  op: DateFilterOp;
  value: Date;
}

interface IContextMenuState {
  target: HTMLElement | null;
  column: IColumn | null;
  fieldInfo?: IFieldInfo | null;
}

export interface IFilterItem {
  columnName: string;
  value: string[];
}

export interface ContextMenuState {
  target: EventTarget & HTMLElement | null;
  column: IColumn | null;
  fieldInfo: IFieldInfo | null;
}

export interface IUnisightProcessDocumentsProps {
  selectedArea: DynamicProperty<any>;
  selectedTerm: any;
  context: WebPartContext;
  absoluteURL: string;
  tabsConfiguration: any[];
  // webpartTitle: string;
  // displayMode: DisplayMode;
  isConnected: boolean;
  // updateProperty: (value: string) => void;
  resetMessage?: () => void;
  termColumnName: string;
}

// TODO:  remove
export interface IUnisightProcessDocumentsState {
  selectedArea: any;
  searchedDocuments?: any[]; 
  searchedLinks?: any[];
  isSorted?: boolean;
  isloaded? : boolean;
  shouldLoad?: boolean;
  nonConfigured?: boolean;
  searchedSitePages?: any[]; 
  isSelectedAreaRecieved?: boolean;
  beskrivning?: string;
  ViewFields: any[];
  activeTab: string | undefined;
  noResultsWithSelectedTerm: boolean;
  contextMenu: IContextMenuState;
  // ephemeral menu UI state
  pendingOp: string | null;
  menuTextValue: string;
  menuDateValue: Date | null;
  sortDirection: 'asc' | 'desc' | null;
  fieldInfos?: Map<string, IFieldInfo>;
  isFilterPanelOpen: boolean;
  filters: string[];
  useFilters: any[];
}

export interface IExtendedState extends IUnisightProcessDocumentsState {
  searchedDocuments?: any[];
  searchedDocumentsFiltered?: any[];
  searchedSitePages?: any[];
  searchedLinks?: any[];
  contextMenu: IContextMenuState;
  // ephemeral menu UI state
  pendingOp: string | null;
  menuTextValue: string;
  menuDateValue: Date | null;
  fieldInfos?: Map<string, IFieldInfo>;
}
