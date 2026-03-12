import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ISourceFieldProps {
  context: WebPartContext;
  value: string | string[] | undefined;

  selectedType?: string;
  selectedSite?: string | string[] | undefined;
  
  placeholder?: string;
  label?: string;

  // Optional view selection support when selectedType === 'documents'
  selectedViewId?: string | undefined;
  
  viewPlaceholder?: string;

  
  
  sitePlaceholder?: string;
  onViewChanged?: (viewId: string | undefined) => void;
  onSiteChanged?: (site: string | string [] | undefined) => void;
  onChanged: (newValue: string | string[] | undefined) => void;
}