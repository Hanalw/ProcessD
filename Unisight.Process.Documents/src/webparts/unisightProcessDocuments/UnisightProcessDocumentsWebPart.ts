import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDynamicField,
  PropertyPaneLabel,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { DynamicProperty } from '@microsoft/sp-component-base';
import * as strings from 'UnisightProcessDocumentsWebPartStrings';
import UnisightProcessDocuments from './components/UnisightProcessDocuments/UnisightProcessDocuments';
import { IUnisightProcessDocumentsProps } from './components/UnisightProcessDocuments/IUnisightProcessDocumentsProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { PropertyFieldTermPicker, IPickerTerms } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
import { getGraph, getSP } from './services';
import SourceField from './components/FormFields/SourceField';
import { fetchSiteColumns } from './services/fetchSiteColumns';

export interface IUnisightProcessDocumentsWebPartProps {
  selectedArea?: DynamicProperty<any>;
  termColumnName: string;
  collectionDataForTabs: any[];
  isConnected: boolean;
  selectedTerm?: IPickerTerms;
}

export default class UnisightProcessDocumentsWebPart extends BaseClientSideWebPart<IUnisightProcessDocumentsWebPartProps> {

  private _ppObserver?: MutationObserver;
  private _siteColumns: any[] = [];

  private _tagManageTabsButton = (): void => {
    try {
      const labels = [strings.ButtonLabelManageTabs, 'Manage Tabs']
        .filter(Boolean)
        .map(s => (s as string).trim().toLowerCase());

      const spans = Array.from(document.querySelectorAll('span'));
      for (const span of spans) {
        const txt = (span.textContent || '').trim().toLowerCase();
        if (labels.includes(txt)) {
          const btn = span.closest('button');
          if (btn && !btn.classList.contains('manage-tabs')) {
            btn.classList.add('manage-tabs');
          }
        }
      }
    } catch { /* noop */ }
  };

  protected onPropertyPaneConfigurationStart(): void {
    this._ppObserver?.disconnect();
    this._tagManageTabsButton();
    this._ppObserver = new MutationObserver(() => this._tagManageTabsButton());
    this._ppObserver.observe(document.body, { childList: true, subtree: true });
  }

  protected onPropertyPaneConfigurationComplete(): void {
    this._ppObserver?.disconnect();
    this._ppObserver = undefined;
  }

  public render(): void {
    const selectedAreaValue = this.properties.selectedArea?.tryGetValue() || {};
    const element: React.ReactElement<IUnisightProcessDocumentsProps> = React.createElement(
      UnisightProcessDocuments,
      {
        selectedTerm: this.properties.selectedTerm,
        selectedArea: selectedAreaValue,
        isConnected: this.properties.isConnected,
        context: this.context,
        absoluteURL: this.context.pageContext.site.absoluteUrl,
        tabsConfiguration: this.properties.collectionDataForTabs,
        termColumnName: this.properties.termColumnName,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    getSP(this.context);
    getGraph(this.context);
    this._siteColumns = await fetchSiteColumns(this.context.pageContext.site.absoluteUrl);
  }

  protected onDispose(): void {
    this._ppObserver?.disconnect();
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const siteColumns = this._siteColumns;

    const termColumns: IPropertyPaneDropdownOption[] = siteColumns.map(column => ({
      key: column.InternalName,
      text: column.Title
    }));

    const dynamicPropertySettings = this.properties.isConnected
      ? [
        PropertyPaneDynamicField("selectedArea", {
          label: " ",
        })
      ]
      : [];

    return {
      pages: [
        {
          groups: [
            {
              groupName: strings.PropertyPaneGroupTerms,
              groupFields: [
                PropertyPaneDropdown('termColumnName', {
                  label: strings.FieldLabelTermColumn,
                  options: termColumns,
                  selectedKey: this.properties.termColumnName
                }),
                PropertyFieldTermPicker('selectedTerm', {
                  label: strings.FieldLabelTermPicker,
                  panelTitle: strings.PanelTitleTermPicker,
                  initialValues: this.properties.selectedTerm,
                  allowMultipleSelections: false,
                  hideTermStoreName: true,
                  excludeSystemGroup: true,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as any,
                  deferredValidationTime: 0,
                  key: 'termPickerFieldId'
                })
              ]
            },
            {
              groupName: strings.PropertyPaneGroupConnectors,
              groupFields: [
                PropertyPaneLabel('selectedArea', {
                  text: strings.PropertyPaneLabelConnectors,
                }),
                PropertyPaneToggle('isConnected', {
                  onText: strings.Yes,
                  offText: strings.No,
                  checked: this.properties.isConnected,
                }),
                ...dynamicPropertySettings
              ]
            },
            {
              groupName: strings.PropertyPaneGroupTabs,
              groupFields: [
                PropertyFieldCollectionData("collectionDataForTabs", {
                  key: "collectionDataForTabs",
                  label: "",
                  panelHeader: strings.PanelTitleManageTabs,
                  manageBtnLabel: strings.ButtonLabelManageTabs,
                  value: this.properties.collectionDataForTabs,
                  disabled: false,
                  fields: [
                    {
                      id: "OrderOfTab",
                      title: strings.TabTitleOrder,
                      type: CustomCollectionFieldType.number,
                      required: true
                    },
                    {
                      id: "TabName",
                      title: strings.TabTitleName,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "Type",
                      title: strings.TabTitleType,
                      type: CustomCollectionFieldType.dropdown,
                      required: true,
                      options: [
                        { key: 'description', text: strings.TypeTextDescription },
                        { key: 'documents', text: strings.TypeTextDocument },
                        { key: 'pages', text: strings.TypeTextPages },
                        { key: 'links', text: strings.TypeTextLinks },
                        { key: 'list', text: strings.TypeTextList }
                      ]
                    },
                    {
                      id: 'Source',
                      title: strings.TabTitleSource,
                      type: CustomCollectionFieldType.custom,
                      required: true,
                      onCustomRender: (field, value, onUpdate, item /* current row */) => {
                        return React.createElement(SourceField, {
                          context: this.context,
                          value: value as string | string[],

                          selectedType: item?.Type as string,

                          selectedSite: item?.SelectedSite as string | string[] | undefined,
                          onSiteChanged: (site) => onUpdate('SelectedSite', site),

                          selectedViewId: item?.SelectedViewId as string | undefined,
                          viewPlaceholder: 'Välj vy...',
                          onViewChanged: (viewId?: string) => {
                            onUpdate('SelectedViewId', viewId);
                          },

                          onChanged: (newValue?: string | string[]) => {
                            onUpdate(field.id, newValue);
                          }
                        });
                      }
                    },
                    {
                      id: 'SelectedSite',
                      title: '',
                      type: CustomCollectionFieldType.custom,
                      required: false,
                      onCustomRender: (field, value, onUpdate, item) => {
                        return React.createElement('input', {
                          type: 'hidden',
                          value: item?.SelectedSite || '',
                          readOnly: true,
                          'aria-hidden': true
                        });
                      }
                    },
                    {
                      id: 'SelectedViewId',
                      title: '',
                      type: CustomCollectionFieldType.custom,
                      required: false,
                      onCustomRender: (field, value, onUpdate, item /* current row */) => {
                        return React.createElement('input', {
                          type: 'hidden',
                          value: item?.SelectedViewId || '',
                          readOnly: true,
                          'aria-hidden': true
                        });
                      }
                    },
                    {
                      id: "Icon",
                      title: strings.TabTitleIcon,
                      type: CustomCollectionFieldType.fabricIcon,
                      iconFieldRenderMode: 'picker',
                    },
                    {
                      id: "ShowIcon",
                      title: strings.TabTitleShowIcon,
                      type: CustomCollectionFieldType.boolean,
                      required: false,
                      defaultValue: false
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}