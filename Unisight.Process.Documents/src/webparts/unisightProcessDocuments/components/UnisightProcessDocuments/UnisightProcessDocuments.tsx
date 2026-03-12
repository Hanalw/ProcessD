import * as React from 'react';
import {useState, useRef, useEffect, useCallback, MouseEvent, FC} from 'react';
import styles from './UnisightProcessDocuments.module.scss';
import type { 
  IUnisightProcessDocumentsProps, 
  ContextMenuState
} from './IUnisightProcessDocumentsProps';
import {
  initializeIcons,
  Pivot, PivotItem, Spinner, MessageBarType, MessageBar,
  DetailsList, DetailsListLayoutMode, SelectionMode,
  IColumn, Icon, Link, 
  TooltipHost,
  ContextualMenu, IContextualMenuItem
} from '@fluentui/react';
import { initializeFileTypeIcons, getFileTypeIconProps } from '@uifabric/file-type-icons';
import { SPFI } from '@pnp/sp';
import {  
  getSPForWeb,
  fetchLibraryFields, 
  fetchLibraryViewFields,
  fetchLibraryFieldsAndInfos
 } from '../../services';

import { buildRenderableColumns } from "../FormFields/BuildColumns";
import { fillMissingTaxonomyLabelsFromHiddenList } from "../../helpers/taxonomyResolvers";
import type { ISearchResult } from '@pnp/sp/search';
import * as strings from 'UnisightProcessDocumentsWebPartStrings';
import FilterPanel from '../FormFields/FilterPanel';
import {trimText} from '../../helpers/textHelpers';

import {IFieldInfo, IViewColumn} from '../../models';
import {hydrateSearchResultsForView} from "../../helpers/spViewHydrator";
import {getDistinctTaxonomyValues, getTaxonomyLabel, getUserIdAndName, resolveKind, TaxoOption, fieldType} from '../../helpers/filterHelpers';

// sort
import { NAME_FIELDS_LOWER, getSortDefFromView, normalizeNameFields, buildComparator } from '../../helpers';

type ISearchResultWithListIds = ISearchResult & {
  ListItemId?: string;
  ListId?: string;
  UniqueId?: string;
};

// type ViewColumn = IColumn & { nameType?: string};

initializeIcons();
initializeFileTypeIcons()

const UnisightProcessDocuments: FC<IUnisightProcessDocumentsProps> = ({
  selectedTerm,
  selectedArea,
  isConnected,
  termColumnName,
  absoluteURL,
  tabsConfiguration,
  context,
  resetMessage
}) => {
  // --- Persistent values ---
  const _allDocuments = useRef<any[]>([]);
  const _filterColumn = useRef<IColumn | null>(null);
  const _filterFieldInfo = useRef<IFieldInfo | null>(null);
  const _hoveredMenuKey = useRef<string | null>(null);

  // --- State ---
  const [, setChosenArea] = useState<any>(null);
  const [isloaded, setIsLoaded] = useState(false); // page-level: active tab loaded
  const [isConfigured, setIsConfigured] = useState(false);
  const [ViewFields, setViewFields] = useState<IColumn[]>([]);
  const [activeTab, setActiveTab] = useState<string>('');
  const [beskrivningar, setBeskrivningar] = useState<Record<string, string>>({}); // key is tab.Type
  const setBeskrivning = useCallback((tabName: string, description: string) => {
    setBeskrivningar(prev => ({
      ...prev,
      [tabName]: description,
    }));
  }, []);
  const [noResultsWithSelectedTerm, setNoResultsWithSelectedTerm] = useState(false);
  const [searchedDocuments, setSearchedDocuments] = useState<any[]>([]);
  const [searchedListItems, setSearchedListItems] = useState<any[]>([]);
  const [searchedSitePages, setSearchedSitePages] = useState<any[]>([]);
  const [searchedLinks, setSearchedLinks] = useState<any[]>([]);
  const [filters, ] = useState<any[]>([]);
  const [contextMenu, setContextMenu] = useState<ContextMenuState>({ target: null, column: null, fieldInfo: null });
  const [pendingOp, setPendingOp] = useState<string | null>(null);
  const [, setMenuTextValue] = useState('');
  const [, setMenuDateValue] = useState<any>(null);
  const [, setSortDirection] = useState<'asc' | 'desc' | null>(null);
  const [isFilterPanelOpen, setIsFilterPanelOpen] = useState(false);
  const [useFilters, setUseFilters] = useState<any[]>([]);
  const [, setForceUpdate] = useState(0);

  // Per-tab load state
  // key is tab.Type: 'documents' | 'pages' | 'links' | 'description'
  const [tabLoadState, setTabLoadState] = useState<Record<string, { loading: boolean; loaded: boolean }>>({});
  const tabLoadStateRef = useRef(tabLoadState);
  useEffect(() => { tabLoadStateRef.current = tabLoadState; }, [tabLoadState]);

  const forceUpdate = useCallback(() => setForceUpdate(n => n + 1), []);

  const resolveSelectedTermId = (): string | null => {
    const selectedTermfromArea = (selectedArea as any)?.term;
    const selectedTermfromAreaId = selectedTermfromArea ? selectedTermfromArea[0]?.id : null;
    const selectedTermArr = (selectedTerm as any[]) ?? [];
    const selectedTermId = selectedTermArr.length > 0 ? selectedTermArr[0]?.key : null;
    const isSelectedAreaEmpty =
      selectedArea == null ||
      (typeof selectedArea === 'object' && Object.keys(selectedArea).length === 0);

    let termId: string | null = null;
    if (isConnected) {
      if (!isSelectedAreaEmpty && selectedTermfromAreaId) termId = selectedTermfromAreaId;
      else if (isSelectedAreaEmpty && selectedTermId) termId = selectedTermId;
    } else {
      termId = selectedTermId;
    }
    return termId ?? null;
  }

  // Checks if we have configuration and a selected term/area to search with
  useEffect(() => {
    const hasSelectedTerm = !!selectedTerm && selectedTerm.length > 0;
    const hasSelectedArea = !!selectedArea && typeof selectedArea === "object" && Object.keys(selectedArea).length > 0;
    setIsConfigured(hasSelectedTerm || hasSelectedArea);
  }, [selectedTerm, selectedArea]);

  // Initialize tabLoadState when tabsConfiguration changes
  useEffect(() => {
    const tabs = tabsConfiguration ?? [];
    if (!Array.isArray(tabs)) return;
    setTabLoadState(prev => {
      const next: Record<string, { loading: boolean; loaded: boolean }> = {};
      for (const t of tabs) {
        const type = t.Type;
        next[type] = prev[type] ?? { loading: false, loaded: false };
      }
      return next;
    });
  }, [tabsConfiguration]);

  // Context menu handlers
  const onOpenColumnMenu = useCallback((ev: MouseEvent<HTMLElement>, column: IColumn, fieldInfo?: IFieldInfo) => {
    setContextMenu({ target: ev.currentTarget, column, fieldInfo: fieldInfo || null });
    setMenuTextValue('');
    setMenuDateValue(null);
  },[]);

  const closeMenu = () => {
    setContextMenu({ target: null, column: null, fieldInfo: null });
    setMenuDateValue(null);
    setMenuTextValue('');
  };

  // Tabs helpers
  const getItems = (source: any) => {
    switch (source) {
      case "documents":
        return searchedDocuments;
      case "list":
        return searchedListItems;
      case "pages":
        return searchedSitePages;
      case "links":
        return searchedLinks;
      case "description":
        return null;
      default:
        return null;
    }
  }

  // Sorting function (documents only)
  const sortDocs = (sortKey: string, fieldName?: string) => {
    if (!fieldName) return;
    setPendingOp(`${fieldName}|${sortKey}`);

    const isDateSort = sortKey === 'oldToNew' || sortKey === 'newToOld';
    const sortFieldType = fieldType(contextMenu.fieldInfo || undefined);
    const compare = (a: any, b: any) => {
      let av = fieldName === 'LinkFilename' ? a?.FieldValuesAsText?.FileLeafRef : a?.FieldValuesAsText?.[fieldName];
      let bv = fieldName === 'LinkFilename' ? b?.FieldValuesAsText?.FileLeafRef : b?.FieldValuesAsText?.[fieldName];
      
      if (sortFieldType === 'User') {
        av = a?.[fieldName].Title;
        bv = b?.[fieldName].Title;
        const astr = (av ?? '').toString().toLowerCase();
        const bstr = (bv ?? '').toString().toLowerCase();
        return sortKey === 'sortasc' ? astr.localeCompare(bstr) : bstr.localeCompare(astr);
      }

      if (isDateSort) {
        const ad = new Date(av);
        const bd = new Date(bv);
        return sortKey === 'oldToNew' ? ad.getTime() - bd.getTime() : bd.getTime() - ad.getTime();
      } else {
        const astr = (av ?? '').toString().toLowerCase();
        const bstr = (bv ?? '').toString().toLowerCase();
        return sortKey === 'sortasc' ? astr.localeCompare(bstr) : bstr.localeCompare(astr);
      }
    }

    const newDocs = [...(searchedDocuments ?? [])].sort(compare);
    let direction: 'asc' | 'desc' | null = null;
    if (sortKey === 'oldToNew' || sortKey === 'sortasc') direction = 'asc';
    else if (sortKey === 'newToOld' || sortKey === 'sortdesc') direction = 'desc';

    const updatedViewFields = (ViewFields ?? []).map(col => {
      if ((col as any).fieldName === fieldName) {
        return {
          ...col,
          isSorted: true,
          isSortedDescending: direction === 'desc'
        };
      }
      const { isSorted, isSortedDescending, ...rest } = col as any;
      
      return { ...rest, isSorted: false };
    });
    setSearchedDocuments(newDocs);
    setSortDirection(direction);
    setViewFields(updatedViewFields);
    closeMenu();
  }

   // Sort reapply if documents view indicates a sorted column
  const applyPersistedSort = useCallback(() => {
    if (pendingOp) return;
    const sortedCol = (ViewFields || []).find(col => col.isSorted === true);
    if (sortedCol && sortedCol.fieldName) {
      setViewFields([...ViewFields]);
      sortDocs(sortedCol.isSortedDescending ? 'sortdesc' : 'sortasc', sortedCol.fieldName);
    }
  },[pendingOp, ViewFields]);

  const reportFirstTabType = (): string | undefined => {
    const sourceTabs = tabsConfiguration ?? [];
    if (!Array.isArray(sourceTabs) || sourceTabs.length === 0) return undefined;
    const tabConfiguration = [...sourceTabs].sort((a: any, b: any) => (a.OrderOfTab ?? 0) - (b.OrderOfTab ?? 0));
    if (!tabConfiguration.length) return undefined;
    const first = tabConfiguration[0];
    return first.Type;
  };

  // -------------------- Data fetching (per tab) --------------------
  const markTabLoading = (type: string) => {
    setTabLoadState(prev => ({ ...prev, [type]: { ...(prev[type] || { loading: false, loaded: false }), loading: true } }));
  };
  const markTabLoaded = (type: string) => {
    setTabLoadState(prev => ({ ...prev, [type]: { ...(prev[type] || { loading: false, loaded: false }), loading: false, loaded: true } }));
  };
  const markTabFailed = (type: string) => {
    setTabLoadState(prev => ({ ...prev, [type]: { ...(prev[type] || { loading: false, loaded: false }), loading: false, loaded: true } }));
  };

  const fetchDocuments = useCallback(async (term: string) => {
    const tabs = tabsConfiguration ?? [];
    const docTab = tabs.find((tab: any) => tab.Type === "documents");

    const libraryId = docTab?.Source;
    const viewId = docTab?.SelectedViewId;
    const siteUrl = docTab?.SelectedSite;

    if (!libraryId || !viewId || !siteUrl) {
      setSearchedDocuments([]);
      setViewFields([]);
      _allDocuments.current = [];
      return;
    }

    const spForSite = getSPForWeb(siteUrl);

    // 1) Get view columns and view query (for OrderBy)
    const { columns: viewDisplayColumns, viewQuery } = 
      await fetchLibraryViewFields(spForSite, libraryId, viewId);

    // 2) Resolve canonical sort definition from the view
    const sortDef = getSortDefFromView(viewQuery);
    if (!sortDef.key) {
      sortDef.key = "FileLeafRef";
      sortDef.originalKey = "FileLeafRef";
      sortDef.originalWasName = true;
      sortDef.ascending = true;
    }
    console.log("Column to sort: ", sortDef)

    // 3) Build view columns (map any "name" variant to FileLeafRef, preserve nameType)
    const viewColumns: IViewColumn[] = viewDisplayColumns.map(col => {
      const orig = col.fieldName ?? col.key ?? "";
      const isName = !!orig && NAME_FIELDS_LOWER.has(orig.toLowerCase());
      if (isName) {
        return {
          ...col,
          fieldName: "FileLeafRef",
          nameType: orig,
          isSorted: false,
          isSortedDescending: false
        } as IViewColumn;
      }
      return {
        ...col,
        fieldName: orig,
        isSorted: false,
        isSortedDescending: false
      } as IViewColumn;
    });

    // 4) Search (no stray asterisk before Path)
    const docSearchQuery = `Path:${siteUrl} AND owstaxId${termColumnName}:GP0|#${term} AND IsDocument:1 AND ListId:${libraryId} AND -FileExtension:aspx`;
    const docSearchResults = await spForSite.search({
      Querytext: docSearchQuery,
      SelectProperties: [
        "Title","Path","ListId","ListItemId","UniqueId","FileType","FileExtension",
        "FileName","FileLeafRef","LinkFilename","LinkFilenameNoMenu",
        "LastModifiedTime","Editor","Author","NormUniqueID","SPWebUrl"
      ],
      RowLimit: 500,
      TrimDuplicates: false
    });
    console.log('docQuery: ', docSearchQuery)

    // 5) Ensure hydration includes the sort field even if it's not visible
    const viewFieldInternalNames = viewColumns.map(f => f.fieldName!).filter(Boolean);
    const fieldsForHydrate = Array.from(new Set([
      ...viewFieldInternalNames,
      "FileLeafRef",
      ...(sortDef.originalWasName ? [] : [sortDef.originalKey!])
    ]));

    const { fieldInfos } = await fetchLibraryFieldsAndInfos(spForSite, libraryId, {
      onlyViewFieldNames: fieldsForHydrate
    });

    // 6) Normalize, hydrate, taxonomy resolution
    const docPrimary = docSearchResults.PrimarySearchResults || [];
    const normalizedDocs = docPrimary.map(normalizeNameFields);

    const hydrated = await hydrateSearchResultsForView(
      spForSite, libraryId, viewColumns, normalizedDocs, fieldInfos
    );
    const withResolvedTaxo = await fillMissingTaxonomyLabelsFromHiddenList(
      spForSite, hydrated.items, viewColumns, hydrated.fieldInfos
    );

    // 7) Format DateTime fields (fix: use col.data, not col.col.data)
    const formatDate = (dateStr: string) => {
      if (!dateStr) return dateStr;
      const d = new Date(dateStr);
      if (isNaN(d.getTime())) return dateStr;
      return d.toISOString().slice(0, 10);
    };
    const dateFields: string[] = viewColumns
      .filter(col => (col as any).data === "DateTime")
      .map(col => col.fieldName!);

    const formattedDocs = withResolvedTaxo.map(item => {
      const newItem = { ...item };
      dateFields.forEach(field => {
        if (newItem[field]) newItem[field] = formatDate(newItem[field]);
      });
      return newItem;
    });

    // 8) Build renderable columns for DetailsList
    let builtColumns = buildRenderableColumns(viewColumns, hydrated.fieldInfos, {
      onOpenMenu: onOpenColumnMenu,
      activeFilters: filters
    });

    // Mark sorted column if visible (UI indication only)
    const originalLower = (sortDef.originalKey || '').toLowerCase();
    const sortedIsName = sortDef.originalWasName;

    builtColumns = builtColumns.map((col: any) => {
      let isSorted = false;
      let isSortedDescending = false;

      if (sortedIsName) {
        const nameType = (col as any).nameType?.toLowerCase?.();
        isSorted = col.fieldName === "FileLeafRef" && nameType === originalLower;
      } else {
        isSorted = col.fieldName === sortDef.key;
      }
      if (isSorted) isSortedDescending = !sortDef.ascending;

      return { ...col, isSorted, isSortedDescending };
    });

    // 9) Sort data regardless of visibility of the sort field
    const comparator = buildComparator(sortDef.key!, sortDef.ascending);
    const initiallySortedDocs = [...formattedDocs].sort(comparator);

    setSearchedDocuments(initiallySortedDocs);
    setViewFields(builtColumns);
    _allDocuments.current = initiallySortedDocs;

    // Apply persisted sort (user override)
    applyPersistedSort?.();
  }, [onOpenColumnMenu, filters, applyPersistedSort]);

  useEffect(() => {
    console.log("Documents: ", _allDocuments.current);
  }, [_allDocuments.current]);

  const fetchList = useCallback(async (term: string) => {
  const tabs = tabsConfiguration ?? [];
  const listTab = tabs.find((tab: any) => tab.Type === "list");

  const listId = listTab?.Source;
  const viewId = listTab?.SelectedViewId;
  const siteUrl = listTab?.SelectedSite;

  if (!listId || !viewId || !siteUrl) {
    setSearchedListItems([]);
    setViewFields([]);
    _allDocuments.current = [];
    return;
  }

  const spForSite = getSPForWeb(siteUrl);

  // 1) Get view columns and view query (for OrderBy)
  const { columns: viewDisplayColumns, viewQuery } =
    await fetchLibraryViewFields(spForSite, listId, viewId);

  // 2) Resolve canonical sort definition from the view
  const sortDef = getSortDefFromView(viewQuery);
  if (!sortDef.key) {
    // Fallback: Title is common for lists; if not present, just use "Title" anyway
    sortDef.key = "Title";
    sortDef.originalKey = "Title";
    sortDef.originalWasName = false;
    sortDef.ascending = true;
  }

  // 3) Build view columns (keep as-is; name-field normalization is mainly for docs)
  const viewColumns: IViewColumn[] = viewDisplayColumns.map(col => {
    const orig = col.fieldName ?? col.key ?? "";
    return {
      ...col,
      fieldName: orig,
      isSorted: false,
      isSortedDescending: false
    } as IViewColumn;
  });

  // 4) Search list items by ListId and taxonomy term
  // NOTE: do NOT include IsDocument:1, and do NOT exclude aspx
  // Keep Path scoped to siteUrl to avoid cross-site noise.
  const listSearchQuery =
    `Path:${siteUrl} AND owstaxId${termColumnName}:GP0|#${term} AND ListId:${listId}`;

  const listSearchResults = await spForSite.search({
    Querytext: listSearchQuery,
    SelectProperties: [
      "Title", "Path", "ListId", "ListItemId", "UniqueId",
      "LastModifiedTime", "Editor", "Author",
      "SPWebUrl"
    ],
    RowLimit: 500,
    TrimDuplicates: false
  });

  console.log(listSearchQuery);
  console.log("List Search Results: ", listSearchResults.PrimarySearchResults);

  // 5) Ensure hydration includes the sort field even if it's not visible
  const viewFieldInternalNames = viewColumns.map(f => f.fieldName!).filter(Boolean);
  console.log(viewFieldInternalNames)
  const fieldsForHydrate = Array.from(new Set([
    ...viewFieldInternalNames,
    ...(sortDef.originalKey ? [sortDef.originalKey] : [])
  ]));

  const { fieldInfos } = await fetchLibraryFieldsAndInfos(spForSite, listId, {
    onlyViewFieldNames: fieldsForHydrate
  });

  // 6) Hydrate + taxonomy resolve
  const primary = listSearchResults.PrimarySearchResults || [];

  const ids = primary.map((r: any) => Number(r.ListItemId ?? r.ListItemID));
  console.log("List ids: ", ids)

  const hydrated = await hydrateSearchResultsForView(
    spForSite, listId, viewColumns, primary, fieldInfos
  );

  const withResolvedTaxo = await fillMissingTaxonomyLabelsFromHiddenList(
    spForSite, hydrated.items, viewColumns, hydrated.fieldInfos
  );

  // 7) Format DateTime fields
  const formatDate = (dateStr: string) => {
    if (!dateStr) return dateStr;
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toISOString().slice(0, 10);
  };

  const dateFields: string[] = viewColumns
    .filter(col => (col as any).data === "DateTime")
    .map(col => col.fieldName!);

  const formattedItems = withResolvedTaxo.map(item => {
    const newItem = { ...item };
    dateFields.forEach(field => {
      if (newItem[field]) newItem[field] = formatDate(newItem[field]);
    });
    return newItem;
  });

  // 8) Build columns
  let builtColumns = buildRenderableColumns(viewColumns, hydrated.fieldInfos, {
    onOpenMenu: onOpenColumnMenu,
    activeFilters: filters
  });

  // mark sorted column if it matches sortDef
  builtColumns = builtColumns.map((col: any) => {
    const isSorted = col.fieldName === sortDef.key;
    const isSortedDescending = isSorted ? !sortDef.ascending : false;
    return { ...col, isSorted, isSortedDescending };
  });

  // 9) Initial sort
  const comparator = buildComparator(sortDef.key!, sortDef.ascending);
  const initiallySorted = [...formattedItems].sort(comparator);

  setSearchedListItems(initiallySorted);
  setViewFields(builtColumns);
  _allDocuments.current = initiallySorted;

  applyPersistedSort?.();
}, [tabsConfiguration, termColumnName, filters, onOpenColumnMenu, applyPersistedSort]);

  const fetchPages = useCallback(async (sp: SPFI, term: string) => {
    const tabs = tabsConfiguration ?? [];
    const pagesTab = tabs.find(tab => tab.Type === "pages");
    const pageSources: string[] = (pagesTab?.Source ?? []) as string[];
    if (!Array.isArray(pageSources) || pageSources.length === 0) {
      setSearchedSitePages([]);
      return;
    }
    const paths = pageSources.map(source => `"${source}/SitePages"`).join(' OR ');
    const pageSearchQuery = `Path:(${paths}) AND owstaxId${termColumnName}:GP0|#${term}`;
    const pageSearchResults = await sp.search({
      Querytext: pageSearchQuery,
      TrimDuplicates: false,
      SelectProperties: [
        "Title",
        "Path",
        "ListId",
        "ListItemId",
        "UniqueId",
        "FileType",
        "FileExtension",
        "FileName",
        "LastModifiedTime",
        "Editor",
        "Author"
      ],
      RowLimit: 50
    });

    const pagePrimary = pageSearchResults.PrimarySearchResults || [];
    console.log("Page Search Query: ", pageSearchQuery);
    console.log("Pages Search Results: ", pagePrimary);
    const mappedPages = pagePrimary.map((r: any) => ({
      Typ: r.FileExtension || 'aspx',
      Namn: r.Title || r.FileName || r.Path,
      FileExtension: r.FileExtension || 'aspx',
      FileType: r.FileType,
      Path: r.Path,
      UniqueId: r.UniqueId,
      ListId: r.ListId,
      ListItemId: r.ListItemId,
    }));
    setSearchedSitePages(mappedPages);
  }, [tabsConfiguration, termColumnName]);

  const fetchDescription = useCallback(async (sp: SPFI, term: string, tab: any) => {
    const desclistId = tab.Source;
    const tabId = tab.uniqueId;

    if (!desclistId) {
      setBeskrivning(tabId, '');
      return;
    }

    try {
      // Fetch all fields in the list
      const fields = await sp.web.lists
        .getById(desclistId)
        .fields.select('InternalName', 'TypeAsString', 'Hidden', 'Title')();

      // Ensure single multiline field exists
      const multilineFields = fields.filter(f => f.TypeAsString === 'Note' && !f.Hidden);
      if (multilineFields.length !== 1) {
        throw new Error(`Expected exactly one visible multiline text field for description, found ${multilineFields.length}`);
      }

      const descFieldInternalName = multilineFields[0].InternalName;

      // Build search query
      const descQueryParts: string[] = [`ListId:${desclistId}`, `owstaxId${termColumnName}:L0|#${term}`];
      const descSearchQuery = descQueryParts.join(' AND ');

      console.log("descSearchQuery:", descSearchQuery);

      const descSearchResults = await sp.search({
        Querytext: descSearchQuery,
        SelectProperties: ['Title', 'ListId', 'ListItemId', 'Path', `owstaxId${termColumnName}`],
        RowLimit: 10,
        TrimDuplicates: false,
        SortList: [{ Property: 'LastModifiedTime', Direction: 1 }],
      });

      const primaryResults = descSearchResults?.PrimarySearchResults as ISearchResultWithListIds[] || [];
      const exact = primaryResults.find(item => {
        const taxo = (item as any)[`owstaxId${termColumnName}`];
        if (!taxo) return false;
        const taxoParts = taxo.split(';').map((s: string) => s.trim());
        return taxoParts.some((t: string) => t.startsWith(`L0|#0${term}|`));
      });

      // Fetch the item based on the exact description term
      const listItemIdStr = exact?.ListItemId;
      if (listItemIdStr) {
        const listItemId = parseInt(listItemIdStr, 10);
        if (!isNaN(listItemId)) {
          try {
            const item = await sp.web.lists.getById(desclistId).items.getById(listItemId)
              .select(descFieldInternalName)();
            const beskrivningRaw: string = item?.[descFieldInternalName] ?? '';
            setBeskrivning(tabId, beskrivningRaw);
          } catch (fieldReadErr) {
            console.warn('Failed to read description field, falling back to empty.', fieldReadErr);
            setBeskrivning(tabId, '');
          }
        }
      }
    } catch (error) {
      console.error('Error fetching description:', error);
      setBeskrivning(tabId, '');
    }
  }, [termColumnName]);

  const fetchLinks = useCallback(async (sp: SPFI, term: string) => {
    const tabs = tabsConfiguration ?? [];
    const linkTab = tabs.find(tab => tab.Type === 'links');
    const linkListId = linkTab?.Source;

    if (!linkListId) {
      setSearchedLinks([]);
      return;
    }

    const listFields = await fetchLibraryFields(sp, linkListId);
    const listUrlField = listFields.find((f: any) => f.fieldName.toLowerCase().includes('url') || f.fieldName.toLowerCase().includes('link'));
    const linkInterName = listUrlField?.fieldName;
    const listNoteField = listFields.find((f: any) => f.data.toLowerCase() === 'note');
    const descInterName = listNoteField?.fieldName;

    let termIds: string[] = [];
    let selectedTermObj: any = null;
    if (Array.isArray(selectedTerm) && selectedTerm.length > 0) {
      selectedTermObj = selectedTerm[0];
    }
    if (selectedTermObj) {
      termIds.push(selectedTermObj.key);
      if (Array.isArray(selectedTermObj.children) && selectedTermObj.children.length > 0) {
        termIds.push(...selectedTermObj.children.map((child: any) => child.key));
      }
    }
    termIds = Array.from(new Set(termIds));

    let linkSearchQuery: string;
    if (termIds.length > 1) {
      const orTerms = termIds.map(tid => `owstaxId${termColumnName}:GP0|#${tid}`).join(' OR ');
      linkSearchQuery = `ListId:${linkListId} AND (${orTerms})`;
    } else {
      linkSearchQuery = `ListId:${linkListId} AND owstaxId${termColumnName}:GP0|#${term}`;
    }

    const linkSearchResults = await sp.search({
      Querytext: linkSearchQuery,
      SelectProperties: ['Title', 'ListId', 'ListItemId', 'Path', 'UniqueId'],
      RowLimit: 200,
      TrimDuplicates: false
    });
    console.log('linkSearchQuery:', linkSearchQuery);

    const linkPrimary = linkSearchResults.PrimarySearchResults || [];
    console.log('Link search results:', linkPrimary);
    const mappedLinks = await Promise.all(
      linkPrimary.map(async (r: any) => {
        const listItemIdNum = parseInt(r.ListItemId, 10);
        let beskrivning: string = '';
        let link: string = '';
        if (!isNaN(listItemIdNum)) {
          try {
            const item = await sp.web.lists.getById(linkListId).items.getById(listItemIdNum).select(`${descInterName}`, `${linkInterName}`)();
            if (descInterName) beskrivning = item?.[descInterName] ?? '';
            if (linkInterName) link = item?.[linkInterName] ?? '';
          } catch (e) {
            console.warn('Failed to read Beskrivning for link item', r?.ListItemId, e);
          }
        }
        return {
          Namn: r.Title || r.Path,
          ListId: r.ListId,
          ListItemId: r.ListItemId,
          UniqueId: r.UniqueId,
          Path: link,
          Beskrivning: beskrivning
        };
      })
    );

    setSearchedLinks(mappedLinks);
  }, [tabsConfiguration, selectedTerm, termColumnName]);  

  const getDataForTabs = useCallback(async (term: string, types: string[]) => {
    try {
      const sp = getSPForWeb(absoluteURL);
      const uniqueTypes = Array.from(new Set(types)).filter(Boolean);

      for (const type of uniqueTypes) {
        const state = tabLoadStateRef.current[type];
        if (state?.loaded || state?.loading) continue;
        markTabLoading(type);
        try {
          if (type === 'documents') {
            await fetchDocuments(term);
          } else if (type === 'list') {
            await fetchList(term);
          } else if (type === 'pages') {
            await fetchPages(sp, term);
          } else if (type === 'links') {
            await fetchLinks(sp, term);
          } else if (type === 'description') {
            const descTabs = (tabsConfiguration ?? []).filter((tab: any) => tab.Type === 'description');
            for (const tab of descTabs) {
              await fetchDescription(sp, term, tab);
            }
          }
          markTabLoaded(type);
        } catch (e) {
          console.error(`Failed to load tab ${type}`, e);
          markTabFailed(type);
        }
      }
    } catch (e) {
      console.error('Failed to initialize data fetch', e);
    }
  // do NOT depend on tabLoadState; we use tabLoadStateRef
  }, [absoluteURL, fetchDocuments, fetchPages, fetchLinks, fetchDescription]);

  const onPivotLinkClick = (item?: PivotItem, ev?: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
    if (!item) return;
    const itemKey = item.props?.itemKey as string | undefined;
    if (!itemKey) return;
    setActiveTab(itemKey);

    // Lazy-load this tab if not already loaded
    const tab = (tabsConfiguration ?? []).find((t: any) => t.uniqueId === itemKey);
    const type = tab?.Type;
    if (!type) return;
    const termId = resolveSelectedTermId();
    if (!termId) return;
    const st = tabLoadState[type];
    if (!st || (!st.loaded && !st.loading)) {
      getDataForTabs(termId, [type]); // do not block other background work
    }
  };

  // --- Filter helpers (no persistence) ---
  const loadPersistedFilters = (): any[] => { return []; }

  // Two-phase initial load: active tab then background
  useEffect(() => {
    const run = async () => {
      if (!isConnected && selectedArea) {
        setChosenArea(null);
      }

      const termId = resolveSelectedTermId();
      const tabs = Array.isArray(tabsConfiguration) ? [...tabsConfiguration] : [];
      if (!termId || tabs.length === 0) {
        setIsLoaded(true);
        setSearchedDocuments([]);
        setSearchedListItems([]);
        setSearchedSitePages([]);
        setSearchedLinks([]);
        setViewFields([]);
        setBeskrivningar({});
        setNoResultsWithSelectedTerm(false);
        return;
      }

      // Determine first (active) tab
      const firstType = reportFirstTabType() || tabs[0]?.Type;
      if (firstType) setActiveTab(firstType);

      // Reset page loaded
      setIsLoaded(false);

      // 1) Load active tab
      await getDataForTabs(termId, [firstType!]);
      setIsLoaded(true);

      // 2) Background load for the remaining tabs
      const rest = tabs.map((t: any) => t.Type).filter((t: string) => t && t !== firstType);
      if (rest.length > 0) {
        // Fire and forget
        getDataForTabs(termId, rest);
      }

      // Restore filters post-load
      const persistedFilters = loadPersistedFilters();
      if (persistedFilters.length > 0) {
        setUseFilters(persistedFilters);
      }
    };
    run();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Reload when selectedArea, selectedTerm, isConnected or tabsConfiguration change
  useEffect(() => {
    let cancelled = false;
    const run = async () => {
      const termId = resolveSelectedTermId();
      const tabs = Array.isArray(tabsConfiguration) ? [...tabsConfiguration] : [];
      if (!termId || tabs.length === 0) {
        if (!cancelled) {
          setIsLoaded(true);
          setSearchedDocuments([]);
          setSearchedListItems([]);
          setSearchedSitePages([]);
          setSearchedLinks([]);
          setViewFields([]);
          setBeskrivningar({});
          setNoResultsWithSelectedTerm(false);
        }
        return;
      }

      // Reset per-tab state (apply to both state and ref so loader guard is consistent)
      const nextState: Record<string, {loading: boolean; loaded: boolean}> = {};
      for (const t of tabs) nextState[t.Type] = {loading: false, loaded: false};
      setTabLoadState(nextState);
      tabLoadStateRef.current = nextState;

      const firstType = reportFirstTabType() || tabs[0]?.Type;
      if (firstType) setActiveTab(firstType);

      setIsLoaded(false);
      await getDataForTabs(termId, [firstType!]);
      if (!cancelled) setIsLoaded(true);

      const rest = tabs.map((t: any) => t.Type).filter((t: string) => t && t !== firstType);
      if (rest.length > 0) getDataForTabs(termId, rest);
    };
    run();
    return () => { cancelled = true; }
    // IMPORTANT: do NOT include getDataForTabs here to avoid effect thrashing from function identity changes
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedArea, selectedTerm, isConnected, tabsConfiguration]);

  const getColumnType =(fieldName?: string): string | undefined => {
    if (!fieldName) return undefined;
    const col = ViewFields?.find(c => c.fieldName === fieldName);
    return (col as any)?.data;
  }

  const openFilterPanel = (column?: IColumn | null, fieldInfo?: IFieldInfo | null) => {
    _filterColumn.current = column ?? null;
    _filterFieldInfo.current = fieldInfo ?? null;
    setIsFilterPanelOpen(true);
  };

  const buildMenuItems =() => {
    const fieldName = contextMenu.column?.fieldName as string;
    const kind = resolveKind(contextMenu.fieldInfo || undefined);

    let config: { key: string; label: string }[] = [];
    if (activeTab === 'documents') {
      config = kind === 'text'
        ? [
            { key: 'sortasc', label: strings.sortasc },
            { key: 'sortdesc', label: strings.sortdesc }
          ]
        : [
            { key: 'oldToNew', label: strings.oldToNew },
            { key: 'newToOld', label: strings.newToOld }
          ];
    }
    config.push({ key: 'filter', label: strings.filter });

    const menuItems: IContextualMenuItem[] = config.map(c => {
      const combinedKey = `${fieldName}|${c.key}`;
      const clickHandler = c.key === 'filter'
        ? () => {
            openFilterPanel(contextMenu.column, contextMenu.fieldInfo);
            closeMenu();
          }
        : () => sortDocs(c.key, fieldName);
      return {
        key: c.key,
        text: c.label,
        onClick: clickHandler,
        onRender: (item) => (
          <div
            className={styles.contextualMenu}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 4,
              cursor: 'pointer',
              backgroundColor: _hoveredMenuKey.current === combinedKey ? '#f3f2f1' : 'transparent',
              padding: '4px 8px',
              borderRadius: 2
            }}
            onMouseEnter={() => { _hoveredMenuKey.current = combinedKey; forceUpdate(); }}
            onMouseLeave={() => { _hoveredMenuKey.current = null; forceUpdate(); }}
            onClick={item.onClick}
          >
            <div className={styles.contextualMenuIcon}>
              {pendingOp === combinedKey ? <Icon iconName="CheckMark" /> : <div></div>}
            </div>
            <span>{c.label}</span>
          </div>
        )
      };
    });

    return (
      <ContextualMenu
        target={contextMenu.target}
        onDismiss={closeMenu}
        items={menuItems}
        shouldFocusOnMount
        calloutProps={{ gapSpace: 2 }}
      />
    );
  }

  const closeFilterPanel = () => {
    setIsFilterPanelOpen(false);
    _filterColumn.current = null;
    _filterFieldInfo.current = null;
  };

  const onClearFilter = () => {
    setUseFilters([]);
    setSearchedDocuments([..._allDocuments.current]);
  }
  
  const handlePanelClear = (fieldNameToClear?: string) => {
    if (!fieldNameToClear) return;
    setUseFilters(prevFilters => prevFilters.filter((f: any) => f.fieldName !== fieldNameToClear));
  };

  const filterDocs = () => {
    if (!Array.isArray(useFilters) || useFilters.length === 0) {
      setSearchedDocuments([..._allDocuments.current]);
      return;
    }

    const normalizeToStrings = (value: any, columnType?: string): string[] => {
      if (value == null) return [];
      const bucket: string[] = [];

      const pushVal = (v: any) => {
        if (v == null) return;
        if (columnType === 'TaxonomyFieldType' || columnType === 'TaxonomyFieldTypeMulti') {
          if (typeof v === 'object' && v.TermGuid) {
            bucket.push(String(v.TermGuid));
          } else if (typeof v === 'string') {
            const parts = v.split('|');
            const guid = parts.length > 1 ? parts[1] : undefined;
            if (guid) bucket.push(guid);
          }
        } else if (
          columnType === 'User' ||
          columnType === 'SP.Data.UserInfoItem' ||
          (typeof columnType === 'string' && columnType.toLowerCase().includes('user'))
        ) {
          if (typeof v === 'object' && v.Id) {
            bucket.push(String(v.Id));
          } else if (typeof v === 'string' && !isNaN(Number(v))) {
            bucket.push(v);
          }
        } else {
          bucket.push(String(v));
        }
      };

      if (Array.isArray(value)) value.forEach(pushVal); else pushVal(value);
      return bucket;
    };

    const filteredDocs: any[] = [];
    for (const doc of _allDocuments.current) {
      let allMatch = true;
      for (const f of useFilters) {
        if (!f || !f.fieldName || !Array.isArray(f.values) || f.values.length === 0) continue;
        let columnType: string | undefined;
        if (ViewFields && ViewFields.length > 0) {
          const col = ViewFields.find(c => c.fieldName === f.fieldName);
          columnType = (col as any)?.data;
        }
        const docFieldRaw = (doc as any)[f.fieldName];
        if (docFieldRaw == null) { allMatch = false; break; }
        const docTokens = normalizeToStrings(docFieldRaw, columnType).map(s => s.toLowerCase());
        if (docTokens.length === 0) { allMatch = false; break; }
        const desired = f.values.map((v: any) => v.toLowerCase());
        if (!desired.some((d: any) => docTokens.includes(d))) { allMatch = false; break; }
      }
      if (allMatch) filteredDocs.push(doc);
    }
    setSearchedDocuments(filteredDocs);
  };

  // Keep only one effect for filterDocs
  useEffect(() => { filterDocs(); }, [useFilters]);

  const removeFilter = (fieldName: string, value: string) => {
    setUseFilters(prevFilters => {
      const existing = Array.isArray(prevFilters) ? prevFilters : [];
      const updated = existing
        .map((f: any) => {
          if (f.fieldName !== fieldName) return f;
          const newValues = (f.values || []).filter((v: string) => v !== value);
          return { ...f, values: newValues };
        })
        .filter((f: any) => Array.isArray(f.values) && f.values.length > 0);
        return updated;
    })
  }

  const getDistinctColumnValues = (fieldName?: string): (string | TaxoOption)[] => {
    if (!fieldName) return [];
    const items = [...(searchedDocuments || []), ...(searchedSitePages || [])];
    let columnType: string | undefined;
    if (ViewFields && ViewFields.length > 0) {
      const col = ViewFields.find(c => c.fieldName === fieldName);
      columnType = (col as any)?.data;
    }
    if (columnType === "TaxonomyFieldType" || columnType === "TaxonomyFieldTypeMulti") {
      return getDistinctTaxonomyValues(items, fieldName);
    } else if (
      columnType === "User" ||
      columnType?.toLowerCase().includes("user") ||
      columnType === "SP.Data.UserInfoItem"
    ) {
      const seen = new Set<string>();
      const results: TaxoOption[] = [];
      for (const item of items) {
        let v = (item as any)[fieldName];
        if (Array.isArray(v)) {
          v.forEach((userEntry: any) => {
            const label = userEntry?.Title || userEntry?.Email || userEntry?.Id;
            const value = userEntry?.Id ? String(userEntry.Id) : userEntry?.Email;
            if (label && value && !seen.has(value)) {
              results.push({ label, value });
              seen.add(value);
            }
          });
        } else if (typeof v === "object" && v !== null) {
          const label = v.Title || v.Email || v.Id;
          const value = v.Id ? String(v.Id) : v.Email;
          if (label && value && !seen.has(value)) {
            results.push({ label, value });
            seen.add(value);
          }
        } else if (typeof v === "string" && v) {
          if (!seen.has(v)) {
            results.push({ label: v, value: v });
            seen.add(v);
          }
        }
      }
      return results.sort((a, b) => a.label.localeCompare(b.label));
    } else {
      const values = new Set<string>();
      for (const item of items) {
        let v = (item as any)[fieldName];
        if (Array.isArray(v)) v.forEach(val => values.add(String(val)));
        else if (v != null) values.add(String(v));
      }
      return Array.from(values).sort((a, b) => a.localeCompare(b));
    }
  }

  const handleFilterSelectedValuesChange = useCallback((values: string[]) => {
    const field = _filterColumn.current?.fieldName;
    if (!field) return;
    const columnType = getColumnType(field);
    setUseFilters(prevFilters => {
      const existing = Array.isArray(prevFilters) ? prevFilters : [];
      const remaining = existing.filter((f: any) => f?.fieldName !== field);
      if (!values || values.length === 0) {
        return remaining;
      }
      let filterValues = values;
      if (
        columnType === "User" ||
        columnType?.toLowerCase().includes("user") ||
        columnType === "SP.Data.UserInfoItem"
      ) {
        filterValues = values.map(id => String(id));
      }
      const updated = [...remaining, { fieldName: field, values: filterValues }];
      return updated;
    });
  }, []);

  const renderEmptyWebpart = (message: string) => {
    return (
      <MessageBar
        messageBarType={MessageBarType.info}
        isMultiline={false}
        onDismiss={resetMessage}
        dismissButtonAriaLabel="Close"
      >
        {message}
      </MessageBar>
    );
  }

  const getDefaultColumns = (): IColumn[] => {
    const cols: IColumn[] = [
      {
        key: 'type',
        name: 'Typ',
        fieldName: 'Typ',
        minWidth: 20,
        maxWidth: 20,
        isMultiline: false,
        isIconOnly: true,
        iconName: 'Page',
        onRender: (item: any) => (
          <Icon {...getFileTypeIconProps({ extension: (item?.FileExtension || item?.Typ), size: 16 })} />
        )
      },
      {
        key: 'title',
        name: 'Namn',
        fieldName: 'Namn',
        minWidth: 100,
        maxWidth: 200,
        isMultiline: false,
        onRender: (item: any) => {
          const href = item?.Path;
          const text = item?.Namn ?? '';
          return href
            ? <Link href={href} target="_blank" rel="noopener noreferrer" data-interception="off" className={`${styles.docLinks} ${styles.links}`}>{text}</Link>
            : <span>{text}</span>;
        }
      },
    ];

    if (pendingOp) {
      const [sortedField, op] = pendingOp.split('|');
      const isDesc = op === 'newToOld' || op === 'sortdesc';
      cols.forEach(c => {
        if (c.fieldName === sortedField) {
          c.isSorted = true;
          c.isSortedDescending = isDesc;
        } else {
          c.isSorted = false;
          c.isSortedDescending = undefined;
        }
      });
    }
    return cols;
  }

  const renderPivotContent = (tab: { Type: string; SourceField?: any; uniqueId: string }) => {
    const items = getItems(tab.Type);

    // If this tab hasn't loaded yet and is currently being viewed, show a spinner
    const tabState = tabLoadState[tab.Type];
    const isTabLoading = tabState?.loading && !tabState?.loaded;
    if (isTabLoading) {
      return <Spinner label={strings.loading} ariaLive="assertive" labelPosition="bottom" />;
    }

    const isEmptyTab =
      (tab.Type === "description" && (!beskrivningar || beskrivningar === null)) ||
      (tab.Type === "documents" && (!items || items.length === 0)) ||
      (tab.Type === "pages" && (!items || items.length === 0)) ||
      (tab.Type === "links" && (!items || items.length === 0)) ||
      (tab.Type === "list" && (!items || items.length === 0));

    if (isEmptyTab && tabState?.loaded) {
      let emptyMsg = strings.EmptyData;
      if (tab.Type === "description") emptyMsg = strings.EmptyDescription;
      return renderEmptyWebpart(emptyMsg);
    }

    if (tab.Type === "description") {
      return (
        <div className={styles.description} dangerouslySetInnerHTML={{ __html: beskrivningar[tab.uniqueId] ?? '' }} />
      );
    } else if (items && items.length > 0) {
      if (tab.Type === 'links') {
        const linkColumns: IColumn[] = [
          {
            key: 'title',
            name: strings.ColumnNameTitle,
            fieldName: strings.ColumnNameTitle,
            minWidth: 100,
            maxWidth: 300,
            isMultiline: false,
            onRender: (item: any) => {
              let href = item.Path;
              if (href && !/^https?:\/\//i.test(href)) {
                href = `https://${href}`;
              }
              return (
                <>
                  <TooltipHost content={item.Beskrivning}>
                    <Link href={href} target="_blank" rel="noopener noreferrer" data-interception="off" className={styles.links}>
                      {item.Namn}
                    </Link>
                  </TooltipHost>
                </>
              );
            }
          },
          {
            key: 'description',
            name: strings.ColumnNameDescription,
            fieldName: strings.ColumnNameDescription,
            minWidth: 100,
            maxWidth: 400,
            isMultiline: true,
            onRender: (item: any) => (
              <TooltipHost content={item.Beskrivning}>
                <span>{trimText(item.Beskrivning, 60)}</span>
              </TooltipHost>
            )
          }
        ];
        return (
          <DetailsList
            items={items}
            columns={linkColumns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
          />
        );
      }

      let columnsToShow: IColumn[] = [];
      switch (tab.Type) {
        case 'documents':
          columnsToShow = ViewFields;
          break;
        case 'pages':
          columnsToShow = getDefaultColumns();
          break;
      }

      if (columnsToShow && columnsToShow.length > 0 && useFilters?.length > 0) {
        const filteredFields = new Set(
          useFilters
            .filter(f => f?.fieldName)
            .map(f => f.fieldName)
        );
        columnsToShow = columnsToShow.map(col => {
          if (filteredFields.has(col.fieldName)) {
            return {
              ...col,
              onRenderColumnHeaderTooltip: undefined,
              onRenderHeader: (props, defaultRender) => (
                <span style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                  <span>{col.name}</span>
                  <Icon iconName="Filter" style={{ color: '#0078d4' }} />
                </span>
              ),
              iconName: undefined,
              onRenderIcon: undefined,
            };
          } else {
            return {
              ...col,
              onRenderColumnHeaderTooltip: undefined,
              onRenderHeader: (props, defaultRender) => (
                <span>{col.name}</span>
              ),
              iconName: undefined,
              onRenderIcon: undefined,
            };
          }
        });
      }

      return (
        <>
          {useFilters?.length > 0 && (
            <div className={styles.filterDiv}>
              <div>
                {Object.entries(
                  useFilters.reduce((acc: Record<string, string[]>, f: any) => {
                    if (!f.fieldName) return acc;
                    if (!acc[f.fieldName]) acc[f.fieldName] = [];
                    if (Array.isArray(f.values)) acc[f.fieldName].push(...f.values);
                    return acc;
                  }, {})
                ).map(([fieldName, values]) => {
                  let columnType: string | undefined;
                  if (ViewFields && ViewFields.length > 0) {
                    const col = ViewFields.find(c => c.fieldName === fieldName);
                    columnType = (col as any)?.data;
                  }
                  let formattedPairs: { original: string, display: string }[] = [];
                  const seen = new Set<string>();
                  formattedPairs = [];
                  for (const v of values as string[]) {
                    if (!seen.has(v)) {
                      seen.add(v);
                      let display = v;
                      if (
                        columnType === "TaxonomyFieldType" ||
                        columnType === "TaxonomyFieldTypeMulti"
                      ) {
                        display = getTaxonomyLabel(fieldName, v, searchedDocuments ?? [] );
                      } else if (
                        columnType === "User" ||
                        columnType?.toLowerCase().includes("user") ||
                        columnType === "SP.Data.UserInfoItem"
                      ) {
                        let userResult = getUserIdAndName(fieldName, v, searchedDocuments ?? []);
                        display = typeof userResult === 'string' ? userResult : userResult?.name ?? '';
                      }
                      formattedPairs.push({ original: v, display });
                    }
                  }
                  return (
                    <span key={fieldName} style={{ marginRight: 12 }}>
                      <span style={{ fontWeight: 600 }}>{fieldName}:</span>
                      {formattedPairs.map((pair, idx: number) => (
                        <span
                          key={`${fieldName}-${pair.original}-${idx}`}
                          onClick={() => removeFilter(fieldName, pair.original)}
                          className={styles.filterChip}
                        >
                          {pair.display}
                          <Icon iconName="Cancel" className={styles.filterChipIcon} />
                        </span>
                      ))}
                    </span>
                  );
                })}
              </div>
              <div className={styles.cleardiv}>
                <span onClick={onClearFilter}>clear all filters</span>
              </div>
            </div>
          )}
        
          <DetailsList
            items={items}
            columns={columnsToShow}
            setKey="set"
            layoutMode={DetailsListLayoutMode.fixedColumns}
            selectionMode={SelectionMode.none}
          />
        </>
      );
    }
    return null;
  };

  // #region Main render and empty state logic
  const tabConfiguration = (
    tabsConfiguration ? (
      [...tabsConfiguration].sort((a: any, b: any) => (a.OrderOfTab ?? 0) - (b.OrderOfTab ?? 0))
    ) : []
  );

  const noTabsConfigured = tabConfiguration.length === 0;

  const isSelectedAreaEmpty =
    selectedArea == null ||
    (typeof selectedArea === 'object' && Object.keys(selectedArea).length === 0);

  const isSelectedTermArray = Array.isArray(selectedTerm);
  const noChosenTerm = isSelectedTermArray && selectedTerm.length === 0;

  // FIX: make this correct; was previously always true when beskrivning was falsy
  const hasData: boolean =
    (searchedDocuments?.length ?? 0) > 0 ||
    (searchedSitePages?.length ?? 0) > 0 ||
    (searchedLinks?.length ?? 0) > 0 ||
    !!beskrivningar;

  const connectedEmpty = isConnected && isSelectedAreaEmpty && !hasData;
  const notConnected = (!isConnected || isConnected === undefined) && noChosenTerm;

  let renderText: string = '';
  if (connectedEmpty) renderText = strings.connectedEmpty; 
  else if (notConnected) renderText = strings.notConnectedEmpty;

  const shouldRenderEmptyState = connectedEmpty || notConnected;

  // Compute whether all configured tabs are loaded to control global empty-state behavior
  const allConfiguredTypes = tabConfiguration.map((t: any) => t.Type);
  const allTabsLoaded = allConfiguredTypes.every(t => tabLoadState[t]?.loaded);
  const shouldLoad = !allTabsLoaded; // while background tabs still loading

  const emptyDocsSearched: boolean = !shouldLoad && tabsConfiguration?.find((tab: any) => tab.Type === 'documents') && searchedDocuments?.length === 0;
  const emptyPagesSearched: boolean = !shouldLoad && tabsConfiguration?.find((tab: any) => tab.Type === 'pages') && searchedSitePages?.length === 0;
  const emptyLinksSearched: boolean = !shouldLoad && tabsConfiguration?.find((tab: any) => tab.Type === 'links') && searchedLinks?.length === 0;
  const emptyDesc: boolean = !shouldLoad && tabsConfiguration?.find((tab: any) => tab.Type === 'description') && beskrivningar === null;
  const renderEmptyData = searchedDocuments?.length === 0 && searchedSitePages?.length === 0 && searchedLinks?.length === 0 && !beskrivningar && (emptyDocsSearched || emptyPagesSearched || emptyLinksSearched || emptyDesc || noResultsWithSelectedTerm);

  return (
    <section className={styles.unisightProcessDocuments}>
      {!isloaded && (
        <Spinner label={strings.loading} ariaLive="assertive" labelPosition="bottom" />
      )}

      <div className={styles.row}>
        {
          isloaded && (
            noTabsConfigured && shouldRenderEmptyState && renderEmptyWebpart(renderText),
            renderEmptyData && renderEmptyWebpart(strings.EmptyData)
          )
        }
        
        {
         !isConfigured && renderEmptyWebpart(strings.EmptyConfigure)
        }
        
        {!noTabsConfigured && !shouldRenderEmptyState && isloaded && (
          <>
            <Pivot 
              selectedKey={activeTab}
              onLinkClick={onPivotLinkClick} 
              className={styles.tabList}>
              {tabConfiguration.map((tab: any) => {
                const type = tab.Type as string;

                const tabState = tabLoadState[type] ?? {loading: false, loaded: false};

                const isCountedTab = type !== 'description';
                const showEllipsis = !tabState.loaded;

                const items = getItems(type);
                const count = Array.isArray(items) ? items.length : 0;

                const suffix = isCountedTab ? (showEllipsis ? '...' : String(count)) : '';
                const headerText = isCountedTab ? `${tab.TabName} (${suffix})` : tab.TabName;

                const reactKey = `${tab.uniqueId}-${suffix}`;

                return (
                  <PivotItem
                    itemKey={tab.uniqueId}
                    key={reactKey}
                    headerText={headerText}
                    itemIcon={tab.ShowIcon ? tab.Icon : undefined}
                  >
                      {renderPivotContent(tab)}
                  </PivotItem>
                );
              })}
            </Pivot>
          </>
        )}
      </div>
      {contextMenu.target && contextMenu.column && buildMenuItems()}
      {isFilterPanelOpen && (
        <FilterPanel
          isOpen={true}
          onDismiss={closeFilterPanel}
          onApply={closeFilterPanel}
          onClear={handlePanelClear}
          fieldName={_filterColumn.current?.fieldName}
          fieldInfo={_filterFieldInfo.current}
          values={
            getColumnType(_filterColumn.current?.fieldName)?.startsWith("TaxonomyFieldType")
              ? getDistinctColumnValues(_filterColumn.current?.fieldName) as { label: string; value: string }[]
              : getDistinctColumnValues(_filterColumn.current?.fieldName) as string[]
          }
          onSelectedValuesChange={handleFilterSelectedValuesChange}
          searchDocs={searchedDocuments}
          useFilters={useFilters}
        />
      )}
    </section>
  );
  // #endregion
}
export default UnisightProcessDocuments;