import * as React from 'react';
import { useEffect, useMemo, useRef, useState } from 'react';

import {
  Dropdown,
  IDropdownOption,
  TextField
} from '@fluentui/react';

import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'UnisightProcessDocumentsWebPartStrings';
import { ISourceFieldProps } from '../../models';
import { getListsOptionsForSite } from '../../services/fetchOptions';
import { normalizeSiteUrl, toSingle } from '../../helpers';

const isListsType = (type?: string) => type === 'description' || type === 'links';
const isDocumentsType = (type?: string) => type === 'documents';
const isSitesType = (type?: string) => type === 'pages';
const isListNewType = (type?: string) => type === 'list';

const needsSitePicker = (type?: string) => isSitesType(type) || isDocumentsType(type) || isListNewType(type);
const needsListPicker = (type?: string) => isListsType(type) || isDocumentsType(type) || isListNewType(type);
const needsViewPicker = (type?: string) => type === 'documents' || type === 'list';

// cache types (unchanged)
export interface CacheState {
  cached: IDropdownOption[] | null;
  inFlight: Promise<IDropdownOption[]> | null;
}
const listsCacheByKey = new Map<string, CacheState>();
const viewsCacheByListId = new Map<string, CacheState>();

const getBaseUrl = (context: WebPartContext) =>
  context.pageContext.web.absoluteUrl.replace(/\/$/, '');

// -------------------- Site model --------------------
type SiteValue = { siteName: string; Url: string };
const isSiteValue = (v: any): v is SiteValue =>
  !!v && typeof v === 'object' && typeof v.Url === 'string';

function siteValueFromLegacyString(url: string): SiteValue {
  const u = normalizeSiteUrl(url);
  return { siteName: u, Url: u };
}

function normalizeSiteValue(s: SiteValue): SiteValue {
  return { siteName: s.siteName, Url: normalizeSiteUrl(s.Url) };
}

function siteKey(s: SiteValue): string {
  return normalizeSiteUrl(s.Url);
}

function mergeUniqueSites(a: SiteValue[], b: SiteValue[]): SiteValue[] {
  const m = new Map<string, SiteValue>();
  for (const s of a) m.set(siteKey(s), normalizeSiteValue(s));
  for (const s of b) m.set(siteKey(s), normalizeSiteValue(s));
  return Array.from(m.values());
}

// pages stores selection in `Source` (required), docs/list stores in `SelectedSite`
function readSelectedSites(isPagesType: boolean, props: ISourceFieldProps): SiteValue[] {
  const mapItem = (item: any): SiteValue | null => {
    if (isSiteValue(item)) return normalizeSiteValue(item);
    if (typeof item === 'string' && item) return siteValueFromLegacyString(item);
    return null;
  };

  if (isPagesType) {
    const v = props.value as any;
    if (Array.isArray(v)) return v.map(mapItem).filter(Boolean) as SiteValue[];
    if (isSiteValue(v)) return [normalizeSiteValue(v)];
    if (typeof v === 'string' && v) return [siteValueFromLegacyString(v)];
    return [];
  }

  const s = (props.selectedSite as any);
  if (Array.isArray(s)) return s.map(mapItem).filter(Boolean) as SiteValue[];
  if (isSiteValue(s)) return [normalizeSiteValue(s)];
  if (typeof s === 'string' && s) return [siteValueFromLegacyString(s)];
  return [];
}

// -------------------- SharePoint Search: baseline sites --------------------
type SearchRow = { Cells: Array<{ Key: string; Value: string }> };

function extractSearchRows(json: any): SearchRow[] {
  const primary = json?.PrimaryQueryResult || json?.d?.query?.PrimaryQueryResult;
  const rowsObj = primary?.RelevantResults?.Table?.Rows;
  if (!rowsObj) return [];

  const rows = Array.isArray(rowsObj?.results) ? rowsObj.results : rowsObj;
  return (rows || []).map((r: any) => {
    const cells = Array.isArray(r?.Cells?.results) ? r.Cells.results : r?.Cells;
    return { Cells: cells || [] };
  });
}

function cellVal(cells: Array<{ Key: string; Value: string }>, key: string): string | undefined {
  return cells.find((c) => c.Key === key)?.Value;
}

async function preloadSitesViaSPSearch(context: WebPartContext, rowLimit: number = 500): Promise<SiteValue[]> {
  const baseUrl = getBaseUrl(context);
  const url =
    `${baseUrl}/_api/search/query?` +
    `querytext='contentclass:STS_Site'&` +
    `rowlimit=${rowLimit}&trimduplicates=false&` +
    `selectproperties='Title,Path'`;

  const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
    headers: { 'odata-version': '3.0', 'accept': 'application/json;odata=nometadata' }
  });

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Failed to load sites. ${res.status} ${res.statusText}. ${text.slice(0, 200)}`);
  }

  const json = await res.json();
  const rows = extractSearchRows(json);

  const sites: SiteValue[] = rows
    .map(r => {
      const title = cellVal(r.Cells, 'Title') || '';
      const path = cellVal(r.Cells, 'Path') || '';
      if (!path) return null;
      return { siteName: title || path, Url: normalizeSiteUrl(path) };
    })
    .filter(Boolean) as SiteValue[];

  const unique = new Map<string, SiteValue>();
  for (const s of sites) unique.set(normalizeSiteUrl(s.Url), s);

  return Array.from(unique.values()).sort((a, b) => a.siteName.localeCompare(b.siteName, 'sv'));
}

// -------------------- Graph search assist --------------------
async function searchSitesViaGraph(context: WebPartContext, query: string): Promise<SiteValue[]> {
  const q = (query || '').trim();
  if (!q) return [];

  const client = await context.msGraphClientFactory.getClient('3');
  const res = await client.api('/sites').search(q).top(50).get();

  const value = (res?.value || []) as Array<{ webUrl?: string; displayName?: string; name?: string }>;
  const raw = value
    .filter(s => !!s.webUrl)
    .map(s => ({
      Url: normalizeSiteUrl(s.webUrl!),
      siteName: (s.displayName || s.name || s.webUrl!) as string
    }));

  const unique = new Map<string, SiteValue>();
  for (const s of raw) unique.set(normalizeSiteUrl(s.Url), s);

  return Array.from(unique.values()).sort((a, b) => a.siteName.localeCompare(b.siteName, 'sv'));
}

// -------------------- Views (unchanged) --------------------
async function getViewsOptions(context: WebPartContext, siteUrl: string, listId: string): Promise<IDropdownOption[]> {
  if (!listId) return [];

  const cacheKey = `${normalizeSiteUrl(siteUrl)}::${listId}`;
  let state = viewsCacheByListId.get(cacheKey);
  if (!state) {
    state = { cached: null, inFlight: null };
    viewsCacheByListId.set(cacheKey, state);
  }
  if (state.cached) return state.cached;
  if (state.inFlight) return state.inFlight;

  const fetcher = async (): Promise<IDropdownOption[]> => {
    const baseUrl = normalizeSiteUrl(siteUrl);
    const guid = listId.match(/^[0-9a-fA-F-]{36}$/) ? `guid'${listId}'` : `'${encodeURIComponent(listId)}'`;
    const url = `${baseUrl}/_api/web/lists(${guid})/Views?$select=Id,Title,Hidden,DefaultView`;

    const res = await context.spHttpClient.get(url, 0 as any, {
      headers: {
        'odata-version': '3.0',
        'ACCEPT': 'application/json;odata=nometadata',
        'X-ClientService-ClientTag': 'NonISV|Bravero|ProcessDocuments',
        'UserAgent': 'NonISV|Bravero|ProcessDocuments'
      }
    });

    if (!res.ok) {
      const text = await res.text().catch(() => '');
      throw new Error(`Failed to load views. ${res.status} ${res.statusText}. ${text.slice(0, 160)}`);
    }

    const json = await res.json();
    const items = (json?.value || []) as Array<{ Id: string; Title: string; Hidden: boolean; DefaultView: boolean }>;

    const filtered = items.filter(v => !v.Hidden);
    const opts: IDropdownOption[] = filtered
      .map(v => ({ key: v.Id, text: v.Title }))
      .sort((a, b) => a.text.localeCompare(b.text, 'sv'));

    const defaultView = filtered.find(v => v.DefaultView);
    if (defaultView) {
      const idx = opts.findIndex(o => o.key === defaultView.Id);
      if (idx > 0) {
        const [def] = opts.splice(idx, 1);
        opts.unshift(def);
      }
    }
    return opts;
  };

  state.inFlight = fetcher()
    .then((opts) => {
      state!.cached = opts;
      state!.inFlight = null;
      return opts;
    })
    .catch((err) => {
      state!.inFlight = null;
      throw err;
    });

  return state.inFlight;
}

// -------------------- Component --------------------
export const SourceField: React.FC<ISourceFieldProps> = (props) => {
  const { context, value, selectedType, onChanged, onViewChanged } = props;

  const isPagesType = isSitesType(selectedType);
  const isDocsType = isDocumentsType(selectedType);
  const isListType = isListNewType(selectedType);

  // Selected sites:
  // - pages => from Source (props.value)
  // - docs/list => from SelectedSite
  const selectedSites = useMemo(
    () => readSelectedSites(isPagesType, props),
    [isPagesType, props.value, props.selectedSite]
  );

  // docs/list need a single site
  const selectedSiteSingleUrl = !isPagesType ? normalizeSiteUrl(selectedSites[0]?.Url || '') : '';

  // site baseline/search state
  const [siteFilterText, setSiteFilterText] = useState('');
  const [baselineSites, setBaselineSites] = useState<SiteValue[]>([]);
  const [graphSites, setGraphSites] = useState<SiteValue[]>([]);
  const [sitesLoading, setSitesLoading] = useState(false);
  const [sitesError, setSitesError] = useState<string | undefined>(undefined);

  const MIN_SITE_CHARS = 3;

  // list/library state
  const [options, setOptions] = useState<IDropdownOption[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  // views state
  const [viewOptions, setViewOptions] = useState<IDropdownOption[]>([]);
  const [viewsLoading, setViewsLoading] = useState(false);
  const [viewsError, setViewsError] = useState<string | undefined>(undefined);

  const latestReqRef = useRef(0);
  const latestViewsReqRef = useRef(0);

  const placeholder = useMemo(() => {
    if (props.placeholder) return props.placeholder;
    if (!selectedType) return strings.TabSourceEmpty;
    if (needsListPicker(selectedType)) return 'Välj lista eller bibiliotek...';
    return 'Välj källa...';
  }, [props.placeholder, selectedType]);

  const sitePlaceholder = useMemo(() => {
    if (props.sitePlaceholder) return props.sitePlaceholder;
    if (!selectedType) return 'Välj webbplats...';
    if (isDocsType || isListType) return 'Välj webbplats (en)...';
    return 'Välj webbplats';
  }, [props.sitePlaceholder, selectedType, isDocsType, isListType]);

  // Baseline preload whenever site picker becomes relevant
  useEffect(() => {
    let cancelled = false;

    const run = async () => {
      setSitesError(undefined);
      setBaselineSites([]);
      setGraphSites([]);
      setSiteFilterText('');

      if (!needsSitePicker(selectedType)) return;

      try {
        setSitesLoading(true);
        const sites = await preloadSitesViaSPSearch(context, 500);
        if (!cancelled) setBaselineSites(sites);
      } catch (e: any) {
        if (!cancelled) {
          setSitesError(e?.message || 'Kunde inte ladda webbplatser.');
          setBaselineSites([]);
        }
      } finally {
        if (!cancelled) setSitesLoading(false);
      }
    };

    run();
    return () => { cancelled = true; };
  }, [context, selectedType]);

  // Graph search assist when typing >= 3 chars
  useEffect(() => {
    let cancelled = false;
    if (!needsSitePicker(selectedType)) return;

    const q = (siteFilterText || '').trim();
    if (q.length < MIN_SITE_CHARS) {
      setGraphSites([]);
      return;
    }

    const t = window.setTimeout(async () => {
      try {
        const sites = await searchSitesViaGraph(context, q);
        if (!cancelled) setGraphSites(sites);
      } catch {
        if (!cancelled) setGraphSites([]);
      }
    }, 250);

    return () => {
      cancelled = true;
      window.clearTimeout(t);
    };
  }, [context, selectedType, siteFilterText]);

  // Sites shown in dropdown:
  // - always include selected sites
  // - if typing >=3 => merge baseline + graph (so selection doesn't disappear)
  // - else => baseline
  const visibleSites: SiteValue[] = useMemo(() => {
    const q = (siteFilterText || '').trim();
    const base = q.length >= MIN_SITE_CHARS ? mergeUniqueSites(baselineSites, graphSites) : baselineSites;
    return mergeUniqueSites(selectedSites, base);
  }, [siteFilterText, baselineSites, graphSites, selectedSites]);

  const siteDropdownOptions: IDropdownOption[] = useMemo(() => {
    const q = (siteFilterText || '').trim().toLowerCase();
    const filtered = !q
      ? visibleSites
      : visibleSites.filter(s =>
          (s.siteName || '').toLowerCase().includes(q) ||
          (s.Url || '').toLowerCase().includes(q)
        );

    return filtered
      .slice()
      .sort((a, b) => (a.siteName || '').localeCompare(b.siteName || '', 'sv'))
      .map(s => ({
        key: siteKey(s),     // key = URL (persisted)
        text: s.siteName || s.Url, // display name
        data: s
      }));
  }, [visibleSites, siteFilterText]);

  const selectedSiteKeys = useMemo(
    () => selectedSites.map(s => siteKey(s)),
    [selectedSites]
  );

  // When a site is selected/unselected
  const onSiteChange = (_ev: any, option?: IDropdownOption) => {
    if (!option || option.disabled) return;

    const url = normalizeSiteUrl(String(option.key));
    const siteName = (option.text || url) as string;

    if (isPagesType) {
      // multi-select; persist into Source (required)
      const current = new Map<string, SiteValue>(
        selectedSites.map(s => [siteKey(s), normalizeSiteValue(s)])
      );

      if (current.has(url)) current.delete(url);
      else current.set(url, { siteName, Url: url });

      // IMPORTANT: persist as string[] (urls) (matches your current "working fine now" state)
      onChanged(Array.from(current.values()).map(s => s.Url));
      return;
    }

    // docs/list: single-select; persist into SelectedSite (and clear dependent selections)
    props.onSiteChanged?.(url);
    onChanged(undefined);
    onViewChanged?.(undefined);
  };

  // Load list/library options when needed (documents => libraries only; list => lists only)
  useEffect(() => {
    let cancelled = false;
    const myReqId = Date.now() + Math.random();
    latestReqRef.current = myReqId;

    const run = async () => {
      setError(undefined);
      setOptions([]);

      if (!needsListPicker(selectedType)) return;

      // docs/list require site selection
      if ((isDocsType || isListType) && !selectedSiteSingleUrl) return;

      const baseSiteUrl =
        (isDocsType || isListType)
          ? selectedSiteSingleUrl
          : getBaseUrl(context);

      try {
        setLoading(true);

        // documents => libraries only, list => lists only, description/links => lists only on current site
        const documentsOnly = isDocsType;
        const opts = await getListsOptionsForSite(context, baseSiteUrl, documentsOnly, listsCacheByKey);

        if (!cancelled && latestReqRef.current === myReqId) setOptions(opts);
      } catch {
        if (!cancelled && latestReqRef.current === myReqId) {
          setError('Kunde inte ladda källor.');
          setOptions([]);
        }
      } finally {
        if (!cancelled && latestReqRef.current === myReqId) setLoading(false);
      }
    };

    run().catch((e) => console.error('Error in SourceField list/library loading', e));
    return () => { cancelled = true; };
  }, [context, selectedType, selectedSiteSingleUrl, isDocsType, isListType]);

  const selectedKey = !isSitesType(selectedType) ? toSingle(props.value) : undefined;

  const handleSourceChange = (_: any, option?: IDropdownOption) => {
    if (!option) return;
    onChanged(option.key as string);
    if (needsViewPicker(selectedType)) onViewChanged?.(undefined);
  };

  const selectedListId = (isDocsType || isListType) ? toSingle(value) : undefined;

  // Load views for documents and lists only
  useEffect(() => {
    let cancelled = false;
    setViewsError(undefined);
    setViewOptions([]);

    if (!(isDocsType || isListType) || !selectedListId || !selectedSiteSingleUrl) {
      setViewsLoading(false);
      return;
    }

    const myReqId = Date.now() + Math.random();
    latestViewsReqRef.current = myReqId;

    const run = async () => {
      try {
        setViewsLoading(true);
        const opts = await getViewsOptions(context, selectedSiteSingleUrl, selectedListId);
        if (!cancelled && latestViewsReqRef.current === myReqId) setViewOptions(opts);
      } catch (e: any) {
        if (!cancelled && latestViewsReqRef.current === myReqId) {
          setViewsError(e?.message || 'Kunde inte ladda vyer.');
          setViewOptions([]);
        }
      } finally {
        if (!cancelled && latestViewsReqRef.current === myReqId) setViewsLoading(false);
      }
    };

    run();
    return () => { cancelled = true; };
  }, [context, isDocsType, selectedListId, selectedSiteSingleUrl]);

  const handleViewChange = (_: any, option?: IDropdownOption) => {
    if (!option) return;
    props.onViewChanged?.(option.key as string);
  };

  const showSitePicker = needsSitePicker(selectedType);
  const showListPicker = needsListPicker(selectedType) && !isSitesType(selectedType);
  const showViewPicker = needsViewPicker(selectedType);

  const siteDisabled = !!sitesError || !selectedType;

  const sourceDisabled =
    loading ||
    !!error ||
    !selectedType ||
    ((isDocsType || isListType) && !selectedSiteSingleUrl);

  return (
    <div style={{ minWidth: 260 }}>
      {showSitePicker && (
        <div style={{ marginBottom: 12 }}>
          {sitesLoading && (
            <div style={{ marginBottom: 8 }}>
              <Spinner size={SpinnerSize.small} label="Laddar webbplatser..." />
            </div>
          )}

          {sitesError && (
            <div style={{ marginBottom: 8, color: 'red' }}>
              {sitesError}
            </div>
          )}

          {/* Search helper (does not overwrite selection) */}
          <TextField
            placeholder={`Sök webbplats... (Graph efter ${MIN_SITE_CHARS} tecken)`}
            value={siteFilterText}
            onChange={(_, v) => setSiteFilterText(v || '')}
            disabled={siteDisabled}
            styles={{ root: { marginBottom: 8 } }}
          />

          <Dropdown
            placeholder={sitePlaceholder}
            options={siteDropdownOptions}
            disabled={siteDisabled}
            multiSelect={isPagesType}
            selectedKey={!isPagesType ? (selectedSiteSingleUrl || undefined) : undefined}
            selectedKeys={isPagesType ? selectedSiteKeys : undefined}
            onChange={onSiteChange}
          />

          {isPagesType && selectedSites.length > 0 && (
            <div style={{ marginTop: 6, fontSize: 12, color: '#605e5c' }}>
              Valda webbplatser: {selectedSites.length}
            </div>
          )}

          {(isDocsType || isListType) && !selectedSiteSingleUrl && (
            <div style={{ marginTop: 6, fontSize: 12, color: '#605e5c' }}>
              Välj en webbplats först för att ladda källor.
            </div>
          )}
        </div>
      )}

      {showListPicker && (
        <>
          {loading && (
            <div style={{ marginBottom: 8 }}>
              <Spinner size={SpinnerSize.small} label="Laddar källor..." />
            </div>
          )}
          {error && (
            <div style={{ marginBottom: 8, color: 'red' }}>
              {error}
            </div>
          )}

          <Dropdown
            placeholder={placeholder}
            options={options}
            disabled={sourceDisabled}
            multiSelect={false}
            selectedKey={selectedKey}
            onChange={handleSourceChange}
          />
        </>
      )}

      {showViewPicker && selectedSiteSingleUrl && selectedListId && (
        <div style={{ marginTop: 12 }}>
          {viewsLoading && (
            <div style={{ marginBottom: 8 }}>
              <Spinner size={SpinnerSize.small} label="Laddar vyer..." />
            </div>
          )}
          {viewsError && (
            <div style={{ marginBottom: 8, color: 'red' }}>
              {viewsError}
            </div>
          )}
          <Dropdown
            placeholder={props.viewPlaceholder || "Välj vy..."}
            options={viewOptions}
            disabled={viewsLoading || !!viewsError}
            selectedKey={props.selectedViewId}
            onChange={handleViewChange}
          />
        </div>
      )}
    </div>
  );
};

export default SourceField;