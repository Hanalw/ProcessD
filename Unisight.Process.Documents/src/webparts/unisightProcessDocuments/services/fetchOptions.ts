import { IDropdownOption } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { CacheState } from "../components/FormFields/SourceField";
import { isSystemListOrLibrary, normalizeSiteUrl, parseJsonOrThrow } from "../helpers";

export async function getListsOptionsForSite(
  context: WebPartContext,
  siteUrl: string,
  documentsOnly: boolean,
  listsCacheByKey: Map<string, CacheState>
): Promise<IDropdownOption[]> {
  const baseUrl = normalizeSiteUrl(siteUrl);
  const cacheKey = `${baseUrl}::${documentsOnly ? 'docs' : 'lists'}`;

  let state = listsCacheByKey.get(cacheKey);
  if (!state) {
    state = { cached: null, inFlight: null };
    listsCacheByKey.set(cacheKey, state);
  }

  if (state.cached) return state.cached;
  if (state.inFlight) return state.inFlight;

  const fetcher = async (): Promise<IDropdownOption[]> => {
    const select =
      '$select=Title,Id,BaseTemplate,BaseType,Hidden,RootFolder/ServerRelativeUrl&$expand=RootFolder';

    const filter = documentsOnly
      ? "$filter=(Hidden eq false) and (BaseType eq 1)" // Document libraries only
      : "$filter=(Hidden eq false) and (BaseType eq 0)"; // Lists only

    const url = `${baseUrl}/_api/web/lists?${select}&${filter}`;

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
      throw new Error(`Failed to load lists. ${res.status} ${res.statusText}. ${text.slice(0, 160)}`);
    }

    const json = await parseJsonOrThrow(res as any, url);
    const items = (json?.value || []) as Array<{
      Title: string;
      Id: string;
      BaseTemplate: number;
      BaseType: number;
      Hidden: boolean;
      RootFolder?: { ServerRelativeUrl: string };
    }>;

    const filteredItems = items.filter((l) => !isSystemListOrLibrary(l));

    const opts: IDropdownOption[] = filteredItems.map((l) => ({
      key: l.Id,
      text: l.Title || l.RootFolder?.ServerRelativeUrl || l.Id
    }));

    const unique = new Map<string, IDropdownOption>();
    for (const o of opts) unique.set(String(o.key), o);
    return Array.from(unique.values()).sort((a, b) => a.text.localeCompare(b.text, 'sv'));
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