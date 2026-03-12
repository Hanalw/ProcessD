import { SPHttpClientResponse } from '@microsoft/sp-http';

export const toArray = (v: string | string[] | undefined): string[] => (Array.isArray(v) ? v : (v ? [v] : []));
export const toSingle = (v: string | string[] | undefined): string | undefined => (Array.isArray(v) ? v[0] : v);

export const normalizeSiteUrl = (siteUrl: string) => (siteUrl || '').replace(/\/$/, '');

export async function parseJsonOrThrow(res: SPHttpClientResponse, url: string): Promise<any> {
  const ct = res.headers.get('content-type')?.toLowerCase() || '';
  if (ct.includes('application/json')) {
    return res.json();
  }
  const text = await res.text();
  const snippet = text.slice(0, 200).replace(/\s+/g, ' ');
  throw new Error(`Non-JSON response from ${url}: ${snippet}`);
}

export const siteTextFromUrl = (url: string) => {
  const u = (url || '').trim();
  if (!u) return '';
  return u;
}