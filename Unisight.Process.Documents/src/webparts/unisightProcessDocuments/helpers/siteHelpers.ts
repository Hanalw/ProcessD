import { systemTitles, systemUrlSegments } from "../components/constants";

export function isSystemListOrLibrary(l: {
  Title: string;
  RootFolder?: { ServerRelativeUrl: string };
  BaseTemplate: number;
}): boolean {
  const title = (l.Title || '').toLowerCase();
  const url = (l.RootFolder?.ServerRelativeUrl || '').toLowerCase();

  if (systemTitles.has(title)) return true;
  if (systemUrlSegments.some(seg => url.includes(seg))) return true;

  return false;
}