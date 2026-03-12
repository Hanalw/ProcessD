import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx as spSPFx  } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";

import "@pnp/sp/webs";
import "@pnp/sp/sites";
import "@pnp/sp/lists";
import "@pnp/sp/views";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/search";

let _sp: SPFI | undefined = undefined;
let _ctx: WebPartContext | undefined = undefined;

export const getSP = (context?: WebPartContext): SPFI => {
  if (context !== undefined) {
    _ctx = context;
    _sp = spfi().using(spSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  if (!_sp) {
    throw new Error('getSP was not initialized with a context');
  }
  return _sp!;
};

export const getSPForWeb = (webAbsoluteUrl: string): SPFI => {
  if (!_ctx) {
    throw new Error('getSPForWeb was not initialized with a context');
  }
  return spfi(webAbsoluteUrl).using(spSPFx(_ctx)).using(PnPLogging(LogLevel.Warning));
}