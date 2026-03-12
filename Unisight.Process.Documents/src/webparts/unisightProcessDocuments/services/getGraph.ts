import { WebPartContext } from "@microsoft/sp-webpart-base";

// import pnp and pnp logging system
import { graphfi, GraphFI, SPFx as graphSPFx } from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/graph"; // Import Graph methods
import "@pnp/graph/taxonomy";

let _graph: GraphFI | undefined = undefined;

export const getGraph = (context?: WebPartContext): GraphFI => {
  if (context !== undefined) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _graph = graphfi().using(graphSPFx(context)).using(PnPLogging(LogLevel.Warning));
  }
  return _graph!;
};