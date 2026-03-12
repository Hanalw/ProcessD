// import {  WebPartContext } from "@microsoft/sp-webpart-base";
// import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";
// import {  sp } from "@pnp/sp";
// import { ISPHttpClientOptions, SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
// import { taxonomy, ITermStore,} from '@pnp/sp-taxonomy'; //ITermData,
// import { ISharePointSearchResponse } from "../models/ISharePointSearchResponse";
// import { ISharePointSearchResult } from "../models/ISharePointSearchResults";
// import { IUnisightProcessDocumentsProps } from "../components/UnisightProcessDocuments/IUnisightProcessDocumentsProps";

// export class DataFetcher {

//   private spHttpClient: SPHttpClient;
//   private _spHttpOptions: any = {
//     getNoMetaData: <ISPHttpClientOptions>{
//         headers: {
//             'ACCEPT': 'application/json; odata.metadata=none'
//         }
//     },
//     updateNoMetadata: <ISPHttpClientOptions>{
//         headers: {
//             'ACCEPT': 'application/json; odata.metadata=none',
//             'CONTENT-TYPE': 'application/json',
//             'X-HTTP-Method': 'PATCH'
//         }
//     }
// };
//   constructor(private context: WebPartContext) {
//     sp.setup({
//         spfxContext: this.context as any
//     });
//     this.spHttpClient = this.context.spHttpClient;
//   }

//     public async SearchDocuments(query: any): Promise<any> {
//       try {
//         const response = await this.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`, SPHttpClient.configurations.v1, {
//             body: JSON.stringify(query),
//             headers: {
//               'odata-version': '3.0',
//               'accept': 'application/json;odata=nometadata',
//               'X-ClientService-ClientTag': 'NonISV|Bravero|ProcessDocuments',
//               'UserAgent': 'NonISV|Bravero|ProcessDocuments'
//             }
//         });
//         const data: ISharePointSearchResponse = await response.json();
//         if(data.PrimaryQueryResult){
//           const resultRows = data?.PrimaryQueryResult?.RelevantResults?.Table?.Rows;
//           let searchResults: ISharePointSearchResult[] = resultRows ? this.getSearchResults(resultRows) : [];
//           return searchResults;
//         }
//       }
//       catch(error) {
//         console.error(error);
//       }
//     }

//     public async getProcessLinks(processObjectList: string | string[], term: any): Promise<any[]> {
//       let processLinks: any[] = [];
//       try {
//         const LIST_API_ENDPOINT = `/_api/web/lists('${processObjectList}')`;
//         const FILTER = `$filter=TaxCatchAll/IdForTerm eq '${term}'`;
//         const SELECT_QUERY = `$select=Title,ServerRedirectedEmbedURL,ServerRedirectedURL,TaxCatchAll/IdForTerm`;
//         const EXPAND_QUERY = `$expand=TaxCatchAll`;

//         let query = `${this.context.pageContext.web.absoluteUrl}${LIST_API_ENDPOINT}/Items?${FILTER}&${SELECT_QUERY}&${EXPAND_QUERY}`;
//         const response = await this.spHttpClient.get(
//           query, 
//           SPHttpClient.configurations.v1,
//           this._spHttpOptions.getNoMetaData
//         );
//         const data = await response.json();
        
//         if(data.value.length > 0){
//           data.value.forEach((item: any) => {
//             processLinks.push({
//               title: item.Title,
//               url: item.ServerRedirectedURL,
//               embedUrl: item.ServerRedirectedEmbedURL,
//               termId: item.TaxCatchAll.IdForTerm
//             });
//           });
//         }
//       } catch (error) {
//         console.error(error);
//       }
//       return processLinks;
//     }

//     public async getAvailableManagedProperties(): Promise<any> {
//       let managedProperties: { name: any; }[] = [];
//       try{
//         let _query = {
//           request:{
//             '__metadata': {
//               'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//            },
//            'QueryTemplate': `*`,
//            'Refiners': 'ManagedProperties(filter=600/0/*)',
//            'RowLimit' : 1
//           }
//         };
//         const response = await this.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`, SPHttpClient.configurations.v1, {
//           body: JSON.stringify(_query),
//           headers: {
//             'odata-version': '3.0',
//             'accept': 'application/json;odata=nometadata',
//             'X-ClientService-ClientTag': 'NonISV|Bravero|ProcessDocuments',
//             'UserAgent': 'NonISV|Bravero|ProcessDocuments'
//           }
//       });

//         const data = await response.json();
//         let refinementResultRows = data?.PrimaryQueryResult?.RefinementResults?.Refiners;
//         refinementResultRows.map((refiner: { Entries: { RefinementName: any; }[]; }) => {
//           refiner.Entries.map((item: { RefinementName: any; }) => {
//             managedProperties.push({
//               name: item.RefinementName
//             });
//           });
//         });
//       }
//       catch(error){
//         console.error(error);
//       }

//       return managedProperties;
//     }

//     public generateQuery = async (_props: IUnisightProcessDocumentsProps, term: any, isSitePages?:boolean) => {
//         let choice = '';
//         let _query: any = {};
//         let documentQuery = '';
//         if(isSitePages){
//           documentQuery = "IsDocument:True AND FileExtension:aspx";
//         } else {
//           documentQuery = "IsDocument:1 AND NOT FileExtension:aspx";
//         }

//         //const documentQuery = _props.showSitePagesAlso ? "SPContentType:'Site Page' AND IsDocument:True AND FileExtension:aspx" : "IsDocument:1";
//         console.log(documentQuery);

//         if (_props.HubSiteId !== undefined || null) {
//           choice = 'hubsite';
//         }
//         if (_props.DocumentLibraryId !== undefined || null) {
//           choice = 'documentlibrary';
//         }
//         if (_props.CurrentSiteUrl !== undefined || null) {
//           choice = 'currentSite';
//         }
//         if (_props.Sites?.length !== 0) {
//           choice = 'sites';
//         }
    
//         let selectProperties = ["owstaxIdepProcess","Filename", "ModifiedBy", "CreatedBy", "FileType", "Title", "ServerRedirectedURL", "ServerRedirectedEmbedURL", "UniqueID", "Path", "ContentType", "ContentTypeId",];


//         let termSearchString = "";
//         const termIds = await this.getTermsForTermSetOrTerm(term);
//         //console.log(termIds);
//         termSearchString = termIds.map((id: any) => `owstaxIdepProcess:"GP0|#${id}"`).join(" OR ");
//         console.log(termSearchString);

//         switch (choice) {
//             case "documentlibrary":
//                 _query = {
//                   request:{
//                     '__metadata': {
//                       'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//                    },
//                    'QueryTemplate': `{searchterms} Path:"${_props.absoluteURL}" AND (${termSearchString}) IsDocument:"True" ListID:${_props.DocumentLibraryId}`,
//                    'RowLimit' : 1000,
//                    'TrimDuplicates': true,
//                     'SelectProperties': {
//                       'results': selectProperties
//                     }
//                   }
//                 };
//               //console.log('generateQuery-documentlibrary');
//               break;
      
//             case "hubsite":
//               _query = {
//                 request:{
//                   '__metadata': {
//                     'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//                  },
//                  'QueryTemplate': `${documentQuery} AND (${termSearchString}) AND (DepartmentId:{${_props.HubSiteId}} OR DepartmentId:${_props.HubSiteId})`,
//                  'RowLimit' : 1000,
//                  'TrimDuplicates': true,
//                   'SelectProperties': {
//                     'results': selectProperties
//                   }
//                 }
//               };
//               break;
      
//               case "currentSite":
//                 _query = {
//                   request:{
//                     '__metadata': {
//                       'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//                    },
//                    'QueryTemplate': `{searchterms} Path:"${_props.CurrentSiteUrl}" AND (${termSearchString}) AND ${documentQuery}`,
//                    'RowLimit' : 1000,
//                    'TrimDuplicates': true,
//                     'SelectProperties': {
//                       'results': selectProperties
//                     }
//                   }
//                 };
//               break;
      
//             case "sites":
//               let siteUrls = "";
//               _props.Sites?.forEach((site, key, arr) => {
//                 if (key !== (arr.length - 1)) {
//                   return siteUrls += `Path:"${site}" OR `;
//                 } else {
//                   return siteUrls += `Path:"${site}"`;
//                 }
//               });
  
//               _query = {
//                 request:{
//                   '__metadata': {
//                     'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//                  },
//                  'QueryTemplate': `(${siteUrls}) AND (${termSearchString}) AND ${documentQuery}"`,
//                  'RowLimit' : 1000,
//                  'TrimDuplicates': true,
//                   'SelectProperties': {
//                     'results': selectProperties
//                   }
//                 }
//               };
//               //console.log('generateQuery-sites');
//               break;
//             default:
//                 throw new Error('choice');
//       }
//           return _query;
//     }

//     public static PopulateDropDownHubSite = (hubSiteId : any):IPropertyPaneDropdownOption[] => {
//       let dropdownhubSite :IPropertyPaneDropdownOption[] = [];
  
//       if(hubSiteId){
//           sp.hubSites.getById(hubSiteId).get().then(response => {
//               dropdownhubSite.push({key: response.SiteId, text: response.Title});
//           });
//       } else(
//           dropdownhubSite.push({key: 0, text: 'No Hub site connected'})
//       );

//       return dropdownhubSite;
//     }

//     public PopulateDropDownDocumentLibrariesSearchResult = (siteUrl : string): IPropertyPaneDropdownOption[] => {
//       let dropdownLists: IPropertyPaneDropdownOption[] = [{ key: 'default', text: 'Choose library' }];

//       const _query = {
//         request:{
//           '__metadata': {
//             'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//          },
//          'QueryTemplate': `{searchterms} Path:"${siteUrl}"contentclass:STS_List_DocumentLibrary`,
//          'RowLimit' : 1000,
//          'TrimDuplicates': true,
//           'SelectProperties': {
//             'results': ['Title', 'Id']
//           }
//         }
//       }

//       this.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`, SPHttpClient.configurations.v1, {
//         body: JSON.stringify(_query),
//         headers: {
//           'odata-version': '3.0',
//           'accept': 'application/json;odata=nometadata',
//           'X-ClientService-ClientTag': 'NonISV|Bravero|ProcessDocuments',
//           'UserAgent': 'NonISV|Bravero|ProcessDocuments'
//       }
//       }).then(response => response.json()).then((r: any) => {
//         const results = r.PrimaryQueryResult.RelevantResults.Table.Rows;
//         results.forEach((row: any) => {
//           const cells = row.Cells;
//           const titleCell = cells.find((cell: any) => cell.Key === "Title");
//           const listIdCell = cells.find((cell: any) => cell.Key === "ListId");
    
//           if (titleCell && listIdCell) {
//             dropdownLists.push({ key: listIdCell.Value, text: titleCell.Value });
//           }
//         });
//       })
//       .catch(error => {
//         console.error(error);
//       });

//       //console.log(dropdownLists);

//       return dropdownLists;
//     }

//     //Get sites in this site collection in a propertypane dropdown
//     public PopulateDropDownSites = (hubSiteId: any): IPropertyPaneDropdownOption[] => {
//       let dropdownSites: IPropertyPaneDropdownOption[] = [];
//       let _query: any = {};
//       let selectProperties = ['Title', 'Url', 'ServerRedirectedEmbedURL', 'OriginalPath', 'Author', 'Path', 'SiteName', 'contentclass'];

//       if (hubSiteId !== null || undefined) {
//         _query = {
//           request:{
//             '__metadata': {
//               'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//            },
//            'QueryTemplate': `contentclass:STS_Site AND NOT WebTemplate:SPSPERS AND (DepartmentId:{${hubSiteId}} OR DepartmentId:${hubSiteId})`,
//            'RowLimit' : 1000,
//            'TrimDuplicates': true,
//             'SelectProperties': {
//               'results': selectProperties
//             }
//           }
//         };
//       } else {
//         _query = {
//           request:{
//             '__metadata': {
//               'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
//            },
//            'QueryTemplate': `contentclass:STS_Site AND NOT WebTemplate:SPSPERS`,
//            'RowLimit' : 1000,
//            'TrimDuplicates': true,
//             'SelectProperties': {
//               'results': selectProperties
//             }
//           }
//         };
//       }

//       this.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`, SPHttpClient.configurations.v1, {
//         body: JSON.stringify(_query),
//         headers: {
//           'odata-version': '3.0',
//           'accept': 'application/json;odata=nometadata',
//           'X-ClientService-ClientTag': 'NonISV|Bravero|ProcessDocuments',
//           'UserAgent': 'NonISV|Bravero|ProcessDocuments'
//       }
//       }).then(response => response.json()).then((r: any) => {
//         const results = r.PrimaryQueryResult.RelevantResults.Table.Rows;
//         results.forEach((row: any) => {
//           const cells = row.Cells;
//           const titleCell = cells.find((cell: any) => cell.Key === "Title");
//           const urlCell = cells.find((cell: any) => cell.Key === "Url");
    
//           if (titleCell && urlCell) {
//             dropdownSites.push({ key: urlCell.Value, text: titleCell.Value });
//           }
//         });
//       }).catch(error => {
//         console.error(error);
//       });

//       return dropdownSites;
//     }

//     public async getTermsForTermSetOrTerm(termId: string): Promise<string[]> {
//       const store: ITermStore = await taxonomy.getDefaultSiteCollectionTermStore().get();
//       const _terms: string[] = [];
      
//       const term = await store.getTermById(termId).select('Id', 'Name', 'Parent', 'PathOfTerm', 'IsRoot', 'LocalCustomProperties', 'CustomSortOrder', 'Labels', 'Description').usingCaching().get();
//       if (term) {
//       if (term.Id) {
//         _terms.push(term.Id.replace('/Guid(', '').replace(')/', ''));
//       }
//       //console.table(term);
//       if (term.IsRoot) {
//         //console.log('IsRoot');
//         const resultTerms = await term.terms.get();
//         //console.table(resultTerms);
//         resultTerms.forEach((resultTerm) => {
//         if (resultTerm.Id) {
//           _terms.push(resultTerm.Id.replace('/Guid(', '').replace(')/', ''));
//         }
//         });
//       }
//       }
//       return _terms;
//     }

//     public getTermNameFromTermId(termId: string): Promise<string> {
//       return taxonomy.getDefaultSiteCollectionTermStore().get().then(store => {
//       return store.getTermById(termId).select('Id', 'Name').usingCaching().get().then(term => {
//         return term.Name || '';
//       });
//       });
//     }

//     private getSearchResults(resultRows: any[]): ISharePointSearchResult[] {
//       // Map search results
//       let searchResults: ISharePointSearchResult[] = resultRows.map((elt: { Cells: { Key: string; Value: string; }[]; }) => {
//           // Build item result dynamically
//           // We can't type the response here because search results are by definition too heterogeneous so we treat them as key-value object
//           let result: ISharePointSearchResult = {
//               Title: "",
//               Path: "",
//               FileType: "",
//               HitHighlightedSummary: "",
//               AuthorOWSUSER: "",
//               owstaxidmetadataalltagsinfo: "",
//               Created: "",
//               UniqueID: "",
//               NormSiteID: "",
//               NormWebID: "",
//               NormListID: "",
//               NormUniqueID: "",
//               ListItemID: "",
//               ToolLink: "",
//               ToolLinkCategory: ""
//           };
//           elt.Cells.map((item: { Key: string; Value: string; }) => {
//               if (item.Key === "HtmlFileType" && item.Value) {
//                   result["FileType"] = item.Value;
//               }
//               else if (!result[item.Key]) {
//                   result[item.Key] = item.Value;
//               }
//           });
//           return result;
//       });
//       return searchResults;
//     }

//     //Get Description for a process
//     public getBeskrivingForProcess(listGuid:string | string [], termId: string){
//       const LIST_API_ENDPOINT = `/_api/web/lists('${listGuid}')`;
//       const FILTER = `$filter=TaxCatchAll/IdForTerm eq '${termId}'`;
//       const SELECT_QUERY = `$select=Beskrivning,TaxCatchAll/IdForTerm`;
//       const EXPAND_QUERY = `$expand=TaxCatchAll`;

//       let promise : Promise<any> = new Promise((resolve, reject) => {
//         let query = `${this.context.pageContext.web.absoluteUrl}${LIST_API_ENDPOINT}/Items?${FILTER}&${SELECT_QUERY}&${EXPAND_QUERY}`;
//         this.spHttpClient.get(
//           query, 
//           SPHttpClient.configurations.v1,
//           this._spHttpOptions.getNoMetaData
//         ).then((response: SPHttpClientResponse): Promise<any> => {
//           return response.json()
//         }).then((data: any) => {
//           if(data.value.length > 0){
//             resolve(data.value[0].Beskrivning);
//           } else {
//             resolve('');
//           }
//         }).catch((error) => {
//           console.error(error);
//           reject(error);
//         });
//       });
//       return promise;
      
//     }
// }
