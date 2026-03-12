export default class DocumentMapper{
    public name?: string;
    public editor?: string;
    public CreatedBy?: string;
    public fileType?: string;
    public docType?: string;
    public link?: string;
    public icon?: string;
    public docId?: number;
    public termId?: string;
    public termPath?: string[];
    public guid?: string;
    public isFavorite?: boolean;
    public itemID?: number;
    public docOwner?: string;
    public validTo?: string;
    public Filename?:string;
    public path? : string;
    public ModifiedBy?: string;
    public owstaxIdMapp?: string;


    /**
     * Constructor
     * @param {object} results 
     */
    constructor(results: { [x: string]: string | undefined; FileType: string | undefined; ServerRedirectedEmbedURL: string | undefined; Title: string | undefined; }){
        this.fileType = results.FileType;
        this.guid = results.UniqueID ? results.UniqueID.replace('{', '').replace('}', '') : undefined;
        this.link = results.ServerRedirectedEmbedURL;
        this.name = results.Title;
        this.Filename = results.Filename;
        this.path = results.Path;
        this.ModifiedBy = results.ModifiedBy;
        this.CreatedBy = results.CreatedBy;
        this.owstaxIdMapp = results.owstaxIdMapp ? (results.owstaxIdMapp.match(/GP0\|#([^;]+);L0/) ? results.owstaxIdMapp.match(/GP0\|#([^;]+);L0/)![1] : "") : "";
    }
}