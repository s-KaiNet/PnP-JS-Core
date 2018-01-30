import { List } from "./lists";
import { TemplateFileType, FileAddResult, File } from "./files";
import { Item, ItemUpdateResult } from "./items";
import { Util } from "../utils/util";
import { TypedHash } from "../collections/collections";

/**
 * Page promotion state
 */
export const enum PromotedState {
    /**
     * Regular client side page
     */
    NotPromoted = 0,
    /**
     * Page that will be promoted as news article after publishing
     */
    PromoteOnPublish = 1,
    /**
     * Page that is promoted as news article
     */
    Promoted = 2,
}

/**
 * Type describing the available page layout types for client side "modern" pages
 */
export type ClientSidePageLayoutType = "Article" | "Home";

/**
 * Column size factor. Max value is 12 (= one column), other options are 8,6,4 or 0
 */
export type CanvasColumnFactorType = 0 | 2 | 4 | 6 | 8 | 12;

/**
 * Represents the data and methods associated with client side "modern" pages
 */
export class ClientSidePage extends File {

    /**
     * Creates a new blank page within the supplied library
     * 
     * @param library The library in which to create the page
     * @param pageName Filename of the page, such as "page.aspx"
     * @param title The display title of the page
     * @param pageLayoutType Layout type of the page to use
     */
    public static create(library: List, pageName: string, title: string, pageLayoutType: ClientSidePageLayoutType = "Article"): Promise<ClientSidePage> {

        // see if file exists, if not create it
        return library.rootFolder.files.select("Name").filter(`Name eq '${pageName}'`).get().then((fs: any[]) => {

            if (fs.length > 0) {
                throw new Error(`A file with the name '${pageName}' already exists in the library '${library.toUrl()}'.`);
            }

            // get our server relative path
            return library.rootFolder.select("ServerRelativePath").get().then(path => {

                const pageServerRelPath = Util.combinePaths("/", path.ServerRelativePath.DecodedUrl, pageName);

                // add the template file
                return library.rootFolder.files.addTemplateFile(pageServerRelPath, TemplateFileType.ClientSidePage).then((far: FileAddResult) => {

                    // get the item associated with the file
                    return far.file.getItem().then((i: Item) => {

                        // update the item to have the correct values to create the client side page
                        return i.update({
                            ContentTypeId: "0x0101009D1CB255DA76424F860D91F20E6C4118",
                            Title: title,
                            ClientSideApplicationId: "b6917cb1-93a0-4b97-a84d-7cf49975d4ec",
                            PageLayoutType: pageLayoutType,
                            PromotedState: PromotedState.NotPromoted,
                            BannerImageUrl: {
                                Url: "/_layouts/15/images/sitepagethumbnail.png",
                            },
                            CanvasContent1: "",
                        }).then((iar: ItemUpdateResult) => new ClientSidePage(iar.item.file));
                    });
                });
            });
        });
    }

    /**
     * Creates a new instance of the ClientSidePage class
     *
     * @param baseUrl The url or SharePointQueryable which forms the parent of this web collection
     */
    constructor(file: File, public sections: CanvasSection[] = []) {
        super(file);
    }

    /**
     * Updates the properties of the underlying ListItem associated with this ClientSidePage
     * 
     * @param properties Set of properties to update
     * @param eTag Value used in the IF-Match header, by default "*"
     */
    public updateProperties(properties: TypedHash<any>, eTag = "*"): Promise<ItemUpdateResult> {
        return this.getItem().then(i => i.update(properties, eTag));
    }

    /**
     * Add a section to this page
     */
    public addSection(): CanvasSection {

        let order = 0;
        for (let i = 0; i < this.sections.length; i++) {
            if (this.sections[i].order > order) {
                order = this.sections[i].order + 1;
            }
        }

        const section = new CanvasSection(this, order);
        this.sections.push(section);
        return section;
    }
}

export class CanvasSection {

    constructor(public page: ClientSidePage, public order: number, public columns: CanvasColumn[] = []) {

    }

    public get defaultColumn(): CanvasColumn {

        if (this.columns.length < 1) {
            this.columns.push(new CanvasColumn(this));
        }

        return this.columns[0];
    }

    public addColumn(): CanvasColumn {

        let order = 0;
        for (let i = 0; i < this.columns.length; i++) {
            if (this.columns[i].order > order) {
                order = this.columns[i].order + 1;
            }
        }

        const column = new CanvasColumn(this, order);
        this.columns.push(column);
        return column;
    }

    public toHtml(): string {

        const html = [];

        for (let i = 0; i < this.columns.length; i++) {
            html.push(this.columns[i].toHtml());
        }

        return html.join("");
    }
}

export class CanvasColumn {

    constructor(
        public section: CanvasSection,
        public order: number,
        public controls: CanvasControl[] = [],
        private _factor: CanvasColumnFactorType = 12,
        private _dataVersion = "1.0") {
    }

    /**
     * Column size factor. Max value is 12 (= one column), other options are 8,6,4 or 0
     */
    public get factor(): CanvasColumnFactorType {
        return this._factor;
    }

    public set factor(value: CanvasColumnFactorType) {
        this._factor = value;
    }

    public addControl(): CanvasControl {
        const control = new CanvasControl(this);
        this.controls.push(control);
        return control;
    }

    public toHtml(): string {
        const html = [];

        if (this.controls.length < 1) {

            // we need to render an empty section
            const data = JSON.stringify({
                Position: {
                    ZoneIndex: this.section.order,
                    SectionIndex: this.order,
                    SectionFactor: this._factor,
                },
            });

            html.push(`<div data-sp-canvascontrol="" data-sp-canvasdataversion="${this._dataVersion}" data-sp-controldata="${data}"></div>`);

        } else {

            // if we have controls, render them
            for (let i = 0; i < this.controls.length; i++) {
                html.push(this.controls[i].toHtml(i + 1));
            }
        }

        return html.join("");
    }
}

export class CanvasControl {

    constructor(public column: CanvasColumn) {

    }

    public toHtml(index: number): string {
        return "";
    }


    //     "<div><div data-sp-canvascontrol="" data-sp-canvasdataversion="1.0" 
    // data-sp-controldata="&#123;&quot;controlType&quot;&#58;4,&quot;displayMode&quot;&#58;2,
    // &quot;id&quot;&#58;&quot;68081bcb-b14c-4a7c-b01b-e3117bc9456c&quot;,
    // &quot;position&quot;&#58;&#123;&quot;zoneIndex&quot;&#58;1,&quot;sectionIndex&quot;&#58;1,&quot;
    // controlIndex&quot;&#58;1&#125;,&quot;innerHTML&quot;&#58;&quot;&lt;p&gt;this is a test.&lt;/p&gt;\n&quot;,&quot;
    // editorType&quot;&#58;&quot;CKEditor&quot;&#125;"><div data-sp-rte=""><p>this is a test.</p>
    // </div></div></div>"
}

/**
 * Client side webpart object (retrieved via the _api/web/GetClientSideWebParts REST call)
 */
export interface ClientSidePageComponent {
    /**
     * Component type for client side webpart object
     */
    ComponentType: number;
    /**
     * Id for client side webpart object
     */
    Id: string;
    /**
     * Manifest for client side webpart object
     */
    Manifest: string;
    /**
     * Manifest type for client side webpart object
     */
    ManifestType: number;
    /**
     * Name for client side webpart object
     */
    Name: string;
    /**
     * Status for client side webpart object
     */
    Status: number;
}
