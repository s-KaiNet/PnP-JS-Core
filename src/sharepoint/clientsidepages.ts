import { List } from "./lists";
import { TemplateFileType, FileAddResult, File } from "./files";
import { ItemUpdateResult } from "./items";
import { Util } from "../utils/util";

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
 * Represents the data and methods associated with client side "modern" pages
 */
export class ClientSidePage {

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

            return library.rootFolder.select("ServerRelativePath").get().then(path => {

                const pageServerRelPath = Util.combinePaths("/", path.ServerRelativePath.DecodedUrl, pageName);

                return library.rootFolder.files.addTemplateFile(pageServerRelPath, TemplateFileType.ClientSidePage).then((far: FileAddResult) => {

                    // get the item associated with the file
                    return far.file.getItem().then(i => {

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
    constructor(private _file: File) {




    }

    public delete(): Promise<void> {
        return this._file.delete();
    }
}
