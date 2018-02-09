import { expect } from "chai";
import pnp from "../../src/pnp";
import { testSettings } from "../test-config.test";
import { Util } from "../../src/utils/util";
import {
    ClientSideText,
    ClientSideWebpart,
    ClientSideWebpartPropertyTypes,
    ClientSidePage,
} from "../../src/sharepoint/clientsidepages";

describe("Client-side Page", () => {

    if (testSettings.enableWebTests) {

        describe("create", () => {

            it("Should create a new page", () => {
                return expect(pnp.sp.web.addClientSidePage(`TestingAdd_${Util.getRandomString(4)}.aspx`)).to.eventually.be.fulfilled;
            });
        });

        describe("load", function () {

            const pageFileName = `TestingAdd_${Util.getRandomString(4)}.aspx`;

            before(done => {
                pnp.sp.web.addClientSidePage(pageFileName).then(_ => {
                    done();
                });
            });

            it("Should load from an existing file", () => {

                // need to make the path relative
                const rel = testSettings.webUrl.substr(testSettings.webUrl.indexOf("/sites/"));
                const promise = ClientSidePage.fromFile(pnp.sp.web.getFileByServerRelativeUrl(Util.combinePaths("/", rel, "SitePages", pageFileName)));
                return expect(promise).to.eventually.be.fulfilled;
            });
        });

        describe("save", () => {

            it("Should update a pages content with a text control", () => {
                return pnp.sp.web.addClientSidePage(`TestingAdd_${Util.getRandomString(4)}.aspx`).then(page => {

                    page.addSection().addControl(new ClientSideText("This is test text!!!"));

                    return expect(page.save()).to.eventually.be.fulfilled;
                });
            });

            it("Should update a pages content with an embed control", () => {
                return pnp.sp.web.getClientSideWebParts().then(parts => {

                    pnp.sp.web.addClientSidePage(`TestingAdd_${Util.getRandomString(4)}.aspx`).then(page => {

                        const part = ClientSideWebpart.fromComponentDef(parts.filter(c => c.Id === "490d7c76-1824-45b2-9de3-676421c997fa")[0]);

                        part.setProperties<ClientSideWebpartPropertyTypes.Embed>({
                            embedCode: "https://www.youtube.com/watch?v=IWQFZ7Lx-rg",
                        });

                        page.addSection().addControl(part);

                        return expect(page.save()).to.eventually.be.fulfilled;
                    });
                });
            });
        });

        describe("Page comments", () => {

            let page: ClientSidePage;

            before(done => {
                pnp.sp.web.addClientSidePage(`TestingAdd_${Util.getRandomString(4)}.aspx`).then(p => {
                    page = p;
                    done();
                });
            });

            it("Should disable", () => {
                return expect(page.disableComments()).to.eventually.be.fulfilled;
            });

            it("Should enable", () => {
                return expect(page.enableComments()).to.eventually.be.fulfilled;
            });
        });
    }
});
