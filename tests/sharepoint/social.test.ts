import { expect } from "chai";
import { SocialQuery } from "../../src/sharepoint/social";

describe("Social", () => {
    let socialQuery: SocialQuery;

    beforeEach(() => {
        socialQuery = new SocialQuery("_api");
    });

    it("Should be an object", () => {
        expect(socialQuery).to.be.a("object");
    });
});
