import { SharePointQueryable, SharePointQueryableInstance } from "./sharepointqueryable";
import { ODataValue } from "../odata/parsers";
import { Util } from "../utils/util";

export class SocialQuery extends SharePointQueryableInstance {

    /**
     * Creates a new instance of the SocialQuery class
     *
     * @param baseUrl The url or SharePointQueryable which forms the parent of this social query
     */
    constructor(baseUrl: string | SharePointQueryable, path = "_api/social.following") {
        super(baseUrl, path);
    }

    /**
     * Makes the current user start following a user, document, site, or tag
     *
     * @param actorInfo The actor to start following
     */
    public follow(actorInfo: SocialActorInfo): Promise<SocialFollowResult> {
        return this.clone(SocialQuery, "follow")
            .postCore({ body: this.createSocialActorInfoRequestBody(actorInfo) }, ODataValue<number>());
    }

    /**
     * Indicates whether the current user is following a specified user, document, site, or tag
     *
     * @param actorInfo The actor to find the following status for
     */
    public isFollowed(actorInfo: SocialActorInfo): Promise<boolean> {
        return this.clone(SocialQuery, "isfollowed")
            .postCore({ body: this.createSocialActorInfoRequestBody(actorInfo) }, ODataValue<boolean>());
    }

    /**
     * Makes the current user stop following a user, document, site, or tag
     *
     * @param actorInfo The actor to stop following
     */
    public stopFollowing(actorInfo: SocialActorInfo): Promise<void> {
        return this.clone(SocialQuery, "stopfollowing")
            .postCore({ body: this.createSocialActorInfoRequestBody(actorInfo) });
    }

    /**
     * Creates SocialActorInfo request body
     *
     * @param actorInfo The actor to create request body
     */
    private createSocialActorInfoRequestBody(actorInfo: SocialActorInfo): string {
        return JSON.stringify({
            "actor":
                Util.extend({
                    Id: null,
                    "__metadata": { "type": "SP.Social.SocialActorInfo" },
                }, actorInfo),
        });
    }
}

/**
 * Social actor info
 *
 */
export interface SocialActorInfo {
    AccountName?: string;
    ActorType: SocialActorType;
    ContentUri?: string;
    Id?: string;
    TagGuid?: string;
}

/**
 * Social actor type
 *
 */
export enum SocialActorType {
    User = 0,
    Document = 1,
    Site = 2,
    Tag = 3,
}

/**
 * Result from following
 *
 */
export enum SocialFollowResult {
    Ok = 0,
    AlreadyFollowing = 1,
    LimitReached = 2,
    InternalError = 3,
}
