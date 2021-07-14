import { Opportunity } from "../models/Opportunity";
import { MSGraphClient } from "@microsoft/sp-http";
import { Config } from "../config/config";
import { Web } from "@pnp/sp/webs";
import { OpportunityFile } from "../models/OpportunityFile";

export default class ActivitiesHelper {

    private graphClient: MSGraphClient;

    constructor(graphClient: MSGraphClient) {
        this.graphClient = graphClient;
    }

    public getOpportunityActivity = async (opp: Opportunity) => {

        try {
            var response = await this.graphClient.api(`${Config.ManagedPath}${Config.Domain},${opp.siteId},${opp.webId}/lists/${opp.listId}/activities`).version("beta").top(5).get();
            opp.activities = response !== [] ? response : { value: [] };
        } catch (e) {
            opp.activities = { value: [] };
        }

        return opp;
    }

    public getOpportunityFileActivity = async (oppFile: OpportunityFile) => {
        var web = Web(oppFile.opportunityUrl);

        var file = await web.getFileByUrl(oppFile.url).expand("ModifiedBy")
            .select("ModifiedBy/Title, ModifiedBy/UserPrincipalName, TimeLastModified")
            .get();

        oppFile.lastAccessed = file.TimeLastModified;
        oppFile.lastModified = file.TimeLastModified;
        oppFile.lastModifiedUserDisplayName = file["ModifiedBy"]["Title"];
        oppFile.lastModifiedUserLoginName = file["ModifiedBy"]["UserPrincipalName"];

        return oppFile;
    }
}