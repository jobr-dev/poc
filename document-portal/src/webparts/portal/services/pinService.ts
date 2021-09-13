import { findIndex } from "@fluentui/react";
import { sp } from "@pnp/sp";
import "@pnp/sp/social";
import { ISocialActor, SocialActorType, SocialActorTypes } from "@pnp/sp/social";
import { Web } from "@pnp/sp/webs";
import { Config } from "../config/config";
import { OpportunityFile } from "../models/OpportunityFile";
import { ApplicationIconList, ApplicationType } from "../utils/BrandIcons";
import "@pnp/sp/sites";

export default class PinService {

    public async pin(url: string, type: SocialActorType): Promise<void> {
        await sp.social.follow({
            ActorType: type,
            ContentUri: url
        });
    }

    public async unpin(url: string, type: SocialActorType): Promise<void> {
        await sp.social.stopFollowing({
            ActorType: type,
            ContentUri: url
        });
    }

    public async getPinnedFiles(): Promise<OpportunityFile[]> {
        const files = await sp.social.my.followed(SocialActorTypes.Document);

        return await this.parseOpportunityFiles(files);
    }

    private async parseOpportunityFiles(files: ISocialActor[]): Promise<OpportunityFile[]> {

        var oppList: OpportunityFile[] = [];
        for (let index = 0; index < files.length; index++) {
            const current = files[index];

            try {
                var web = Web(current.Uri.split("/Shared Documents/")[0]);

                var opportunity = await web.expand("AllProperties").select(`Title,AllProperties/${Config.Keys.SolutionKey},Url,Id`).get();

                var file = await web.getFileByUrl(current.Uri).expand("ModifiedBy,ListItemAllFields")
                    .select("ListItemAllFields/File_x0020_Type, LinkingUrl, ModifiedBy/Title, ModifiedBy/UserPrincipalName, TimeLastModified, Name, UniqueId, ServerRelativeUrl")
                    .get();

                const appIdx = findIndex(ApplicationIconList, item => { return item.extensions.indexOf(file["ListItemAllFields"]["File_x0020_Type"].toLowerCase()) !== -1; });

                if (file["ListItemAllFields"]["File_x0020_Type"] == "url") {
                    var text = await web.getFileById(file["UniqueId"]).getText();
                    file["LinkingUrl"] = text.replace("[InternetShortcut]\nURL=", "");
                }
                if (opportunity["AllProperties"][Config.Keys.SolutionKey] == Config.SolutionConstant) {
                    oppList.push({
                        type: ApplicationType[appIdx],
                        name: file.Name.split(".")[0],
                        editUrl: file.LinkingUrl !== "" ? file.LinkingUrl : `${new URL(opportunity.Url).origin}${file.ServerRelativeUrl}`,
                        url: `${new URL(opportunity.Url).origin}${file.ServerRelativeUrl}`,
                        lastAccessed: file.TimeLastModified,
                        lastModified: file.TimeLastModified,
                        lastModifiedUserDisplayName: file["ModifiedBy"]["Title"],
                        lastModifiedUserLoginName: file["ModifiedBy"]["UserPrincipalName"],
                        opportunityUrl: opportunity.Url,
                        opportunityName: opportunity.Title.toUpperCase(),
                        pinned: true
                    });
                }
            }
            catch { }
        }

        return oppList;
    }
}