import { Opportunity } from "../models/Opportunity";
import { sp, ISearchQuery, ISearchResult, Web, SocialActorType } from "@pnp/sp/presets/all";
import { findIndex, uniqBy } from "@microsoft/sp-lodash-subset";
import { Config } from "../config/config";
import Sorter from "../utils/Sorter";
const data = require('../data/opportunities.json');
import { OpportunityFile } from "../models/OpportunityFile";
import { ApplicationIconList, ApplicationType, BrandIcons } from "../utils/BrandIcons";

export default class SearchService {
    public async getAllOpportunitites(): Promise<Opportunity[]> {
        const appSearchSettings: ISearchQuery = {
            RowLimit: 100,
            SelectProperties: ["Title", "Description", "SiteName", "OriginalPath", "ListId", "SiteId", "WebId"],
            Querytext: `${Config.Keys.SolutionKey}:${Config.SolutionConstant}`,
            TrimDuplicates: false
        };

        const results = await sp.search(appSearchSettings);

        var searchResults = results.PrimarySearchResults;

        var page = 2;
        do {
            var currentPageResults = await results.getPage(page);

            if (currentPageResults !== null) {
                searchResults = searchResults.concat(currentPageResults.PrimarySearchResults);
                page++;
            }
        } while (currentPageResults !== null);

        return (await this.mapPNPSearchResultsToOpportunity(searchResults)).sort(Sorter.sortOpportunityByName);
    }

    public async getAllOpportunititesFromStructure(): Promise<Opportunity[]> {
        var opps: Opportunity[] = data as Opportunity[];
        var userOpps = await this.getAllOpportunitites();

        opps.forEach(x => {
            var existing = userOpps.filter(u => u.url == x.url);

            x.userHasAccess = false;
            x.pinned = false;
            x.activities = {};

            if (existing.length > 0) {
                x.userHasAccess = true;
                x.pinned = existing[0].pinned;
                x.activities = existing[0].activities;
                x.listId = existing[0].listId;
                x.siteId = existing[0].siteId;
                x.webId = existing[0].webId;
            }
        });

        return Promise.resolve(opps.sort(Sorter.sortOpportunityByName));
    }

    public async search(query: string, rowLimit: number): Promise<OpportunityFile[]> {

        const appSearchSettings: ISearchQuery = {
            RowLimit: rowLimit,
            HiddenConstraints: `((isDocument:true OR FileType:jpg OR FileType:jpeg OR FileType:gif OR FileType:png) AND NOT Title:__siteIcon__ AND SiteName:https://${Config.Domain}${Config.ManagedPath}${Config.SolutionConstant}*)`,
            SelectProperties: ["Title", "SiteName", "LinkingUrl", "OriginalPath", "LastModifiedTime", "Created", "Modified", "Filename", "EditorOWSUSER", "DocumentLink", "FileType", "ModifiedOWSDATE", "UniqueId"],
            Querytext: query.length > 0 ? `*${query}*` : "*",
            TrimDuplicates: false
        };

        const results = await sp.search(appSearchSettings);

        var searchResults = results.PrimarySearchResults;

        return this.mapPNPSearchResultsToOpportunityFile(searchResults);
    }

    private async mapPNPSearchResultsToOpportunity(searchResults: ISearchResult[]): Promise<Opportunity[]> {

        var listId = "";

        return await Promise.all(uniqBy(searchResults, "Title").map(async x => {

            var web = Web(x.OriginalPath);
            var allProperties = await web.expand("AllProperties").select(`AllProperties/${Config.Keys.ClientKey},AllProperties/${Config.Keys.IdKey}`).get();

            if (listId == "") {
                var list = await web.lists.getByTitle("Documents").select("Id").get();
                listId = list.Id;
            }

            var opportunity = new Opportunity();
            opportunity.id = allProperties["AllProperties"][Config.Keys.IdKey];
            opportunity.title = this.generateOpportunityTitle(allProperties["AllProperties"][Config.Keys.IdKey]);
            opportunity.client = allProperties["AllProperties"][Config.Keys.ClientKey];
            opportunity.url = x.OriginalPath;
            opportunity.activities = { value: [] };
            opportunity.webId = x["WebId"];
            opportunity.siteId = x["SiteId"];
            opportunity.listId = listId;
            opportunity.pinned = await
                sp.social.isFollowed({
                    ContentUri: x.OriginalPath,
                    ActorType: SocialActorType.Site,
                });

            return opportunity;
        }));
    }

    private async mapPNPSearchResultsToOpportunityFile(searchResults: ISearchResult[]): Promise<OpportunityFile[]> {

        const files: OpportunityFile[] = [];
        for (let index = 0; index < searchResults.length; index++) {
            const result = searchResults[index];

            const appIdx = findIndex(ApplicationIconList, item => { return item.extensions.indexOf(result["FileType"].toLowerCase()) !== -1; });

            var defaultLink = result["DocumentLink"];
            if (result["FileType"] == "url") {
                var web = Web(result.SiteName);
                var text = await web.getFileById(result["UniqueId"]).getText();
                defaultLink = text.replace("[InternetShortcut]\nURL=", "");
            }

            const oppFile: OpportunityFile = {
                type: ApplicationType[appIdx],
                name: result.Title,
                editUrl: result["LinkingUrl"] ? result["LinkingUrl"] : defaultLink,
                url: result["DocumentLink"],
                lastAccessed: result["ModifiedOWSDATE"].toString(),
                lastModified: result["ModifiedOWSDATE"].toString(),
                lastModifiedUserDisplayName: result["EditorOWSUSER"].split('|')[1].trim(),
                lastModifiedUserLoginName: result["EditorOWSUSER"].split('|')[0].trim(),
                opportunityUrl: result.SiteName,
                opportunityName: result.SiteName.replace(`https://${Config.Domain}${Config.ManagedPath}${Config.SolutionConstant}`, "").toUpperCase(),
                pinned: false
            };

            files.push(oppFile);
        }
        return files;
    }

    private generateOpportunityTitle(oppId: string): string {
        const idSplit = oppId.split("-");
        const id = idSplit[0].replace("&", "").toUpperCase().repeat(2).substring(0, 4);
        const idNumber = `00${idSplit[2]}`;

        const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        const idInt = parseInt(idSplit[2].substring(1, 3));
        const index = idInt % alphabet.length;
        const index2 = (idInt * 3) % alphabet.length;

        return `${id}-${idNumber} - ${alphabet[index]}${alphabet[index2]}`;
    }
}