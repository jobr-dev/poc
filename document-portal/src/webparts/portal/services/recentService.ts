import { graph } from "@pnp/graph";
import "@pnp/graph/insights";
import "@pnp/graph/users";
import "@pnp/sp/files";
import { UsedInsight } from "@microsoft/microsoft-graph-types";
import { OpportunityFile } from "../models/OpportunityFile";
import Sorter from "../utils/Sorter";
import { Config } from "../config/config";

export default class RecentService {
    private _excludeTypes: string[] = ["Web", "spsite", "Folder", "Archive", "Image", "Other"];

    public async getRecentFiles(): Promise<OpportunityFile[]> {
        const filter = this._excludeTypes.map(type => `resourceVisualization/type ne '${type}'`).join(' and ');

        const recentFiles: UsedInsight[] = await graph.me
            .insights
            .used
            .orderBy("LastUsed/LastAccessedDateTime", false)
            .filter(`ResourceVisualization/containerType eq 'Site' and ${filter}`).get();

        var currentUser = await graph.me.get();

        return recentFiles.filter(x => x.resourceVisualization.containerWebUrl && x.resourceVisualization.containerWebUrl.indexOf(`https://${Config.Domain}${Config.ManagedPath}${Config.SolutionConstant}`) > -1).sort(Sorter.sortRecentDocumentByDate).map((file) => {
            const fullUrl = `${this._updateSitePath(file.resourceVisualization.containerWebUrl)}/Shared Documents/${this.getFileNameWithExtension(file.resourceReference.webUrl)}`;

            return {
                name: file.resourceVisualization.title,
                lastAccessed: file.lastUsed.lastAccessedDateTime,
                lastModified: file.lastUsed.lastAccessedDateTime,
                type: file.resourceVisualization.type,
                editUrl: file.resourceReference.webUrl,
                opportunityUrl: this._updateSitePath(file.resourceVisualization.containerWebUrl),
                opportunityName: file.resourceVisualization.containerDisplayName.toUpperCase(),
                lastModifiedUserDisplayName: currentUser.displayName,
                lastModifiedUserLoginName: currentUser.userPrincipalName,
                url: fullUrl,
                pinned: false
            };
        });
    }

    private _updateSitePath(path: string): string {
        if (path) {
            const pathSplit = path.split(Config.ManagedPath);
            if (pathSplit.length === 2) {
                const siteUrlPath = pathSplit[1].substring(0, pathSplit[1].indexOf("/"));

                return `${pathSplit[0]}${Config.ManagedPath}${siteUrlPath}`;
            } else {
                const matches = path.match(/^https?\:\/\/([^\/?#]+)(?:[\/?#]|$)/i);
                if (matches && matches.length > 0) {
                    return matches[0];
                }
            }
        }
        return path;
    }

    private getFileNameWithExtension(url: string) {
        var parsedUrl = new URL(url);

        if (parsedUrl.search.length > 0) {
            return parsedUrl.search.split('&').filter(x => x.includes("file="))[0].split('=')[1];
        } else
            return parsedUrl.href.split("/").pop();
    }
}