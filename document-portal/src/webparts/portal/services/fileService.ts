import { Web } from "@pnp/sp/webs";
import { LocalFile } from "../models/LocalFile";
import * as XSLX from "xlsx";
import { sp } from "@pnp/sp";
import "@pnp/sp/social";

import { findIndex } from "@fluentui/react";
import { ApplicationIconList, ApplicationType } from "../utils/BrandIcons";
import { SocialActorType } from "@pnp/sp/social";
import { OpportunityFile } from "../models/OpportunityFile";
import { Config } from "../config/config";

export default class FileService {

    public async getFilesFromOpportunity(url: string): Promise<OpportunityFile[]> {
        var web = Web(url);
        const oppId = new URL(url).pathname.replace(`${Config.ManagedPath}${Config.SolutionConstant}`, "").toUpperCase();

        var items = await web.getList(`${Config.ManagedPath}${Config.SolutionConstant}${oppId}/Shared Documents/`).items.expand("Editor")
            .select("DocIcon,File_x0020_Type,FileRef, Editor/Title,Editor/Name,Modified_x0020_By,Modified,FileLeafRef,FileDirRef,Last_x0020_Modified,ServerRedirectedEmbedUri,UniqueId")
            .get();

        var oppItemList: OpportunityFile[] = [];
        for (let index = 0; index < items.length; index++) {
            const x = items[index];

            const appIdx = x["File_x0020_Type"]
                ? findIndex(ApplicationIconList, item => { return item.extensions.indexOf(x["File_x0020_Type"].toLowerCase()) !== -1; }) : -1;

            if (appIdx == ApplicationType.OneNote) {
                x["ServerRedirectedEmbedUri"] = `${url}/_layouts/15/Doc.aspx?sourcedoc={${x["UniqueId"]}}&action=edit`;
            }

            if (x["File_x0020_Type"] == "url") {
                var text = await web.getFileById(x["UniqueId"]).getText();
                x["ServerRedirectedEmbedUri"] = text.replace("[InternetShortcut]\nURL=", "");
            }

            oppItemList.push({
                type: appIdx !== -1 ? ApplicationType[appIdx] : "Folder",
                name: x["FileLeafRef"],
                editUrl: x["ServerRedirectedEmbedUri"] && x["File_x0020_Type"] !== "pdf" ? x["ServerRedirectedEmbedUri"].replace("interactivepreview", "edit") : x["FileRef"],
                url: `${new URL(url).origin}${x["FileRef"]}`,
                lastAccessed: x["Modified"],
                lastModified: x["Modified"],
                lastModifiedUserDisplayName: x["Editor"]["Title"],
                lastModifiedUserLoginName: x["Editor"]["Name"].replace("i:0#.f|membership|", ""),
                opportunityUrl: url,
                opportunityName: oppId,
                pinned: appIdx !== -1 ?
                    await sp.social.isFollowed({
                        ContentUri: `${new URL(url).origin}${x["FileRef"]}`,
                        ActorType: SocialActorType.Document,
                    }) : false
            });
        }

        return oppItemList;
    }

    public async addNewFileToOpportunity(url: string, fileExtension: string, folder: string): Promise<boolean> {
        var fileName = fileExtension == ".docx" ? "Document" : fileExtension == ".xlsx" ? "Book" : "Presentation";

        var web = Web(url);

        let counter = 0;
        let exists = true;

        while (exists) {
            counter++;
            var currentFile = web.getFolderByServerRelativePath(folder).files.getByName(`${fileName}${counter}${fileExtension}`);
            exists = await currentFile.exists();
        }

        var blob = this.getFileContent(fileExtension, `${fileName}${counter}${fileExtension}`);

        await web
            .getFolderByServerRelativePath(folder)
            .files.addChunked(`${fileName}${counter}${fileExtension}`, blob);

        return true;
    }

    public async addNewFolderToOpportunity(url: string, name: string, folder: string): Promise<boolean> {
        var web = Web(url);

        var existingFolder = await web.getFolderByServerRelativePath(`${folder}/${name}`).select("Exists").get();
        if (existingFolder.Exists) return true;

        await web.folders.addUsingPath
            (`${folder}/${name}`);

        return true;
    }

    public async uploadFiles(url: string, files: LocalFile[], parent: string = "Shared Documents"): Promise<boolean> {
        var web = Web(url);

        const originalParent = parent;
        for (let i = 0; i < files.length; i++) {
            const current = files[i];

            if (current.file["fullPath"]) {
                const folders = current.file["fullPath"].substr(1).split("/");
                folders.pop();

                for (let j = 0; j < folders.length; j++) {
                    const element = folders[j];
                    var result = await this.addNewFolderToOpportunity(url, element, parent);

                    if (result) {
                        parent += `/${element}`;
                    }
                }
            }
            await web
                .getFolderByServerRelativePath(parent)
                .files.addChunked(current.file.name, current.file);
            parent = originalParent;
        }

        return true;
    }

    public async deleteFile(url: string, fileName: string, folder: string): Promise<boolean> {
        var web = Web(url);

        try {
            var currentFile = web.getFolderByServerRelativePath(folder)
                .files.getByName(fileName);
            await currentFile.delete();
        }
        catch (e) {
            return await this.deleteFolder(url, fileName, folder);
        }

        return true;
    }

    public async deleteFolder(url: string, fileName: string, folder: string): Promise<boolean> {
        var web = Web(url);

        try {
            var currentFolder = web.getFolderByServerRelativePath(folder)
                .folders.getByName(fileName);
            await currentFolder.delete();
        }
        catch (e) {
            return false;
        }

        return true;
    }

    private getFileContent(fileExtension: string, fileName?: string): Blob {
        if (fileExtension === ".xlsx") {
            var wb = XSLX.utils.book_new();
            wb.Props = {
                Title: fileName
            };
            wb.SheetNames.push("Sheet1");

            var wbout = XSLX.write(wb, { type: "array", bookType: 'xlsx' });
            return new Blob([wbout], {
                type: "application/binary"

            });
        }

        return new Blob();
    }
}