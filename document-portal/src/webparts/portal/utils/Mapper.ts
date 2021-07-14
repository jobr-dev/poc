import { LocalFile } from "../models/LocalFile";

export default class Mapper {
    public static inputFilesToLocalFiles(fileList: FileList): LocalFile[] {
        var result: LocalFile[] = [];
        for (let index = 0; index < fileList.length; index++) {
            var current = new LocalFile();
            current.selected = true;
            current.file = fileList[index];
            result.push(current);
        }

        return result;
    }
}
