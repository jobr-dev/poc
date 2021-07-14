import { UsedInsight } from "@microsoft/microsoft-graph-types";
import { Opportunity } from "../models/Opportunity";

const _getTime = (date?: Date) => {
    return date !== null ? date.getTime() : 0;
};

export default class Sorter {
    public static sortOpportunityByName = (a: Opportunity, b: Opportunity) => {
        let fa = a.title.toLowerCase(),
            fb = b.title.toLowerCase();

        if (fa < fb) {
            return -1;
        }
        if (fa > fb) {
            return 1;
        }
        return 0;
    }
    public static sortRecentDocumentByDate = (a: UsedInsight, b: UsedInsight): number => {
        const aDate = new Date(a.lastUsed.lastAccessedDateTime);
        const bDate = new Date(b.lastUsed.lastAccessedDateTime);
        return _getTime(bDate) - _getTime(aDate);
    }
}
