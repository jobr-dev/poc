import { ItemActivity } from '@microsoft/microsoft-graph-types';

export class Opportunity {
    public id: string;
    public title: string;
    public client: string;
    public url: string;
    public pinned: boolean;
    public userHasAccess: boolean;
    public activities: {};
    public listId: string;
    public siteId: string;
    public webId: string;
}