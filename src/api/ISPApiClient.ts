import { LibraryName, ListName } from "../enums";
import { IAttachment, ISiteGroup, ISPColumn } from "../interfaces";
import IApiResult from "./IApiResult";

export interface ISPApiClient {
    readonly internalSiteName: string;
    readonly templateSiteName: string;
    checkAllListExistence: () => Promise<{ [key in ListName | LibraryName.Requests]: boolean }>;
    ensureList: (listName: ListName | LibraryName.Requests) => Promise<void>;
    deleteList: (listName: ListName) => Promise<boolean>;
    updateItemInList: (uniqueId: string, body: object, listName?: string) => Promise<boolean>;
    updateListItem(itemId: number, body: object, listName: ListName | LibraryName): Promise<boolean>
    deleteListItem: (id: number, listName: ListName) => Promise<boolean>;
    getFieldChoices: (listName: ListName, columnTitle: string) => Promise<string[]>;
    createListItemUsingPath: (body: object, listName: ListName, path?: string, channelSpecific?: boolean) => Promise<IApiResult>
    updateListItemsBatch: (updates: { uniqueId: string, body: object }[], listName: ListName) => Promise<void>
    updateList: (body: object, listName?: string) => Promise<void>
    updateColumn: (body: object, listName: string, columnName: string) => Promise<void>
    ensureColumns: (listName: ListName | LibraryName.Requests) => Promise<ISPColumn[]>
    getUserSPId: (userEmail: string) => Promise<number>
    init: (ctx?: string, teamId?: string, channelId?: string) => Promise<void>
    getUserRole: () => number
    getTeamId: () => string;
    getChannelId: () => string;
    getAbsoluteUrl: () => string;
    getChannelName: () => string;
    getRelativeUrl: () => string;
    getTeamName: () => Promise<string>;
    getSiteGroups: () => Promise<ISiteGroup[]>;
    getTemplatesAttachments: (listName: ListName) => Promise<IAttachment[]>;
    moveFolder: (path: string, newPath: string) => Promise<boolean>;
    hideLists: (listNames: ListName[]) => Promise<void>;
    log: (message: string, data?: any) => Promise<void>;
    getSiteGroupByName: (groupName: string, webUrl?: string) => Promise<ISiteGroup | undefined>;
    isPrivateChannel: () => Promise<boolean>;
}