import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx, SPFI } from "@pnp/sp";
import { FileInfo } from '@syncfusion/ej2-inputs';
import LZString from 'lz-string';
import { z } from "zod";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/files";

import { BaseApiClient } from "./BaseApiClient";
import IApiResult from './IApiResult';
import { ISPApiClient } from "./ISPApiClient";
import { ChannelType, LibraryName, ListName, SPGroupExternal, SPGroupInternal, SPGroupRO, TemplatesListName } from '../enums';
import { IAttachment, IChannelDetail, ISiteGroup, ISPColumn } from '../interfaces';
import { listColumns } from './listColumns';
import { getFileExtension, getFileNameWOExt } from '../utils/utils';

type Context = {
    teamId: string,
    channelId: string,
    channelName: string,
    siteUrl: string
}

const ContextSchema = z.object({
    c: z.string(),
    t: z.string(),
    n: z.string(),
    s: z.string().url(),
});

const parseContext = (value?: unknown): Context | undefined => {
    if (!value) {
        return undefined;
    }

    try {
        if (typeof value !== 'string') {
            return undefined;
        }

        const parsedJson = JSON.parse(LZString.decompressFromEncodedURIComponent(value));
        const parsedContactData = ContextSchema.safeParse(parsedJson);
        if (!parsedContactData.success) {
            return undefined;
        }


        return {
            teamId: parsedContactData.data.t,
            channelId: parsedContactData.data.c,
            channelName: parsedContactData.data.n,
            siteUrl: parsedContactData.data.s
        };
    } catch {
        return undefined;
    }
};

export default class SPApiClient extends BaseApiClient implements ISPApiClient {
    protected channelId: string;
    protected teamId: string;
    protected absoluteUrl: string;
    protected relativeUrl: string;
    protected userRole: number;
    protected channelName: string;
    protected teamName: string;
    protected channelType: ChannelType;

    // REFACTOR: Hardcoded Czech string — move to a named constant or config value
    // so it can be changed without hunting through implementation files.
    protected readonly documentFolder: string = 'Sdilene dokumenty';

    private listMappings: Partial<Record<ListName, TemplatesListName>> = {
        [ListName.RequestTypes]: TemplatesListName.RequestTypes,
        [ListName.Categories]: TemplatesListName.DocumentStructure,
    };
    protected libraryName: string = LibraryName.Documents;

    public sp: SPFI;

    constructor(
        webPartContext: WebPartContext,
        public internalSiteName: string,
        public templateSiteName: string,
    ) {
        super(webPartContext, () => this.getAbsoluteUrl());
    }


    /**
     * Checks existence of all required SharePoint lists and libraries.
     * @returns Object mapping each list/library name to a boolean indicating existence.
     */
    public async checkAllListExistence(): Promise<{ [key in ListName | LibraryName.Requests]: boolean }> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists?$select=Title`;
        const response = await this.spGet(apiUrl);

        const listItems = Array.isArray(response.data) ? response.data : [];
        const titles = listItems.map((l: any) => l.Title);
        const allNames: (ListName | LibraryName.Requests)[] = [
            ListName.Details,
            ListName.Requests,
            ListName.QA,
            ListName.Categories,
            ListName.Duties,
            ListName.Forms,
            LibraryName.Requests,
            ListName.RequiredDocuments,
            ListName.DocumentTypes,
            ListName.RequestTypes,
            ListName.Notifications,
            ListName.WorkflowActions,
            ListName.ActionToTrigger,
            ListName.Logs
        ];

        // Create result object with exact mapped typing
        const result = {} as {
            [key in ListName | LibraryName.Requests]: boolean
        };

        allNames.forEach(name => {
            result[name] = titles.includes(name);
        });
        return result;
    }

    /**
     * Ensures a SharePoint list exists, creates it if missing, and ensures required columns and views.
     * @param listName The name of the list or library.
     */
    public async ensureList(listName: ListName | LibraryName.Requests): Promise<void> {
        const listExists = await this.checkListExistence(listName);
        if (!listExists) {
            const result = await this.createList(listName);
            if (!result) {
                return;
            }
        }
        const missingColumns = await this.ensureColumns(listName);
        if (listName !== LibraryName.Requests && !listExists) {
            await this.ensureListView(listName, missingColumns);
        }
    }

    /**
     * Deletes a SharePoint list and all its items and folders.
     * @param listName The name of the list to delete.
     * @returns Promise resolving to true if deleted successfully.
     */
    public async deleteList(listName: ListName): Promise<boolean> {
        const itemsFilter = `startswith(ContentTypeId,'0x0100')`;

        const itemsUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        const itemsResponse = await this.spGet(`${itemsUrl}?$filter=${itemsFilter}`);
        const items = Array.isArray(itemsResponse.data) ? itemsResponse.data : [];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            await this.spDelete(`${itemsUrl}(${item.Id})`);
        }

        const foldersResponse = await this.spGet(itemsUrl);
        const folders = Array.isArray(foldersResponse.data) ? foldersResponse.data : [];

        for (let i = 0; i < folders.length; i++) {
            const item = folders[i];
            await this.spDelete(`${itemsUrl}(${item.Id})`);
        }

        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
        const result = await this.spDelete(apiUrl);
        return !!result.ok;
    }

    /**
     * Updates a SharePoint list with the provided body.
     * @param body Update payload.
     * @param listName The name of the list (defaults to libraryName).
     */
    public async updateList(body: object, listName: string = this.libraryName): Promise<void> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
        await this.spPatch(apiUrl, body);
    }

    /**
     * Updates a column in a SharePoint list.
     * @param body Update payload.
     * @param listName The name of the list.
     * @param columnName The column to update.
     */
    public async updateColumn(body: object, listName: string, columnName: string): Promise<void> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/fields/getByInternalNameOrTitle('${columnName}')`;
        const headers = {
            "Accept": "application/json;odata=nometadata",
            "Content-type": "application/json;odata=nometadata",
            "OData-Version": ""
        };
        await this.spPatch(apiUrl, body, headers);
    }

    /**
     * Updates an item in a SharePoint list by unique ID.
     * @param uniqueId The unique identifier of the item.
     * @param body Update payload.
     * @param listName The name of the list (defaults to libraryName).
     * @returns Promise resolving to true if updated successfully.
     */
    public async updateItemInList(uniqueId: string, body: object = {}, listName: string = this.libraryName): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${uniqueId}')`;
        const result = await this.spPatch(apiUrl, body);
        return result.ok ?? false;
    }

    /**
     * Deletes an item from the specified SharePoint list by its ID.
     * 
     * @param id - The ID of the item to delete.
     * @param listName - The name of the SharePoint list.
     * @returns A promise resolving to true if the item was successfully deleted, otherwise false.
     */
    public async deleteListItem(id: number, listName: ListName): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;
        const result = await this.spDelete(apiUrl);
        return result.ok === true;
    }

    /**
     * Updates a list item by item ID.
     * @param itemId The item ID.
     * @param body Update payload.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if updated successfully.
     */
    public async updateListItem(itemId: number, body: object = {}, listName: ListName | LibraryName): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
        const result = await this.spPatch(apiUrl, body);
        return result.ok ?? false;
    }

    /**
     * Retrieves choice options for a field in a SharePoint list.
     * @param listName The name of the list.
     * @param columnTitle The column title.
     * @returns Array of choice strings.
     */
    public async getFieldChoices(listName: ListName, columnTitle: string): Promise<string[]> {
        const apiUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/fields/getbyinternalnameortitle('${columnTitle}')`;
        const result = await this.spGet(apiUrl);
        return result.data && (result.data as any).Choices ? (result.data as any).Choices : [];
    }

    /**
     * Creates a list item using a specified path in SharePoint.
     * @param body Item creation payload.
     * @param listName The name of the list.
     * @param path Optional path for item placement.
     * @param channelSpecific Whether to use channel-specific path.
     * @returns API result of the operation.
     */
    public async createListItemUsingPath(body: object = {}, listName: ListName, path?: string, channelSpecific: boolean = true): Promise<IApiResult> {
        let folderPath = `${this.absoluteUrl}/lists/${listName}`;
        if (channelSpecific) {
            folderPath += `/${this.channelName}`;
        }
        if (path) {
            folderPath += `${path}`;
        }
        const bodyWithPath = {
            "listItemCreateInfo": {
                "FolderPath": {
                    "DecodedUrl": folderPath
                },
                "UnderlyingObjectType": 0
            },
            "bNewDocumentUpdate": false,
            ...body
        };
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/AddValidateUpdateItemUsingPath`;
        const result = await this.spPost(apiUrl, bodyWithPath);
        return result;
    }

    /**
     * Updates multiple list items in batch.
     * @param updates Array of update objects with uniqueId and body.
     * @param listName The name of the list.
     */
    public async updateListItemsBatch(updates: { uniqueId: string, body: object }[], listName: ListName): Promise<void> {
        const chunkSize = 50;
        const chunks = this.chunkArray(updates, chunkSize);
        const results: IApiResult[] = [];

        for (const chunk of chunks) {
            const result = await this.sendItemsUpdateBatchChunk(listName, chunk);
            results.push(result);
        }
    }

    /**
     * Splits an array into chunks of specified size.
     * @param arr The array to chunk.
     * @param size The chunk size.
     * @returns Array of chunked arrays.
     */
    public chunkArray<T>(arr: T[], size: number): T[][] {
        const chunks: T[][] = [];
        for (let i = 0; i < arr.length; i += size) {
            chunks.push(arr.slice(i, i + size));
        }
        return chunks;
    }

    /**
     * Sends a batch update chunk for list items.
     * @param listName The name of the list.
     * @param updates Array of update objects.
     * @returns API result of the batch operation.
     */
    private async sendItemsUpdateBatchChunk(
        listName: ListName,
        updates: { uniqueId: string, body: object }[],
    ): Promise<IApiResult> {

        const boundary = "batch_" + Date.now();
        const changeBoundary = "changeset_" + Date.now();

        const batchParts = updates.map((u, index) => {
            const itemUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${u.uniqueId}')`;
            //const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${uniqueId}')`;

            return [
                `--${changeBoundary}`,
                `Content-Type: application/http`,
                `Content-Transfer-Encoding: binary`,
                `Content-ID: ${index}`,  // <-- first empty line
                '',
                `PATCH ${itemUrl} HTTP/1.1`,
                `Content-Type: application/json;odata=nometadata`,
                `IF-MATCH: *`,
                //`X-HTTP-Method: MERGE`,             // <-- second empty line
                '',
                JSON.stringify(u.body),
                '',
                ''
            ].join('\r\n');
        });

        const batchBody = `--${boundary}\r\n`
            + `Content-Type: multipart/mixed; boundary=${changeBoundary}\r\n\r\n`
            + batchParts.join('\r\n')
            + `--${changeBoundary}--`
            + `\r\n--${boundary}--\r\n`;

        const response = await this.spPost(
            `${this.absoluteUrl}/_api/$batch`,
            batchBody,
            {
                "Content-Type": `multipart/mixed; boundary=${boundary}`,
                "Accept": "application/json;odata=nometadata",
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0"
            }
        );

        return response;
    }

    /**
     * Checks if a SharePoint list exists.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if exists.
     */
    public async checkListExistence(listName: ListName | LibraryName): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
        const response = await this.spGet(apiUrl, {}, false);
        return response.code !== 404;
    }

    /**
     * Creates a SharePoint list or library.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if created successfully.
     */
    public async createList(listName: ListName | LibraryName.Requests): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists`;
        const body = {
            Title: listName,
            BaseTemplate: listName === LibraryName.Requests ? 101 : 100,
            EnableFolderCreation: true
        };
        const response = await this.spPost(apiUrl, body);
        return response.ok === true;
    }

    /**
     * Ensures required columns exist in a SharePoint list.
     * @param listName The name of the list or library.
     * @returns Array of missing columns.
     */
    public async ensureColumns(listName: ListName | LibraryName.Requests): Promise<ISPColumn[]> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields`;
        const result = await this.spGet(apiUrl);
        const fields = result.data as any[];
        const columns = listColumns[listName];
        const missingColumns: ISPColumn[] = [];
        for (let i = 0; i < columns.length; i++) {
            const column = columns[i];
            const field = fields.find(f => f.InternalName === column.Title);
            if (!field) {
                missingColumns.push(column);
            }
            if (!field) {
                await this.createColumn(listName, column);
            } else if (column.FieldTypeKind === '7') {
                const lookupListUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${column.lookUpListTitle}')`;
                const lookupListResult = await this.spGet(lookupListUrl);
                const lookupListId = (lookupListResult.data as any).Id;
                const fieldType = column.allowMultiple ? 'LookupMulti' : 'Lookup';
                const multAttr = column.allowMultiple ? "Mult='TRUE'" : '';
                const lookupFieldXml = `<Field Type='${fieldType}' DisplayName='${column.Title}' ${multAttr} Name='${column.Title}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
                const existingLookupListId = field.LookupList;
                if (existingLookupListId !== lookupListId) {
                    const body = { 'SchemaXml': lookupFieldXml };
                    const updateApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
                    await this.spPatch(updateApiUrl, body);
                }
            }
        }
        return missingColumns;
    }

    /**
     * Ensures the default list view contains required columns.
     * @param listName The name of the list.
     * @param columns Array of columns to ensure in the view.
     */
    public async ensureListView(listName: ListName, columns: ISPColumn[]): Promise<void> {
        const viewId = await this.getDefaultListViewId(listName);
        //const columns = listColumns[listName];
        for (let i = 0; i < columns.length; i++) {
            const column = columns[i];
            const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/views('${viewId}')/ViewFields/AddViewField('${column.Title}')`;
            await this.spPost(apiUrl);
        }
    }

    /**
     * Retrieves the default view ID for a SharePoint list.
     * @param listName The name of the list.
     * @returns The default view ID as a string.
     */
    private async getDefaultListViewId(listName: ListName): Promise<string> {
        const apiUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/DefaultView`;
        const result = await this.spGet(apiUrl);
        return result.data ? (result.data as any).Id : '';
    }

    /**
     * Creates a column in a SharePoint list.
     * @param listName The name of the list or library.
     * @param column Column definition object.
     * @returns Promise resolving to true if created successfully.
     */
    private async createColumn(listName: ListName | LibraryName.Requests, column: ISPColumn): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields`;
        if (column.FieldTypeKind !== '7') {
            const { displayName, allowMultiple, ...baseColumnProps } = column;
            const body: { [key: string]: object | string | boolean } = { ...baseColumnProps };
            if (column.FieldTypeKind === '6' || column.FieldTypeKind === '15') {
                body['@odata.type'] = column.FieldTypeKind === '15' ? 'SP.FieldMultiChoice' : 'SP.FieldChoice';
                const templatesListName = listName !== LibraryName.Requests ? this.listMappings[listName] : undefined;
                if (!column.Choices && templatesListName) {
                    const webUrl = `${this.absoluteUrl.replace(this.relativeUrl, '')}/sites/${this.templateSiteName}`;
                    const columnsApi = `${webUrl}/_api/lists/getbytitle('${templatesListName}')/fields/getbyinternalnameortitle('${column.Title}')`;
                    const col = await this.spGet(columnsApi);
                    body.Choices = (col.data as any).Choices ?? [];
                }
            }
            if (column.FieldTypeKind === '20') {
                body['@odata.type'] = 'SP.FieldUser';
                body.AllowMultipleValues = !!allowMultiple;
            }
            const result = await this.spPost(apiUrl, body);
            if (result.ok) {
                const fieldApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
                const body = { Title: displayName };
                await this.spPatch(fieldApiUrl, body);
            }
            return !!result.ok;
        } else {
            const lookupListUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${column.lookUpListTitle}')`;
            const lookupListResult = await this.spGet(lookupListUrl);
            const lookupListId = (lookupListResult.data as any).Id;
            const fieldType = column.allowMultiple ? 'LookupMulti' : 'Lookup';
            const multAttr = column.allowMultiple ? "Mult='TRUE'" : '';
            const lookupFieldXml = `<Field Type='${fieldType}' DisplayName='${column.Title}' ${multAttr} Name='${column.Title}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
            const xmlApiUrl = `${apiUrl}/CreateFieldAsXml`;
            if (lookupListId) {
                const body = { "parameters": { 'SchemaXml': lookupFieldXml } };
                const result = await this.spPost(xmlApiUrl, body);
                if (result.ok) {
                    const fieldApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
                    const body = { Title: column.displayName };
                    await this.spPatch(fieldApiUrl, body);
                }
                return !!result.ok;
            }

        }
        return false;
    }

    /**
     * Creates an item in a SharePoint list.
     * @param body Item creation payload.
     * @param listName The name of the list or library.
     * @returns API result of the operation.
     */
    public async createListItem(body: object = {}, listName: ListName | LibraryName): Promise<IApiResult> {
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        const result = await this.spPost(apiUrl, body);
        return result;
    }

    /**
     * Ensures a user exists in SharePoint and returns their ID.
     * @param userEmail The user's email address.
     * @param webUrl Optional web URL (defaults to absoluteUrl).
     * @returns The SharePoint user ID.
     */
    public async getUserSPId(userEmail: string, webUrl: string = this.absoluteUrl): Promise<number> {
        const result: IApiResult = await this.spPost(`${webUrl}/_api/web/ensureUser`, {
            logonName: userEmail
        });
        return result.data && (result.data as any).Id ? (result.data as any).Id : 0;
    }

    public async init(ctx?: string, teamId?: string, channelId?: string): Promise<void> {
        if (ctx) {
            await this.initWithCtx(ctx);
        } else if (teamId && channelId) {
            await this.initWithIds(teamId, channelId);
        }
    }

    public async initWithCtx(ctx: string): Promise<void> {
        const parsed = parseContext(ctx);
        if (!parsed) {
            throw new Error('Invalid ctx format');
        }
        const { teamId, channelId, channelName, siteUrl } = parsed;
        const teamSiteUrl = siteUrl;
        this.absoluteUrl = teamSiteUrl;
        this.relativeUrl = new URL(teamSiteUrl).pathname;
        this.teamId = teamId;
        this.channelId = channelId;
        this.userRole = await this.getUserChannelRole(teamId, channelId);
        this.channelName = channelName;
        this.sp = spfi(teamSiteUrl).using(SPFx(this.webPartContext));
    }

    public async initWithIds(teamId: string, channelId: string): Promise<void> {
        const channelDetail = await this.getChannelDetail(teamId, channelId);
        const teamSiteUrl = channelDetail.absoluteUrl ? channelDetail.absoluteUrl : await this.getTeamSiteUrl(teamId);
        this.absoluteUrl = teamSiteUrl;
        this.relativeUrl = new URL(teamSiteUrl).pathname;
        this.teamId = teamId;
        this.channelId = channelId;
        this.userRole = channelDetail.userRole;
        this.channelName = channelDetail.displayName;
        this.sp = spfi(teamSiteUrl).using(SPFx(this.webPartContext));
        this.upgradeLegacyUrl();
    }

    private upgradeLegacyUrl(): void {
        const newCtx = LZString.compressToEncodedURIComponent(JSON.stringify({
            t: this.teamId,
            c: this.channelId,
            n: this.channelName,
            s: this.absoluteUrl
        }));

        const newUrl = `${window.location.pathname}?ctx=${newCtx}`;

        window.history.replaceState(null, '', newUrl);
    }

    private async getChannelDetail(teamId: string, channelId: string): Promise<IChannelDetail> {
        const apiUrl = `teams/${teamId}/channels/${channelId}/filesFolder`;
        const result = await this.graphGet(apiUrl);
        const displayName = await this.getChannelDisplayName(teamId, channelId);
        const relativeUrl = result.data ? new URL((result.data as any).webUrl).pathname : '';
        const origin = result.data ? new URL((result.data as any).webUrl).origin : '';
        const parts = relativeUrl ? relativeUrl.replace('/sites/', '').split('/') : '';
        const absoluteUrl = parts && parts[0] && origin ? `${origin}/sites/${parts[0]}` : '';
        const userRole = await this.getUserChannelRole(teamId, channelId);
        return { displayName: displayName, absoluteUrl: absoluteUrl, userRole: userRole };
    }

    private async getChannelDisplayName(teamId: string, channelId: string): Promise<string> {
        const apiUrl = `teams/${teamId}/channels/${channelId}`;
        const result = await this.graphGet(apiUrl);
        return result.data ? (result.data as any).displayName : '';
    }


    private async getTeamSiteUrl(id: string): Promise<string> {
        const teamSite = await this.graphGet(`groups/${id}/sites/root`);
        return teamSite.data ? (teamSite.data as any).webUrl : '';
    }

    /**
     * Gets the current user's role in the channel.
     * @returns The user role as a number.
     */
    public getUserRole(): number {
        return this.userRole;
    }

    /**
     * Gets the current team ID.
     * @returns The team ID as a string.
     */
    public getTeamId(): string {
        return this.teamId;
    }

    /**
     * Gets the current channel ID.
     * @returns The channel ID as a string.
     */
    public getChannelId(): string {
        return this.channelId;
    }

    /**
     * Gets the absolute URL for the current site.
     * @returns The absolute URL as a string.
     */
    public getAbsoluteUrl(): string {
        return this.absoluteUrl;
    }

    /**
     * Gets the current channel name.
     * @returns The channel name as a string.
     */
    public getChannelName(): string {
        return this.channelName;
    }

    /**
     * Gets the relative URL for the current site.
     * @returns The relative URL as a string.
     */
    public getRelativeUrl(): string {
        return this.relativeUrl;
    }

    private async getChannelType(): Promise<ChannelType> {
        if (this.channelType) {
            return this.channelType;
        }
        const channelDetails = await this.graphGet(`/teams/${this.getTeamId()}/channels/${this.getChannelId()}`);
        const channelType = (channelDetails.data as any)?.membershipType === "private" ? ChannelType.Private : ChannelType.Standard;
        this.channelType = channelType;
        return channelType;
    }

    public async isPrivateChannel(): Promise<boolean> {
        const channelType = await this.getChannelType();
        return channelType === ChannelType.Private;
    }

    /**
     * Gets the current team name, loading it if not already set.
     * @returns Promise resolving to the team name.
     */
    public async getTeamName(): Promise<string> {
        if (this.teamName) {
            return this.teamName;
        }
        const teamName = await this.loadTeamName();
        return teamName;
    }

    /**
     * Gets the user's role in a Teams channel.
     * @param teamId The team ID.
     * @param channelId The channel ID.
     * @returns Role as a number.
     */
    private async getUserChannelRole(teamId: string, channelId: string): Promise<number> {
        const memberApiUrl = `/teams/${teamId}/channels/${channelId}/members`;
        const members = await this.graphGet(memberApiUrl);
        const userRole = Array.isArray(members.data)
            && members.data.some((item: any) => item.email === this.webPartContext.pageContext.user.email && item.roles.some((role: string) => role.toLowerCase() === 'owner'))
            ? 0
            : 1;
        return userRole;
    }

    /**
     * Loads the team name from Microsoft Graph if not already set.
     * @returns Promise resolving to the team name.
     */
    protected async loadTeamName(): Promise<string> {
        if (this.teamName) {
            return this.teamName;
        }
        const teamSite = await this.graphGet(`groups/${this.teamId}/sites/root`);
        const teamName = teamSite.data ? (teamSite.data as any).displayName : '';
        this.teamName = teamName;
        return teamName;
    }

    /**
     * Ensures a folder exists at the specified path, creating it if necessary.
     * @param path The server-relative path.
     * @param name The folder name.
     * @returns UniqueId of the folder or undefined.
     */
    public async ensureFolder(path: string, name: string): Promise<string | undefined> {
        const checkApiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}/${name}')`;
        //await this.log('Ensuring folder:', { path, name });
        const check = await this.spGet(checkApiUrl, {}, false);
        if (check.ok && check.data) {
            return (check.data as any).UniqueId;
        }
        const apiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')/folders/add('${name}')`;
        const result = await this.spPost(apiUrl);
        return (result.data as any).UniqueId;
    }

    /**
     * Creates multiple list items in batch.
     * @param itemBodies Array of item creation payloads.
     * @param listName The name of the list or library.
     */
    public async createListItemsBatch(itemBodies: object[], listName: ListName | LibraryName): Promise<void> {
        const chunkSize = 50;
        const chunks = this.chunkArray(itemBodies, chunkSize);
        const results: IApiResult[] = [];

        for (const chunk of chunks) {
            const result = await this.sendItemsPostBatchChunk(listName, chunk);
            results.push(result);
        }
    }

    /**
     * Sends a batch chunk for creating list items.
     * @param listName The name of the list or library.
     * @param itemBodies Array of item creation payloads.
     * @returns API result of the batch operation.
     */
    private async sendItemsPostBatchChunk(
        listName: ListName | LibraryName,
        itemBodies: object[],
    ): Promise<IApiResult> {
        const boundary = "batch_" + Date.now();

        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;

        const batchParts = itemBodies.map(body => {
            return [
                `--${boundary}`,
                'Content-Type: application/http',
                'Content-Transfer-Encoding: binary',
                '',
                `POST ${apiUrl} HTTP/1.1`,
                'Content-Type: application/json',
                'Accept: application/json',
                '',
                JSON.stringify(body),
                ''
            ].join('\r\n');
        });

        const batchBody = [...batchParts, `--${boundary}--\r\n`].join('\r\n');

        const response = await this.spPost(
            `${this.absoluteUrl}/_api/$batch`,
            batchBody,
            {
                'Content-Type': `multipart/mixed; boundary=${boundary}`
            }
        );

        return response;
    }

    /**
     * Moves a list item to a new folder within the document structure.
     * @param id The item ID.
     * @param fileDirRef The target folder's file directory reference.
     * @param listName The name of the list.
     */
    public async moveListItem(id: number, fileDirRef: string, listName: ListName): Promise<void> {
        const getUrl = `${this.absoluteUrl}/_api/web/GetListByTitle('${listName}')/items(${id})?$select=FileRef,FileDirRef,FileLeafRef`;
        const item = await this.spGet(getUrl);
        if (item.data) {
            const newFileDirRef = encodeURIComponent((item.data as any).FileRef.replace((item.data as any).FileDirRef, fileDirRef));
            const moveUrl = `${this.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${(item.data as any).FileRef}')/moveto(newurl='${newFileDirRef}',flags=1)`;
            await this.spPost(moveUrl);
            // await this.moveFile((item.data as any).FileRef, fileDirRef, (item.data as any).FileLeafRef);
        }
    }

    /**
     * Generates a batch string for setting permissions.
     * @param boundary Batch boundary string.
     * @param groupId Group ID.
     * @param roledefid Role definition ID.
     * @param apiUrlBase Base API URL.
     * @returns Batch string for permissions.
     */
    public getPermissionBatchString(boundary: string, groupId: number, roledefid: number, apiUrlBase: string): string {
        return [
            `--${boundary}`,
            'Content-Type: application/http',
            'Content-Transfer-Encoding: binary',
            '',
            `POST ${apiUrlBase}(principalid=${groupId},roledefid=${roledefid}) HTTP/1.1`,
            'Accept: application/json',
            '',
            ''
        ].join('\r\n');
    }

    /**
     * Retrieves site groups from the current SharePoint site.
     * @param webUrl Optional web URL.
     * @returns Array of site group objects.
     */
    public async getSiteGroups(webUrl?: string): Promise<ISiteGroup[]> {
        const apiUrl = `${webUrl ? webUrl : this.absoluteUrl}/_api/web/sitegroups`;
        const result = await this.spGet(apiUrl);
        const sitegroups: ISiteGroup[] = [];
        const groups: any[] = Array.isArray(result.data) ? result.data : [];
        for (let i = 0; i < groups.length; i++) {
            const group = groups[i];
            if (Object.values(SPGroupInternal).includes(group.Title) || Object.values(SPGroupExternal).includes(group.Title) || Object.values(SPGroupRO).includes(group.Title)) {
                //TODO map function
                sitegroups.push({
                    id: group.Id,
                    title: group.Title,
                    users: []
                });
            }
        }
        return sitegroups;
    }

    public async getSiteGroupByName(groupName: string, webUrl?: string): Promise<ISiteGroup | undefined> {
        const apiUrl = `${webUrl ? webUrl : this.absoluteUrl}/_api/web/sitegroups/getByName('${encodeURIComponent(groupName)}')`;
        const result = await this.spGet(apiUrl);
        const group = result.data as any;
        if (!group) {
            return undefined;
        }
        return {
            id: group.Id,
            title: group.Title,
            users: []
        };
    }

    /**
     * Attaches a file to a list item.
     * @param id The item ID.
     * @param file The file to attach.
     * @param listName The name of the list.
     * @returns API result of the operation.
     */
    public async addAttachmentFile(id: number, file: FileInfo, listName: ListName): Promise<IApiResult> {
        const body = await (file.rawFile as File).arrayBuffer();
        return await this.addAttachmentArrayBuffer(id, file.name, body, listName);
    }

    /**
     * Attaches an array buffer as a file to a list item, optionally skipping if file exists.
     * @param id The item ID.
     * @param fileName The file name.
     * @param buffer The file content as ArrayBuffer.
     * @param listName The name of the list.
     * @param skipIfExists Whether to skip if file exists.
     * @returns API result of the operation.
     */
    public async addAttachmentArrayBuffer(
        id: number,
        fileName: string,
        buffer: ArrayBuffer,
        listName: ListName,
        skipIfExists: boolean = false
    ): Promise<IApiResult> {
        const ext = getFileExtension(fileName);
        const name = getFileNameWOExt(fileName);
        let finalName = fileName;
        let count = 1;

        while (await this.checkAttachmentExists(id, finalName, listName)) {
            if (skipIfExists) {
                return { ok: true };
            }
            finalName = `${name} (${count}).${ext}`;
            count++;
        }

        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})/AttachmentFiles/add(FileName='${finalName}')`;
        const response = await this.spPost(apiUrl, buffer);
        return response;
    }

    /**
     * Checks if an attachment exists for a list item.
     * @param itemId The item ID.
     * @param fileName The file name.
     * @param listName The name of the list.
     * @returns Promise resolving to true if exists.
     */
    private async checkAttachmentExists(itemId: number, fileName: string, listName: ListName): Promise<boolean> {
        const apiUrl = `${this.relativeUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/AttachmentFiles?$filter=FileName eq '${fileName}'`;

        const result = await this.spGet(apiUrl);
        return !!result.ok && Array.isArray(result.data) && result.data.length > 0;
    }

    /**
     * Retrieves template attachments from a SharePoint list.
     * @param listName The name of the list.
     * @returns Array of attachment objects.
     */
    public async getTemplatesAttachments(listName: ListName): Promise<IAttachment[]> {

        const folderPath = `${this.relativeUrl}/lists/${listName}`;
        const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100') and Title eq 'šablony'`;
        const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=${filter}`;
        const result = await this.spGet(apiUrl);
        const id = Array.isArray(result.data) ? result.data[0]?.Id : undefined;
        if (!id) {
            return [];
        }
        const attachUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/items(${id})/AttachmentFiles`;
        const attachResult = await this.spGet(attachUrl);
        const attachments: IAttachment[] = (Array.isArray(attachResult.data) ? attachResult.data : []).map((att: any) => ({
            name: att.FileName,
            serverRelativeUrl: att.ServerRelativeUrl
        }));
        return attachments;
    }

    /**
     * Ensures a folder exists in a SharePoint list, creating and moving it if necessary.
     * @param name The folder name.
     * @param listName The name of the list or library.
     * @param path Optional path for folder placement.
     */
    public async ensureListFolder(name: string, listName: ListName | LibraryName, path?: string): Promise<void> {
        const fileDirRef = `${this.relativeUrl}/Lists/${listName}${path ? `/${path}` : ''}`;
        const fileRef = `${fileDirRef}/${name}`;
        const checkApiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${fileRef}')`;
        const folderCheck = await this.spGet(checkApiUrl, {}, false);
        if (!folderCheck.ok) {
            const createBody = {
                "ContentTypeId": "0x0120",
                'Title': name,
            };
            const result = await this.createListItem(createBody, listName);
            const renameBody = {
                'Title': name,
                'FileLeafRef': name,
            };
            await this.updateListItem((result.data as any).Id, renameBody, listName);
            if (path) {
                await this.moveFolder(`${this.relativeUrl}/Lists/${listName}/${name}`, fileRef);
            }
        }
    }

    /**
     * Moves a folder to a new path in SharePoint.
     * @param path The current folder path.
     * @param newPath The new folder path.
     * @returns Promise resolving to true if moved successfully.
     */
    public async moveFolder(path: string, newPath: string): Promise<boolean> {
        const apiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')/moveTo(newurl='${newPath}')`;
        const result = await this.spPost(apiUrl);
        return result.ok ?? false;
    }

    /**
     * Logs a message and data to the SharePoint Logs list.
     * @param message The log message.
     * @param data Optional log data.
     */
    public async log(message: string, data?: any): Promise<void> {
        const body = {
            formValues: [
                { FieldName: "Title", FieldValue: message },
                { FieldName: 'Data', FieldValue: JSON.stringify(data) },]
        };
        await this.ensureListFolder(this.channelName, ListName.Logs);
        await this.createListItemUsingPath(body, ListName.Logs, '', true);
    }

    /**
     * Hides specified SharePoint lists.
     * @param listNames Array of list names to hide.
     */
    public async hideLists(listNames: ListName[]): Promise<void> {
        for (const listName of listNames) {
            await this.sp.web.lists.getByTitle(listName).update({ Hidden: true });
        }
    }
    // public async test(): Promise<void> {
    //     await this.graphGet(`groups/${SystemAccount.EUPAdminsId}/members`);
    //     await this.graphGet(`groups/${SystemAccount.EUPMembersId}/members`);
    // }
}