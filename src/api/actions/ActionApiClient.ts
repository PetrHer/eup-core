import { IActionApiClient } from "./IActionApiClient";
import { ColumnsTrigger, ListName } from "../../enums";
import { IScheduledAction, IWorkflowAction } from "../../interfaces";
import { mapScheduledAction, mapScheduledActions, mapWorkflowActions } from "../mappers";
import SPApiClient from "../SPApiClient";

export default class ActionApiClient implements IActionApiClient {
    /**
     * Constructs an ActionApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    /**
     * Retrieves workflow actions from the SharePoint WorkflowActions list.
     * Filters items by content type and maps them to structured workflow action objects.
     * @returns A Promise resolving to an array of IWorkflowAction objects, or an empty array if no data is found.
     */
    public async getWorkflowActions(fromTemplates?: boolean): Promise<IWorkflowAction[]> {
        const absoluteUrl = fromTemplates
            ? this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), `/sites/${this.sp.templateSiteName}`)
            : this.sp.getAbsoluteUrl();
        const filter = `startswith(ContentTypeId,'0x0100')`;
        const apiUrl = `${absoluteUrl}/_api/web/lists/getbytitle('${ListName.WorkflowActions}')/items?$filter=${filter}&$select=*,FileDirRef`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) ? mapWorkflowActions(result.data) : [];
    }

    /**
     * Retrieves scheduled actions from the SharePoint ActionToTrigger list.
     * Filters items by folder and processed status, then maps them to structured scheduled action objects.
     * @returns A Promise resolving to a Map of scheduled actions, or an empty Map if no data is found.
     */
    public async getScheduledActions(): Promise<Map<number, IScheduledAction>> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.ActionToTrigger}/${this.sp.getChannelName()}`;
        const filter = `startswith(FileDirRef,'${fileDirRef}') and ${ColumnsTrigger.Processed} eq null`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.ActionToTrigger}')/items?$filter=${filter}&$select=*`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) ? mapScheduledActions(result.data) : new Map();
    }

    public async getScheduledActionByFileId(fileId?: number): Promise<undefined | IScheduledAction> {
        if (!fileId) {
            return undefined;
        }
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.ActionToTrigger}/${this.sp.getChannelName()}`;
        const filter = `startswith(FileDirRef,'${fileDirRef}') and ${ColumnsTrigger.Processed} eq null and ${ColumnsTrigger.FileId} eq ${fileId}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.ActionToTrigger}')/items?$filter=${filter}&$select=*`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) && result.data[0] ? mapScheduledAction(result.data[0]) : undefined;
    }

    /**
     * Creates a new scheduled action in the ActionToTrigger list.
     * @param action The scheduled action object to create.
     * @param title The title for the scheduled action.
     * @returns Promise resolving to true if creation succeeded, otherwise false.
     */
    public async createScheduledAction(action: IScheduledAction, title: string, comment?: string): Promise<boolean> {
        const body = {
            "formValues": [
                {
                    "FieldName": "Title",
                    "FieldValue": title,
                },
                {
                    "FieldName": ColumnsTrigger.FileId,
                    "FieldValue": action.fileId.toString(),
                },
                {
                    "FieldName": ColumnsTrigger.ActionId,
                    "FieldValue": action.actionId.toString()
                },
                {
                    "FieldName": ColumnsTrigger.Comment,
                    "FieldValue": comment ?? ''
                },
            ],
        };
        const result = await this.sp.createListItemUsingPath(body, ListName.ActionToTrigger);
        return !!result.ok;
    }
}