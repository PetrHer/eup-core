import { WebPartContext } from '@microsoft/sp-webpart-base';
import { FileInfo } from '@syncfusion/ej2-inputs';
import { z } from 'zod';

import { INotificationApiClient } from './INotificationApiClient';
import { ColumnsCategories, ListName } from '../../enums';
import { INewNotification, INotification } from '../../interfaces';
import { mapNotifications } from '../mappers';
import SPApiClient from "../SPApiClient";

const NotificationItemSchema = z.object({
    Id: z.number(),
    Title: z.string().nullable().optional(),
    Title_en: z.string().nullable().optional(),
    Task: z.boolean().nullable().optional(),
    Completed: z.boolean().nullable().optional(),
    Created: z.string(),
    Modified: z.string().nullable().optional(),
    Assigned: z.string().nullable().optional(),
    ReadId: z.string().nullable().optional(),
    TargetId: z.string().nullable().optional(),
    Status: z.string().nullable().optional(),
    Comment: z.string().nullable().optional(),
    Link: z.string().nullable().optional(),
    UniqueId: z.string().nullable().optional(),
    FileLink: z.string().nullable().optional(),
    Attachments: z.boolean().nullable().optional(),
    Author: z.object({
        Title: z.string().nullable().optional(),
    }).nullish().transform(author => ({
        Title: author?.Title ?? ''
    })),
    Editor: z.object({
        Title: z.string().nullable().optional(),
    }).nullable().optional(),
    Category: z.object({
        Id: z.number().nullable().optional(),
        Title: z.string().nullable().optional(),
        [ColumnsCategories.CategoryEn]: z.string().nullable().optional(),
        [ColumnsCategories.Stage]: z.string().nullable().optional(),
        [ColumnsCategories.StageEn]: z.string().nullable().optional(),
        [ColumnsCategories.Phase]: z.string().nullable().optional(),
        [ColumnsCategories.PhaseEn]: z.string().nullable().optional(),
    }).nullable().optional(),
});

const parseNotificationItem = (value?: unknown): z.infer<typeof NotificationItemSchema> | undefined => {
    if (!value) {
        return undefined;
    }

    const parsed = NotificationItemSchema.safeParse(value);
    return parsed.success ? parsed.data : undefined;
};

const parseNotificationItems = (value?: unknown): z.infer<typeof NotificationItemSchema>[] => {
    if (!value || !Array.isArray(value)) {
        return [];
    }

    return value
        .map(item => parseNotificationItem(item))
        .filter((item): item is z.infer<typeof NotificationItemSchema> => item !== undefined);
};

export default class NotificationApiClient implements INotificationApiClient {
    constructor(private readonly sp: SPApiClient, private webPartContext: WebPartContext) { }

    /**
         * Retrieves and filters notifications assigned to the specified user groups from the SharePoint Notifications list.
         * @param groups - An array of group identifiers to filter notifications by assignment.
         * @returns A Promise resolving to a sorted array of notifications relevant to the user, mapped with user-specific data.
         */
    public async getAllNotifications(groups: string[]): Promise<INotification[]> {
        const folderPath = `${this.sp.getRelativeUrl()}/lists/${ListName.Notifications}/${this.sp.getChannelName()}`;
        const id = await this.sp.getUserSPId(this.webPartContext.pageContext.user.email);
        const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
        const expand = 'Category,Author';
        const select = `*,Category/Title,Category/${ColumnsCategories.Phase},Category/${ColumnsCategories.Stage},Author/Title,Category/Id,Category/${ColumnsCategories.CategoryEn},Category/${ColumnsCategories.PhaseEn},Category/${ColumnsCategories.StageEn},UniqueId`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Notifications}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
        const result = await this.sp.spGet(apiUrl);

        const parsedItems = parseNotificationItems(result.data);
        const data = parsedItems.filter(item => item.Assigned && item.Assigned.split(';').some((e: string) => groups.includes(e)));
        data.sort((a, b) => {
            return new Date(b.Created).getTime() - new Date(a.Created).getTime();
        });
        const returnData = mapNotifications(data, id);
        return returnData;
    }

    /**
     * Creates a new notification item in the SharePoint Notifications list using provided form values.
     * @param notification - The notification data including title, task, category, assigned groups, and file reference.
     * @returns A Promise resolving to true if the item was successfully created, or false otherwise.
     */
    public async createNotification(notification: INewNotification): Promise<boolean> {

        const body = {
            "formValues": [
                {
                    "FieldName": "Title",
                    "FieldValue": notification.title,
                },
                {
                    "FieldName": "Title_en",
                    "FieldValue": notification.titleEn,
                },
                {
                    "FieldName": 'Task',
                    "FieldValue": `${notification.task}`
                },
                {
                    "FieldName": 'Category',
                    "FieldValue": notification.categoryId.toString()
                },
                {
                    "FieldName": 'Assigned',
                    "FieldValue": notification.userGroups
                },
                {
                    "FieldName": 'TargetId',
                    "FieldValue": notification.targetId
                },
                {
                    "FieldName": 'Link',
                    "FieldValue": notification.link
                },
                {
                    "FieldName": 'Comment',
                    "FieldValue": notification.comment
                },
                {
                    "FieldName": 'FileLink',
                    "FieldValue": notification.fileLink
                },
            ],
        };
        const result = await this.sp.createListItemUsingPath(body, ListName.Notifications);
        return !!result.ok;
    }

    /**
     * Marks a task as completed in the SharePoint Notifications list and optionally attaches a file.
     * @param id - The ID of the notification item to update.
     * @param body - An object containing the updated field values for the task.
     * @param file - (Optional) A file to be attached to the notification item.
     * @returns A Promise resolving to true if the update (and optional file attachment) was successful, or false otherwise.
     */
    public async completeTask(id: number, body: object, files: FileInfo[]): Promise<boolean> {
        const updated = await this.sp.updateListItem(id, body, ListName.Notifications);
        if (!updated || files.length === 0) {
            return updated;
        }
        let result = true;
        for (const file of files) {
            const addedFile = await this.sp.addAttachmentFile(id, file, ListName.Notifications);
            if (!addedFile.ok) {
                result = false;
            }
        }
        return result;
    }

    public async deleteNotifications(notifications: INotification[]): Promise<void> {
        for (let i = 0; i < notifications.length; i++) {
            const notification = notifications[i];
            await this.sp.deleteListItem(notification.id, ListName.Notifications);
        }
    }

    public async deleteNotificationsByTargetId(targetId: string, link?: string): Promise<void> {
        const filter = `TargetId eq '${targetId}' and Task ne 1${link ? ` and Link eq '${link}'` : ''}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.Notifications}')/items?$filter=${filter}`;
        const result = await this.sp.spGet(apiUrl);
        const data = Array.isArray(result.data) ? result.data : [];
        for (let i = 0; i < data.length; i++) {
            const id = data[i].Id;
            await this.sp.deleteListItem(id, ListName.Notifications);
        }
    }

    public async completeTasksByTargetId(targetId: string, link?: string, resultStatus?: string): Promise<void> {
        const filter = `TargetId eq '${targetId}' and Task eq 1${link ? ` and Link eq '${link}'` : ''}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.Notifications}')/items?$filter=${filter}`;
        const result = await this.sp.spGet(apiUrl);
        const data = Array.isArray(result.data) ? result.data : [];
        for (let i = 0; i < data.length; i++) {
            const id = data[i].Id;
            const body = { Completed: true, Status: resultStatus ? resultStatus : null };
            await this.sp.updateListItem(id, body, ListName.Notifications);
        }
    }


    public async getTaskByTargetId(targetId: string): Promise<INotification | undefined> {
        const filter = `TargetId eq '${targetId}' and Task eq 1 and Completed ne 1`;
        const expand = 'Category,Author,Editor';
        const select = `*,UniqueId,Category/Title,Category/${ColumnsCategories.Phase},Category/${ColumnsCategories.Stage},Author/Title,Category/Id,Editor/Title`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.Notifications}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const result = await this.sp.spGet(apiUrl);
        const parsedItems = parseNotificationItems(result.data);
        const mapped = mapNotifications(parsedItems);
        return mapped[0];
    }


    /**
     * Retrieves the history of completed tasks associated with a specific file ID from the SharePoint Notifications list.
     * @param fileId - The ID of the file to filter task history by.
     * @returns A Promise resolving to a chronologically sorted array of notifications with mapped attachments and metadata.
     */
    public async getTaskHistory(fileId: string): Promise<INotification[]> {
        const filter = `TargetId eq '${fileId}' and Task eq 1`;
        const expand = 'Category,Author,Editor';
        const select = `*,UniqueId,Category/Title,Category/${ColumnsCategories.Phase},Category/${ColumnsCategories.Stage},Author/Title,Category/Id,Editor/Title`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.Notifications}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const result = await this.sp.spGet(apiUrl);
        const parsedItems = parseNotificationItems(result.data);
        const sorted = parsedItems.sort((a, b) => {
            return new Date(a.Created).getTime() - new Date(b.Created).getTime();
        });

        const attachmentsMap = await this.getAttachmentsMap(sorted);

        return mapNotifications(sorted, undefined, attachmentsMap);
    }

    /**
 * Builds a map of attachment metadata for each item in the provided list.
 * @param items - An array of SharePoint list items to retrieve attachments for.
 * @returns A Promise resolving to a map where each key is an item ID and the value is an array of attachment objects containing file names and URLs.
 */
    private async getAttachmentsMap(items: any[]): Promise<Map<number, { fileName: string; url: string }[]>> {
        const attachmentsMap = new Map<number, { fileName: string; url: string }[]>();

        await Promise.all(items.map(async item => {
            if (item.Attachments) {
                const attachUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.Notifications}')/items(${item.Id})/AttachmentFiles`;
                const attachResult = await this.sp.spGet(attachUrl);
                const attachments = (Array.isArray(attachResult.data) ? attachResult.data : []).map((att: any) => ({
                    fileName: att.FileName,
                    url: att.ServerRelativeUrl
                }));
                attachmentsMap.set(item.Id, attachments);
            }
        }));

        return attachmentsMap;
    }

}