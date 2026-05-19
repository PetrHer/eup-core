import { Language, ColumnsCategories } from "../../enums";
import { INotification } from "../../interfaces";

const mapNotification = (item: any, id?: number, attachmentsMap?: Map<number, { fileName: string; url: string }[]>): INotification => {
    return {
        id: item.Id,
        title: {
            [Language.CS]: item.Title,
            [Language.EN]: item.Title_en
        },
        task: item.Task,
        completed: item.Completed,
        category: {
            [Language.CS]: item.Category ? item.Category.Title : '',
            [Language.EN]: item.Category ? item.Category[ColumnsCategories.CategoryEn] : ''
        },
        stage: {
            [Language.CS]: item.Category ? item.Category[ColumnsCategories.Stage] : '',
            [Language.EN]: item.Category ? item.Category[ColumnsCategories.StageEn] : ''
        },
        phase: {
            [Language.CS]: item.Category ? item.Category[ColumnsCategories.Phase] : '',
            [Language.EN]: item.Category ? item.Category[ColumnsCategories.PhaseEn] : ''
        },
        created: new Date(item.Created).toLocaleDateString(),
        author: item.Author.Title,
        unread: !item.Task && ((item.ReadId && !item.ReadId.includes(id)) || !item.ReadId),
        read: item.ReadId ? item.ReadId : [],
        targetId: item.TargetId,
        categoryId: item.Category && item.Category.Id ? item.Category.Id : 0,
        status: item.Status,
        comment: item.Comment,
        modified: item.Modified,
        editor: item.Editor ? item.Editor.Title : '',
        attachments: attachmentsMap?.get(item.Id) || [],
        link: item.Link,
        uniqueId: item.UniqueId,
        fileLink: item.FileLink
        // assingedTo: item.AssingedTo,
    };
};

/**
     * Maps raw SharePoint notification items to a structured array of INotification objects.
     * @param data - The raw array of SharePoint list items to be transformed.
     * @param id - (Optional) The current user's SharePoint ID, used to determine unread status.
     * @param attachmentsMap - (Optional) A map linking item IDs to their associated attachments.
     * @returns An array of structured INotification objects with localized fields and metadata.
     */
export const mapNotifications = (data: any[], id?: number, attachmentsMap?: Map<number, { fileName: string; url: string }[]>): INotification[] => {
    return data.map(item => mapNotification(item, id, attachmentsMap));
};