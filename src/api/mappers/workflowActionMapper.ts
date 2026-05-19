import { Language, ColumnsActions, FileValidityAction, ColumnsTrigger } from "../../enums";
import { IScheduledAction, IWorkflowAction } from "../../interfaces";
import { expandRoles } from "../../utils/utils";

/**
 * Maps raw SharePoint list data to structured workflow action objects.
 * @param data - An array of list items containing workflow action definitions.
 * @returns An array of IWorkflowAction objects with localized titles, statuses, file types, and configuration details.
 */
export const mapWorkflowActions = (data: any[]): IWorkflowAction[] => {
    return data.map(item => ({
        id: item.Id.toString(),
        title: {
            [Language.CS]: item.Title,
            [Language.EN]: item[ColumnsActions.TitleEn]
        },
        fileTypes: item[ColumnsActions.DocumentType] ? item[ColumnsActions.DocumentType].split('\n') : [],
        statuses: item[ColumnsActions.DocumentStatus] ? item[ColumnsActions.DocumentStatus].split('\n') : [],
        notification: item[ColumnsActions.Notification]
            ? {
                [Language.CS]: item[ColumnsActions.Notification],
                [Language.EN]: item[ColumnsActions.NotificationEn]
            }
            : undefined,
        resultStatus: item[ColumnsActions.ResultStatus],
        publish: item[ColumnsActions.Publish],
        allowComment: item[ColumnsActions.AllowComment],
        actionGroups: item[ColumnsActions.ActionGroups] ? expandRoles(item[ColumnsActions.ActionGroups]) : [],
        notificationGroups: item[ColumnsActions.NotificationGroups] ? expandRoles(item[ColumnsActions.NotificationGroups]) : [],
        task: item[ColumnsActions.Task]
            ? {
                [Language.CS]: item[ColumnsActions.Task],
                [Language.EN]: item[ColumnsActions.TaskEn]
            }
            : undefined,
        taskGroups: item[ColumnsActions.TaskGroups] ? expandRoles(item[ColumnsActions.TaskGroups]) : [],
        phase: item[ColumnsActions.Phase],
        readPermissions: item[ColumnsActions.ReadPermissions],
        writePermissions: item[ColumnsActions.WritePermissions],
        validity: Object.values(FileValidityAction).find(val => val === item[ColumnsActions.Validity]),
        reupload: item[ColumnsActions.Reupload],
        fileLink: item[ColumnsActions.FileLink],
        type: item.FileDirRef.split('/').pop(),
        onMajorVersion: item[ColumnsActions.OnMajorVersion]
    }));
};

export const mapScheduledAction = (item: any): IScheduledAction => {
    return {
        actionId: item[ColumnsTrigger.ActionId] ? Number(item[ColumnsTrigger.ActionId]) : 0,
        fileId: item[ColumnsTrigger.FileId] ? Number(item[ColumnsTrigger.FileId]) : 0,
        processed: item[ColumnsTrigger.Processed] ? new Date(item[ColumnsTrigger.Processed]) : undefined,
        id: item.Id
    };
};

export const mapScheduledActions = (data: any[]): Map<number, IScheduledAction> => {
    const actionMap: Map<number, IScheduledAction> = new Map();
    data.forEach(item => {
        const fileId = item[ColumnsTrigger.FileId] ? Number(item[ColumnsTrigger.FileId]) : 0;
        if (fileId) {
            actionMap.set(fileId, mapScheduledAction(item));
        }
    });
    return actionMap;
};