import { FileValidityAction } from "../enums";
import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IWorkflowAction {
    id: string,
    title: ILocalizedStrings,
    fileTypes: string[],
    statuses: string[],
    notification?: ILocalizedStrings,
    resultStatus?: string,
    publish?: boolean,
    allowComment?: boolean,
    actionGroups: string[],
    notificationGroups: string[],
    task?: ILocalizedStrings,
    taskGroups: string[],
    phase?: string,
    readPermissions?: string,
    writePermissions?: string,
    validity?: FileValidityAction,
    reupload: boolean,
    fileLink?: boolean,
    type: string,
    onMajorVersion: boolean,
}