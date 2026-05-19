import { FileInfo } from '@syncfusion/ej2-react-inputs';

import { INewNotification, INotification } from '../../interfaces';

export interface INotificationApiClient {
    getAllNotifications(groups: string[]): Promise<INotification[]>;
    createNotification(notification: INewNotification): Promise<boolean>;
    completeTask(id: number, body: object, files: FileInfo[]): Promise<boolean>;
    deleteNotifications(notifications: INotification[]): Promise<void>;
    deleteNotificationsByTargetId(targetId: string, link?: string): Promise<void>;
    completeTasksByTargetId(targetId: string, link?: string, resultStatus?: string): Promise<void>;
    getTaskByTargetId(targetId: string): Promise<INotification | undefined>;
    getTaskHistory(fileId: string): Promise<INotification[]>;
}
