import { IScheduledAction, IWorkflowAction } from '../../interfaces';

export interface IActionApiClient {
    getWorkflowActions(fromTemplates?: boolean): Promise<IWorkflowAction[]>;
    getScheduledActions(): Promise<Map<number, IScheduledAction>>;
    getScheduledActionByFileId(fileId?: number): Promise<IScheduledAction | undefined>;
    createScheduledAction(action: IScheduledAction, title: string, comment?: string): Promise<boolean>;
}
