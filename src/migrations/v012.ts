import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, ColumnsDocTypes, Language, ListName } from "../enums";

export const version = 12;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const fileTypes = await apiClient.files.getFileTypes();
    const typeToUpdate = fileTypes.find(t => t.loc[Language.CS] === 'Administrativní - Žádost');
    if (typeToUpdate) {
        await apiClient.sp.updateListItem(typeToUpdate.id, { [ColumnsDocTypes.FolderPathAfter]: null }, ListName.DocumentTypes);
    }
    const workflowActions = await apiClient.actions.getWorkflowActions();
    const actionsToUpdate = workflowActions.filter(a => a.type === 'WF typ 1' && a.title[Language.CS] === 'Odvolat dokument');
    for (const action of actionsToUpdate) {
        const body = {
            [ColumnsActions.Notification]: "Dokument '{filename}' byl odvolán.",
            [ColumnsActions.NotificationEn]: "Document '{filename}' was revoked.",
            [ColumnsActions.NotificationGroups]: 'client'
        };
        await apiClient.sp.updateListItem(Number(action.id), body, ListName.WorkflowActions);
    }
    return await apiClient.install.updateLoanDetailsVersion(version);
};
