import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, Language, ListName } from "../enums";

export const version = 23;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const actions = await apiClient.actions.getWorkflowActions();
    const acceptAction = actions.find(a => a.title[Language.EN].toLowerCase().trim() === 'accept' && a.type === 'WF typ Dotaznik CRS FATCA');
    if (acceptAction) {
        const body = {
            [ColumnsActions.NotificationGroups]: 'UUO'
        };
        await apiClient.sp.updateListItem(Number(acceptAction.id), body, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};