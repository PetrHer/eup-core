import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, Language, ListName } from "../enums";

export const version = 16;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const actions = await apiClient.actions.getWorkflowActions();

    const acceptActions = actions.filter(a => a.title[Language.EN] === 'Accept' && (a.type === 'WF typ smlouva' || a.type === 'WF typ Hot-TS') && a.onMajorVersion);
    for (const action of acceptActions) {
        await apiClient.sp.updateListItem(Number(action.id), { [ColumnsActions.ActionGroups]: 'client' }, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
