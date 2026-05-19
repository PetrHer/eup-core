import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, Language, ListName } from "../enums";

export const version = 6;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureColumns(ListName.WorkflowActions);
    const actions = await apiClient.actions.getWorkflowActions();
    const amlActions = actions.filter(a => a.title[Language.EN].includes('AML'));
    for (let i = 0; i < amlActions.length; i++) {
        const id = amlActions[i].id;
        await apiClient.sp.updateListItem(Number(id), { [ColumnsActions.FileLink]: true }, ListName.WorkflowActions);
    }
    await apiClient.sp.ensureColumns(ListName.Notifications);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
