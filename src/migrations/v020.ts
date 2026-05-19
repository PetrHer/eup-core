import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, FileStatus, Language, ListName } from "../enums";

export const version = 20;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.install.setListReadOnly(ListName.DocumentTypes);
    await apiClient.install.setListReadOnly(ListName.RequestTypes);
    // await apiClient.sp.hideLists(Object.values(ListName));

    const actions = await apiClient.actions.getWorkflowActions();
    const remindActions = actions.filter(a => a.title[Language.EN].toLowerCase() === 'remind' && a.type === 'WF typ 1');
    for (const action of remindActions) {
        const body = {
            [ColumnsActions.DocumentStatus]: `${FileStatus.Requested}\n${FileStatus.NotProvided}`
        };
        await apiClient.sp.updateListItem(Number(action.id), body, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};

