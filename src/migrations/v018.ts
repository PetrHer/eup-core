import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, Language, ListName } from "../enums";

export const version = 18;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const actions = await apiClient.actions.getWorkflowActions();

    const returnActions = actions.filter(a => a.title[Language.EN].toLowerCase() === 'return to client' && a.type === 'WF typ 1');
    for (const action of returnActions) {
        const body = {
            [ColumnsActions.ReadPermissions]: 'External Auditor;External Adviser;Risk analyst;GCC analyst;Risk deputy;Risk director;WorkOut Specialist;WorkOut Deputy;WorkOut Director;AML Specialist;AML Director;Approver;Interní Auditor;Business Admin;Admin',
            [ColumnsActions.WritePermissions]: 'Relationship manager;RSM Analyst;Back office;UUO Deputy;UUO Director;Credit controller;Lawyer;OAO Deputy;OAO Director;Client - lower access;Client - full access;'
        };
        await apiClient.sp.updateListItem(Number(action.id), body, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
