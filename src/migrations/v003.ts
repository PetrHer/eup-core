import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, ColumnsQA, Language, ListName, QACategory } from "../enums";

export const version = 3;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const body = { "DraftVersionVisibility": 1 };
    await apiClient.sp.updateList(body);

    await apiClient.sp.ensureColumns(ListName.RequiredDocuments);

    const columnBody = {
        "Choices": Object.values(QACategory)
    };
    await apiClient.sp.updateColumn(columnBody, ListName.QA, ColumnsQA.Category);

    await apiClient.sp.updateColumn({ FieldTypeKind: '3' }, "Dokumenty", 'ScheduledNotificationRecipients');
    await apiClient.sp.updateColumn({ FieldTypeKind: '2' }, "Dokumenty", 'ContractValidity');

    const actions = await apiClient.actions.getWorkflowActions();
    const publishActions = actions.filter(a => a.title[Language.EN] === 'Publish' && (a.fileTypes.includes('Obchodní - nabídka') || a.fileTypes.includes('Právní - smlouva')));
    for (let i = 0; i < publishActions.length; i++) {
        const action = publishActions[i];
        const body: any = {};
        if (action.fileTypes.includes('Obchodní - nabídka')) {
            body[ColumnsActions.ReadPermissions] = `Risk analyst;GCC analyst;Risk deputy;Risk director;WorkOut Specialist;WorkOut Deputy;WorkOut Director;AML Specialist;AML Director;Approver;Interní Auditor;Business Admin;Admin;External Auditor;External Adviser;Credit controller;Lawyer;OAO Deputy;OAO Director`;
            body[ColumnsActions.WritePermissions] = `Relationship manager;RSM Analyst;Back office;UUO Deputy;UUO Director;Client - lower access;Client - full access`;
        } else if (action.fileTypes.includes('Právní - smlouva')) {
            body[ColumnsActions.ReadPermissions] = `Risk analyst;GCC analyst;Risk deputy;Risk director;WorkOut Specialist;WorkOut Deputy;WorkOut Director;AML Specialist;AML Director;Approver;Interní Auditor;Business Admin;Admin;External Auditor;External Adviser`;
            body[ColumnsActions.WritePermissions] = `Relationship manager;RSM Analyst;Back office;UUO Deputy;UUO Director;Client - lower access;Client - full access;Credit controller;Lawyer;OAO Deputy;OAO Director`;
        }
        await apiClient.sp.updateListItem(Number(action.id), body, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
