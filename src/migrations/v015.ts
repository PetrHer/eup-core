import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, FileStatus, Language, ListName, SPGroupExternal } from "../enums";

export const version = 15;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const actions = await apiClient.actions.getWorkflowActions();

    const acceptActions = actions.filter(a => a.title[Language.EN] === 'Accept' && (a.type === 'WF typ smlouva' || a.type === 'WF typ Hot-TS'));
    for (const action of acceptActions) {
        await apiClient.sp.updateListItem(Number(action.id), { [ColumnsActions.ActionGroups]: action.type === 'WF typ smlouva' ? 'UUO;OAO' : 'UUO' }, ListName.WorkflowActions);
    }

    const createAcceptAction = async (type: string, path: string): Promise<void> => {
        const body = {
            formValues: [
                { FieldName: "Title", FieldValue: "Akceptovat" },
                { FieldName: ColumnsActions.TitleEn, FieldValue: "Accept" },
                { FieldName: ColumnsActions.DocumentStatus, FieldValue: FileStatus.ClientRevision },
                { FieldName: ColumnsActions.ResultStatus, FieldValue: FileStatus.FinalVersion },
                { FieldName: ColumnsActions.DocumentType, FieldValue: type },
                { FieldName: ColumnsActions.Publish, FieldValue: 'true' },
                { FieldName: ColumnsActions.ActionGroups, FieldValue: "client" },
                { FieldName: ColumnsActions.ReadPermissions, FieldValue: 'Relationship manager;RSM Analyst;Back office;UUO Deputy;UUO Director;Risk analyst;GCC analyst;Risk deputy;Risk director;Credit controller;Lawyer;OAO Deputy;OAO Director;WorkOut Specialist;WorkOut Deputy;WorkOut Director;AML Specialist;AML Director;Approver;Interní Auditor;Business Admin;Admin;External Auditor;External Adviser;Client - lower access;Client - full access' },
                { FieldName: ColumnsActions.Validity, FieldValue: 'update' },
                { FieldName: ColumnsActions.Reupload, FieldValue: '' },
                { FieldName: ColumnsActions.FileLink, FieldValue: '' },
                { FieldName: ColumnsActions.OnMajorVersion, FieldValue: 'true' },
            ]
        };

        await apiClient.sp.createListItemUsingPath(body, ListName.WorkflowActions, path, false);
    };

    const clientContractAccept = actions.find(a => a.type === 'WF typ smlouva' && a.title[Language.EN] === 'Accept' && a.actionGroups.includes(SPGroupExternal.EXT_CLIFULL));
    if (!clientContractAccept) {
        await createAcceptAction("Právní - smlouva", '/WF typ smlouva');
    }

    const clientOfferAccept = actions.find(a => a.type === 'WF typ Hot-TS' && a.title[Language.EN] === 'Accept' && a.actionGroups.includes(SPGroupExternal.EXT_CLIFULL));
    if (!clientOfferAccept) {
        await createAcceptAction("Obchodní - nabídka", '/WF typ Hot-TS');
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
