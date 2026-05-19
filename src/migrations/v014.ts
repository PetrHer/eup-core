import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, FileStatus, Language, ListName, SPGroupExternal, SystemAccount, TeamRole } from "../enums";

export const version = 14;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureList(ListName.WorkflowActions);
    const actions = await apiClient.actions.getWorkflowActions();

    const acceptAction = actions.find(a => a.type === 'WF typ Dotaznik CRS FATCA' && a.title[Language.CS] === 'Akceptovat');
    if (acceptAction) {
        await apiClient.sp.updateListItem(Number(acceptAction.id), { [ColumnsActions.ActionGroups]: 'UUO' }, ListName.WorkflowActions);
    }

    const signAction = actions.find(a => a.type === 'WF typ Dotaznik CRS FATCA' && a.title[Language.EN] === 'Upload signed scan');
    if (signAction) {
        await apiClient.sp.updateListItem(Number(signAction.id),
            { [ColumnsActions.Notification]: "Dokument '{filename}' byl podepsán.", [ColumnsActions.NotificationEn]: "Document '{filename}' was signed." },
            ListName.WorkflowActions);
    }

    const returnAction = actions.find(a => a.type === 'WF typ Dotaznik CRS FATCA' && a.title[Language.EN] === 'Return to client');
    if (returnAction) {
        await apiClient.sp.updateListItem(
            Number(returnAction.id), { [ColumnsActions.Notification]: "Dokument '{filename}' byl vrácen.", [ColumnsActions.NotificationEn]: "Document '{filename}' was returned." },
            ListName.WorkflowActions);
    }

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

    await apiClient.install.updateDateColumns();

    await apiClient.sp.ensureList(ListName.RequiredDocuments);
    await apiClient.users.addUserByEmailToTeamAndChannel(SystemAccount.EUPNoReply, TeamRole.Owner);
    await apiClient.users.addEUPAdminsAsOwners();

    return await apiClient.install.updateLoanDetailsVersion(version);
};
