import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, FileStatus, Language, ListName } from "../enums";

export const version = 2;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureColumns(ListName.WorkflowActions);
    const actions = await apiClient.actions.getWorkflowActions();
    const internalActions = actions.filter(a => a.title[Language.EN] === 'Internally prepared');
    for (let i = 0; i < internalActions.length; i++) {
        const id = internalActions[i].id;
        await apiClient.sp.updateListItem(Number(id), { [ColumnsActions.Reupload]: true }, ListName.WorkflowActions);
    }

    const submitActions = actions.filter(a => a.title[Language.EN] === 'Submit revisions to bank');
    for (let i = 0; i < submitActions.length; i++) {
        const id = submitActions[i].id;
        await apiClient.sp.updateListItem(Number(id), { [ColumnsActions.Reupload]: true, [ColumnsActions.TitleEn]: 'Submit revisions back to bank', Title: 'Předat bance zpět revize' }, ListName.WorkflowActions);
    }

    const uploadActions = actions.filter(a => a.title[Language.EN] === 'Upload revisions for bank');
    for (let i = 0; i < uploadActions.length; i++) {
        const id = uploadActions[i].id;
        await apiClient.sp.deleteListItem(Number(id), ListName.WorkflowActions);
    }


    const body = {
        formValues: [
            { FieldName: "Title", FieldValue: "AML akceptovat s výhradou" },
            { FieldName: ColumnsActions.TitleEn, FieldValue: "AML accept with objection" },
            { FieldName: ColumnsActions.DocumentStatus, FieldValue: FileStatus.Verification },
            { FieldName: ColumnsActions.Notification, FieldValue: "Dokument '{filename}' prošel AML verifikací s výhradou." },
            { FieldName: ColumnsActions.NotificationEn, FieldValue: "Document '{filename}' passed AML verification with objection." },
            { FieldName: ColumnsActions.ResultStatus, FieldValue: FileStatus.AcceptedConditional },
            { FieldName: ColumnsActions.DocumentType, FieldValue: "Administrativní - AML a KYC" },
            { FieldName: ColumnsActions.Publish, FieldValue: '' },
            { FieldName: ColumnsActions.AllowComment, FieldValue: 'true' },
            { FieldName: ColumnsActions.NotificationGroups, FieldValue: "AML" },
            { FieldName: ColumnsActions.Phase, FieldValue: '' },
            { FieldName: ColumnsActions.Task, FieldValue: '' },
            { FieldName: ColumnsActions.TaskEn, FieldValue: '' },
            { FieldName: ColumnsActions.TaskGroups, FieldValue: '' },
            { FieldName: ColumnsActions.ReadPermissions, FieldValue: '' },
            { FieldName: ColumnsActions.WritePermissions, FieldValue: '' },
            { FieldName: ColumnsActions.Validity, FieldValue: '' },
            { FieldName: ColumnsActions.Reupload, FieldValue: '' },
        ]
    };

    await apiClient.sp.createListItemUsingPath(body, ListName.WorkflowActions, '/WF typ AML', false);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
