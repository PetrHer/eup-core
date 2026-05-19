import { IApiClient } from '../api/IApiClient';
import { Language, ColumnsActions, ListName, ColumnsReqTypes, TeamRole, SPGroupInternal } from "../enums";

export const version = 22;
/**
 * migration for deleting and recreating workflow actions list
 */
export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const actionsFromTemplates = await apiClient.actions.getWorkflowActions(true);
    const actions = await apiClient.actions.getWorkflowActions();
    for (const action of actions) {
        const template = actionsFromTemplates.find(a =>
            a.type === action.type
            && a.actionGroups.join(';') === action.actionGroups.join(';')
            && a.title[Language.EN].toLowerCase().trim() === action.title[Language.EN].toLowerCase().trim()
            && a.fileTypes.join(';') === action.fileTypes.join(';')
        );
        if (template
            && (action.notification?.[Language.EN] !== template.notification?.[Language.EN]
                || action.task?.[Language.EN] !== template.task?.[Language.EN])) {
            const body = {
                [ColumnsActions.TaskEn]: template.task?.[Language.EN],
                [ColumnsActions.NotificationEn]: template.notification?.[Language.EN]
            };
            await apiClient.sp.updateListItem(Number(action.id), body, ListName.WorkflowActions);
        }
    }

    const hotAction = actions.find(a => a.type === 'WF typ Hot-TS' && a.title[Language.EN].toLowerCase().trim() === 'accept' && a.actionGroups.includes(SPGroupInternal.UUO_DEP));
    if (hotAction) {
        const body = {
            [ColumnsActions.ActionGroups]: 'UUO',
            [ColumnsActions.Notification]: "Dokument '{filename}' byl akceptován.",
            [ColumnsActions.NotificationEn]: "The document '{filename}' was accepted.",
            [ColumnsActions.NotificationGroups]: 'UUO'
        };
        await apiClient.sp.updateListItem(Number(hotAction.id), body, ListName.WorkflowActions);
    }

    const contractAction = actions.find(a => a.type === 'WF typ smlouva' && a.title[Language.EN].toLowerCase().trim() === 'accept' && a.actionGroups.includes(SPGroupInternal.UUO_DEP));
    if (contractAction) {
        const body = {
            [ColumnsActions.ActionGroups]: 'UUO;OAO',
            [ColumnsActions.Notification]: "Dokument '{filename}' byl akceptován.",
            [ColumnsActions.NotificationEn]: "The document '{filename}' was accepted.",
            [ColumnsActions.NotificationGroups]: 'UUO;OAO'
        };
        await apiClient.sp.updateListItem(Number(contractAction.id), body, ListName.WorkflowActions);
    }

    const hotAction2 = actions.find(a => a.type === 'WF typ Hot-TS' && a.title[Language.EN].toLowerCase().trim() === 'accept' && !a.actionGroups.includes(SPGroupInternal.UUO_DEP));
    if (hotAction2) {
        const body = {
            [ColumnsActions.Notification]: "Dokument '{filename}' byl akceptován.",
            [ColumnsActions.NotificationEn]: "The document '{filename}' was accepted.",
            [ColumnsActions.NotificationGroups]: 'UUO'
        };
        await apiClient.sp.updateListItem(Number(hotAction2.id), body, ListName.WorkflowActions);
    }

    const contractAction2 = actions.find(a => a.type === 'WF typ smlouva' && a.title[Language.EN].toLowerCase().trim() === 'accept' && !a.actionGroups.includes(SPGroupInternal.UUO_DEP));
    if (contractAction2) {
        const body = {
            [ColumnsActions.Notification]: "Dokument '{filename}' byl akceptován.",
            [ColumnsActions.NotificationEn]: "The document '{filename}' was accepted.",
            [ColumnsActions.NotificationGroups]: 'UUO;OAO'
        };
        await apiClient.sp.updateListItem(Number(contractAction2.id), body, ListName.WorkflowActions);
    }

    const requests = await apiClient.requests.getRequestsOptions();
    const req1 = requests.find(r => r.loc[Language.EN].toLowerCase().trim() === 'request for a credit product');
    if (req1) {
        await apiClient.sp.updateListItem(req1.id, { [ColumnsReqTypes.TitleEn]: 'Request for credit product' }, ListName.RequestTypes);
    }

    const req2 = requests.find(r => r.loc[Language.EN].toLowerCase().trim() === 'requests for credit line reduction');
    if (req2) {
        await apiClient.sp.updateListItem(req2.id, { [ColumnsReqTypes.TitleEn]: 'Request for credit line reduction', Title: 'Žádost o snížení úvěrového rámce' }, ListName.RequestTypes);
    }

    await apiClient.users.addUserByEmailToTeamAndChannel(apiClient.users.eupNoReplyMail, TeamRole.Owner);

    return await apiClient.install.updateLoanDetailsVersion(version);
};
