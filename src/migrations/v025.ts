import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, ListName, SPGroupRO } from '../enums';

export const version = 25;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.users.createSiteGroups([SPGroupRO.RORole], () => { /* no-op */ });
    const result = await apiClient.users.addEUPMembersAsMembers();
    if (!result) { return false; }

    const actions = await apiClient.actions.getWorkflowActions();
    for (const action of actions) {
        if ((!action.readPermissions && !action.writePermissions) || (action.readPermissions && action.readPermissions.includes(SPGroupRO.RORole))) { continue; }
        const updatedReadPermissions = action.readPermissions
            ? `${action.readPermissions};${SPGroupRO.RORole}`
            : SPGroupRO.RORole;
        await apiClient.sp.updateListItem(Number(action.id), { [ColumnsActions.ReadPermissions]: updatedReadPermissions }, ListName.WorkflowActions);
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
