import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, ListName, TemplatesListName } from "../enums";

export const version = 8;
/**
 * migration for deleting and recreating workflow actions list
 */
export const migrate = async (apiClient: IApiClient, version: number = 8): Promise<boolean> => {
    const deleted = await apiClient.sp.deleteList(ListName.WorkflowActions);
    if (!deleted) {
        return false;
    }
    await apiClient.sp.ensureList(ListName.WorkflowActions);
    await apiClient.install.copyFoldersAndItemsWithStructure(TemplatesListName.WorkflowActions, ListName.WorkflowActions, Object.values(ColumnsActions) as string[]);
    await apiClient.install.setListReadOnly(ListName.WorkflowActions);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
