import { IApiClient } from '../api/IApiClient';
import { ColumnsActions, ListName, TemplatesListName } from "../enums";

export const version = 4;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const deleted = await apiClient.sp.deleteList(ListName.WorkflowActions);
    if (!deleted) {
        return false;
    }
    await apiClient.sp.ensureList(ListName.WorkflowActions);
    await apiClient.install.copyFoldersAndItemsWithStructure(TemplatesListName.WorkflowActions, ListName.WorkflowActions, Object.values(ColumnsActions) as string[]);
    await apiClient.install.setListReadOnly(ListName.WorkflowActions);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
