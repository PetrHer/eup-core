import { IApiClient } from '../api/IApiClient';
import { ListName, TemplatesListName } from "../enums";

export const version = 13;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const templates = await apiClient.sp.getTemplatesAttachments(ListName.RequiredDocuments);
    if (templates.length === 0) {
        await apiClient.install.copyTemplateAttachments(TemplatesListName.RequiredDocuments, ListName.RequiredDocuments);
    }
    return await apiClient.install.updateLoanDetailsVersion(version);
};
