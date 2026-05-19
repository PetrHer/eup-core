import { IApiClient } from '../api/IApiClient';
import { ListName } from "../enums";

export const version = 7;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureColumns(ListName.RequiredDocuments);
    await apiClient.sp.ensureColumns(ListName.Requests);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
