import { IApiClient } from '../api/IApiClient';
import { ListName } from "../enums";

export const version = 1;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureColumns(ListName.Details);
    return await apiClient.install.updateLoanDetailsVersion(version);
};
