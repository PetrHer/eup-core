import { IApiClient } from '../api/IApiClient';

export const version = 5;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.install.ensureFolderForAllChannels('', 'Temporary');
    return await apiClient.install.updateLoanDetailsVersion(version);
};
