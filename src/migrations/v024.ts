import * as v008 from './v008';
import { IApiClient } from '../api/IApiClient';
import { ListName } from '../enums';

export const version = 24;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureColumns(ListName.ActionToTrigger);
    return await v008.migrate(apiClient, version);
};
