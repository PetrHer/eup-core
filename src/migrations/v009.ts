import * as v008 from './v008';
import { IApiClient } from '../api/IApiClient';

export const version = 9;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    return await v008.migrate(apiClient, version);
};
