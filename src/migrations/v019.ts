import { IApiClient } from '../api/IApiClient';
import { ColumnsQA, ListName } from "../enums";

export const version = 19;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.updateColumn({ FieldTypeKind: '3' }, ListName.QA, ColumnsQA.Answer);

    return await apiClient.install.updateLoanDetailsVersion(version);
};
