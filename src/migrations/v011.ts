import { IApiClient } from '../api/IApiClient';
import { ColumnsReqTypes, ListName, SPGroupInternal } from "../enums";

export const version = 11;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const requestTypes = await apiClient.requests.getRequestsOptions();
    const notifGroups = [SPGroupInternal.UUO_BO,
    SPGroupInternal.UUO_DEP,
    SPGroupInternal.UUO_DIR,
    SPGroupInternal.UUO_RM,
    SPGroupInternal.UUO_RSMA,
    SPGroupInternal.OAO_CC,
    SPGroupInternal.OAO_DEP,
    SPGroupInternal.OAO_DIR,
    SPGroupInternal.OAO_LAW].join(';');
    for (const reqType of requestTypes) {
        await apiClient.sp.updateListItem(reqType.id, { [ColumnsReqTypes.Notification]: notifGroups }, ListName.RequestTypes);

    }
    return await apiClient.install.updateLoanDetailsVersion(version);
};
