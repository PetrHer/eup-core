import * as v008 from './v008';
import { IApiClient } from '../api/IApiClient';
import { ColumnsDocTypes, ColumnsReqTypes, Language, ListName } from '../enums';

export const version = 10;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    const requestTypes = await apiClient.requests.getRequestsOptions();
    const requestToUpdate = requestTypes.find(r => r.loc[Language.EN] === 'Request for a credit product');
    if (requestToUpdate) {
        await apiClient.sp.updateListItem(requestToUpdate.id, {
            [ColumnsReqTypes.Stage]: [
                '01 Před schválením',
                '02 Po schválením'
            ]
        }, ListName.RequestTypes);
    }

    const newDoctType = {
        formValues: [
            { FieldName: "Title", FieldValue: 'Obchodní - nabídka (podepsané)' },
            { FieldName: ColumnsDocTypes.TitleEn, FieldValue: 'Business - Offer (signed)' },
            { FieldName: ColumnsDocTypes.UploadStatus, FieldValue: 'provided' },
            { FieldName: ColumnsDocTypes.Permissions, FieldValue: 'UUO;client' },
            { FieldName: ColumnsDocTypes.FolderPath, FieldValue: 'KLIENT - EXT/PRO KLIENTA/Nabídky/Podepsané nabídky' },
        ]
    };
    await apiClient.sp.createListItemUsingPath(newDoctType, ListName.DocumentTypes, '', false);
    return await v008.migrate(apiClient, version);
};