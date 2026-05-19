import { IApiClient } from '../api/IApiClient';
import { ColumnsDocTypes, Language, ListName, TeamName } from "../enums";

export const version = 17;

export const migrate = async (apiClient: IApiClient): Promise<boolean> => {
    await apiClient.sp.ensureList(ListName.Logs);
    await apiClient.install.ensureListFolder(apiClient.sp.getChannelName(), ListName.Logs);

    await apiClient.sp.ensureColumns(ListName.DocumentTypes);

    const teamName = TeamName.JTFG;
    const webUrl = `${apiClient.sp.getAbsoluteUrl().replace(apiClient.sp.getRelativeUrl(), '')}/sites/${teamName}`;
    const templateDocTypes = await apiClient.files.getFileTypes(webUrl);
    const docTypes = await apiClient.files.getFileTypes();
    for (const docType of docTypes) {
        const template = templateDocTypes.find(t => t.loc[Language.CS] === docType.loc[Language.CS]);
        if (template) {
            await apiClient.sp.updateListItem(docType.id, {
                [ColumnsDocTypes.FolderPathInternal]: template.folderPathInternal,
                [ColumnsDocTypes.FolderPathInternalAfter]: template.folderPathInternalAfter,
            },
                ListName.DocumentTypes);
        }
    }

    return await apiClient.install.updateLoanDetailsVersion(version);
};
