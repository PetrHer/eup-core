import { Language, ColumnsRequired, FileStatus, ColumnsDocTypes } from "../../enums";
import { IFile, IFileType, IFileVersion, IFolder, IRequiredFile } from "../../interfaces";
import { expandRoles, getFileTypeTitle } from "../../utils/utils";

/**
 * Maps raw data items to an array of `IRequiredFile` objects.
 *
 * @param data - The raw data array fetched from the API.
 * @returns An array of mapped `IRequiredFile` objects with localization, folder path, description, type, and permitted formats.
 */
export const mapRequiredFiles = (data: any[]): IRequiredFile[] => {
    return data.map(item => {
        return {
            loc: {
                [Language.CS]: item.Title,
                [Language.EN]: item[ColumnsRequired.TitleEn] ? item[ColumnsRequired.TitleEn] : item.Title,
            },
            id: item.Id,
            //folderPath: item[ColumnsRequired.FolderPath],
            description: item[ColumnsRequired.FileDesc],
            type: item[ColumnsRequired.DocumentType]
                ? {
                    loc: {
                        [Language.EN]: getFileTypeTitle(item[ColumnsRequired.DocumentType][ColumnsDocTypes.TitleEn]),
                        [Language.CS]: getFileTypeTitle(item[ColumnsRequired.DocumentType].Title),
                    },
                    id: item.DocumentType.Id,
                    uploadStatus: item[ColumnsRequired.DocumentType][ColumnsDocTypes.UploadStatus],
                    permissions: item[ColumnsRequired.DocumentType][ColumnsDocTypes.Permissions]
                        ? expandRoles(item[ColumnsRequired.DocumentType][ColumnsDocTypes.Permissions])
                        : [],
                    folderPath: item[ColumnsRequired.DocumentType][ColumnsDocTypes.FolderPath],
                    folderPathAfter: item[ColumnsRequired.DocumentType][ColumnsDocTypes.FolderPathAfter],
                    validityStatuses: []
                }
                : undefined,
            format: item[ColumnsRequired.PermittedFormats],
            uploadedFile: item[ColumnsRequired.UploadedFile],
            deadline: item[ColumnsRequired.Deadline] ? new Date(item[ColumnsRequired.Deadline]) : undefined,
            status: Object.values(FileStatus).includes(item[ColumnsRequired.Status]) ? item[ColumnsRequired.Status] : undefined
        };
    });
};

/**
 * Recursively maps raw folder data into a hierarchical array of `IFolder` objects.
 * Filters out items that are children (i.e., folders inside other folders) to avoid duplicates,
 * then constructs folder objects with localized names and nested subfolders.
 *
 * @param data - Raw folder data array retrieved from SharePoint API.
 * @returns An array of `IFolder` objects representing the folder hierarchy.
 */
export const mapFolders = (data: any[]): IFolder[] => {
    return data
        .filter(item => !data.find(it => it.FileRef === item.FileDirRef))
        .map(item => {
            return {
                name: item.FileLeafRef,
                serverRelativeUrl: item.FileRef,
                localization: {
                    [Language.CS]: item.FileLeafRef,
                    [Language.EN]: item.Name_en
                },
                folders: mapFolders(data
                    .filter(it => (it.FileDirRef as string).startsWith(item.FileRef)))
            };
        });
};

/**
 * Maps a raw SharePoint list item to an `IFile` object.
 *
 * Extracts relevant metadata such as name, URL, status, type, version, and creation date.
 * Handles localization of document type titles and sets default status if not present.
 *
 * @param item - A raw SharePoint list item object.
 * @returns A mapped `IFile` object.
 */
export const mapItemAsFile = (item: any): IFile => {
    return {
        name: item.FileLeafRef,
        serverRelativeUrl: item.FileRef,
        status: item.Status && item.Status !== '' ? item.Status as FileStatus : FileStatus.Processing,
        uniqueId: item.UniqueId,
        version: item.OData__UIVersionString,
        type: item.DocumentType
            ? {
                loc: {
                    [Language.EN]: getFileTypeTitle(item.DocumentType.Title_en),
                    [Language.CS]: getFileTypeTitle(item.DocumentType.Title),
                },
                id: item.DocumentType.Id,
                folderPath: item.DocumentType.FolderPath,
                folderPathAfter: item.DocumentType.FolderPathAfter,
                validityStatuses: item.DocumentType[ColumnsDocTypes.ValidityStatuses] ? item.DocumentType[ColumnsDocTypes.ValidityStatuses].split(';') : [],
                folderPathInternal: item.DocumentType[ColumnsDocTypes.FolderPathInternal],
                folderPathInternalAfter: item.DocumentType[ColumnsDocTypes.FolderPathInternalAfter],
            }
            : undefined,
        created: item.Created ? new Date(item.Created) : undefined,
        id: item.Id,
        nodeId: item.StrukturaDokumentaceId,
        documentation: item.Documentation,
        provided: item.UploadedBy,
        contractValidity: item.ContractValidity,
        validTo: item.ValidTo ? new Date(item.ValidTo) : undefined,
        qaId: item.QAId ? item.QAId : '',
        author: item.Author && item.Author.Title ? item.Author.Title : '',
        modified: item.Modified ? new Date(item.Modified) : undefined,
    };
};

export const mapFile = (file: any): IFile => {
    return {
        name: file.FileLeafRef,
        serverRelativeUrl: file.FileRef,
        status: file.Status,
        uniqueId: file.UniqueId,
        version: file.OData__UIVersionString,
        nodeId: file.StrukturaDokumentaceId,
        typeId: file.DocumentTypeId,
    };
};

/**
     * Maps an array of raw SharePoint list items to an array of `IFile` objects.
     * Delegates the mapping of individual items to `mapItemAsFile`.
     * @param items - Array of raw list item objects from SharePoint.
     * @returns An array of mapped `IFile` objects.
     */
export const mapItemsAsFiles = (items: any[]): IFile[] => {
    return items.map(item => mapItemAsFile(item)
    );
};

const mapFileVersion = (item: any, teamName: string): IFileVersion => {
    const relativeUrl = item.FileRef.replace(`/sites/${teamName}`, '');
    return {
        editor: item.Editor ? item.Editor.LookupValue : '',
        modified: item.Modified,
        status: item.Status,
        versionLabel: item.VersionLabel,
        relativeUrl: item.IsCurrentVersion ? item.FileRef : `/sites/${teamName}/_vti_history/${item.VersionId}${relativeUrl}`,
        isCurrentVersion: item.IsCurrentVersion,
        id: item.VersionId
    };
};

/**
* Maps raw version data from SharePoint into an array of IFileVersion objects.
* @param data - The raw version data array from SharePoint.
* @returns An array of formatted file version objects.
*/
export const mapFileVersions = (data: any[], teamName: string): IFileVersion[] => {
    return data.map(item => mapFileVersion(item, teamName));
};

/**
     * Maps raw SharePoint document type items into a structured array of file types with localized titles
     * and hierarchical (parent-child) relationships.
     *
     * - Top-level items do **not** contain `'>'` in the title.
     * - Child items **do** contain `'>'` and are associated with their parent by title prefix.
     *
     * @param data - Raw data array returned from SharePoint list containing document type items.
     * @returns An array of structured `IFileType` objects with localization and nested children.
     */
export const mapFileTypes = (data: any[]): IFileType[] => {
    return data
        .filter(item => !(item.Title as string).includes('>'))
        .map(item => {
            const titleEn = typeof item[ColumnsDocTypes.TitleEn] === 'string' ? item[ColumnsDocTypes.TitleEn] : '';

            return {
                loc: {
                    [Language.CS]: item.Title,
                    [Language.EN]: titleEn,
                },
                // children: data
                //     .filter(type => (type.Title as string).startsWith(item.Title) && (type.Title as string).includes('>'))
                //     .map(type => {
                //         const csTitle = typeof type.Title === 'string' ? type.Title : '';
                //         const enTitle = typeof type[ColumnsDocTypes.TitleEn] === 'string' ? type[ColumnsDocTypes.TitleEn] : '';

                //         const csParts = csTitle.split('>').map((p: string) => p.trim());
                //         const enParts = enTitle.split('>').map((p: string) => p.trim());
                //         return {
                //             loc: {
                //                 [Language.CS]: csParts[1] || '',
                //                 [Language.EN]: enParts[1] || '',
                //             },
                //             id: type.Id
                //         };
                //     }),
                id: item.Id,
                uploadStatus: item[ColumnsDocTypes.UploadStatus],
                permissions: item[ColumnsDocTypes.Permissions] ? expandRoles(item[ColumnsDocTypes.Permissions]) : [],
                folderPath: item[ColumnsDocTypes.FolderPath],
                folderPathAfter: item[ColumnsDocTypes.FolderPathAfter],
                validityStatuses: item[ColumnsDocTypes.ValidityStatuses] ? item[ColumnsDocTypes.ValidityStatuses].split(';') : [],
                folderPathInternal: item[ColumnsDocTypes.FolderPathInternal],
                folderPathInternalAfter: item[ColumnsDocTypes.FolderPathInternalAfter],
            };
        });
};



