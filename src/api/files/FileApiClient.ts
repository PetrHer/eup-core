import { FileInfo } from '@syncfusion/ej2-react-inputs';
import { z } from "zod";
import "@pnp/sp/files";
import "@pnp/sp/webs";

import { IFileApiClient } from './IFileApiClient';
import { ColumnsDocTypes, FileStatus, FinancialDoc, Language, LibraryName, ListName } from '../../enums';
import { IFile, IFileInput, IFileType, IFileUploadResult, IFileVersion, IFolder, IRequiredFile, ISiteGroup, IStructureNode, IWorkflowAction } from '../../interfaces';
import { expandRoles, getCleanBaseName, getFileExtension, getFileTypeTitle, transformPath } from '../../utils/utils';
import IApiResult from '../IApiResult';
import { mapFile, mapFileTypes, mapFileVersions, mapFolders, mapItemAsFile, mapItemsAsFiles, mapRequiredFiles } from '../mappers';
import SPApiClient from "../SPApiClient";

const FileItemSchema = z.object({
    FileLeafRef: z.string(),
    FileRef: z.string(),
    Status: z.nativeEnum(FileStatus).nullable().optional(),
    UniqueId: z.string(),
    OData__UIVersionString: z.string(),
    DocumentType: z.object({
        Title: z.string(),
        [ColumnsDocTypes.TitleEn]: z.string(),
        Id: z.number(),
        [ColumnsDocTypes.ValidityStatuses]: z.string().nullable().optional(),
        [ColumnsDocTypes.FolderPathInternal]: z.string().nullable().optional(),
        [ColumnsDocTypes.FolderPathInternalAfter]: z.string().nullable().optional(),
        [ColumnsDocTypes.FolderPath]: z.string().nullable().optional(),
        [ColumnsDocTypes.FolderPathAfter]: z.string().nullable().optional(),
    }).nullable().optional(),
    Id: z.number(),
    StrukturaDokumentaceId: z.number().nullable().optional(),
    Documentation: z.boolean().nullable().optional(),
    UploadedBy: z.string().nullable().optional(),
    ContractValidity: z.string().nullable().optional(),
    ValidTo: z.string().nullable().optional(),
    QAId: z.string().nullable().optional(),
    Created: z.string().nullable().optional(),
    Modified: z.string().nullable().optional(),
    Author: z.object({
        Title: z.string().nullable().optional(),
    }).nullable().optional(),
});

const FileItemsSchema = z.array(FileItemSchema);

const parseFileItem = (value?: unknown): z.infer<typeof FileItemSchema> | undefined => {
    if (!value) {
        return undefined;
    }
    const parsedFileItem = FileItemSchema.safeParse(value);
    return parsedFileItem.success ? parsedFileItem.data : undefined;
};

const parseFileItems = (value?: unknown): z.infer<typeof FileItemsSchema> => {
    if (!value || !Array.isArray(value)) {
        return [];
    }
    const items = value.map(item => parseFileItem(item)).filter((item): item is z.infer<typeof FileItemSchema> => item !== undefined);
    return items;
};

const parseAndMapFileItems = (value?: unknown): IFile[] => {
    const data = parseFileItems(value);
    return mapItemsAsFiles(data);
};

export default class FileApiClient implements IFileApiClient {
    protected readonly documentFolder: string = 'Sdilene dokumenty';

    private readonly trashPath: string = 'KLIENT - EXT/KOŠ';

    /**
     * Constructs a FileApiClient instance.
     * @param spClient The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly spClient: SPApiClient) { }


    /**
     * Moves a file to the bin (trash) folder.
     * @param fileRef The file reference (server-relative URL).
     * @param fileName The file name.
     * @param uniqueId Optional unique ID for the file.
     * @returns Promise resolving to true if moved successfully.
     */
    public async moveFileToBin(fileRef: string, fileName: string): Promise<boolean> {
        const newFileDirRef = this.trashPath;
        return await this.moveFile(fileRef, fileName, newFileDirRef);
    }

    /**
    * Loads files from a SharePoint document library filtered by a structure ID,
    * excluding deleted files and those in the trash folder
    * @param id - The structure ID used to filter files
    * @returns A promise resolving to an array of IFile objects matching the filter criteria
    */
    public async loadFiles(id: number): Promise<IFile[]> {
        const select = `*,FileRef,FileLeafRef,UniqueId,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${ColumnsDocTypes.ValidityStatuses},DocumentType/${ColumnsDocTypes.FolderPathInternal},DocumentType/${ColumnsDocTypes.FolderPath},DocumentType/${ColumnsDocTypes.FolderPathAfter},DocumentType/${ColumnsDocTypes.FolderPathInternalAfter}`;
        const filter = `StrukturaDokumentaceId eq ${id} and Status ne '${FileStatus.Deleted}'`;
        const expand = 'DocumentType,Author';
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/lists/getbytitle('${LibraryName.Documents}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const result = await this.spClient.spGet(apiUrl);
        return parseAndMapFileItems(result.data);
    }

    /**
     * Loads files from SharePoint by their item IDs.
     * @param ids Array of item IDs.
     * @returns Promise resolving to an array of IFile objects.
     */
    public async loadFilesByIds(ids: number[]): Promise<IFile[]> {
        if (ids.length === 0) {
            return [];
        }
        const select = `*,FileRef,FileLeafRef,UniqueId,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${ColumnsDocTypes.ValidityStatuses},DocumentType/${ColumnsDocTypes.FolderPathInternal},DocumentType/${ColumnsDocTypes.FolderPath},DocumentType/${ColumnsDocTypes.FolderPathAfter},DocumentType/${ColumnsDocTypes.FolderPathInternalAfter}`;
        const filter = `${ids.map(id => `Id eq ${id}`).join(' or ')} and Status ne '${FileStatus.Deleted}'`;
        const expand = 'DocumentType,Author';
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/lists/getbytitle('${LibraryName.Documents}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const result = await this.spClient.spGet(apiUrl);
        return parseAndMapFileItems(result.data);
    }

    /**
     * Loads files from SharePoint by their unique IDs.
     * @param uniqueIds Array of unique IDs.
     * @returns Promise resolving to an array of IFile objects.
     */
    public async loadFilesByUniqueIds(uniqueIds: string[]): Promise<IFile[]> {
        const result = await Promise.all(
            uniqueIds.map(id =>
                this.loadFileByUniqueId(id)
            )
        );

        return result.filter((file): file is IFile => file !== undefined);
    }

    /**
     * Loads a file from SharePoint by its server-relative URL.
     * @param serverRelativeUrl The server-relative URL of the file.
     * @returns Promise resolving to an IFile object or undefined.
     */
    public async loadFileByRelativeUrl(serverRelativeUrl: string): Promise<IFile | undefined> {
        const baseUrl = this.spClient.getAbsoluteUrl();
        const result = await this.spClient.spGet(
            `${baseUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/ListItemAllFields?$select=*,StrukturaDokumentaceId,FileRef,FileLeafRef,UIVersionLabel,UniqueId`
        );

        return result.data ? mapFile(result.data) : undefined;
    }

    /**
     * Loads a file from SharePoint by its unique ID (private).
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to an IFile object or undefined.
     */
    private async loadFileByUniqueId(uniqueId: string): Promise<IFile | undefined> {
        const baseUrl = this.spClient.getAbsoluteUrl();
        const select = `*,StrukturaDokumentaceId,FileRef,FileLeafRef,UIVersionLabel,UniqueId`;
        const result = await this.spClient.spGet(`${baseUrl}/_api/web/getfilebyid('${uniqueId}')/ListItemAllFields?$select=${select}`);
        return result.data ? mapFile(result.data) : undefined;
    }

    /**
    * Loads all files from a channel folder path within the SharePoint document library,
    * optionally filtering by documentation flag, excluding deleted files and those in the trash folder
    * @param documentation - If true, filters files marked as documentation
    * @returns A promise resolving to an array of IFile objects matching the criteria
    */
    public async loadAllFiles(documentation: boolean): Promise<IFile[]> {
        const folderPath = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}`;
        // const deletedPath = `${this.getRelativeUrl()}/${this.trashPath}`;
        const filter = `startswith(ContentTypeId,'0x0101') and startswith(FileDirRef,'${folderPath}') and Status ne '${FileStatus.Deleted}' ${documentation ? `and Documentation eq 1` : ''}`; //0x101
        const select = `*,UniqueId,FileRef,FileLeafRef,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${ColumnsDocTypes.ValidityStatuses},DocumentType/${ColumnsDocTypes.FolderPathInternal},DocumentType/${ColumnsDocTypes.FolderPathAfter},DocumentType/${ColumnsDocTypes.FolderPath},DocumentType/${ColumnsDocTypes.FolderPathInternalAfter}`;
        const expand = 'DocumentType,Author';
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/GetListByTitle('${LibraryName.Documents}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
        const result = await this.spClient.spGet(apiUrl);
        return parseAndMapFileItems(result.data);
    }

    /**
     * publishs file by using its uniqueId
     * @param uniqueId 
     * @returns boolean
     */
    public async publishFile(uniqueId: string): Promise<boolean> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/GetItemByUniqueId('${uniqueId}')/File/publish('')`;
        const result = await this.spClient.spPost(apiUrl);
        return result.ok ?? false;
    }

    private async assignFilePermissions(uniqueId: string, readGroupsIds: number[], writeGroupIds: number[]): Promise<void> {
        const boundary = "batch_" + Date.now();
        const webUrl = this.spClient.getAbsoluteUrl();
        const apiUrlBase = `${webUrl}/_api/web/GetFileById('${uniqueId}')/ListItemAllFields/roleassignments/addroleassignment`;

        const readPart = readGroupsIds.map(id =>
            this.spClient.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
        ).join('');

        const writePart = writeGroupIds.map(id =>
            this.spClient.getPermissionBatchString(boundary, id, 1073741827, apiUrlBase)
        ).join('');

        const batchBody = `${readPart}${writePart}--${boundary}--\r\n`;
        await this.spClient.spPost(
            `${webUrl}/_api/$batch`,
            batchBody,
            {
                "Content-Type": `multipart/mixed; boundary=${boundary}`,
            }
        );
    }

    public async setFilePermissions(action: IWorkflowAction, siteGroups: ISiteGroup[], uniqueId: string): Promise<void> {
        const writeGroups = action.writePermissions ? expandRoles(action.writePermissions) : [];
        const readGroupsAll = action.readPermissions ? expandRoles(action.readPermissions) : [];
        const writeGroupIds = siteGroups.filter(g => writeGroups.includes(g.title)).map(g => g.id);
        await this.resetAndBreakInheritance(uniqueId);
        const readGroups = [...readGroupsAll].filter(g => !writeGroups.includes(g));
        const readGroupIds = siteGroups.filter(g => readGroups.includes(g.title)).map(g => g.id);
        await this.assignFilePermissions(uniqueId, readGroupIds, writeGroupIds);
    }

    /**
     * upload array of files using serverRelativeUrl of folder
     * @param folderPath 
     * @param files 
     * @returns array of fileNames whose upload failed
     */
    public async uploadFiles(files: IFileInput[], secondStage: boolean, uploadedBy?: string, folderId?: number): Promise<IFileUploadResult[]> {
        const result: IFileUploadResult[] = await Promise.all(
            files.map(async (file) => {
                const folderPath = !file.type
                    ? '/Temporary'
                    : secondStage && file.type.folderPathAfter ? file.type.folderPathAfter : file.type.folderPath;
                return await this.uploadFile(folderPath, file.file, uploadedBy, folderId, file.type?.id, file.type?.uploadStatus);
            })
        );
        return result;
    }

    /**
    * Retrieves the list of available file types from the SharePoint 'DocumentTypes' list
    * @returns A promise resolving to an array of IFileType objects
    */
    public async getFileTypes(webUrl: string = this.spClient.getAbsoluteUrl()): Promise<IFileType[]> {
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${ListName.DocumentTypes}')/items`;
        const result = await this.spClient.spGet(apiUrl);
        return Array.isArray(result.data) ? mapFileTypes(result.data) : [];
    }

    /**
     * Undoes the check-out of a file by unique ID.
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to true if successful.
     */
    public async undoFileCheckOut(uniqueId: string): Promise<boolean> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyid(guid'${uniqueId}')/UndoCheckOut()`;
        const result = await this.spClient.spPost(apiUrl);
        return !!result.ok;
    }

    /**
    * Retrieves a single file by its unique ID, including its document type details if available
    * @param uniqueId - The GUID of the file to retrieve
    * @returns A promise resolving to the file object or undefined if not found
    */
    public async getFileByUniqueId(uniqueId: string): Promise<IFile | undefined> {
        const select = '*,FileRef,FileLeafRef,UniqueId';
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyid(guid'${uniqueId}')/ListItemAllFields?$select=${select}`;
        const result = await this.spClient.spGet(apiUrl);
        const item = result.data as any;

        if (!item) return undefined;

        const documentTypeId = item.DocumentTypeId;

        const file = mapItemAsFile(item);
        if (documentTypeId) {
            const lookupApiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.DocumentTypes}')/items(${documentTypeId})?$select=Id,Title,Title_en`;
            const lookupResult = await this.spClient.spGet(lookupApiUrl);
            const documentType = lookupResult.data as any;
            file.type = documentType ? {
                loc: {
                    [Language.EN]: getFileTypeTitle(documentType.Title_en),
                    [Language.CS]: getFileTypeTitle(documentType.Title),
                },
                id: documentType.Id,
                folderPath: documentType[ColumnsDocTypes.FolderPath],
                folderPathAfter: documentType[ColumnsDocTypes.FolderPathAfter],
                validityStatuses: documentType[ColumnsDocTypes.ValidityStatuses] ? documentType[ColumnsDocTypes.ValidityStatuses].split(';') : []
            } : undefined;
        }

        return file;
    }

    /**
    * Retrieves the version history for a file based on its list item ID.
    * @param id - The list item ID of the file.
    * @returns A promise resolving to an array of file version objects.
    */
    public async getFileVersionHistory(id: number): Promise<IFileVersion[]> {
        const expand = `Version`;
        const select = `*,Version/Id`;
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/items(${id})/versions?$expand=${expand}&$select=${select}`;
        const result = await this.spClient.spGet(apiUrl);
        const teamName = this.spClient.getAbsoluteUrl().split('/').pop();
        return Array.isArray(result.data) ? mapFileVersions(result.data, teamName ?? '') : [];
    }

    /**
     * Retrieves a specific version of a file as an ArrayBuffer.
     * @param relativeUrl The server-relative URL of the file.
     * @param versionId The version ID.
     * @returns Promise resolving to the file content as ArrayBuffer.
     */
    public async getFileVersion(relativeUrl: string, versionId: number): Promise<ArrayBuffer | undefined> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${relativeUrl}')/versions(${versionId})/$value`;
        const result = await this.spClient.spGetBinary(apiUrl);
        return result.ok && result.data instanceof ArrayBuffer ? result.data : undefined;
    }

    /**
     * Retrieves a file for download as an ArrayBuffer.
     * @param relativeUrl The server-relative URL of the file.
     * @returns Promise resolving to the file content as ArrayBuffer.
     */
    public async getFileToDownload(relativeUrl: string): Promise<ArrayBuffer | undefined> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${relativeUrl}')/$value`;
        const result = await this.spClient.spGetBinary(apiUrl);
        return result.ok && result.data instanceof ArrayBuffer ? result.data : undefined;
    }

    /**
     * Archives a file to a designated location, updating metadata.
     * @param file The file object to archive.
     * @param stage Optional stage node.
     * @param phase Optional phase node.
     * @returns Promise resolving to true if archived successfully.
     */
    public async archiveFile(file: IFile, stage?: IStructureNode, phase?: IStructureNode): Promise<boolean> {


        const sourceSiteUrl = new URL(this.spClient.getAbsoluteUrl()).origin;
        const endpoint = `${sourceSiteUrl}/_api/SP.MoveCopyUtil.CopyFileByPath`;

        const destinationUrl = `${this.spClient.getAbsoluteUrl().replace(this.spClient.getRelativeUrl(), '')}/sites/${this.spClient.internalSiteName}`;

        const archivedFileRef = await this.getArchivedFile(file.uniqueId);

        if (archivedFileRef) {
            const fileBuffer = await this.spClient.spGetBinary(
                `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyid(guid'${file.uniqueId}')/$value`,
                { responseType: "arraybuffer" }
            );
            if (!(fileBuffer.ok && fileBuffer.data instanceof ArrayBuffer)) {
                return false;
            }
            await this.spClient.spPut(
                `${destinationUrl}/_api/web/getfilebyserverrelativeurl('${archivedFileRef}')/$value`,
                fileBuffer.data
            );
            const updateUrl = `${destinationUrl}/_api/web/getfilebyserverrelativeurl('${archivedFileRef}')/ListItemAllFields`;
            await this.spClient.spPatch(updateUrl, { externalFileId: file.uniqueId });
            return true;
        }
        const teamName = await this.spClient.getTeamName();

        const updatedTeamName = teamName.replace('_ext', '');

        const folderPath = !stage?.name.startsWith('02')
            ? file.type?.folderPathInternal
            : file.type?.loc[Language.EN].startsWith('Legal - Fulfillment of Contractual') && (phase?.name?.startsWith('01') || phase?.name?.startsWith('02'))
                ? file.type?.folderPathInternal
                : file.type?.folderPathInternalAfter;

        const basePath = `/sites/${this.spClient.internalSiteName}/${this.documentFolder}/${updatedTeamName}/${this.spClient.getChannelName()}/${folderPath ?? file.type?.folderPathInternal}`;

        const uniqueName = await this.getUniqueFileName(basePath, file.name, 0, destinationUrl);

        const body = {
            srcPath: {
                DecodedUrl: `${sourceSiteUrl}${file.serverRelativeUrl}`
            },
            destPath: {
                DecodedUrl: `${destinationUrl}/${this.documentFolder}/${updatedTeamName}/${this.spClient.getChannelName()}/${file.type?.folderPathInternal}/${uniqueName}`
            },
            overwrite: true
        };
        const result = await this.spClient.spPost(endpoint, body);

        const fileRef = `/sites/${this.spClient.internalSiteName}/${this.documentFolder}/${updatedTeamName}/${this.spClient.getChannelName()}/${file.type?.folderPathInternal}/${uniqueName}`;

        const updateUrl = `${destinationUrl}/_api/web/getfilebyserverrelativeurl('${fileRef}')/ListItemAllFields`;
        await this.spClient.spPatch(updateUrl, { externalFileId: file.uniqueId });
        return !!result.ok;
    }

    /**
     * Retrieves the archived file reference by unique ID (private).
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to the file reference string or undefined.
     */
    private async getArchivedFile(uniqueId: string): Promise<string | undefined> {
        const apiUrl = `${this.spClient.getAbsoluteUrl().replace(this.spClient.getRelativeUrl(), '')}/sites/${this.spClient.internalSiteName}/_api/web/lists/getbytitle('Dokumenty')/items?$filter=externalFileId eq '${uniqueId}'&$select=FileRef&$top=1`;
        const result = await this.spClient.spGet(apiUrl);
        return result.ok && Array.isArray(result.data) ? result.data[0]?.FileRef : undefined;
    }

    // private async getUniqueInternalFileName(
    //     folderPath: string,
    //     fileName: string,
    //     siteUrl: string = this.sp.getAbsoluteUrl(),
    //     uniqueId: string,
    // ): Promise<string> {
    //     const ext = getFileExtension(fileName);
    //     const baseName = getCleanBaseName(fileName);

    //     let count = 0;
    //     let uniqueName = fileName;
    //     let uniquePath = `${folderPath}/${uniqueName}`;

    //     let result = await this.checkInternalFileExists(uniquePath, siteUrl);
    //     console.log(result);

    //     while (result.ok && (result.data as any).externalFileId !== uniqueId) {
    //         count++;
    //         uniqueName = `${baseName} (${count}).${ext}`;
    //         uniquePath = `${folderPath}/${uniqueName}`;
    //         result = await this.checkInternalFileExists(uniquePath, siteUrl);
    //     }

    //     return uniqueName;
    // }

    // private async checkInternalFileExists(path: string, siteUrl: string = this.sp.getAbsoluteUrl()): Promise<IApiResult> {
    //     const apiUrl = `${siteUrl}/_api/web/getfilebyserverrelativeurl('${path}')/ListItemAllFields`;
    //     const result = await this.sp.spGet(apiUrl, {}, false);
    //     return result;
    // }

    /**
     * Moves a file to a new directory, renaming it if a file with the same name already exists.
     * Ensures uniqueness by appending a numeric suffix if necessary.
     *
     * @param fileRef - The current server-relative URL of the file to move.
     * @param fileName - The name of the file (including extension).
     * @param newFileDirRef - The relative path (from the document root) to the new directory.
     * @returns A promise that resolves to true if the move was successful, false otherwise.
     */
    /**
     * Moves a file to a new directory, renaming if necessary to ensure uniqueness.
     * @param fileRef The current server-relative URL of the file.
     * @param fileName The file name.
     * @param newFileDirRef The new directory reference.
     * @param uniqueId Optional unique ID for the file.
     * @returns Promise resolving to true if moved successfully.
     */
    // public async moveFile(fileRef: string, fileName: string, newFileDirRef: string, uniqueId?: string): Promise<boolean> {
    //     const basePath = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}`;
    //     const path = `${basePath}/${newFileDirRef}`;
    //     const ext = getFileExtension(fileName);

    //     const uniqueName = await this.getUniqueFileName(path, fileName);

    //     const shortTempName = `${Date.now()}.${ext}`;
    //     const newUrl = `${basePath}/${uniqueId ? shortTempName : uniqueName}`;
    //     const moveUrl = `${this.sp.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${fileRef}')/moveto(newurl='${newUrl}',flags=1)`;
    //     const result = await this.sp.spPost(moveUrl);
    //     if (!result.ok) {
    //         return false;
    //     }
    //     const moveUrl2 = `${this.sp.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${newUrl}')/moveto(newurl='${path}/${uniqueName}',flags=1)`;
    //     const result2 = await this.sp.spPost(moveUrl2);
    //     return !!result2.ok;
    // }

    public async moveFile(fileRef: string, fileName: string, newFileDirRef: string): Promise<boolean> {
        const basePath = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}`;
        const path = `${basePath}/${newFileDirRef}`;

        const uniqueName = await this.getUniqueFileName(path, fileName);

        try {

            await this.spClient.sp.web
                .getFileByServerRelativePath(fileRef)
                .moveByPath(`${path}/${uniqueName}`, false);
            return true;
        } catch (e) {
            console.error(e);
            return false;
        }
    }

    /**
     * Updates a file's metadata by unique ID.
     * @param uniqueId The unique ID of the file.
     * @param body Metadata update payload.
     * @returns API result object.
     */
    public async updateFileByUniqueId(uniqueId: string, body: object = {}): Promise<IApiResult> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/GetItemByUniqueId('${uniqueId}')`;
        const result = await this.spClient.spPatch(apiUrl, body);
        return result;
    }

    /**
     * Updates a file's metadata by server-relative URL.
     * @param fileRef The server-relative URL of the file.
     * @param body Metadata update payload.
     * @returns API result object.
     */
    public async updateFileByUrl(fileRef: string, body: object = {}): Promise<IApiResult> {
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${fileRef}')/ListItemAllFields`;
        const result = await this.spClient.spPatch(apiUrl, body);
        return result;
    }

    private async resetAndBreakInheritance(uniqueId: string): Promise<void> {
        const baseUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/GetItemByUniqueId('${uniqueId}')`;
        await this.spClient.spPost(`${baseUrl}/resetroleinheritance`);
        await this.spClient.spPost(`${baseUrl}/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`);
        await this.clearRoleAssignments(uniqueId);
    }

    private async clearRoleAssignments(uniqueId: string): Promise<void> {
        const webUrl = this.spClient.getAbsoluteUrl();

        // This gives you the list item behind the file identified by UniqueId (always correct)
        const apiRoot = `${webUrl}/_api/web/GetFileById('${uniqueId}')/ListItemAllFields/RoleAssignments`;

        // Expand to get principal and role definitions
        const result = await this.spClient.spGet(`${apiRoot}?$expand=Member,RoleDefinitionBindings`);
        const roles = Array.isArray(result.data) ? result.data : [];

        for (const ra of roles) {
            await this.spClient.spPost(
                `${apiRoot}/removeroleassignment(principalid=${ra.Member.Id},roledefid=${ra.RoleDefinitionBindings[0].Id})`
            );
        }
    }

    /**
    * upload file using serverRelativeUrl of folder
    * @param folderPath 
    * @param file 
    * @returns fileName if upload fail
    */
    public async uploadFile(folderPath: string, file: FileInfo, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult> {
        const path = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}/${folderPath}`;
        const body = await (file.rawFile as File).arrayBuffer();


        const maxAttempts = 3;
        let attempt = 0;
        let result: IApiResult = { ok: true };
        while (attempt < maxAttempts) {
            const uniqueName = await this.getUniqueFileName(path, file.name, attempt);
            const url = `${this.spClient.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${path}')/Files/add(url='${uniqueName}',overwrite=false)`;
            result = await this.spClient.spPost(url, body);
            if (result.code === 403) {
                attempt++;
                continue;
            } else {
                break;
            }
        }

        if (result.data) {
            const uniqueId = (result.data as any).UniqueId;
            await this.resetFileMetadata(uniqueId, uploadedBy, folderId, documentTypeId, status);
            const uploadedFile = await this.getFileByUniqueId(uniqueId);
            if (!uploadedFile) {
                return { fileName: file.name, success: false };
            }
            return {
                fileName: file.name,
                success: true,
                file: uploadedFile,
                typeId: documentTypeId
            };
        }
        return { fileName: file.name, success: false };
    }

    private async resetFileMetadata(uniqueId: string, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<void> {
        const body: any = {
            Status: status ? status : '',
            StrukturaDokumentaceId: folderId ?? null,
            DocumentTypeId: documentTypeId,
            ContractValidity: null,
            ValidTo: null,
            ScheduledNotificationDate: null,
            ScheduledNotificationMessage: null,
            ScheduledNotificationMessageEn: null,
            ScheduledNotificationRecipients: null,
            Documentation: null,
            ActionToTrigger: null,
        };
        if (uploadedBy) {
            body.UploadedBy = uploadedBy;
        }
        await this.spClient.updateItemInList(uniqueId, body);
    }

    /**
     * Reuploads a file to a folder, overwriting if it exists.
     * @param folderPath The folder path.
     * @param file The file to upload.
     * @param fileName The file name.
     * @param uploadedBy Optional uploader.
     * @param folderId Optional folder ID.
     * @param documentTypeId Optional document type ID.
     * @param status Optional file status.
     * @returns File upload result object.
     */
    public async reuploadFile(folderPath: string, file: FileInfo, fileName: string, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult> {
        const url = `${this.spClient.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files/add(url='${fileName}',overwrite=true)`;
        const body = await (file.rawFile as File).arrayBuffer();
        const result = await this.spClient.spPost(url, body);
        if (result.data) {
            const uniqueId = (result.data as any).UniqueId;
            await this.resetFileMetadata(uniqueId, uploadedBy, folderId, documentTypeId, status);
            const uploadedFile = await this.getFileByUniqueId(uniqueId);
            return {
                fileName: file.name,
                success: true,
                file: uploadedFile,
            };
        }
        return { fileName: file.name, success: false };
    }

    /**
     * Retrieves required files for a folder by ID.
     * @param id The folder ID.
     * @returns Array of required files.
     */
    public async getRequiredFiles(id: number): Promise<IRequiredFile[]> {
        const select = `*,FileDirRef,DocumentType/Title,DocumentType/${ColumnsDocTypes.TitleEn},DocumentType/Id,DocumentType/${ColumnsDocTypes.UploadStatus},DocumentType/${ColumnsDocTypes.FolderPath},DocumentType/${ColumnsDocTypes.FolderPathAfter}`;
        const expand = 'DocumentType';
        const fileDirRef = `${this.spClient.getRelativeUrl()}/Lists/${ListName.RequiredDocuments}/${this.spClient.getChannelName()}/${id}`;
        const filter = `FileDirRef eq '${fileDirRef}'`;
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.RequiredDocuments}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const result = await this.spClient.spGet(apiUrl);
        return Array.isArray(result.data) ? mapRequiredFiles(result.data) : [];
    }

    /**
     * Retrieves all required files for the current channel, grouped by folder number.
     * @returns Map of folder number to array of required files.
     */
    public async getAllRequiredFiles(): Promise<Map<number, IRequiredFile[]>> {
        const select = `*,FileDirRef`;
        const fileDirRef = `${this.spClient.getRelativeUrl()}/Lists/${ListName.RequiredDocuments}/${this.spClient.getChannelName()}`;
        const filter = `startswith(FileDirRef,'${fileDirRef}')`;
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.RequiredDocuments}')/items?$select=${select}&$filter=${filter}`;
        const result = await this.spClient.spGet(apiUrl);
        const data = Array.isArray(result.data) ? result.data.filter(item => (item.ContentTypeId as string).startsWith('0x010')) : [];
        const map = new Map<number, IRequiredFile[]>();

        for (const item of data) {
            const parts = item.FileDirRef.split('/');
            const folderName = parts[parts.length - 1]; // last part is folder number
            const folderNumber = parseInt(folderName, 10);
            if (isNaN(folderNumber)) continue; // skip if last part is not a number

            if (!map.has(folderNumber)) {
                map.set(folderNumber, []);
            }

            // Map raw item to IRequiredFile using your mapping function
            const mappedItems = mapRequiredFiles([item]);
            map.get(folderNumber)?.push(...mappedItems);
        }

        return map;
    }

    /**
     * Generates a unique file name in a folder, appending a numeric suffix if needed (private).
     * @param folderPath The folder path.
     * @param fileName The file name.
     * @param countStart Optional starting count.
     * @param siteUrl Optional site URL.
     * @returns Unique file name string.
     */
    private async getUniqueFileName(
        folderPath: string,
        fileName: string,
        countStart: number = 0,
        siteUrl: string = this.spClient.getAbsoluteUrl(),
    ): Promise<string> {
        const ext = getFileExtension(fileName);
        const baseName = getCleanBaseName(fileName);

        let count = 0 + countStart;
        let uniqueName = count > 0 ? `${baseName} (${count}).${ext}` : fileName;
        let uniquePath = `${folderPath}/${uniqueName}`;

        while (await this.checkFileExists(uniquePath, siteUrl)) {
            count++;
            uniqueName = `${baseName} (${count}).${ext}`;
            uniquePath = `${folderPath}/${uniqueName}`;
        }

        return uniqueName;
    }

    /**
     * Checks if a file exists at a given path (private).
     * @param path The file path.
     * @param siteUrl Optional site URL.
     * @returns Promise resolving to true if exists.
     */
    private async checkFileExists(path: string, siteUrl: string = this.spClient.getAbsoluteUrl()): Promise<boolean> {
        const apiUrl = `${siteUrl}/_api/web/getfilebyserverrelativeurl('${path}')/Exists`;
        const result = await this.spClient.spGet(apiUrl, {}, false);
        return !!result.ok;
    }

    /**
     * Retrieves folders from the document library filtered by the current channel and site context.
     * The folder name adapts based on the site URL (supports localization).
     * Filters items that start with the specified folder path and are of folder content type.
     *
     * @returns A promise that resolves to an array of folders (`IFolder[]`), or an empty array if none found.
     */
    public async getFolders(): Promise<IFolder[]> {
        const folderName = this.spClient.getAbsoluteUrl().includes('appgestio') ? 'Shared documents' : 'Sdilene dokumenty';
        const fileDirRef = `${this.spClient.getRelativeUrl()}/${folderName}/${this.spClient.getChannelName()}`;
        const filter = `startswith(FileDirRef,'${fileDirRef}') and startswith(ContentTypeId,'0x0120')`;
        const select = '*,FileDirRef,FileLeafRef,FileRef';
        const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/items?$filter=${filter}&$select=${select}`;
        const result = await this.spClient.spGet(apiUrl);
        return Array.isArray(result.data) ? mapFolders(result.data) : [];
    }

    public async sendFiletoExtraction(file: IFile, type: FinancialDoc): Promise<boolean> {
        const transformedPath = transformPath(file.serverRelativeUrl.replace(this.spClient.getRelativeUrl(), ''));

        const origin = new URL(this.spClient.getAbsoluteUrl()).origin;

        const apiUrl = `${origin}/sites/${this.spClient.internalSiteName}/_api/web/lists/getbytitle('Extrakce')/items`;
        const body = {
            Title: transformedPath,
            SiteUrl: this.spClient.getAbsoluteUrl(),
            DocumentType: type,
            FolderPath: file.serverRelativeUrl.split('/').slice(0, -1).join('/').replace(this.spClient.getRelativeUrl(), ''),
            FileId: file.id
        };
        const result = await this.spClient.spPost(apiUrl, body);
        return !!result.ok;
    }
}