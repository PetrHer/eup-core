import { IInstallApiClient } from './IInstallApiClient';
import { ColumnsCategories, ColumnsDetails, ColumnsReqTypes, ColumnsRequired, LibraryName, ListName, TemplatesListName } from '../../enums';
import { IFolderTemplate, ISiteGroup, ITemplateContent, ITemplateItemRelation, ITemplateReference } from '../../interfaces';
import { delay, expandRoles, normalizeDocType } from '../../utils/utils';
import DetailApiClient from '../detail/DetailApiClient';
import IApiResult from '../IApiResult';
import { listColumns } from '../listColumns';
import { mapFolderTemplates, mapTemplateReferences, mapTemplates } from '../mappers';
import SPApiClient from "../SPApiClient";

export default class InstallApiClient implements IInstallApiClient {
    protected documentFolder: string = 'Sdilene dokumenty';

    /**
     * Constructs an InstallApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    /**
 * Initializes request-related folders and copies request templates.
 * 
 * Creates a request folder for the current channel and an 'options' folder,
 * calling the provided progress callback after each step.
 * Finally, copies the request templates.
 * 
 * @returns Promise<void> - A promise that resolves when all initialization steps are complete.
 */
    public async initRequests(): Promise<void> {
        await this.createRequestFolder(this.sp.getChannelName());
        await this.createRequestFolder('options');
        await this.copyRequestTemplates();
    }

    public async ensureIUPData(): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        // await this.ensureIUPTeam(webUrl);
        await this.ensureIUPChannel(webUrl);
        await this.ensureIUPFolders(webUrl, this.sp.internalSiteName);
    }

    public async updateIUPChannel(body: object): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        const baseApiUrl = `${webUrl}/_api/web/lists/getbytitle('Channels')/items`;
        const channelRes = await this.sp.spGet(`${baseApiUrl}?$filter=ChannelId eq '${this.sp.getChannelId()}'`);
        const id = Array.isArray(channelRes.data) && typeof channelRes.data[0]?.Id === 'number'
            ? channelRes.data[0].Id
            : undefined;
        if (id === undefined) {
            return;
        }
        await this.sp.spPatch(`${baseApiUrl}(${id})`, body);
    }

    public async updateIUPChannelContacts(email?: string, secondaryEmail?: string): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        const userId = email ? await this.sp.getUserSPId(email, webUrl) : null;
        const secondaryId = secondaryEmail ? await this.sp.getUserSPId(secondaryEmail, webUrl) : null;
        const body = {
            ContactId: userId,
            SecondaryContactId: secondaryId,
        };
        await this.updateIUPChannel(body);
    }

    public async updateIUPGroupColumns(oaoEmails: string[], uuoEmails: string[], amlEmails: string[]): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        const oaoIds = await Promise.all(oaoEmails.map(async (email) => await this.sp.getUserSPId(email, webUrl)));
        const uuoIds = await Promise.all(uuoEmails.map(async (email) => await this.sp.getUserSPId(email, webUrl)));
        const amlIds = await Promise.all(amlEmails.map(async (email) => await this.sp.getUserSPId(email, webUrl)));
        const body = {
            OAOId: oaoIds,
            UUOId: uuoIds,
            AMLId: amlIds
        };
        await this.updateIUPChannel(body);
    }

    /**
     * Moves a folder from the specified current path to a new path,
     * then updates the folder's title in the specified SharePoint list.
     *
     * @param path - The current server-relative URL of the folder to move.
     * @param newPath - The target server-relative URL where the folder should be moved.
     * @param newTitle - The new title to set on the folder after moving.
     * @param listName - The SharePoint list name where the folder metadata resides.
     * @returns A Promise that resolves when the move and update operations are complete.
     */
    public async moveFolderAndChangeTitle(path: string, newPath: string, newTitle: string, listName: ListName): Promise<void> {
        const moved = await this.sp.moveFolder(path, newPath);
        if (!moved) {
            return;
        }
        const getItemUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${newPath}')`;
        const itemResponse = await this.sp.spGet(getItemUrl);
        const uniqueId = itemResponse.data ? (itemResponse.data as any).UniqueId : null;

        const body = {
            Title: newTitle
        };
        await this.sp.updateItemInList(uniqueId, body, listName);
    }

    public async ensureFolderForAllChannels(basePath: string, name: string): Promise<void> {
        const topLevelFolders = await this.getTopLevelFolders();

        // 2. ensure folder inside each
        for (const channelName of topLevelFolders) {
            const channelFolderPath = `${basePath}/${channelName}`;
            await this.sp.ensureFolder(channelFolderPath, name);
        }
    }

    /**
     * ensures document library and then creates folders
     * @param folderTemplates 
     */
    public async createFoldersFromTemplates(templateItems: ITemplateContent[]): Promise<ITemplateItemRelation[]> {
        // const folderTemplates = await this.getFolderTemplates(selectedTemplate);
        //await this.ensureDocumentLibrary();
        await this.createMainFolderInDocumentStructure();
        const ids = await this.createCategories(templateItems);
        return ids;
    }

    /**
         * Updates the 'RequiredDocuments' links for all category items in the current channel,
         * ensuring the description contains the correct relative path.
         *
         * @returns A promise that resolves when all category items have been updated.
         */
    public async updateLinksInCategories(): Promise<void> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.Categories}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${fileDirRef}'`;
        const select = `Id,${ColumnsCategories.RequiredDocuments}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.Categories}')/items?$select=${select}&$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        const items = response.data;
        if (Array.isArray(items)) {
            for (let i = 0; i < items.length; i++) {
                const item = items[i];
                const body = {
                    [ColumnsCategories.RequiredDocuments]: {
                        Url: item[ColumnsCategories.RequiredDocuments].Url,
                        Description: `${ListName.RequiredDocuments}/${this.sp.getChannelName()}/${item.Id}` // Optional description
                    }
                };
                await this.sp.updateListItem(item.Id, body, ListName.Categories);
            }
        }
    }

    /**
         * Creates a new category item in the SharePoint list for the current channel,
         * moves the item to the appropriate folder path, and updates its required documents link.
         *
         * @param templateItem - The template content containing category details.
         * @returns The ID of the created item or undefined if creation failed.
         */
    public async createCategoryItem(templateItem: ITemplateContent): Promise<number | undefined> {
        const path = `/${this.sp.getRelativeUrl()}/Lists/${ListName.Categories}/${this.sp.getChannelName()}`;
        const { stage, phase, category, description, stageEn, categoryEn, phaseEn, descriptionEn, order } = templateItem;
        const body = {
            [ColumnsCategories.Stage]: stage,
            [ColumnsCategories.Phase]: phase,
            'Title': category,
            //[ColumnsCategories.Folder]: folder,
            [ColumnsCategories.DescriptionEn]: descriptionEn,
            [ColumnsCategories.Description]: description,
            [ColumnsCategories.StageEn]: stageEn,
            [ColumnsCategories.PhaseEn]: phaseEn,
            [ColumnsCategories.CategoryEn]: categoryEn,
            //[ColumnsCategories.Deadline]: deadline ? deadline : null,
            [ColumnsCategories.Order]: order ? order : null,
        };
        const result = await this.sp.createListItem(body, ListName.Categories);
        if (result.data) {
            const itemId = (result.data as any).Id;
            await this.sp.moveListItem(itemId, path, ListName.Categories);
            const body = {
                [ColumnsCategories.RequiredDocuments]: {
                    Url: `${this.sp.getAbsoluteUrl()}/Lists/${ListName.RequiredDocuments}/${this.sp.getChannelName()}/${itemId}`,
                    Description: `${ListName.RequiredDocuments}/${this.sp.getChannelName()}/${itemId}` // Optional description
                }
            };
            await this.sp.updateListItem(itemId, body, ListName.Categories);
            return itemId;
        }
    }

    public async updateLoanDetailsVersion(version: number): Promise<boolean> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Details}')/items?$select=*`;
        const response = await this.sp.spGet(apiUrl);
        const data = Array.isArray(response.data) ? response.data : [];
        const details = data.filter(item => item[ColumnsDetails.Version] < version);
        for (let i = 0; i < details.length; i++) {
            const detail = details[i];
            const updated = await this.sp.updateListItem(detail.Id, { [ColumnsDetails.Version]: version }, ListName.Details);
            if (!updated) {
                return false;
            }
        }
        return true;
    }

    /**
     * Retrieves templates from the Templates list based on the specified file reference.
     * @param fileRef The file reference path to filter templates.
     * @returns An array of template content objects.
     */
    public async getTemplates(fileRef: string): Promise<ITemplateContent[]> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const apiUrl = `${webUrl}/_api/web/GetListByTitle('${TemplatesListName.DocumentStructure}')/items?$select=*,FileRef,FileLeafRef,FileDirRef&$filter=FileDirRef eq '${fileRef}'`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) ? mapTemplates(response.data) : [];
    }

    /**
     * Retrieves the names of templates from the Templates list.
     * @returns An array of template reference objects containing file reference and leaf file reference.
     */
    public async getTemplatesNames(): Promise<ITemplateReference[]> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const apiUrl = `${webUrl}/_api/web/GetListByTitle('${TemplatesListName.DocumentStructure}')/items?$select=FileRef,FileLeafRef&$filter=FSObjType eq 1`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) ? mapTemplateReferences(response.data) : [];
    }

    /**
         * Ensures the existence of necessary folders for different lists under the current channel,
         * invoking a progress increment callback after each folder is created or confirmed.
         *
         * @returns Promise that resolves when all folders have been ensured.
         */
    public async createListFolders(): Promise<void> {
        await this.ensureListFolder(this.sp.getChannelName(), ListName.Requests);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.Forms);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.Duties);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.RequiredDocuments);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.Notifications);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.ActionToTrigger);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.Logs);
    }

    public async setListReadOnly(listName: ListName): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web?$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup`;
        const groupsResult = await this.sp.spGet(apiUrl);
        if (!groupsResult.data) {
            return;
        }
        const data = groupsResult.data as any;
        const ownersGroupId = data.AssociatedOwnerGroup && data.AssociatedOwnerGroup.Id ? data.AssociatedOwnerGroup.Id : undefined;
        const membersGroupId = data.AssociatedMemberGroup && data.AssociatedMemberGroup.Id ? data.AssociatedMemberGroup.Id : undefined;
        const visitorsGroupId = data.AssociatedVisitorGroup && data.AssociatedVisitorGroup.Id ? data.AssociatedVisitorGroup.Id : undefined;

        const breakApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
        await this.sp.spPost(breakApiUrl);

        const boundary = "batch_" + Date.now();
        const apiUrlBase = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment`;

        const readPart = [membersGroupId, visitorsGroupId]
            .filter(id => id !== undefined)
            .map(id =>
                this.sp.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
            ).join('');

        const writePart = ownersGroupId ? [ownersGroupId].map(id =>
            this.sp.getPermissionBatchString(boundary, id, 1073741829, apiUrlBase)
        ).join('') : '';

        const batchBody = `${readPart}${writePart}--${boundary}--\r\n`;
        await this.sp.spPost(
            `${this.sp.getAbsoluteUrl()}/_api/$batch`,
            batchBody,
            {
                "Content-Type": `multipart/mixed; boundary=${boundary}`,
            }
        );
    }

    public async ensureListFolder(name: string, listName: ListName | LibraryName, path?: string): Promise<void> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${listName}${path ? `/${path}` : ''}`;
        const fileRef = `${fileDirRef}/${name}`;
        const checkApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${fileRef}')`;
        const folderCheck = await this.sp.spGet(checkApiUrl, {}, false);
        if (!folderCheck.ok) {
            const createBody = {
                "ContentTypeId": "0x0120",
                'Title': name,
            };
            const result = await this.sp.createListItem(createBody, listName);
            const renameBody = {
                'Title': name,
                'FileLeafRef': name,
            };
            await this.sp.updateListItem((result.data as any).Id, renameBody, listName);
            if (path) {
                await this.sp.moveFolder(`${this.sp.getRelativeUrl()}/Lists/${listName}/${name}`, fileRef);
            }
        }
    }

    /**
         * check existence of all necessary columns in document library and create them if needed
         */
    public async ensureDocumentLibrary(selectedTemplate: string): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')`;
        const fieldsApiUrl = `${apiUrl}/fields`;
        const result = await this.sp.spGet(`${fieldsApiUrl}`);
        const fields = result.data as any[];
        const requiredFields = [
            { title: 'Status', type: '2' },
            { title: 'FileType', type: '2' },
            { title: 'StrukturaDokumentaceId', type: '9' },
            { title: 'Name_en', type: '2' },
            { title: 'Documentation', type: '8' },
            { title: 'UploadedBy', type: '2' },
            { type: '4', title: 'ValidTo' },
            { title: 'ContractValidity', type: '2' },
            { title: 'ScheduledNotificationDate', type: '4' },
            { title: 'ScheduledNotificationMessage', type: '3' },
            { title: 'ScheduledNotificationMessageEn', type: '3' },
            { title: 'ScheduledNotificationRecipients', type: '3' },
            { title: 'ActionToTrigger', type: '9' },
            { title: 'QAId', type: '2' }
        ];
        for (let i = 0; i < requiredFields.length; i++) {
            const reqField = requiredFields[i];
            const field = fields?.find(f => f.Title === reqField.title);
            if (!field) {
                const body: { [key: string]: object | string } = { 'FieldTypeKind': reqField.type, 'InternalName': reqField.title, 'Title': reqField.title };
                await this.sp.spPost(fieldsApiUrl, body);
            }
        }

        const lookupFieldTitle = "DocumentType";
        const lookupField = fields.find(f => f.Title === lookupFieldTitle);
        const lookupListUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.DocumentTypes}')`;
        const lookupListResult = await this.sp.spGet(lookupListUrl);

        const lookupListId = (lookupListResult.data as any).Id;
        if (!lookupField) {
            if (lookupListId) {
                const lookupFieldXml = `<Field Type='Lookup' DisplayName='${lookupFieldTitle}' Name='${lookupFieldTitle}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
                const lookupBody = { "parameters": { 'SchemaXml': lookupFieldXml } };
                const xmlApiUrl = `${fieldsApiUrl}/CreateFieldAsXml`;
                await this.sp.spPost(xmlApiUrl, lookupBody);
            }
        } else {
            const existingLookupListId = lookupField.LookupList;
            if (existingLookupListId !== lookupListId) {
                const updateFieldXml = `<Field Type='Lookup' DisplayName='${lookupFieldTitle}' Name='${lookupFieldTitle}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
                const updateBody = { 'SchemaXml': updateFieldXml };
                const updateApiUrl = `${fieldsApiUrl}/getbytitle('${lookupFieldTitle}')`;
                await this.sp.spPatch(updateApiUrl, updateBody);
            }
        }
        const headers = {
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE",
        };
        const body = { 'EnableVersioning': true, 'EnableMinorVersions': true, "DraftVersionVisibility": 1 };
        await this.sp.spPost(apiUrl, body, headers);
        await this.ensureFoldersInDocumentLibrary(selectedTemplate);
    }

    /**
         * Adds required files for categories under the current channel and copies template-related files.
         * Ensures each category folder exists in the RequiredDocuments list and copies files based on relations.
         * Calls the provided progress increment callback once the process completes.
         * @param relations - Array of template item relations to copy required files for.
         * @param templateName - Name of the template used for copying files.
         * @returns Promise that resolves when all required files are processed.
         */
    public async addRequiredFiles(relations: ITemplateItemRelation[], templateName: string): Promise<void> {
        const folderPath = `${this.sp.getRelativeUrl()}/lists/${ListName.Categories}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${folderPath}'`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Categories}')/items?$select=Id&$filter=${filter}`;
        const result = await this.sp.spGet(apiUrl);
        const idArray = Array.isArray(result.data) ? result.data.map(item => item.Id) : [];
        for (let i = 0; i < idArray.length; i++) {
            const id: number = idArray[i];
            await this.ensureListFolder(id.toString(), ListName.RequiredDocuments, this.sp.getChannelName());
        }
        for (let i = 0; i < relations.length; i++) {
            const relation = relations[i];
            await this.copyRequiredfiles(relation, templateName);
        }
    }

    /**
         * Ensures a SharePoint list exists and copies items from a template list if the list was newly created.
         * - Checks if the target list exists; creates it if missing.
         * - Ensures necessary columns and list view are set up.
         * - If the list was created, copies items from the specified template list, adjusting file template paths as needed.
         *
         * @param listName - The name of the list to ensure/create.
         * @param templatesListName - The template list name to copy items from if the list is newly created.
         * @returns Promise that resolves when the list and items setup is complete.
         */
    public async ensureListAndCopyItems(listName: ListName, templatesListName: TemplatesListName): Promise<void> {
        const listExists = await this.sp.checkListExistence(listName);
        if (!listExists) {
            const result = await this.sp.createList(listName);
            if (!result) {
                return;
            }
        }
        const missingColumns = await this.sp.ensureColumns(listName);
        if (!listExists) {
            const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
            const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items`;
            const result = await this.sp.spGet(apiUrl);
            if (Array.isArray(result.data)) {
                const itemBodies: object[] = [];
                for (let i = 0; i < result.data.length; i++) {
                    const item = result.data[i];
                    const body: { [key: string]: any } = {
                        Title: item.Title,
                    };
                    listColumns[listName].forEach(col => {
                        if ((col.Title === ColumnsReqTypes.FileTemplate || col.Title === ColumnsReqTypes.FileTemplateEn) && listName === ListName.RequestTypes) {
                            const fileTemplate = item[col.Title]
                                ? item[col.Title]
                                    .split(';')
                                    .map((part: string) =>
                                        part.replace(
                                            `/sites/${this.sp.templateSiteName}/Requests`,
                                            `${this.sp.getRelativeUrl()}/Requests/options`
                                        )
                                    )
                                    .join(';')
                                : '';
                            body[col.Title] = fileTemplate;
                        } else if (col.Title === ColumnsReqTypes.Fields && listName === ListName.RequestTypes) {
                            body[col.Title] = item[col.Title] ? item[col.Title] : [];
                        } else {
                            body[col.Title] = item[col.Title];
                        }
                    });
                    itemBodies.push(body);
                    //await this.createListItem(body, listName);
                }
                await this.sp.createListItemsBatch(itemBodies, listName);
            }
            await this.sp.ensureListView(listName, missingColumns);
        }
    }

    /**
         * Copies folders and items from a source SharePoint list to a target list while preserving the folder structure.
         * 
         * @param sourceListName - The name of the source list from which folders and items will be copied.
         * @param targetListName - The name of the target list where folders and items will be created.
         * @param fields - An array of field names to copy from the source items to the target items.
         * @returns Promise<void> - A promise that resolves when the copying process is complete.
         */
    public async copyFoldersAndItemsWithStructure(sourceListName: TemplatesListName, targetListName: ListName, fields: string[]): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const sourceApiUrl = `${webUrl}/_api/web/lists/getbytitle('${sourceListName}')/items?$select=*,Id,Title,FileSystemObjectType,FileDirRef,FileLeafRef,Folder&$orderby=ID&$top=5000`;
        const result = await this.sp.spGet(sourceApiUrl);

        if (!Array.isArray(result.data)) return;

        const baseSourceFolder = `/sites/${this.sp.templateSiteName}/Lists/${sourceListName}`;

        for (const item of result.data.filter(i => i.FileSystemObjectType === 1)) {
            const folderName = item.FileLeafRef;
            await this.ensureListFolder(folderName, targetListName);
        }

        for (const item of result.data.filter(i => i.FileSystemObjectType !== 1)) {
            const sourcePath = item.FileDirRef;
            const relativePath =
                sourcePath.toLowerCase().startsWith(baseSourceFolder.toLowerCase())
                    ? sourcePath.slice(baseSourceFolder.length)
                    : sourcePath;
            const formValues = [{
                'FieldName': 'Title',
                'FieldValue': item.Title
            }];
            fields.forEach(field => {
                let value = item[field];

                if (typeof value === 'boolean') {
                    value = value ? 'true' : 'false';
                } else if (value === null || value === undefined) {
                    value = '';
                } else {
                    value = String(value);
                }

                formValues.push({
                    'FieldName': field,
                    'FieldValue': value
                });
            });

            const body = { "formValues": formValues };
            await this.sp.createListItemUsingPath(body, targetListName, relativePath, false);
        }
    }

    public async ensureListAndCopyItemsIntoChannelFolder(listName: ListName, templatesListName: TemplatesListName, templateName: string): Promise<void> {
        const listExists = await this.sp.checkListExistence(listName);
        if (!listExists) {
            const result = await this.sp.createList(listName);
            if (!result) {
                return;
            }
        }
        const missingColumns = await this.sp.ensureColumns(listName);
        await this.ensureListFolder(this.sp.getChannelName(), ListName.QA);
        const itemsExists = await this.checkCopiedItemsExistence(listName);
        if (!itemsExists) {
            const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
            const folderPath = `/sites/${this.sp.templateSiteName}/lists/${templatesListName}/${templateName}`;
            const filter = `FileDirRef eq '${folderPath}'`;
            const select = '*,FileDirRef';
            const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items?$select=${select}&$filter=${filter}`;
            const result = await this.sp.spGet(apiUrl);
            if (Array.isArray(result.data)) {
                const itemBodies: object[] = [];
                for (let i = 0; i < result.data.length; i++) {
                    const item = result.data[i];
                    // const body: { [key: string]: any } = {
                    //     Title: item.Title,
                    // };
                    const formValues = [{
                        'FieldName': 'Title',
                        'FieldValue': item.Title
                    }];
                    listColumns[listName].forEach(col => {
                        let value = item[col.Title];

                        if (typeof value === 'boolean') {
                            value = value ? 'true' : 'false';
                        } else if (value === null || value === undefined) {
                            value = '';
                        } else {
                            value = String(value);
                        }

                        formValues.push({
                            'FieldName': col.Title,
                            'FieldValue': value
                        });
                    });

                    const body = { "formValues": formValues };
                    //await this.createListItemUsingPath(body, listName);
                    itemBodies.push(body);
                }
                await this.createListItemsBatchUsingPath(listName, itemBodies);
            }
            await this.sp.ensureListView(listName, missingColumns);
        }
    }

    public async copyTemplateAttachments(templatesListName: TemplatesListName, listName: ListName): Promise<void> {
        const itemTitle = 'šablony';
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items?$filter=Title eq '${itemTitle}'`;
        const result = await this.sp.spGet(apiUrl);
        const id = Array.isArray(result.data) && result.data[0] ? result.data[0].Id : undefined;
        if (!id) {
            return;
        }
        const attachUrl = `${webUrl}/_api/lists/getbytitle('${templatesListName}')/items(${id})/AttachmentFiles`;
        const attachResult = await this.sp.spGet(attachUrl);
        const files = Array.isArray(attachResult.data)
            ? await Promise.all(attachResult.data.map(async (item) => {
                const result = await this.sp.spGetBinary(`${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}${item.ServerRelativeUrl}`); await this.sp.spGetBinary(`${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}${item.ServerRelativeUrl}`);
                const arrayBuffer = result.ok && result.data instanceof ArrayBuffer ? result.data : undefined;
                return {
                    fileName: item.FileName,
                    arrayBuffer: arrayBuffer
                };
            }))
            : [];
        const filter = `Title eq '${itemTitle}' and FileDirRef eq '${this.sp.getRelativeUrl()}/Lists/${listName}'`;
        const checkApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/items?$filter=${filter}&$select=*,FileDirRef`;
        const checkResult = await this.sp.spGet(checkApiUrl);
        let itemid = Array.isArray(checkResult.data) && checkResult.data[0] ? checkResult.data[0].Id : undefined;
        if (!itemid) {
            const body = { Title: itemTitle };
            const createResult = await this.sp.createListItem(body, listName);
            itemid = createResult.data ? (createResult.data as any).Id : undefined;
        }
        if (!itemid) {
            return;
        }
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            if (!file.arrayBuffer) {
                continue;
            }
            await this.sp.addAttachmentArrayBuffer(itemid, file.fileName, file.arrayBuffer, listName, true);
        }
    }


    /**
     * Ensures a folder with the specified name exists within the 'Requests' library.
     * 
     * @param name - The name of the folder to create or verify.
     * @returns Promise<boolean> - A promise that resolves to true if the folder exists or was created successfully, otherwise false.
     */
    private async createRequestFolder(name: string): Promise<boolean> {
        const path = `${this.sp.getRelativeUrl()}${LibraryName.Requests}/`;
        return await this.sp.ensureFolder(path, name) !== undefined;
    }

    private async createListItemsBatchUsingPath(
        listName: ListName,
        itemBodies: object[],
        path?: string,
        channelSpecific: boolean = true
    ): Promise<IApiResult> {
        const chunkSize = 50;
        const chunks = this.sp.chunkArray(itemBodies, chunkSize);
        const results: IApiResult[] = [];

        for (const chunk of chunks) {
            const result = await this.createListItemsBatchChunkUsingPath(listName, chunk, path, channelSpecific);
            results.push(result);
            await delay(200);
        }

        return {
            ok: results.every(r => r.ok),
            data: results.flatMap(r => r.data ?? []),
            error: results.find(r => !r.ok)?.error ?? undefined
        };
    }

    private async createListItemsBatchChunkUsingPath(
        listName: ListName,
        itemBodies: object[],
        path?: string,
        channelSpecific: boolean = true
    ): Promise<IApiResult> {
        const boundary = "batch_" + Date.now();
        let folderPath = `${this.sp.getAbsoluteUrl()}/lists/${listName}`;

        if (channelSpecific) {
            folderPath += `/${this.sp.getChannelName()}`;
        }
        if (path) {
            folderPath += `${path}`;
        }

        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/AddValidateUpdateItemUsingPath`;

        const batchParts = itemBodies.map(body => {
            const bodyWithPath = {
                listItemCreateInfo: {
                    FolderPath: { DecodedUrl: folderPath },
                    UnderlyingObjectType: 0
                },
                bNewDocumentUpdate: false,
                ...body
            };

            return [
                `--${boundary}`,
                'Content-Type: application/http',
                'Content-Transfer-Encoding: binary',
                '',
                `POST ${apiUrl} HTTP/1.1`,
                'Content-Type: application/json',
                'Accept: application/json',
                '',
                JSON.stringify(bodyWithPath),
                ''
            ].join('\r\n');
        });

        const batchBody = [...batchParts, `--${boundary}--\r\n`].join('\r\n');

        const response = await this.sp.spPost(
            `${this.sp.getAbsoluteUrl()}/_api/$batch`,
            batchBody,
            {
                'Content-Type': `multipart/mixed; boundary=${boundary}`
            }
        );

        return response;
    }

    private async checkCopiedItemsExistence(listName: ListName): Promise<boolean> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${listName}/${this.sp.getChannelName()}`;
        const filter = `startswith(FileDirRef,'${fileDirRef}')`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${listName}')/items?$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) && response.data.length > 0;
    }

    /**
         * Copies required files related to a given template from the source 'RequiredDocuments' template list
         * to the target list, preserving metadata and folder structure.
         * 
         * Fetches items filtered by the template name and relation templateId, expands DocumentType info,
         * then for each required document:
         * - Finds the corresponding DocumentType ID in the target list,
         * - Creates a new list item with relevant fields,
         * - Moves the created item to the appropriate folder path based on the current channel and relation itemId.
         * 
         * @param relation - An object containing the relation info including templateId and itemId.
         * @param templateName - The name of the template to filter required files.
         * @returns Promise<void> - A promise that resolves once all required files are copied and moved.
         */
    private async copyRequiredfiles(relation: ITemplateItemRelation, templateName: string): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const fileDirRef = `/sites/${this.sp.templateSiteName}/Lists/${TemplatesListName.RequiredDocuments}/${templateName}/${relation.templateId}`;
        const filter = `FileDirRef eq '${fileDirRef}'`;
        const select = '*,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id';
        const expand = 'DocumentType';
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${TemplatesListName.RequiredDocuments}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const templates = await this.sp.spGet(apiUrl);
        const path = `${this.sp.getRelativeUrl()}/Lists/${ListName.RequiredDocuments}/${this.sp.getChannelName()}/${relation.itemId}`;
        const docTypesUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.DocumentTypes}')/items`;
        const docTypes = await this.sp.spGet(docTypesUrl);
        const docTypeMap = new Map(
            (docTypes.data as any[]).map(t => [
                normalizeDocType(t.Title_en),
                t.Id
            ])
        );
        if (Array.isArray(templates.data)) {
            for (let i = 0; i < templates.data.length; i++) {
                const item = templates.data[i];
                const templateKey = normalizeDocType(item.DocumentType.Title_en);
                const typeId = docTypeMap.get(templateKey) ?? null;
                const body = {
                    Title: item.Title,
                    //[ColumnsRequired.FolderPath]: item[ColumnsRequired.FolderPath],
                    [ColumnsRequired.TitleEn]: item[ColumnsRequired.TitleEn],
                    [ColumnsRequired.FileDesc]: item[ColumnsRequired.FileDesc],
                    [ColumnsRequired.DocumentTypeId]: typeId,
                };
                const result = await this.sp.createListItem(body, ListName.RequiredDocuments);
                if (result.data) {
                    await this.sp.moveListItem((result.data as any).Id, path, ListName.RequiredDocuments);
                }
            }
        }
    }

    /**
 * Ensures that all necessary folders based on the selected template exist in the document library.
 * Calls folder creation for each folder template and updates progress via callback.
 * 
 * @param selectedTemplate - The name or identifier of the selected folder template.
 */
    private async ensureFoldersInDocumentLibrary(selectedTemplate: string): Promise<void> {
        const path = `${this.sp.getRelativeUrl()}/${this.documentFolder}`;
        await this.sp.ensureFolder(path, this.sp.getChannelName());
        const folderTemplates = await this.getFolderTemplates(selectedTemplate);
        const siteGroups = await this.sp.getSiteGroups(this.sp.getAbsoluteUrl());
        await this.generateFoldersFromTemplates(folderTemplates, '', siteGroups);
    }

    /**
     * Retrieves folder templates for the given selected template from a specific team's document library.
     *
     * @param selectedTemplate - The name of the selected folder template to retrieve.
     * @returns Promise resolving to an array of folder templates ({@link IFolderTemplate}).
     */
    private async getFolderTemplates(selectedTemplate: string): Promise<IFolderTemplate[]> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const fileDirRef = `/sites/${this.sp.templateSiteName}/${this.documentFolder}/${selectedTemplate}`;
        const apiUrl = `${webUrl}/_api/web/GetListByTitle('${LibraryName.Documents}')/items?$select=*,FileRef,FileLeafRef,FileDirRef&$filter=startswith(FileDirRef,'${fileDirRef}')`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) ? mapFolderTemplates(result.data) : [];
    }

    /**
         * recursively generate folders from templates loaded from site EUP-Sablony
         * @param folderTemplates 
         * @param parentFolderPath 
         */
    private async generateFoldersFromTemplates(folderTemplates: IFolderTemplate[], parentFolderPath: string, siteGroups: ISiteGroup[]): Promise<void> {
        for (let i = 0; i < folderTemplates.length; i++) {
            const folderTemplate = folderTemplates[i];
            const uniqueId = await this.createFolder(folderTemplate, parentFolderPath);
            if (folderTemplate.nameEn && uniqueId) {
                const body = {
                    Name_en: folderTemplate.nameEn
                };
                await this.sp.updateItemInList(uniqueId, body, LibraryName.Documents);
            }
            if (uniqueId && (folderTemplate.write || folderTemplate.read)) {
                const folderRelativeUrl = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}${parentFolderPath}/${folderTemplate.name}`;
                await this.setFolderPermissions(folderTemplate, folderRelativeUrl, siteGroups, uniqueId, this.sp.getAbsoluteUrl());
            }
            if (folderTemplate.templates) {
                const path = `${parentFolderPath}/${folderTemplate.name}`;
                await this.generateFoldersFromTemplates(folderTemplate.templates, path, siteGroups);
            }
        }
    }

    private async setFolderPermissions(folderTemplate: IFolderTemplate, folderRelativeUrl: string, siteGroups: ISiteGroup[], uniqueId: string, webUrl: string): Promise<void> {
        // const channelName = await this.getChannelName();
        // const folderRelativeUrl = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${channelName}${parentFolderPath}/${folderTemplate.name}`;
        const writeGroups = folderTemplate.write ? expandRoles(folderTemplate.write) : [];
        const readGroupsAll = folderTemplate.read ? expandRoles(folderTemplate.read) : [];
        const writeGroupIds = siteGroups.filter(g => writeGroups.includes(g.title)).map(g => g.id);
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${LibraryName.Documents}')/GetItemByUniqueId('${uniqueId}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
        await this.sp.spPost(apiUrl);

        const readGroups = [...readGroupsAll].filter(g => !writeGroups.includes(g));
        const readGroupIds = siteGroups.filter(g => readGroups.includes(g.title)).map(g => g.id);
        await this.assignFolderPermissions(folderRelativeUrl, readGroupIds, writeGroupIds, webUrl);
    }

    private async assignFolderPermissions(folderRelativeUrl: string, readGroupsIds: number[], writeGroupIds: number[], webUrl: string): Promise<void> {
        const boundary = "batch_" + Date.now();
        const apiUrlBase = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${folderRelativeUrl}')/ListItemAllFields/roleassignments/addroleassignment`;

        const readPart = readGroupsIds.map(id =>
            this.sp.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
        ).join('');

        const writePart = writeGroupIds.map(id =>
            this.sp.getPermissionBatchString(boundary, id, 1073741827, apiUrlBase)
        ).join('');

        const batchBody = `${readPart}${writePart}--${boundary}--\r\n`;
        await this.sp.spPost(
            `${webUrl}/_api/$batch`,
            batchBody,
            {
                "Content-Type": `multipart/mixed; boundary=${boundary}`,
            }
        );
    }

    /**
 * creates new subfolder with custom attributes in folder using serverRelativeUrl
 * @param folder 
 * @param parentFolderPath 
 * @returns boolean
 */
    private async createFolder(folder: IFolderTemplate, parentFolderPath: string): Promise<string | undefined> {
        const path = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}/${parentFolderPath}/`;
        return await this.sp.ensureFolder(path, folder.name);
    }

    /**
 * Creates a folder structure based on the provided template items and updates progress.
 * @param templateItems The template content items used to create the folder structure.
 * @returns A promise indicating the completion of the folder structure creation.
 */
    private async createCategories(templateItems: ITemplateContent[]): Promise<ITemplateItemRelation[]> {
        const ids: ITemplateItemRelation[] = [];
        for (let i = 0; i < templateItems.length; i++) {
            const itemId = await this.createCategoryItem(templateItems[i]);
            if (itemId) {
                ids.push({ itemId: itemId, templateId: templateItems[i].id });
            }
        }
        return ids;
    }

    private async getTopLevelFolders(): Promise<string[]> {
        const path = `${this.sp.getRelativeUrl()}/${this.documentFolder}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}` +
            `/_api/web/GetFolderByServerRelativeUrl('${path}')/Folders?$select=Name,ServerRelativeUrl`;

        const result = await this.sp.spGet(apiUrl);

        if (!result.ok || !Array.isArray(result.data)) {
            return [];
        }

        return result.data
            .filter((f: any) => !f.Name.startsWith("_")) // remove system folders
            .map((f: any) => f.ServerRelativeUrl);
    }

    /**
     * Creates the main folder in the document structure for the current channel.
     * @returns A promise indicating the completion of the folder creation.
     */
    private async createMainFolderInDocumentStructure(): Promise<void> {
        const folderPath = encodeURIComponent(`${this.sp.getRelativeUrl()}/Lists/${ListName.Categories}`);
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1='${folderPath}'&@a2='${this.sp.getChannelName()}'&siteUrl=${this.sp.getAbsoluteUrl()}`;
        await this.sp.spPost(apiUrl);
    }

    private async ensureIUPFolders(webUrl: string, iupTeamName: string): Promise<void> {
        //const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${this.documentFolder}')`;
        const siteGroups = await this.sp.getSiteGroups(webUrl);
        const teamName = await this.sp.getTeamName();
        const updatedTeamName = teamName.replace('_ext', '');
        const path = `/sites/${iupTeamName}/${this.documentFolder}/${updatedTeamName}`;
        const teamFoldercheck = await this.sp.spGet(`${webUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')`);
        if (!teamFoldercheck.ok) {
            const apiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('/sites/${iupTeamName}/${this.documentFolder}')/folders/add('${updatedTeamName}')`;
            const result = await this.sp.spPost(apiUrl);
            if (!result.ok) {
                return;
            }
        }
        const checkApiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${path}/${this.sp.getChannelName()}')`;
        const check = await this.sp.spGet(checkApiUrl);
        if (!check.ok) {
            const apiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')/folders/add('${this.sp.getChannelName()}')`;
            const result = await this.sp.spPost(apiUrl);
            if (result.ok) {
                const templates = await this.getFolderTemplates('Interní uložiště');
                await this.createIUPFolders(`${path}/${this.sp.getChannelName()}`, templates, webUrl, iupTeamName, siteGroups);
            }
        }
        //}
    }

    private async createIUPFolders(parentFolderPath: string, templates: IFolderTemplate[], webUrl: string, iupTeamName: string, siteGroups: ISiteGroup[]): Promise<void> {
        for (let i = 0; i < templates.length; i++) {
            const template = templates[i];
            const apiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${parentFolderPath}')/folders/add('${template.name}')`;
            /*const result =*/ await this.sp.spPost(apiUrl);
            // const uniqueId = (result.data as any).UniqueId;
            // if (uniqueId && (template.write || template.read)) {
            //     const folderRelativeUrl = `${parentFolderPath}/${template.name}`;
            //     await this.setFolderPermissions(template, folderRelativeUrl, siteGroups, uniqueId, webUrl);
            // }
            if (template.templates) {
                await this.createIUPFolders(`${parentFolderPath}/${template.name}`, template.templates, webUrl, iupTeamName, siteGroups);
            }
        }
    }

    private async ensureIUPChannel(webUrl: string): Promise<void> {
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('Channels')/items`;
        const channelDataExists = await this.sp.spGet(`${apiUrl}?$filter=ChannelId eq '${this.sp.getChannelId()}'`);
        const detail = await (new DetailApiClient(this.sp).getLoanDetail());
        if (!channelDataExists.ok || (channelDataExists.data as any[]).length === 0) {
            const body = {
                Title: this.sp.getChannelName(),
                TeamId: this.sp.getTeamId(),
                ChannelId: this.sp.getChannelId(),
                TeamSiteUrl: this.sp.getAbsoluteUrl(),
                ChannelStatus: 'Active',
                Stage: detail?.currentStatus || '',
                Phase: detail?.currentFolder || '',
                //ChannelType: isPrivate ? 'Private' : 'Standart'
            };
            await this.sp.spPost(apiUrl, body);
            await this.updateIUPChannelContacts(detail?.contact?.email, detail?.secondaryContact?.email);
        } else if (channelDataExists.ok && Array.isArray(channelDataExists.data) && channelDataExists.data[0]) {
            const body = {
                ChannelStatus: 'Active',
                Stage: detail?.currentStatus || '',
                Phase: detail?.currentFolder || '',
                TeamSiteUrl: this.sp.getAbsoluteUrl(),
                //ChannelType: isPrivate ? 'Private' : 'Standart'
            };
            await this.updateIUPChannel(body);
        }
    }

    /**
    * Copies all request template files from the 'Requests' list in the source site to the target location.
    * 
    * Retrieves all items including file references and file names, then calls copyRequestTemplate
    * for each item to perform the actual file copying.
    * 
    * @returns Promise<void> - A promise that resolves when all request templates have been copied.
    */
    private async copyRequestTemplates(): Promise<void> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}`;
        const apiUrl = `${webUrl}/_api/web/lists/getbytitle('Requests')/items?$select=*,FileRef,FileLeafRef`;
        const result = await this.sp.spGet(apiUrl);
        if (Array.isArray(result.data)) {
            for (let i = 0; i < result.data.length; i++) {
                const item = result.data[i];
                await this.copyRequestTemplate(item.FileRef, item.FileLeafRef, webUrl);
            }
        }
    }

    /**
     * Copies a single request template file by downloading it from the source URL and uploading it
     * to the target 'Requests/options' folder, overwriting if the file already exists.
     * 
     * @param fileRef - The server-relative URL of the source file to download.
     * @param fileLeafRef - The file name to use when uploading the file.
     * @param webUrl - The base URL of the source SharePoint site.
     * @returns Promise<void> - A promise that resolves once the file has been uploaded successfully.
     */
    private async copyRequestTemplate(fileRef: string, fileLeafRef: string, webUrl: string): Promise<void> {
        const downloadUrl = `${webUrl}/_api/web/getfilebyserverrelativeurl('${fileRef}')/$value`;
        const result = await this.sp.spGetBinary(downloadUrl);
        const fileContent = result.ok && result.data instanceof ArrayBuffer ? result.data : undefined;
        const uploadUrl = `${this.sp.getAbsoluteUrl()}/_api/web/getfolderbyserverrelativeurl('${this.sp.getRelativeUrl()}/Requests/options')/files/add(overwrite=true, url='${fileLeafRef}')`;
        if (fileContent) {
            await this.sp.spPostBinary(uploadUrl, fileContent);
        }
    }

    public async updateDateColumns(): Promise<void> {
        const docApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${LibraryName.Documents}')/fields/GetByTitle('ScheduledNotificationDate')`;
        await this.sp.spPatch(docApiUrl, { 'DisplayFormat': 0 });

        const reqApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.RequiredDocuments}')/fields/GetByTitle('${ColumnsRequired.Deadline}')`;
        await this.sp.spPatch(reqApiUrl, { 'DisplayFormat': 0 });
    }
}