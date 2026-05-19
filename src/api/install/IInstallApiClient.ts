import { LibraryName, ListName, TemplatesListName } from '../../enums';
import { ITemplateContent, ITemplateItemRelation, ITemplateReference } from '../../interfaces';

export interface IInstallApiClient {
    initRequests(): Promise<void>;
    ensureIUPData(): Promise<void>;
    updateIUPChannel(body: object): Promise<void>;
    updateIUPChannelContacts(email?: string, secondaryEmail?: string): Promise<void>;
    updateIUPGroupColumns(oaoEmails: string[], uuoEmails: string[], amlEmails: string[]): Promise<void>;
    moveFolderAndChangeTitle(path: string, newPath: string, newTitle: string, listName: ListName): Promise<void>;
    ensureFolderForAllChannels(basePath: string, name: string): Promise<void>;
    createFoldersFromTemplates(templateItems: ITemplateContent[]): Promise<ITemplateItemRelation[]>;
    updateLinksInCategories(): Promise<void>;
    createCategoryItem(templateItem: ITemplateContent): Promise<number | undefined>;
    updateLoanDetailsVersion(version: number): Promise<boolean>;
    getTemplates(fileRef: string): Promise<ITemplateContent[]>;
    getTemplatesNames(): Promise<ITemplateReference[]>;
    createListFolders(): Promise<void>;
    setListReadOnly(listName: ListName): Promise<void>;
    ensureListFolder(name: string, listName: ListName | LibraryName, path?: string): Promise<void>;
    ensureDocumentLibrary(selectedTemplate: string): Promise<void>;
    addRequiredFiles(relations: ITemplateItemRelation[], templateName: string): Promise<void>;
    ensureListAndCopyItems(listName: ListName, templatesListName: TemplatesListName): Promise<void>;
    copyFoldersAndItemsWithStructure(sourceListName: TemplatesListName, targetListName: ListName, fields: string[]): Promise<void>;
    ensureListAndCopyItemsIntoChannelFolder(listName: ListName, templatesListName: TemplatesListName, templateName: string): Promise<void>;
    copyTemplateAttachments(templatesListName: TemplatesListName, listName: ListName): Promise<void>;
    updateDateColumns(): Promise<void>;
}
