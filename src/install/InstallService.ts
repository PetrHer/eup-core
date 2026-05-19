import ApiClient from "../api/ApiClient";
import { ColumnsActions, LibraryName, ListName, Language, SPGroupExternal, SPGroupInternal, SPGroupRO, TeamRole, TemplatesListName } from "../enums";
import { ITemplateContent } from "../interfaces";
import { lastVersion } from "../migrations";

const strings = {
    [Language.EN]: {
        Actions: 'Creating list of workflow actions...',
        Categories: 'Creating categories...',
        Detail: "Creating detail...",
        Details: "Creating details list...",
        Duties: "Creating duties list...",
        Folders: "Creating folders...",
        Forms: 'Creating forms list...',
        Notifications: 'Creating list of notifications...',
        QA: "Creating Q&A list...",
        Renaming: 'Renaming folders...',
        Requests: "Creating requests...",
        RequestTypes: 'Creating list of request types...',
        RequiredDocs: 'Creating required documents list...',
        SiteGroups: "Creating user groups...",
        Structure: "Creating list of categories...",
        Trigger: 'Creating list of scheduled actions...',
        Types: 'Creating list of document types...',
        IUP: 'Creating folders in internal storage...',
        Logs: 'Creating logs list...',
        CopyReq: 'Copying required files...',
        ListFolders: 'Creating folders in lists...',
        ListSettings: 'Updating list settings...',
        Permissions: 'Setting permissions...'
    },
    [Language.CS]: {
        Actions: 'Vytvářím seznam workflow akcí...',
        Categories: "Vytvářím kategorie...",
        Detail: "Vytvářím detail...",
        Details: "Vytvářím seznam detailů...",
        Duties: "Vytvářím seznam povinností...",
        Folders: "Vytvářím složky...",
        Forms: 'Vytvářím seznam formulářů...',
        Notifications: 'Vytvářím seznam notifikací...',
        QA: "Vytvářím Q&A seznam...",
        Renaming: 'Přejmenovávám složky...',
        Requests: "Vytvářím žádosti...",
        RequestTypes: 'Vytvářím seznam typů žádostí...',
        RequiredDocs: 'Vytvářím seznam požadovaných dokumentů...',
        SiteGroups: "Vytvářím uživatelské skupiny...",
        Structure: "Vytvářím seznam kategorií...",
        Trigger: 'Vytvářím seznam naplánovaných akcí...',
        Types: 'Vytvářím seznam typů dokumentů...',
        IUP: 'Vytvářím složky v interním úložišti...',
        Logs: 'Vytvářím seznam logů...',
        CopyReq: 'Kopíruji požadované soubory...',
        ListFolders: 'Vytvářím složky v seznamech...',
        ListSettings: 'Aktualizuji nastavení seznamů...',
        Permissions: 'Nastavuji oprávnění...'
    }
}

export default class InstallService {
    constructor(private apiClient: ApiClient) { }

    private getMissingSiteGroups = async (): Promise<(SPGroupExternal | SPGroupInternal | SPGroupRO)[]> => {
        const siteGroups = await this.apiClient.sp.getSiteGroups();
        const missingSiteGroups = [
            ...Object.values(SPGroupExternal).filter(groupName => !siteGroups.find(group => group.title === groupName)),
            ...Object.values(SPGroupInternal).filter(groupName => !siteGroups.find(group => group.title === groupName)),
            ...Object.values(SPGroupRO).filter(groupName => !siteGroups.find(group => group.title === groupName)),
        ];
        return missingSiteGroups;
    }

    public async install(
        increaseProgressCount: () => void,
        locale: Language,
        setProgressText: (text: string) => void,
        setMaximum: (max: number) => void
    ): Promise<void> {
        const missingSiteGroups = await this.getMissingSiteGroups();
        const listExists = await this.apiClient.sp.checkListExistence(ListName.WorkflowActions);
        let maximum = 22;
        if (missingSiteGroups.length > 0) {
            maximum++;
        }
        if (!listExists) {
            maximum++;
        }
        setMaximum(maximum);
        const templateNames = await this.apiClient.install.getTemplatesNames();
        const selectedTemplate = templateNames[0].fileRef;
        const templates: ITemplateContent[] = await this.apiClient.install.getTemplates(selectedTemplate);
        if (missingSiteGroups.length > 0) {
            setProgressText(strings[locale].SiteGroups);
            await this.apiClient.users.createSiteGroups(missingSiteGroups);
            increaseProgressCount();
        }

        //if (!listExists[ListName.DocumentStructure]) {
        setProgressText(strings[locale].Structure);
        await this.apiClient.sp.ensureList(ListName.Categories);
        increaseProgressCount();
        //}
        //if (!listExists[ListName.Details]) {
        setProgressText(strings[locale].Details);
        await this.apiClient.sp.ensureList(ListName.Details);
        increaseProgressCount();
        //}
        setProgressText(strings[locale].Detail);
        templates.sort((a, b) => {
            if (a.stage === b.stage) {
                return a.phase.localeCompare(b.phase);
            } else {
                return a.stage.localeCompare(b.stage);
            }
        });
        const status = templates[0]
            ? `${templates[0].stage}/${templates[0].phase}`
            : '';
        await this.apiClient.detail.createDetail(status, lastVersion);
        increaseProgressCount();

        const template = templateNames.find(t => t.fileRef === selectedTemplate);
        const templateName = template ? template.fileLeafRef : templateNames[0].fileLeafRef;

        setProgressText(strings[locale].RequestTypes);
        await this.apiClient.install.ensureListAndCopyItems(ListName.RequestTypes, TemplatesListName.RequestTypes);
        await this.apiClient.install.setListReadOnly(ListName.RequestTypes);
        increaseProgressCount();

        //if (!listExists[ListName.Requests]) {
        setProgressText(strings[locale].Requests);
        await this.apiClient.sp.ensureList(ListName.Requests);
        increaseProgressCount();
        //}

        //if (!listExists[ListName.QA]) {
        setProgressText(strings[locale].QA);
        await this.apiClient.install.ensureListAndCopyItemsIntoChannelFolder(ListName.QA, TemplatesListName.QA, templateName);
        increaseProgressCount();
        //}

        //if (!listExists[ListName.Duties]) {
        setProgressText(strings[locale].Duties);
        await this.apiClient.sp.ensureList(ListName.Duties);
        await this.apiClient.install.copyTemplateAttachments(TemplatesListName.Duties, ListName.Duties);
        increaseProgressCount();
        //}


        //if (!listExists[LibraryName.Requests]) {
        setProgressText(strings[locale].Requests);
        await this.apiClient.sp.ensureList(LibraryName.Requests);
        increaseProgressCount();
        // }

        //if (!listExists[ListName.DocumentTypes]) {
        setProgressText(strings[locale].Types);
        await this.apiClient.install.ensureListAndCopyItems(ListName.DocumentTypes, TemplatesListName.DocumentTypes);
        await this.apiClient.install.setListReadOnly(ListName.DocumentTypes);
        increaseProgressCount();
        //}

        if (!listExists) {
            setProgressText(strings[locale].Actions);
            await this.apiClient.sp.ensureList(ListName.WorkflowActions);
            await this.apiClient.install.copyFoldersAndItemsWithStructure(TemplatesListName.WorkflowActions, ListName.WorkflowActions, Object.values(ColumnsActions) as string[]);
            await this.apiClient.install.setListReadOnly(ListName.WorkflowActions);
            increaseProgressCount();
        }

        //if (!listExists[ListName.RequiredDocuments]) {
        setProgressText(strings[locale].RequiredDocs);
        await this.apiClient.sp.ensureList(ListName.RequiredDocuments);
        await this.apiClient.install.copyTemplateAttachments(TemplatesListName.RequiredDocuments, ListName.RequiredDocuments);
        increaseProgressCount();
        //}
        setProgressText(strings[locale].RequiredDocs);
        await this.apiClient.sp.ensureList(ListName.Notifications);
        increaseProgressCount();

        setProgressText(strings[locale].Trigger);
        await this.apiClient.sp.ensureList(ListName.ActionToTrigger);
        increaseProgressCount();

        setProgressText(strings[locale].Forms);
        await this.apiClient.sp.ensureList(ListName.Forms);
        increaseProgressCount();

        setProgressText(strings[locale].Logs);
        await this.apiClient.sp.ensureList(ListName.Logs);
        increaseProgressCount();

        setProgressText(strings[locale].ListFolders);
        await this.apiClient.install.createListFolders();
        increaseProgressCount();

        setProgressText(strings[locale].Categories);
        const relations = await this.apiClient.install.createFoldersFromTemplates(templates);
        increaseProgressCount();

        setProgressText(strings[locale].CopyReq);
        await this.apiClient.install.addRequiredFiles(relations, templateName);
        increaseProgressCount();


        setProgressText(strings[locale].Folders);
        await this.apiClient.install.ensureDocumentLibrary(templateName);
        increaseProgressCount();

        setProgressText(strings[locale].Requests);
        await this.apiClient.install.initRequests();
        increaseProgressCount();

        setProgressText(strings[locale].IUP);
        await this.apiClient.install.ensureIUPData();
        increaseProgressCount();

        setProgressText(strings[locale].Permissions);
        await this.apiClient.users.ensureRole();
        await this.apiClient.users.addUserByEmailToTeamAndChannel(this.apiClient.users.eupNoReplyMail, TeamRole.Owner);
        await this.apiClient.users.addEUPAdminsAsOwners();
        await this.apiClient.users.addEUPMembersAsMembers();
        const rmGroup = await this.apiClient.sp.getSiteGroupByName(SPGroupInternal.UUO_RM);
        if (rmGroup) {
            await this.apiClient.users.addUserToSiteGroup(rmGroup.id, this.apiClient.sp.getUserEmail());
        }
        increaseProgressCount();

        setProgressText(strings[locale].ListSettings);
        await this.apiClient.sp.hideLists(Object.values(ListName));
        await this.apiClient.install.updateDateColumns();
        increaseProgressCount();

    }
}