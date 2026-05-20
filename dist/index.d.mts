import * as React from 'react';
import { FieldValues } from 'react-hook-form';
import { FileInfo } from '@syncfusion/ej2-react-inputs';
import { MenuItemModel, MenuEventArgs } from '@syncfusion/ej2-react-navigations';
import { FileInfo as FileInfo$1 } from '@syncfusion/ej2-inputs';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from '@pnp/sp';

declare enum AppTab {
    Requests = "requests",
    Forms = "forms",
    Calendar = "calendar",
    Documentation = "documentation",
    QnA = "qnA",
    AllDocuments = "allDocuments",
    Chat = "chat"
}

declare enum ColumnsDetails {
    Contact = "contact",
    ChannelId = "channelId",
    CurrentStatus = "currentStatus",
    LoanText = "loanText",
    SecondaryContact = "secondaryContact",
    ChannelStatus = "ChannelStatus",
    ContactData = "contactData",
    SecondaryData = "secondaryData",
    Company = "company",
    Version = "version"
}

declare enum ColumnsForms {
    FormData = "FormData",
    TitleEn = "Title_en",
    Status = "Status",
    FormType = "FormType",
    Category = "Category",
    Language = "Language",
    Completed = "Completed",
    FileId = "FileId"
}

declare enum ColumnsQA {
    Answer = "answer",
    Order = "QAOrder",
    Answered = "Answered",
    Priority = "Priority",
    Category = "Category",
    FilesData = "FilesData",
    SubCategory = "SubCategory"
}

declare enum ColumnsRequests {
    FormData = "FormData",
    RequestType = "RequestType",
    RequestStatus = "RequestStatus",
    FileTemplate = "FileTemplate",
    CategoryId = "categoryId",
    FileIds = "FileIds",
    ExportData = "ExportData"
}

declare enum ColumnsCategories {
    Stage = "Stage",
    Phase = "Phase",
    Description = "FolderDescription",
    DescriptionEn = "FolderDescription_en",
    StageEn = "Stage_en",
    PhaseEn = "Phase_en",
    CategoryEn = "Category_en",
    RequiredDocuments = "RequiredDocuments",
    Order = "CategoryOrder",
    CategoryStatus = "CategoryStatus"
}

declare enum ContextMenuAction {
    NewRequests = "newRequests",
    CompletedRequest = "completedRequests",
    ActualDocuments = "actualDocuments",
    AcceptedDocuments = "acceptedDocuments"
}

declare enum FileActionOptions {
    Delete = "delete",
    Download = "download",
    Open = "open",
    Accept = "accept",
    Return = "return",
    Process = "process",
    Publish = "publish",
    Decline = "decline",
    AddToDoc = "addToDoc",
    RemoveFromDoc = "removeFromDoc",
    Archive = "archive",
    CopyLink = "copyLink",
    Extract = "extract"
}

declare enum FileStatus {
    Processing = "processing",
    Accepted = "accepted",
    Declined = "declined",
    Waiting = "waiting",
    Requested = "requested",
    Deleted = "deleted",
    Published = "published",
    AmlReturned = "aml-returned",
    Provided = "provided",
    PreparedForClient = "prepared-for-client",
    Verification = "progress-verification",
    Returned = "returned",
    NotValid = "not-valid",
    BoVerified = "bo-verified",
    Prepared = "prepared",
    BankRevision = "bank-revision",
    ClientRevision = "client-revision",
    FinalVersion = "final-version",
    AcceptedConditional = "accepted-conditional",
    NoType = "no-type",
    Exported = "exported",
    Signed = "signed",
    ReturnedToSign = "returned-to-sign",
    NotProvided = "not-provided"
}

declare enum Filter {
    Actual = "actual",
    Accepted = "accepted"
}

declare enum FolderStatus {
    Accepted = "accepted",
    Default = "default"
}

declare enum FormStatus {
    New = "new",
    Submitted = "submitted",
    ReturnedToClient = "returnedToClient",
    Processing = "processing",
    Completed = "completed",
    Exported = "exported",
    Signed = "signed",
    ReturnedToSign = "returned-to-sign"
}

declare enum ListName {
    Details = "Detaily",
    Requests = "Zadosti",
    QA = "QA",
    Categories = "Kategorie",
    Duties = "Povinnosti",
    Forms = "Formulare",
    RequiredDocuments = "PozadovaneDokumenty",
    DocumentTypes = "TypyDokumentu",
    RequestTypes = "TypyZadosti",
    Notifications = "Notification",
    WorkflowActions = "WorkflowActions",
    ActionToTrigger = "ActionToTrigger",
    Logs = "Logs"
}

declare enum LoadingStatus {
    Loading = "loading",
    Loaded = "loaded",
    NoTeam = "NoTeam",
    Preparing = "preparing",
    Unauthorized = "unauthorized",
    WrongVersion = "wrong-version"
}

declare enum Permissions {
    Owner = "owner",
    Member = "member",
    External = "external",
    Admin = "admin"
}

declare enum RequestStatus {
    New = "new",
    Processing = "processing",
    Completed = "completed",
    Provided = "provided"
}

declare enum RequestType {
    Activation = "activation",
    Repayment = "repayment",
    Statement = "statement",
    Template = "template"
}

declare enum SettingsMenuKey {
    Folders = "folders",
    Contact = "contact",
    Templates = "templates",
    Status = "status",
    Install = "install",
    Users = "users",
    Sharepoint = "sharepoint",
    Internal = "internal",
    Migration = "migration"
}

declare enum TeamName {
    JTFG = "EUP-sablony",
    IUP = "IUP",
    JTFGIUP = "Uvery-Dev"
}

declare enum ToastType {
    Error = "error",
    Success = "success"
}

declare enum ColumnsDuties {
    Date = "Date",
    EventType = "EventType",
    Amount = "Amount",
    Currency = "Currency",
    Instruction = "Instruction",
    Category = "Category",
    EventDescription = "EventDescription",
    Note = "Note",
    Period = "Period",
    EndDate = "EndDate"
}

declare enum UserType {
    External = "03external",
    Internal = "02internal",
    Invited = "01invited"
}

declare enum Language {
    CS = "cs-CZ",
    EN = "en-US"
}

declare enum SPGroupInternal {
    UUO_RM = "Relationship manager",
    UUO_RSMA = "RSM Analyst",
    UUO_BO = "Back office",
    UUO_DEP = "UUO Deputy",
    UUO_DIR = "UUO Director",
    ORKR_RA = "Risk analyst",
    ORKR_GCCA = "GCC analyst",
    ORKR_RD = "Risk deputy",
    ORKR_RAPP = "Risk director",
    OAO_CC = "Credit controller",
    OAO_LAW = "Lawyer",
    OAO_DEP = "OAO Deputy",
    OAO_DIR = "OAO Director",
    WO_S = "WorkOut Specialist",
    WO_DEP = "WorkOut Deputy",
    WO_DIR = "WorkOut Director",
    AML_S = "AML Specialist",
    AML_DIR = "AML Director",
    APPR = "Approver",
    INT_AUD = "Intern\u00ED Auditor",
    BUS_ADMIN = "Business Admin",
    ADMIN = "Admin"
}

declare enum SPGroupExternal {
    EXT_AUD = "External Auditor",
    EXT_ADVISER = "External Adviser",
    EXT_CLILOW = "Client - lower access",
    EXT_CLIFULL = "Client - full access"
}

declare enum TemplatesListName {
    Invitations = "Pozvanky",
    DocumentStructure = "Kategorie",
    Forms = "Forms",
    DocumentTypes = "TypyDokumentu",
    RequiredDocuments = "PozadovaneDokumenty",
    RequestTypes = "TypyZadosti",
    WorkflowActions = "WorkflowActions",
    QA = "QA",
    Duties = "Povinnosti"
}

declare enum LibraryName {
    Documents = "Dokumenty",
    DocumentsAppGestion = "Documents",
    Requests = "Requests"
}

declare enum ColumnsRequired {
    TitleEn = "Title_en",
    FileDesc = "FileDescription",
    DocumentTypeId = "DocumentTypeId",
    DocumentType = "DocumentType",
    PermittedFormats = "PermittedFormats",
    UploadedFile = "uploadedFile",
    Deadline = "Deadline",
    Status = "Status"
}

declare enum ColumnsDocTypes {
    TitleEn = "Title_en",
    UploadStatus = "UploadStatus",
    Permissions = "Permissions",
    FolderPath = "FolderPath",
    FolderPathAfter = "FolderPathAfter",
    ValidityStatuses = "ValidityStatuses",
    FolderPathInternal = "FolderPathInternal",
    FolderPathInternalAfter = "FolderPathInternalAfter"
}

declare enum EventType {
    Covenant = "covenant",
    Repayment = "repayment",
    Meeting = "meeting"
}

declare enum RequestAction {
    Accept = "accept",
    Download = "download",
    Open = "open",
    DownloadAttachments = "DownloadAttachments"
}

declare enum ColumnsReqTypes {
    TitleEn = "Title_en",
    FileTemplate = "FileTemplate",
    FileTemplateEn = "FileTemplateEn",
    Description = "Description",
    DescriptionEn = "Description_en",
    Stage = "Etapa",
    Fields = "Fields",
    FileAttachments = "FileAttachments",
    Notification = "Notification",
    FileReply = "FileReply"
}

declare enum RequestFields {
    Active = "Aktivn\u00ED \u00FA\u010Det",
    Amount = "\u010C\u00E1stka",
    Blocked = "Blokovan\u00FD \u00FA\u010Det",
    ReqSpec = "\u017D\u00E1dost",
    Period = "Obdob\u00ED",
    LoanSpec = "\u00DAv\u011Br",
    FeeSpec = "Poplatek",
    Date = "Datum",
    Currency = "M\u011Bna",
    ActiveLoan = "Aktivn\u00ED \u00FAv\u011Br"
}

declare enum RequestFormFields {
    Active = "active",
    Amount = "amount",
    Blocked = "blocked",
    ReqSpec = "reqSpec",
    Period = "period",
    LoanSpec = "loanSpec",
    FeeSpec = "feeSpec",
    Date = "date",
    Currency = "currency",
    ActiveLoan = "activeLoan",
    Comment = "comment"
}

declare enum DataDispatchAction {
    SET_DETAIL = "SET_DETAIL",
    SET_PERMISSIONS = "SET_PERMISSIONS",
    SET_STRUCTURE = "SET_STRUCTURE",
    SET_TEAM_MEMBERS = "SET_TEAM_MEMBERS",
    SET_FILE_TYPES = "SET_FILE_TYPES",
    SET_SITE_GROUPS = "SET_SITE_GROUPS",
    SET_LIST_EXISTS = "SET_LISTS_EXISTS",
    SET_NOTIFICATIONS = "SET_NOTIFICATIONS",
    SET_WORKFLOW_ACTIONS = "SET_WORKFLOW_ACTIONS",
    SET_CURRENT_USER_GROUPS = "SET_CURRENT_USER_GROUPS",
    UPDATE_STRUCTURE_NODE = "UPDATE_STRUCTURE_NODE",
    ADD_FILES_TO_NODE = "ADD_FILES_TO_NODE",
    SET_FORM_RECORDS = "SET_FORM_RECORDS",
    SET_SELECTED_FORM = "SET_SELECTED_FORM",
    SET_SCHEDULED_ACTIONS = "SET_SCHEDULED_ACTIONS",
    SET_QA_ITEMS = "SET_QA_ITEMS",
    SET_REQUESTS = "SET_REQUESTS",
    ADD_FILES_TO_NODES = "ADD_FILES_TO_NODES"
}

declare enum ColumnsActions {
    DocumentType = "DocumentType",
    DocumentStatus = "DocumentStatus",
    Notification = "Notification",
    ResultStatus = "ResultStatus",
    Publish = "Publish",
    AllowComment = "AllowComment",
    ActionGroups = "ActionGroups",
    Phase = "Phase",
    NotificationGroups = "NotificationGroups",
    Task = "Task",
    TaskGroups = "TaskGroups",
    TitleEn = "Title_en",
    NotificationEn = "Notification_en",
    TaskEn = "Task_en",
    ReadPermissions = "ReadPermissions",
    WritePermissions = "WritePermissions",
    Validity = "validity",
    Reupload = "Reupload",
    FileLink = "FileLink",
    OnMajorVersion = "onMajorVersion"
}

declare enum FormType {
    KYC = "KYC",
    AML = "AML",
    Default = "Default"
}

declare enum SharedDocuments {
    AppGestio = "Shared Documents",
    JT = "Sdilene dokumenty"
}

declare enum RoleDefinition {
    Edit = "\u00DApravy",
    CustomEdit = "\u00DApravy bez maz\u00E1n\u00ED",
    Read = "\u010Cten\u00ED"
}

declare enum ChannelStatus {
    Active = "Active",
    Inactive = "Inactive"
}

declare enum FileValidityAction {
    Start = "start",
    Update = "update",
    End = "end",
    StartFinal = "start+final"
}

declare enum ColumnsTrigger {
    FileId = "FileId",
    ActionId = "ActionId",
    Processed = "Processed",
    Result = "Result",
    Comment = "Comment"
}

declare enum EventPeriodicity {
    Monthly = "Monthly",
    Quarterly = "Quarterly",
    Semiannually = "Semiannually",
    Annually = "Annually"
}

declare enum TeamRole {
    Member = "member",
    Owner = "owner",
    Guest = "guest"
}

declare enum SystemAccount {
    EUPNoReply = "eup_noreply@jtfg.com",
    EUPNoReplyDev = "eup_noreply_dev@jtfg.com",
    EUPAdminsId = "7397659e-48e7-48fa-84f0-bd1d6f290538",
    EUPMembersId = "6834f589-81b2-4e47-9ec0-a41243a85aab"
}

declare enum NavDispatchAction {
    SET_CURRENT_TAB = "SET_CURRENT_TAB",
    SET_ACTIVE_STEPS = "SET_ACTIVE_STEPS",
    SET_SCROLL_CATEGORY_ID = "SET_SCROLL_CATEGORY_ID",
    SET_NOTIFICATION_TARGET_ID = "SET_NOTIFICATION_TARGET_ID",
    SET_ACTIVE_STEP = "SET_ACTIVE_STEP"
}

declare enum DocumentTypeName {
    AdministrativeExtract = "Administrativn\u00ED - V\u00FDpis",
    AdministrativeRequest = "Administrativn\u00ED - \u017D\u00E1dost",
    QA = "Q&A"
}

declare enum QAPriority {
    Min = "min",
    Max = "max",
    Normal = "normal"
}

declare enum QASortType {
    Newest = "newest",
    Oldest = "oldest",
    Priority = "priority"
}

declare enum QAItemType {
    Question = "Question",
    Answer = "Answer",
    Comment = "Comment"
}

declare enum QACategory {
    Basic = "Z\u00E1kladn\u00ED data o spole\u010Dnosti",
    Project = "Projektov\u00E1 dokumentace",
    Financial = "Finan\u010Dn\u00ED data",
    Legal = "Pr\u00E1vn\u00ED dokumentace",
    Information = "Informa\u010Dn\u00ED povinnost",
    Request = "\u017D\u00E1dosti na banku",
    Form = "Formul\u00E1\u0159e",
    Other = "Ostatn\u00ED"
}

declare enum QAStatus {
    Completed = "completed",
    New = "new"
}

declare enum ContactDesc {
    Sales = "Sales",
    BackOffice = "BackOffice",
    LoanAdministration = "LoanAdministration"
}

declare enum FinancialDoc {
    PnL = "VZZ",
    BalanceSheet = "Rozvaha"
}

declare enum RoleAlias {
    UUO = "UUO",
    OAO = "OAO",
    AML = "AML",
    WO = "WO",
    ORKR = "ORKR",
    Admin = "admin",
    Internal = "internal",
    Client = "client",
    External = "external"
}

declare enum SPGroupRO {
    RORole = "Read Only role"
}

declare enum ChannelType {
    Standard = "Standard",
    Private = "Private"
}

interface ILocalizedStrings {
    [Language.CS]: string;
    [Language.EN]: string;
}

interface IFileType {
    loc: ILocalizedStrings;
    id: number;
    uploadStatus?: string;
    permissions?: string[];
    folderPath: string;
    folderPathAfter?: string;
    validityStatuses: string[];
    folderPathInternal?: string;
    folderPathInternalAfter?: string;
}

/**
 * Represents a remaped file loaded from SharePoint.
 */
interface IFile {
    /** The name of the file. */
    name: string;
    /** The current status of the file. */
    status: FileStatus;
    /** The URL or path to the file. */
    serverRelativeUrl: string;
    /** A unique identifier for the file. */
    uniqueId: string;
    version: string;
    type?: IFileType;
    created?: Date;
    id?: number;
    nodeId?: number;
    documentation?: boolean;
    provided?: string;
    contractValidity?: string;
    validTo?: Date;
    qaId?: string;
    author?: string;
    typeId?: number;
    modified?: Date;
}

interface IRequiredFile {
    loc: ILocalizedStrings;
    id: number;
    description?: string;
    type?: IFileType;
    format?: string;
    uploadedFile?: string;
    deadline?: Date;
    status?: FileStatus;
}

interface IStructureNode {
    uploadedFiles: IFile[];
    requiredFiles?: IRequiredFile[];
    nodes: IStructureNode[];
    name: string;
    description?: ILocalizedStrings;
    status: FolderStatus;
    localization?: ILocalizedStrings;
    isCategory?: boolean;
    id?: number;
    order?: number;
}

interface IFolderTemplate {
    name: string;
    nameEn?: string;
    templates?: IFolderTemplate[];
    write?: string;
    read?: string;
}

interface IPerson {
    name: string;
    email: string;
    phone: string;
    photo: string;
    firstName: string;
    lastName: string;
    userType: UserType;
    id?: string;
    loadedPhoto: boolean;
    channelMemberId?: string;
}

interface ILoanDetail {
    contactDesc?: ContactDesc;
    secondaryDesc?: ContactDesc;
    id: number;
    currentStatus?: string;
    currentFolder?: string;
    text?: string;
    title: string;
    contact?: IPerson;
    secondaryContact?: IPerson;
    channelStatus: ChannelStatus;
    iupLink: string;
    version: number;
}

interface IModalState {
    body: React.ReactNode;
    title: string;
    width: number | string;
}

interface IRequest {
    status: RequestStatus;
    created: Date;
    id: number;
    type: IRequestOption;
    templateRef?: string;
    formData: FieldValues;
    categoryId?: string;
    fileIds: string[];
}

interface IToast {
    message: string;
    type: ToastType;
}

/**
 * props used in multiple components
 */
interface ICommonProps {
    refreshCallback: () => Promise<void>;
}

interface IRequestData {
    type?: RequestType;
    typeId?: number;
    status: RequestStatus;
    templateRef?: boolean;
    templateName?: string;
    formData?: FieldValues;
    exportData?: string;
}

interface ISPColumn {
    FieldTypeKind: string;
    Title: string;
    Choices?: string[];
    lookUpListTitle?: ListName;
    displayName: string;
    allowMultiple?: boolean;
}

/**
 * used for file attachments from requests list
 */
interface IAttachment {
    name: string;
    serverRelativeUrl: string;
}

interface IQnAItem {
    question: string;
    answer?: string;
    comments: ISPComment[];
    uniqueId: string;
    id: number;
    status: string;
    category?: QACategory;
    phase: string;
    order?: number;
    stage: string;
    createdText: string;
    answeredText?: string;
    priority?: QAPriority;
    created: Date;
    answeredUser: string;
    answeredDate: string;
    createdUser: string;
    createdDate: string;
    filesData?: IQnAFilesData;
    subCategory?: string;
    modified: string;
}

interface INotification {
    id: number;
    title: ILocalizedStrings;
    task: boolean;
    completed: boolean;
    category: ILocalizedStrings;
    stage: ILocalizedStrings;
    phase: ILocalizedStrings;
    created: string;
    author: string;
    unread: boolean;
    read: number[];
    targetId?: string;
    categoryId: number;
    comment?: string;
    status?: string;
    modified?: string;
    editor?: string;
    attachments?: {
        fileName: string;
        url: string;
    }[];
    link?: string;
    uniqueId: string;
    fileLink?: string;
}

interface ISPComment {
    text: string;
    author: {
        email: string;
        name: string;
        isExternal: boolean;
    };
    createdDate: string;
    id: number;
}

interface IFormField {
    type: string;
    label: string;
    field: string;
    options?: {
        label: string;
        value: string | number;
    }[];
}

interface ITemplateReference {
    fileRef: string;
    fileLeafRef: string;
}

interface ITemplateContent {
    stage: string;
    phase: string;
    category: string;
    description: string;
    descriptionEn: string;
    stageEn: string;
    phaseEn: string;
    categoryEn: string;
    id: number;
    order?: string;
}

interface ICalendarEvent {
    uniqueId: string;
    id: number;
    subject: string;
    time: Date | undefined;
    startDate: Date | undefined;
    type: string;
    amount?: number;
    currency?: string;
    instruction?: string;
    category?: string;
    description?: string;
    note?: string;
    node?: IStructureNode;
    period?: EventPeriodicity;
    endDate?: Date;
}

interface ISiteGroup {
    id: number;
    title: SPGroupExternal | SPGroupInternal | SPGroupRO;
    users: ISiteGroupUser[];
}

interface IInvitation {
    name: string;
    email: string;
    userExists: boolean;
}

interface ISiteGroupUser {
    name: string;
    email: string;
    id: number;
}

interface IFormRecord {
    id: number;
    title: string;
    formData: FieldValues;
    status: FormStatus;
    localization: ILocalizedStrings;
    categoryId?: number;
    type: FormType;
    language?: Language;
    completed?: string;
    fileId?: number;
    file?: IFile;
}

interface IListNotificationCount {
    [ListName.Requests]: number;
    [ListName.Forms]: number;
    'Total': number;
}

interface IRequestOption {
    id: number;
    type: string;
    fileTemplates: {
        [Language.CS]?: string;
        [Language.EN]?: string;
    }[];
    loc: ILocalizedStrings;
    stage?: string[];
    desc?: ILocalizedStrings;
    fields?: string[];
    fileAttachments?: boolean;
    notifGroups?: string;
    fileReply?: boolean;
}

interface IFolder {
    name: string;
    localization: ILocalizedStrings;
    serverRelativeUrl: string;
    folders: IFolder[];
}

interface ITemplateItemRelation {
    templateId: number;
    itemId: number;
}

interface IFileInput {
    file: FileInfo;
    type?: IFileType;
}

interface ITeamDetail {
    teamSiteUrl: string;
}

interface IChannelDetail {
    displayName: string;
    absoluteUrl: string;
    userRole: number;
}

interface IExtendedFile extends IFile {
    description?: string;
    format?: string;
    stage?: string;
    phase?: string;
    category?: string;
    reupload?: boolean;
    requiredFile?: IRequiredFile;
    deadline?: Date;
}

interface IWorkflowAction {
    id: string;
    title: ILocalizedStrings;
    fileTypes: string[];
    statuses: string[];
    notification?: ILocalizedStrings;
    resultStatus?: string;
    publish?: boolean;
    allowComment?: boolean;
    actionGroups: string[];
    notificationGroups: string[];
    task?: ILocalizedStrings;
    taskGroups: string[];
    phase?: string;
    readPermissions?: string;
    writePermissions?: string;
    validity?: FileValidityAction;
    reupload: boolean;
    fileLink?: boolean;
    type: string;
    onMajorVersion: boolean;
}

interface INewNotification {
    title: string;
    task: boolean;
    categoryId: number;
    userGroups: string;
    targetId?: string;
    titleEn: string;
    link?: string;
    comment?: string;
    fileLink?: string;
}

interface IFileUploadResult {
    fileName: string;
    success: boolean;
    file?: IFile;
    typeId?: number;
}

interface IFileVersion {
    versionLabel: string;
    status: string;
    editor: string;
    modified: string;
    relativeUrl: string;
    isCurrentVersion: boolean;
    id: number;
}

interface IContextMenuProps {
    rect: DOMRect;
    items: MenuItemModel[];
    contextMenuSelect: (args: MenuEventArgs) => Promise<void>;
}

interface IScheduledAction {
    actionId: number;
    fileId: number;
    processed?: Date;
    id?: number;
}

interface IExtendedPerson extends IPerson {
    spMembership?: (SPGroupExternal | SPGroupInternal)[];
    userExists?: boolean;
}

interface IQnAAttachment extends IAttachment {
    commentId?: number;
}

interface IQnAFilesData extends Partial<{
    [K in QAItemType]: IQnAAttachment[];
}> {
}

interface IActionApiClient {
    getWorkflowActions(fromTemplates?: boolean): Promise<IWorkflowAction[]>;
    getScheduledActions(): Promise<Map<number, IScheduledAction>>;
    getScheduledActionByFileId(fileId?: number): Promise<IScheduledAction | undefined>;
    createScheduledAction(action: IScheduledAction, title: string, comment?: string): Promise<boolean>;
}

interface ICalendarApiClient {
    getCalendarEvents(): Promise<ICalendarEvent[]>;
    createCalendarEvent(event: ICalendarEvent): Promise<boolean>;
}

interface IChatApiClient {
    getMessages: () => Promise<any[]>;
}

interface IDetailApiClient {
    createDetail(currentStatus: string, version: number): Promise<boolean>;
    getLoanDetail(): Promise<ILoanDetail | undefined>;
    updateDetail(itemId: number, body: object): Promise<boolean>;
    updateContact(itemId: number, contactData: string, secondaryData: string, email?: string, secondaryEmail?: string): Promise<boolean>;
}

interface IApiResult {
    ok: boolean;
    code?: number;
    data?: unknown;
    error?: string;
    exception?: unknown;
}

interface IFileApiClient {
    moveFileToBin(fileRef: string, fileName: string): Promise<boolean>;
    loadFiles(id: number): Promise<IFile[]>;
    loadFilesByIds(ids: number[]): Promise<IFile[]>;
    loadFilesByUniqueIds(uniqueIds: string[]): Promise<IFile[]>;
    loadFileByRelativeUrl(serverRelativeUrl: string): Promise<IFile | undefined>;
    loadAllFiles(documentation: boolean): Promise<IFile[]>;
    publishFile(uniqueId: string): Promise<boolean>;
    setFilePermissions(action: IWorkflowAction, siteGroups: ISiteGroup[], uniqueId: string): Promise<void>;
    uploadFiles(files: IFileInput[], secondStage: boolean, uploadedBy?: string, folderId?: number): Promise<IFileUploadResult[]>;
    getFileTypes(webUrl?: string): Promise<IFileType[]>;
    undoFileCheckOut(uniqueId: string): Promise<boolean>;
    getFileByUniqueId(uniqueId: string): Promise<IFile | undefined>;
    getFileVersionHistory(id: number): Promise<IFileVersion[]>;
    getFileVersion(relativeUrl: string, versionId: number): Promise<ArrayBuffer | undefined>;
    getFileToDownload(relativeUrl: string): Promise<ArrayBuffer | undefined>;
    archiveFile(file: IFile, stage?: IStructureNode, phase?: IStructureNode): Promise<boolean>;
    moveFile(fileRef: string, fileName: string, newFileDirRef: string): Promise<boolean>;
    updateFileByUniqueId(uniqueId: string, body?: object): Promise<IApiResult>;
    updateFileByUrl(fileRef: string, body?: object): Promise<IApiResult>;
    uploadFile(folderPath: string, file: FileInfo, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    reuploadFile(folderPath: string, file: FileInfo, fileName: string, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    getRequiredFiles(id: number): Promise<IRequiredFile[]>;
    getAllRequiredFiles(): Promise<Map<number, IRequiredFile[]>>;
    getFolders(): Promise<IFolder[]>;
    sendFiletoExtraction(file: IFile, type: FinancialDoc): Promise<boolean>;
}

interface IFormApiClient {
    getFormRecords(): Promise<IFormRecord[]>;
    saveFormData(id: number, formData: FieldValues, language: Language, asDraft?: boolean): Promise<boolean>;
}

interface IInstallApiClient {
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

interface ISPApiClient {
    readonly internalSiteName: string;
    readonly templateSiteName: string;
    checkAllListExistence: () => Promise<{
        [key in ListName | LibraryName.Requests]: boolean;
    }>;
    ensureList: (listName: ListName | LibraryName.Requests) => Promise<void>;
    deleteList: (listName: ListName) => Promise<boolean>;
    updateItemInList: (uniqueId: string, body: object, listName?: string) => Promise<boolean>;
    updateListItem(itemId: number, body: object, listName: ListName | LibraryName): Promise<boolean>;
    deleteListItem: (id: number, listName: ListName) => Promise<boolean>;
    getFieldChoices: (listName: ListName, columnTitle: string) => Promise<string[]>;
    createListItemUsingPath: (body: object, listName: ListName, path?: string, channelSpecific?: boolean) => Promise<IApiResult>;
    updateListItemsBatch: (updates: {
        uniqueId: string;
        body: object;
    }[], listName: ListName) => Promise<void>;
    updateList: (body: object, listName?: string) => Promise<void>;
    updateColumn: (body: object, listName: string, columnName: string) => Promise<void>;
    ensureColumns: (listName: ListName | LibraryName.Requests) => Promise<ISPColumn[]>;
    getUserSPId: (userEmail: string) => Promise<number>;
    init: (ctx?: string, teamId?: string, channelId?: string) => Promise<void>;
    getUserRole: () => number;
    getTeamId: () => string;
    getChannelId: () => string;
    getAbsoluteUrl: () => string;
    getChannelName: () => string;
    getRelativeUrl: () => string;
    getTeamName: () => Promise<string>;
    getSiteGroups: () => Promise<ISiteGroup[]>;
    getTemplatesAttachments: (listName: ListName) => Promise<IAttachment[]>;
    moveFolder: (path: string, newPath: string) => Promise<boolean>;
    hideLists: (listNames: ListName[]) => Promise<void>;
    log: (message: string, data?: any) => Promise<void>;
    getSiteGroupByName: (groupName: string, webUrl?: string) => Promise<ISiteGroup | undefined>;
    isPrivateChannel: () => Promise<boolean>;
    getUserEmail: () => string;
}

interface INotificationApiClient {
    getAllNotifications(groups: string[]): Promise<INotification[]>;
    createNotification(notification: INewNotification): Promise<boolean>;
    completeTask(id: number, body: object, files: FileInfo[]): Promise<boolean>;
    deleteNotifications(notifications: INotification[]): Promise<void>;
    deleteNotificationsByTargetId(targetId: string, link?: string): Promise<void>;
    completeTasksByTargetId(targetId: string, link?: string, resultStatus?: string): Promise<void>;
    getTaskByTargetId(targetId: string): Promise<INotification | undefined>;
    getTaskHistory(fileId: string): Promise<INotification[]>;
}

interface IQnAApiClient {
    createQnAItem(question: string, category: string, stage: string, phase: string, priority: string, uploadedBy: string, fileType?: IFileType, order?: number, files?: IFileInput[], subCategory?: string): Promise<string | undefined>;
    addComment(qnAItem: IQnAItem, listName: string, text: string, files: IFileInput[], uploadedBy: string): Promise<{
        newComment: ISPComment | undefined;
        filesData: IQnAFilesData | undefined;
    }>;
    uploadQAFiles(files: IFileInput[], uploadedBy: string, qnAItemId: number, fileType?: IFileType, commentId?: number): Promise<IQnAAttachment[]>;
    deleteComment(listName: ListName, uniqueId: string, commentId: number): Promise<boolean>;
    getQnAItems(stage?: string, phase?: string): Promise<IQnAItem[]>;
    getQnACategories(): Promise<string[]>;
}

interface IRequestApiClient {
    createRequest(data: IRequestData, files: FileInfo$1[], fileType: IFileType, secondStage: boolean, uploadedBy: string): Promise<{
        result: IApiResult;
        failedUploads: string[];
    }>;
    getRequests(): Promise<IRequest[]>;
    getRequestsOptions(): Promise<IRequestOption[]>;
    updateRequest(body: object, itemId: number, uploadedBy?: string, secondStage?: boolean, files?: IFileInput[], category?: IStructureNode): Promise<{
        files: IFile[];
        result: boolean;
    }>;
}

interface IStructureApiClient {
    getStructure(): Promise<IStructureNode[]>;
}

interface IUserApiClient {
    readonly eupNoReplyMail: string;
    getTeamMembers(): Promise<IPerson[]>;
    getExistingUser(mail: string): Promise<IPerson | undefined>;
    getUserPhoto(userIdOrEmail: string): Promise<string>;
    getCurrentUserGroups(): Promise<(SPGroupInternal | SPGroupExternal)[]>;
    getSiteGroupUsersBatch(groupIds: number[]): Promise<Map<number, ISiteGroupUser[]>>;
    ensureRole(): Promise<void>;
    createSiteGroups(groupNames: (SPGroupExternal | SPGroupInternal | SPGroupRO)[], increaseProgressCount: () => void): Promise<void>;
    getInvitations(teamMembers: IPerson[]): Promise<IInvitation[]>;
    addUserToSiteGroup(id: number, userEmail: string): Promise<void>;
    removeUserFromsiteGroup(groupId: number, userId: number): Promise<void>;
    createInvitation(email: string, name: string, roles: string, language: string, existingUser?: boolean): Promise<boolean>;
    checkTemplatesAccess(): Promise<boolean>;
    checkUserAndCreateInvitation(email: string, name: string, roles: string, language: string): Promise<boolean>;
    addUserByEmailToTeamAndChannel(email: string, role?: TeamRole): Promise<boolean>;
    addEUPAdminsAsOwners(): Promise<boolean>;
    addEUPMembersAsMembers(): Promise<boolean>;
    getUserName(userEmail: string): Promise<string>;
    getCurrentUser(): Promise<IPerson | undefined>;
    getExistingUsers(mail: string): Promise<IPerson[]>;
    removeUserFromTeamOrChannel(person: IExtendedPerson): Promise<boolean>;
}

interface IApiClient {
    users: IUserApiClient;
    structure: IStructureApiClient;
    requests: IRequestApiClient;
    qnA: IQnAApiClient;
    notifications: INotificationApiClient;
    install: IInstallApiClient;
    forms: IFormApiClient;
    files: IFileApiClient;
    detail: IDetailApiClient;
    calendar: ICalendarApiClient;
    sp: ISPApiClient;
    actions: IActionApiClient;
    chat: IChatApiClient;
}

/**
 * Adds a class to all child elements within a dropdown item to set its width correctly
 * @param element - The parent HTML element containing the dropdown items
 * @param className - The class name to be added to each child element
 * @returns void
 */
declare const addClassToDropdownItem: (element: HTMLElement, className: string) => void;
declare const addClassToComboBoxItem: (element: HTMLElement | Element, className: string, width: number) => void;
/**
 * Creates a unique file name by appending a counter to the base name if the file already exists in the uploaded files list.
 * @param fileName - The original file name
 * @param uploadedFiles - The list of already uploaded files
 * @returns The unique file name
 */
declare const createUniqueFileName: (fileName: string, uploadedFiles: IFile[]) => string;
/**
 * Extracts the initials from a full name by taking the first letter of each part of the name.
 * @param name - The full name (e.g., "John Doe")
 * @returns A string containing the initials (e.g., "JD")
 */
declare const getInitials: (name: string) => string;
/**
* Converts a UI version number into a human-readable version label.
* @param uiVersion - The UI version number (e.g., 513)
* @returns A string representing the version label (e.g., "1.1")
*/
declare const convertUIVersionToLabel: (uiVersion: number) => string;
/**
 * Extracts the file extension from a given file name.
 * @param fileName - The file name (e.g., "document.pdf")
 * @returns The file extension (e.g., "pdf") or an empty string if no extension exists.
 */
declare const getFileExtension: (fileName?: string) => string;
/**
 * Extracts the file name without its extension.
 * @param fileName - The file name (e.g., "document.pdf")
 * @returns The file name without the extension (e.g., "document")
 */
declare const getFileNameWOExt: (fileName: string) => string;
declare const getFileName: (relativeUrl?: string) => string;
/**
 * Removes the file extension and trailing numeric suffix (e.g., " (1)") from a file name
 * @param fileName - The full name of the file including extension and optional suffix
 * @returns The cleaned base name of the file without extension or numeric suffix
 */
declare const getCleanBaseName: (fileName: string) => string;
/**
 * Cleans the folder name by removing a leading two-digit number and a space, if present.
 * @param folderName - The folder name (e.g., "01 FolderName")
 * @returns The cleaned folder name without the leading number (e.g., "FolderName")
 */
declare const cleanFolderName: (folderName: string) => string;
/**
     * Calculate the height of a context menu based on the number of items
     * @param items - The list of menu items in the context menu
     * @returns number - The calculated height of the context menu
     */
declare const getContextMenuHeight: (items: MenuItemModel[]) => number;
/**
 * Recursively searches for a node with a specific ID within a tree structure
 * @param nodes - The array of structure nodes to search through
 * @param id - The ID of the node to find
 * @returns The matching node if found, otherwise undefined
 */
declare const findNodeById: (nodes: IStructureNode[], id: number) => IStructureNode | undefined;
/**
 * Retrieves the localized stage, phase, and category names for a node by its ID
 * @param nodes - The hierarchical array of structure nodes (stages → phases → categories)
 * @param id - The ID of the target category node
 * @param locale - The language code used to retrieve localized names
 * @returns An object containing the localized stage, phase, and category names, or undefined if not found
 */
declare const getLocalizationPathById: (nodes: IStructureNode[], id: number, locale: Language) => {
    stage: string;
    phase: string;
    category: string;
} | undefined;
declare const getStageByCategoryId: (nodes: IStructureNode[], id?: number) => string | undefined;
declare const getStageNodeByCategoryId: (nodes: IStructureNode[], id?: number) => IStructureNode | undefined;
declare const getPhaseNodeByCategoryId: (nodes: IStructureNode[], id?: number) => IStructureNode | undefined;
/**
 * Parses the form data from a string to an object.
 * @param formData - The form data in string format
 * @returns FieldValues - The parsed form data object, or an empty object if parsing fails
 */
declare const parseFormData: (formData: string) => FieldValues;
declare const parseSharePointDate: (value?: string) => Date | undefined;
/**
 * Type guard to check if an item is a FormRecord based on presence of 'formData' property
 * @param item - The item to check, which can be either IExtendedFile or IFormRecord
 * @returns True if the item is an IFormRecord, otherwise false
 */
declare const isFormRecord: (item: IExtendedFile | IFormRecord) => item is IFormRecord;
/**
 * Opens a file URL in a new browser tab, appending a query parameter to indicate web view
 * @param url - The URL of the file to open
 * @returns void
 */
declare const openFileInTab: (url: string) => void;
declare const openFileInDesktopApp: (url: string) => void;
/**
 * Extracts the leaf (most specific) part of a hierarchical file type title.
 *
 * If the title includes a '>' delimiter (e.g., "Parent > Child"), it returns the child part.
 * Otherwise, returns the original title (trimmed).
 *
 * @param title - The full title string, optionally containing a '>' hierarchy.
 * @returns The simplified or leaf title, or an empty string if input is undefined or empty.
 */
declare const getFileTypeTitle: (title?: string) => string;
declare const formatDateForSharePoint: (date: Date) => string;
declare const expandRoles: (roles: string) => string[];
declare const expandRoleAliases: (aliases: RoleAlias[]) => string[];
declare const delay: (ms: number) => Promise<void>;
declare const downloadFile: (name: string, serverRelativeUrl: string, apiClient: IApiClient) => Promise<void>;
declare const classifyUploader: (currentUserGroups: (SPGroupInternal | SPGroupExternal)[]) => string;
declare const splitFolderPath: (fullPath: string) => {
    path: string;
    name: string;
};
declare const createCategoryOptions: (nodes: IStructureNode[], locale: Language) => {
    [key: string]: any;
}[];
declare const downloadFiles: (files: IFile[], apiClient: IApiClient, zipName?: string) => Promise<void>;
declare const parseNameAndDate: (input?: string) => {
    name: string;
    date: string;
};
declare const getLatestDate: (dates: string[]) => string;
declare const copyLink: (fileLink: string) => Promise<void>;
declare const normalizeDocType: (s?: string) => string;
declare const transformPath: (path: string) => string;

declare abstract class BaseApiClient {
    protected webPartContext: WebPartContext;
    protected getUrl: () => string;
    constructor(webPartContext: WebPartContext, getUrl: () => string);
    /**
     * Retrieves a fresh request digest value from SharePoint for authentication.
     * @returns Promise resolving to the FormDigestValue string.
     */
    protected getFreshDigest(): Promise<string>;
    /**
     * Sends a POST request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object, string, or ArrayBuffer (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    spPost(url: string, body?: object | string | ArrayBuffer, headers?: object): Promise<IApiResult>;
    /**
     * Sends a DELETE request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    spDelete(url: string, body?: object, headers?: object): Promise<IApiResult>;
    /**
     * Sends a PATCH request (MERGE method) to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    spPatch(url: string, body?: object | string, headers?: object): Promise<IApiResult>;
    /**
     * Sends a PUT request to a specified SharePoint API endpoint with the given request body and headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param body The body of the request, which can be an object or string (default is an empty object).
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
     */
    spPut(url: string, body?: object | string, headers?: object): Promise<IApiResult>;
    /**
     * Sends a GET request to a specified SharePoint API endpoint with the given headers.
     * @param url The URL of the SharePoint API endpoint.
     * @param headers Additional headers to include in the request (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request, any response data, or errors.
     */
    spGet(url: string, headers?: object, logError?: boolean, responseType?: "json" | "text"): Promise<IApiResult>;
    /**
     * Sends a GET request to retrieve binary data (as an ArrayBuffer) from a specified SharePoint API URL.
     * @param url The API endpoint to fetch the binary content from.
     * @param headers (Optional) Additional headers to include in the request.
     * @returns A Promise resolving to an ArrayBuffer containing the binary data, or throws an error if the request fails.
     */
    spGetBinary(url: string, headers?: object): Promise<IApiResult>;
    /**
     * Sends a POST request to upload binary data (as an ArrayBuffer) to a specified SharePoint API URL.
     * @param url The API endpoint to which the binary content should be uploaded.
     * @param fileContent The binary content to be uploaded, provided as an ArrayBuffer.
     * @returns A Promise that resolves when the upload is successful, or throws an error if the request fails.
     */
    spPostBinary(url: string, fileContent: ArrayBuffer): Promise<IApiResult>;
    private getGraphClient;
    /**
     * Sends a GET request to the Microsoft Graph API and retrieves data.
     * @param url The URL of the Microsoft Graph API endpoint.
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
     */
    graphGet(url: string, attempt?: number): Promise<IApiResult>;
    /**
     * Sends a POST request to the Microsoft Graph API with a specified body.
     * @param url The URL of the Microsoft Graph API endpoint.
     * @param body The body of the POST request, which is an object (default is an empty object).
     * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
     */
    graphPost(url: string, body?: object, attempt?: number): Promise<IApiResult>;
    graphDelete(url: string, attempt?: number): Promise<IApiResult>;
}

declare class SPApiClient extends BaseApiClient implements ISPApiClient {
    internalSiteName: string;
    templateSiteName: string;
    protected channelId: string;
    protected teamId: string;
    protected absoluteUrl: string;
    protected relativeUrl: string;
    protected userRole: number;
    protected channelName: string;
    protected teamName: string;
    protected channelType: ChannelType;
    protected readonly documentFolder: string;
    private listMappings;
    protected libraryName: string;
    sp: SPFI;
    constructor(webPartContext: WebPartContext, internalSiteName: string, templateSiteName: string);
    /**
     * Checks existence of all required SharePoint lists and libraries.
     * @returns Object mapping each list/library name to a boolean indicating existence.
     */
    checkAllListExistence(): Promise<{
        [key in ListName | LibraryName.Requests]: boolean;
    }>;
    /**
     * Ensures a SharePoint list exists, creates it if missing, and ensures required columns and views.
     * @param listName The name of the list or library.
     */
    ensureList(listName: ListName | LibraryName.Requests): Promise<void>;
    /**
     * Deletes a SharePoint list and all its items and folders.
     * @param listName The name of the list to delete.
     * @returns Promise resolving to true if deleted successfully.
     */
    deleteList(listName: ListName): Promise<boolean>;
    /**
     * Updates a SharePoint list with the provided body.
     * @param body Update payload.
     * @param listName The name of the list (defaults to libraryName).
     */
    updateList(body: object, listName?: string): Promise<void>;
    /**
     * Updates a column in a SharePoint list.
     * @param body Update payload.
     * @param listName The name of the list.
     * @param columnName The column to update.
     */
    updateColumn(body: object, listName: string, columnName: string): Promise<void>;
    /**
     * Updates an item in a SharePoint list by unique ID.
     * @param uniqueId The unique identifier of the item.
     * @param body Update payload.
     * @param listName The name of the list (defaults to libraryName).
     * @returns Promise resolving to true if updated successfully.
     */
    updateItemInList(uniqueId: string, body?: object, listName?: string): Promise<boolean>;
    /**
     * Deletes an item from the specified SharePoint list by its ID.
     *
     * @param id - The ID of the item to delete.
     * @param listName - The name of the SharePoint list.
     * @returns A promise resolving to true if the item was successfully deleted, otherwise false.
     */
    deleteListItem(id: number, listName: ListName): Promise<boolean>;
    /**
     * Updates a list item by item ID.
     * @param itemId The item ID.
     * @param body Update payload.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if updated successfully.
     */
    updateListItem(itemId: number, body: object | undefined, listName: ListName | LibraryName): Promise<boolean>;
    /**
     * Retrieves choice options for a field in a SharePoint list.
     * @param listName The name of the list.
     * @param columnTitle The column title.
     * @returns Array of choice strings.
     */
    getFieldChoices(listName: ListName, columnTitle: string): Promise<string[]>;
    /**
     * Creates a list item using a specified path in SharePoint.
     * @param body Item creation payload.
     * @param listName The name of the list.
     * @param path Optional path for item placement.
     * @param channelSpecific Whether to use channel-specific path.
     * @returns API result of the operation.
     */
    createListItemUsingPath(body: object | undefined, listName: ListName, path?: string, channelSpecific?: boolean): Promise<IApiResult>;
    /**
     * Updates multiple list items in batch.
     * @param updates Array of update objects with uniqueId and body.
     * @param listName The name of the list.
     */
    updateListItemsBatch(updates: {
        uniqueId: string;
        body: object;
    }[], listName: ListName): Promise<void>;
    /**
     * Splits an array into chunks of specified size.
     * @param arr The array to chunk.
     * @param size The chunk size.
     * @returns Array of chunked arrays.
     */
    chunkArray<T>(arr: T[], size: number): T[][];
    /**
     * Sends a batch update chunk for list items.
     * @param listName The name of the list.
     * @param updates Array of update objects.
     * @returns API result of the batch operation.
     */
    private sendItemsUpdateBatchChunk;
    /**
     * Checks if a SharePoint list exists.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if exists.
     */
    checkListExistence(listName: ListName | LibraryName): Promise<boolean>;
    /**
     * Creates a SharePoint list or library.
     * @param listName The name of the list or library.
     * @returns Promise resolving to true if created successfully.
     */
    createList(listName: ListName | LibraryName.Requests): Promise<boolean>;
    /**
     * Ensures required columns exist in a SharePoint list.
     * @param listName The name of the list or library.
     * @returns Array of missing columns.
     */
    ensureColumns(listName: ListName | LibraryName.Requests): Promise<ISPColumn[]>;
    /**
     * Ensures the default list view contains required columns.
     * @param listName The name of the list.
     * @param columns Array of columns to ensure in the view.
     */
    ensureListView(listName: ListName, columns: ISPColumn[]): Promise<void>;
    /**
     * Retrieves the default view ID for a SharePoint list.
     * @param listName The name of the list.
     * @returns The default view ID as a string.
     */
    private getDefaultListViewId;
    /**
     * Creates a column in a SharePoint list.
     * @param listName The name of the list or library.
     * @param column Column definition object.
     * @returns Promise resolving to true if created successfully.
     */
    private createColumn;
    /**
     * Creates an item in a SharePoint list.
     * @param body Item creation payload.
     * @param listName The name of the list or library.
     * @returns API result of the operation.
     */
    createListItem(body: object | undefined, listName: ListName | LibraryName): Promise<IApiResult>;
    /**
     * Ensures a user exists in SharePoint and returns their ID.
     * @param userEmail The user's email address.
     * @param webUrl Optional web URL (defaults to absoluteUrl).
     * @returns The SharePoint user ID.
     */
    getUserSPId(userEmail: string, webUrl?: string): Promise<number>;
    init(ctx?: string, teamId?: string, channelId?: string): Promise<void>;
    initWithCtx(ctx: string): Promise<void>;
    initWithIds(teamId: string, channelId: string): Promise<void>;
    private upgradeLegacyUrl;
    private getChannelDetail;
    private getChannelDisplayName;
    private getTeamSiteUrl;
    /**
     * Gets the current user's role in the channel.
     * @returns The user role as a number.
     */
    getUserRole(): number;
    /**
     * Gets the current team ID.
     * @returns The team ID as a string.
     */
    getTeamId(): string;
    /**
     * Gets the current channel ID.
     * @returns The channel ID as a string.
     */
    getChannelId(): string;
    /**
     * Gets the absolute URL for the current site.
     * @returns The absolute URL as a string.
     */
    getAbsoluteUrl(): string;
    /**
     * Gets the current channel name.
     * @returns The channel name as a string.
     */
    getChannelName(): string;
    /**
     * Gets the relative URL for the current site.
     * @returns The relative URL as a string.
     */
    getRelativeUrl(): string;
    private getChannelType;
    isPrivateChannel(): Promise<boolean>;
    /**
     * Gets the current team name, loading it if not already set.
     * @returns Promise resolving to the team name.
     */
    getTeamName(): Promise<string>;
    getUserEmail(): string;
    /**
     * Gets the user's role in a Teams channel.
     * @param teamId The team ID.
     * @param channelId The channel ID.
     * @returns Role as a number.
     */
    private getUserChannelRole;
    /**
     * Loads the team name from Microsoft Graph if not already set.
     * @returns Promise resolving to the team name.
     */
    protected loadTeamName(): Promise<string>;
    /**
     * Ensures a folder exists at the specified path, creating it if necessary.
     * @param path The server-relative path.
     * @param name The folder name.
     * @returns UniqueId of the folder or undefined.
     */
    ensureFolder(path: string, name: string): Promise<string | undefined>;
    /**
     * Creates multiple list items in batch.
     * @param itemBodies Array of item creation payloads.
     * @param listName The name of the list or library.
     */
    createListItemsBatch(itemBodies: object[], listName: ListName | LibraryName): Promise<void>;
    /**
     * Sends a batch chunk for creating list items.
     * @param listName The name of the list or library.
     * @param itemBodies Array of item creation payloads.
     * @returns API result of the batch operation.
     */
    private sendItemsPostBatchChunk;
    /**
     * Moves a list item to a new folder within the document structure.
     * @param id The item ID.
     * @param fileDirRef The target folder's file directory reference.
     * @param listName The name of the list.
     */
    moveListItem(id: number, fileDirRef: string, listName: ListName): Promise<void>;
    /**
     * Generates a batch string for setting permissions.
     * @param boundary Batch boundary string.
     * @param groupId Group ID.
     * @param roledefid Role definition ID.
     * @param apiUrlBase Base API URL.
     * @returns Batch string for permissions.
     */
    getPermissionBatchString(boundary: string, groupId: number, roledefid: number, apiUrlBase: string): string;
    /**
     * Retrieves site groups from the current SharePoint site.
     * @param webUrl Optional web URL.
     * @returns Array of site group objects.
     */
    getSiteGroups(webUrl?: string): Promise<ISiteGroup[]>;
    getSiteGroupByName(groupName: string, webUrl?: string): Promise<ISiteGroup | undefined>;
    /**
     * Attaches a file to a list item.
     * @param id The item ID.
     * @param file The file to attach.
     * @param listName The name of the list.
     * @returns API result of the operation.
     */
    addAttachmentFile(id: number, file: FileInfo$1, listName: ListName): Promise<IApiResult>;
    /**
     * Attaches an array buffer as a file to a list item, optionally skipping if file exists.
     * @param id The item ID.
     * @param fileName The file name.
     * @param buffer The file content as ArrayBuffer.
     * @param listName The name of the list.
     * @param skipIfExists Whether to skip if file exists.
     * @returns API result of the operation.
     */
    addAttachmentArrayBuffer(id: number, fileName: string, buffer: ArrayBuffer, listName: ListName, skipIfExists?: boolean): Promise<IApiResult>;
    /**
     * Checks if an attachment exists for a list item.
     * @param itemId The item ID.
     * @param fileName The file name.
     * @param listName The name of the list.
     * @returns Promise resolving to true if exists.
     */
    private checkAttachmentExists;
    /**
     * Retrieves template attachments from a SharePoint list.
     * @param listName The name of the list.
     * @returns Array of attachment objects.
     */
    getTemplatesAttachments(listName: ListName): Promise<IAttachment[]>;
    /**
     * Ensures a folder exists in a SharePoint list, creating and moving it if necessary.
     * @param name The folder name.
     * @param listName The name of the list or library.
     * @param path Optional path for folder placement.
     */
    ensureListFolder(name: string, listName: ListName | LibraryName, path?: string): Promise<void>;
    /**
     * Moves a folder to a new path in SharePoint.
     * @param path The current folder path.
     * @param newPath The new folder path.
     * @returns Promise resolving to true if moved successfully.
     */
    moveFolder(path: string, newPath: string): Promise<boolean>;
    /**
     * Logs a message and data to the SharePoint Logs list.
     * @param message The log message.
     * @param data Optional log data.
     */
    log(message: string, data?: any): Promise<void>;
    /**
     * Hides specified SharePoint lists.
     * @param listNames Array of list names to hide.
     */
    hideLists(listNames: ListName[]): Promise<void>;
}

declare class ActionApiClient implements IActionApiClient {
    private readonly sp;
    /**
     * Constructs an ActionApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    /**
     * Retrieves workflow actions from the SharePoint WorkflowActions list.
     * Filters items by content type and maps them to structured workflow action objects.
     * @returns A Promise resolving to an array of IWorkflowAction objects, or an empty array if no data is found.
     */
    getWorkflowActions(fromTemplates?: boolean): Promise<IWorkflowAction[]>;
    /**
     * Retrieves scheduled actions from the SharePoint ActionToTrigger list.
     * Filters items by folder and processed status, then maps them to structured scheduled action objects.
     * @returns A Promise resolving to a Map of scheduled actions, or an empty Map if no data is found.
     */
    getScheduledActions(): Promise<Map<number, IScheduledAction>>;
    getScheduledActionByFileId(fileId?: number): Promise<undefined | IScheduledAction>;
    /**
     * Creates a new scheduled action in the ActionToTrigger list.
     * @param action The scheduled action object to create.
     * @param title The title for the scheduled action.
     * @returns Promise resolving to true if creation succeeded, otherwise false.
     */
    createScheduledAction(action: IScheduledAction, title: string, comment?: string): Promise<boolean>;
}

declare class CalendarApiClient implements ICalendarApiClient {
    private readonly sp;
    /**
     * Constructs a CalendarApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    /**
     * Retrieves calendar events from the Duties SharePoint list.
     * Filters items by folder and content type, then maps them to structured calendar event objects.
     * @returns Promise resolving to an array of ICalendarEvent objects, or an empty array if no events are found.
     */
    getCalendarEvents(): Promise<ICalendarEvent[]>;
    /**
     * Creates a new calendar event item in the Duties SharePoint list.
     * Normalizes event dates to noon and formats as ISO strings before saving.
     * @param event The calendar event data to create.
     * @returns Promise resolving to true if creation succeeded, otherwise false.
     */
    createCalendarEvent(event: ICalendarEvent): Promise<boolean>;
}

declare class ChatApiClient implements IChatApiClient {
    private readonly sp;
    /**
     * Constructs a ChatApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    getMessages(): Promise<any[]>;
}

declare class DetailApiClient implements IDetailApiClient {
    private readonly sp;
    private documentFolder;
    /**
     * Constructs a DetailApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    /**
         * create detail of loan when installing
         * @param channelId
         * @param title
         * @param currentStatus
         */
    createDetail(currentStatus: string, version: number): Promise<boolean>;
    /**
     * @param channelId
     * @returns detail {@link ILoanDetail}
     */
    getLoanDetail(): Promise<ILoanDetail | undefined>;
    /**
   * updates detail using {@link updateListItem}
   * @param itemId
   * @param body
   * @returns boolean
   */
    updateDetail(itemId: number, body: object): Promise<boolean>;
    /**
     * ensure user and update detail using {@link updateListItem}
     * @param itemId
     * @param email
     * @returns boolean
     */
    updateContact(itemId: number, contactData: string, secondaryData: string, email?: string, secondaryEmail?: string): Promise<boolean>;
    private getLoanDetailContacts;
    private getIUPLink;
}

declare class FileApiClient implements IFileApiClient {
    private readonly spClient;
    protected readonly documentFolder: string;
    private readonly trashPath;
    /**
     * Constructs a FileApiClient instance.
     * @param spClient The SPApiClient instance for SharePoint operations.
     */
    constructor(spClient: SPApiClient);
    /**
     * Moves a file to the bin (trash) folder.
     * @param fileRef The file reference (server-relative URL).
     * @param fileName The file name.
     * @param uniqueId Optional unique ID for the file.
     * @returns Promise resolving to true if moved successfully.
     */
    moveFileToBin(fileRef: string, fileName: string): Promise<boolean>;
    /**
    * Loads files from a SharePoint document library filtered by a structure ID,
    * excluding deleted files and those in the trash folder
    * @param id - The structure ID used to filter files
    * @returns A promise resolving to an array of IFile objects matching the filter criteria
    */
    loadFiles(id: number): Promise<IFile[]>;
    /**
     * Loads files from SharePoint by their item IDs.
     * @param ids Array of item IDs.
     * @returns Promise resolving to an array of IFile objects.
     */
    loadFilesByIds(ids: number[]): Promise<IFile[]>;
    /**
     * Loads files from SharePoint by their unique IDs.
     * @param uniqueIds Array of unique IDs.
     * @returns Promise resolving to an array of IFile objects.
     */
    loadFilesByUniqueIds(uniqueIds: string[]): Promise<IFile[]>;
    /**
     * Loads a file from SharePoint by its server-relative URL.
     * @param serverRelativeUrl The server-relative URL of the file.
     * @returns Promise resolving to an IFile object or undefined.
     */
    loadFileByRelativeUrl(serverRelativeUrl: string): Promise<IFile | undefined>;
    /**
     * Loads a file from SharePoint by its unique ID (private).
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to an IFile object or undefined.
     */
    private loadFileByUniqueId;
    /**
    * Loads all files from a channel folder path within the SharePoint document library,
    * optionally filtering by documentation flag, excluding deleted files and those in the trash folder
    * @param documentation - If true, filters files marked as documentation
    * @returns A promise resolving to an array of IFile objects matching the criteria
    */
    loadAllFiles(documentation: boolean): Promise<IFile[]>;
    /**
     * publishs file by using its uniqueId
     * @param uniqueId
     * @returns boolean
     */
    publishFile(uniqueId: string): Promise<boolean>;
    private assignFilePermissions;
    setFilePermissions(action: IWorkflowAction, siteGroups: ISiteGroup[], uniqueId: string): Promise<void>;
    /**
     * upload array of files using serverRelativeUrl of folder
     * @param folderPath
     * @param files
     * @returns array of fileNames whose upload failed
     */
    uploadFiles(files: IFileInput[], secondStage: boolean, uploadedBy?: string, folderId?: number): Promise<IFileUploadResult[]>;
    /**
    * Retrieves the list of available file types from the SharePoint 'DocumentTypes' list
    * @returns A promise resolving to an array of IFileType objects
    */
    getFileTypes(webUrl?: string): Promise<IFileType[]>;
    /**
     * Undoes the check-out of a file by unique ID.
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to true if successful.
     */
    undoFileCheckOut(uniqueId: string): Promise<boolean>;
    /**
    * Retrieves a single file by its unique ID, including its document type details if available
    * @param uniqueId - The GUID of the file to retrieve
    * @returns A promise resolving to the file object or undefined if not found
    */
    getFileByUniqueId(uniqueId: string): Promise<IFile | undefined>;
    /**
    * Retrieves the version history for a file based on its list item ID.
    * @param id - The list item ID of the file.
    * @returns A promise resolving to an array of file version objects.
    */
    getFileVersionHistory(id: number): Promise<IFileVersion[]>;
    /**
     * Retrieves a specific version of a file as an ArrayBuffer.
     * @param relativeUrl The server-relative URL of the file.
     * @param versionId The version ID.
     * @returns Promise resolving to the file content as ArrayBuffer.
     */
    getFileVersion(relativeUrl: string, versionId: number): Promise<ArrayBuffer | undefined>;
    /**
     * Retrieves a file for download as an ArrayBuffer.
     * @param relativeUrl The server-relative URL of the file.
     * @returns Promise resolving to the file content as ArrayBuffer.
     */
    getFileToDownload(relativeUrl: string): Promise<ArrayBuffer | undefined>;
    /**
     * Archives a file to a designated location, updating metadata.
     * @param file The file object to archive.
     * @param stage Optional stage node.
     * @param phase Optional phase node.
     * @returns Promise resolving to true if archived successfully.
     */
    archiveFile(file: IFile, stage?: IStructureNode, phase?: IStructureNode): Promise<boolean>;
    /**
     * Retrieves the archived file reference by unique ID (private).
     * @param uniqueId The unique ID of the file.
     * @returns Promise resolving to the file reference string or undefined.
     */
    private getArchivedFile;
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
    moveFile(fileRef: string, fileName: string, newFileDirRef: string): Promise<boolean>;
    /**
     * Updates a file's metadata by unique ID.
     * @param uniqueId The unique ID of the file.
     * @param body Metadata update payload.
     * @returns API result object.
     */
    updateFileByUniqueId(uniqueId: string, body?: object): Promise<IApiResult>;
    /**
     * Updates a file's metadata by server-relative URL.
     * @param fileRef The server-relative URL of the file.
     * @param body Metadata update payload.
     * @returns API result object.
     */
    updateFileByUrl(fileRef: string, body?: object): Promise<IApiResult>;
    private resetAndBreakInheritance;
    private clearRoleAssignments;
    /**
    * upload file using serverRelativeUrl of folder
    * @param folderPath
    * @param file
    * @returns fileName if upload fail
    */
    uploadFile(folderPath: string, file: FileInfo, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    private resetFileMetadata;
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
    reuploadFile(folderPath: string, file: FileInfo, fileName: string, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    /**
     * Retrieves required files for a folder by ID.
     * @param id The folder ID.
     * @returns Array of required files.
     */
    getRequiredFiles(id: number): Promise<IRequiredFile[]>;
    /**
     * Retrieves all required files for the current channel, grouped by folder number.
     * @returns Map of folder number to array of required files.
     */
    getAllRequiredFiles(): Promise<Map<number, IRequiredFile[]>>;
    /**
     * Generates a unique file name in a folder, appending a numeric suffix if needed (private).
     * @param folderPath The folder path.
     * @param fileName The file name.
     * @param countStart Optional starting count.
     * @param siteUrl Optional site URL.
     * @returns Unique file name string.
     */
    private getUniqueFileName;
    /**
     * Checks if a file exists at a given path (private).
     * @param path The file path.
     * @param siteUrl Optional site URL.
     * @returns Promise resolving to true if exists.
     */
    private checkFileExists;
    /**
     * Retrieves folders from the document library filtered by the current channel and site context.
     * The folder name adapts based on the site URL (supports localization).
     * Filters items that start with the specified folder path and are of folder content type.
     *
     * @returns A promise that resolves to an array of folders (`IFolder[]`), or an empty array if none found.
     */
    getFolders(): Promise<IFolder[]>;
    sendFiletoExtraction(file: IFile, type: FinancialDoc): Promise<boolean>;
}

declare class FormApiClient implements IFormApiClient {
    private readonly sp;
    /**
     * Constructs a FormApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    /**
     * Retrieves form records from the SharePoint list.
     * @returns Promise<IFormRecord[]> - A promise that resolves to an array of form records
     */
    getFormRecords(): Promise<IFormRecord[]>;
    /**
     * Saves the form data to a SharePoint list and updates its status.
     * @param id - The ID of the form record to update
     * @param formData - The form data to be saved
     * @param asDraft - A flag indicating whether the form data should be saved as a draft (optional, defaults to false)
     * @returns Promise<boolean> - A promise that resolves to true if the data is successfully saved, otherwise false
     */
    saveFormData(id: number, formData: FieldValues, language: Language, asDraft?: boolean): Promise<boolean>;
}

declare class InstallApiClient implements IInstallApiClient {
    private readonly sp;
    protected documentFolder: string;
    /**
     * Constructs an InstallApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(sp: SPApiClient);
    /**
 * Initializes request-related folders and copies request templates.
 *
 * Creates a request folder for the current channel and an 'options' folder,
 * calling the provided progress callback after each step.
 * Finally, copies the request templates.
 *
 * @returns Promise<void> - A promise that resolves when all initialization steps are complete.
 */
    initRequests(): Promise<void>;
    ensureIUPData(): Promise<void>;
    updateIUPChannel(body: object): Promise<void>;
    updateIUPChannelContacts(email?: string, secondaryEmail?: string): Promise<void>;
    updateIUPGroupColumns(oaoEmails: string[], uuoEmails: string[], amlEmails: string[]): Promise<void>;
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
    moveFolderAndChangeTitle(path: string, newPath: string, newTitle: string, listName: ListName): Promise<void>;
    ensureFolderForAllChannels(basePath: string, name: string): Promise<void>;
    /**
     * ensures document library and then creates folders
     * @param folderTemplates
     */
    createFoldersFromTemplates(templateItems: ITemplateContent[]): Promise<ITemplateItemRelation[]>;
    /**
         * Updates the 'RequiredDocuments' links for all category items in the current channel,
         * ensuring the description contains the correct relative path.
         *
         * @returns A promise that resolves when all category items have been updated.
         */
    updateLinksInCategories(): Promise<void>;
    /**
         * Creates a new category item in the SharePoint list for the current channel,
         * moves the item to the appropriate folder path, and updates its required documents link.
         *
         * @param templateItem - The template content containing category details.
         * @returns The ID of the created item or undefined if creation failed.
         */
    createCategoryItem(templateItem: ITemplateContent): Promise<number | undefined>;
    updateLoanDetailsVersion(version: number): Promise<boolean>;
    /**
     * Retrieves templates from the Templates list based on the specified file reference.
     * @param fileRef The file reference path to filter templates.
     * @returns An array of template content objects.
     */
    getTemplates(fileRef: string): Promise<ITemplateContent[]>;
    /**
     * Retrieves the names of templates from the Templates list.
     * @returns An array of template reference objects containing file reference and leaf file reference.
     */
    getTemplatesNames(): Promise<ITemplateReference[]>;
    /**
         * Ensures the existence of necessary folders for different lists under the current channel,
         * invoking a progress increment callback after each folder is created or confirmed.
         *
         * @returns Promise that resolves when all folders have been ensured.
         */
    createListFolders(): Promise<void>;
    setListReadOnly(listName: ListName): Promise<void>;
    ensureListFolder(name: string, listName: ListName | LibraryName, path?: string): Promise<void>;
    /**
         * check existence of all necessary columns in document library and create them if needed
         */
    ensureDocumentLibrary(selectedTemplate: string): Promise<void>;
    /**
         * Adds required files for categories under the current channel and copies template-related files.
         * Ensures each category folder exists in the RequiredDocuments list and copies files based on relations.
         * Calls the provided progress increment callback once the process completes.
         * @param relations - Array of template item relations to copy required files for.
         * @param templateName - Name of the template used for copying files.
         * @returns Promise that resolves when all required files are processed.
         */
    addRequiredFiles(relations: ITemplateItemRelation[], templateName: string): Promise<void>;
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
    ensureListAndCopyItems(listName: ListName, templatesListName: TemplatesListName): Promise<void>;
    /**
         * Copies folders and items from a source SharePoint list to a target list while preserving the folder structure.
         *
         * @param sourceListName - The name of the source list from which folders and items will be copied.
         * @param targetListName - The name of the target list where folders and items will be created.
         * @param fields - An array of field names to copy from the source items to the target items.
         * @returns Promise<void> - A promise that resolves when the copying process is complete.
         */
    copyFoldersAndItemsWithStructure(sourceListName: TemplatesListName, targetListName: ListName, fields: string[]): Promise<void>;
    ensureListAndCopyItemsIntoChannelFolder(listName: ListName, templatesListName: TemplatesListName, templateName: string): Promise<void>;
    copyTemplateAttachments(templatesListName: TemplatesListName, listName: ListName): Promise<void>;
    /**
     * Ensures a folder with the specified name exists within the 'Requests' library.
     *
     * @param name - The name of the folder to create or verify.
     * @returns Promise<boolean> - A promise that resolves to true if the folder exists or was created successfully, otherwise false.
     */
    private createRequestFolder;
    private createListItemsBatchUsingPath;
    private createListItemsBatchChunkUsingPath;
    private checkCopiedItemsExistence;
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
    private copyRequiredfiles;
    /**
 * Ensures that all necessary folders based on the selected template exist in the document library.
 * Calls folder creation for each folder template and updates progress via callback.
 *
 * @param selectedTemplate - The name or identifier of the selected folder template.
 */
    private ensureFoldersInDocumentLibrary;
    /**
     * Retrieves folder templates for the given selected template from a specific team's document library.
     *
     * @param selectedTemplate - The name of the selected folder template to retrieve.
     * @returns Promise resolving to an array of folder templates ({@link IFolderTemplate}).
     */
    private getFolderTemplates;
    /**
         * recursively generate folders from templates loaded from site EUP-Sablony
         * @param folderTemplates
         * @param parentFolderPath
         */
    private generateFoldersFromTemplates;
    private setFolderPermissions;
    private assignFolderPermissions;
    /**
 * creates new subfolder with custom attributes in folder using serverRelativeUrl
 * @param folder
 * @param parentFolderPath
 * @returns boolean
 */
    private createFolder;
    /**
 * Creates a folder structure based on the provided template items and updates progress.
 * @param templateItems The template content items used to create the folder structure.
 * @returns A promise indicating the completion of the folder structure creation.
 */
    private createCategories;
    private getTopLevelFolders;
    /**
     * Creates the main folder in the document structure for the current channel.
     * @returns A promise indicating the completion of the folder creation.
     */
    private createMainFolderInDocumentStructure;
    private ensureIUPFolders;
    private createIUPFolders;
    private ensureIUPChannel;
    /**
    * Copies all request template files from the 'Requests' list in the source site to the target location.
    *
    * Retrieves all items including file references and file names, then calls copyRequestTemplate
    * for each item to perform the actual file copying.
    *
    * @returns Promise<void> - A promise that resolves when all request templates have been copied.
    */
    private copyRequestTemplates;
    /**
     * Copies a single request template file by downloading it from the source URL and uploading it
     * to the target 'Requests/options' folder, overwriting if the file already exists.
     *
     * @param fileRef - The server-relative URL of the source file to download.
     * @param fileLeafRef - The file name to use when uploading the file.
     * @param webUrl - The base URL of the source SharePoint site.
     * @returns Promise<void> - A promise that resolves once the file has been uploaded successfully.
     */
    private copyRequestTemplate;
    updateDateColumns(): Promise<void>;
}

declare class NotificationApiClient implements INotificationApiClient {
    private readonly sp;
    private webPartContext;
    constructor(sp: SPApiClient, webPartContext: WebPartContext);
    /**
         * Retrieves and filters notifications assigned to the specified user groups from the SharePoint Notifications list.
         * @param groups - An array of group identifiers to filter notifications by assignment.
         * @returns A Promise resolving to a sorted array of notifications relevant to the user, mapped with user-specific data.
         */
    getAllNotifications(groups: string[]): Promise<INotification[]>;
    /**
     * Creates a new notification item in the SharePoint Notifications list using provided form values.
     * @param notification - The notification data including title, task, category, assigned groups, and file reference.
     * @returns A Promise resolving to true if the item was successfully created, or false otherwise.
     */
    createNotification(notification: INewNotification): Promise<boolean>;
    /**
     * Marks a task as completed in the SharePoint Notifications list and optionally attaches a file.
     * @param id - The ID of the notification item to update.
     * @param body - An object containing the updated field values for the task.
     * @param file - (Optional) A file to be attached to the notification item.
     * @returns A Promise resolving to true if the update (and optional file attachment) was successful, or false otherwise.
     */
    completeTask(id: number, body: object, files: FileInfo$1[]): Promise<boolean>;
    deleteNotifications(notifications: INotification[]): Promise<void>;
    deleteNotificationsByTargetId(targetId: string, link?: string): Promise<void>;
    completeTasksByTargetId(targetId: string, link?: string, resultStatus?: string): Promise<void>;
    getTaskByTargetId(targetId: string): Promise<INotification | undefined>;
    /**
     * Retrieves the history of completed tasks associated with a specific file ID from the SharePoint Notifications list.
     * @param fileId - The ID of the file to filter task history by.
     * @returns A Promise resolving to a chronologically sorted array of notifications with mapped attachments and metadata.
     */
    getTaskHistory(fileId: string): Promise<INotification[]>;
    /**
 * Builds a map of attachment metadata for each item in the provided list.
 * @param items - An array of SharePoint list items to retrieve attachments for.
 * @returns A Promise resolving to a map where each key is an item ID and the value is an array of attachment objects containing file names and URLs.
 */
    private getAttachmentsMap;
}

declare class QnAApiClient implements IQnAApiClient {
    private readonly files;
    private readonly sp;
    protected documentFolder: string;
    constructor(files: FileApiClient, sp: SPApiClient);
    /**
         * Creates a new Q&A item in the QA list with the specified question, category, stage, phase, and order.
         *
         * @param question - The question text to be added.
         * @param category - The category to which the question belongs.
         * @param stage - The stage related to the question.
         * @param phase - The phase related to the question.
         * @param order - The display order of the question.
         * @returns A promise that resolves to true if the item was created successfully, otherwise false.
         */
    createQnAItem(question: string, category: string, stage: string, phase: string, priority: string, uploadedBy: string, fileType?: IFileType, order?: number, files?: IFileInput[], subCategory?: string): Promise<string | undefined>;
    /**
     * Adds a comment to a list item based on its unique ID.
     * @param uniqueId The unique identifier of the list item.
     * @param listName The name of the list where the item exists.
     * @param text The comment text to add.
     * @returns A boolean indicating whether the comment was successfully added.
     */
    addComment(qnAItem: IQnAItem, listName: string, text: string, files: IFileInput[], uploadedBy: string): Promise<{
        newComment: ISPComment | undefined;
        filesData: IQnAFilesData | undefined;
    }>;
    /**
     * Uploads Q&A files to the specified folder and updates file metadata.
     * @param files Array of files to upload.
     * @param uploadedBy User uploading the files.
     * @param qnAItemId Q&A item ID for association.
     * @param fileType File type information (optional).
     * @param commentId Comment ID for association (optional).
     * @returns Promise resolving to an array of IQnAAttachment objects.
     */
    uploadQAFiles(files: IFileInput[], uploadedBy: string, qnAItemId: number, fileType?: IFileType, commentId?: number): Promise<IQnAAttachment[]>;
    /**
     * Deletes a comment from a list item by its unique ID and comment ID.
     * @param listName The name of the list.
     * @param uniqueId The unique identifier of the list item.
     * @param commentId The ID of the comment to delete.
     * @returns Promise resolving to true if deleted successfully, otherwise false.
     */
    deleteComment(listName: ListName, uniqueId: string, commentId: number): Promise<boolean>;
    /**
 * Retrieves Q&A items from the specified folder in the QA list.
 * @param folderPAth The relative path of the folder within the QA list.
 * @returns An array of Q&A items.
 */
    getQnAItems(stage?: string, phase?: string): Promise<IQnAItem[]>;
    /**
     * Retrieves the choice options for the 'Category' field from the QnA list.
     *
     * @returns An array of category choice strings available in the QnA list.
     */
    getQnACategories(): Promise<any>;
    /**
 * Maps raw comment data to a structured array of comment objects.
 * @param data The raw data array of comments.
 * @returns An array of mapped comment objects.
 */
    private mapItemComments;
    /**
         * Maps raw Q&A data to a structured array of Q&A items, including comments for each item.
         * @param data The raw data array of Q&A items.
         * @returns A promise that resolves to an array of mapped Q&A item objects.
         */
    private mapQnAItems;
    /**
     * Maps a single raw Q&A item to a structured IQnAItem object, including comments and metadata.
     * @param item The raw Q&A item data.
     * @returns Promise resolving to a mapped IQnAItem object.
     */
    private mapQnAItem;
    /**
 * Retrieves comments for a specific item in the QA list based on its ID.
 * @param id The ID of the item to retrieve comments for.
 * @returns An array of comments associated with the item.
 */
    private getItemComments;
}

declare class RequestApiClient implements IRequestApiClient {
    private readonly files;
    private readonly sp;
    constructor(files: FileApiClient, sp: SPApiClient);
    /**
         * create request in specific folder using /AddValidateUpdateItemUsingPath endpoint and attach files to it
         * @param data
         * @param files
         * @returns boolean
         */
    createRequest(data: IRequestData, files: FileInfo$1[], fileType: IFileType, secondStage: boolean, uploadedBy: string): Promise<{
        result: IApiResult;
        failedUploads: string[];
    }>;
    /**
     * loads request from SP list from folder with name of channel
     * @returns array of {@link IRequest}
     */
    getRequests(): Promise<IRequest[]>;
    /**
     * Loads all request options from the SharePoint list specified by {@link ListName.RequestTypes}.
     * @returns A promise that resolves to an array of {@link IRequestOption} objects.
     */
    getRequestsOptions(): Promise<IRequestOption[]>;
    /**
         * update request and attach new files to it
         * @param data
         * @param itemId
         * @param files
         * @returns boolean
         */
    updateRequest(body: object, itemId: number, uploadedBy?: string, secondStage?: boolean, files?: IFileInput[], category?: IStructureNode): Promise<{
        files: IFile[];
        result: boolean;
    }>;
}

declare class StructureApiClient implements IStructureApiClient {
    private readonly sp;
    constructor(sp: SPApiClient);
    /**
     * Retrieves all folders within the document structure for the current channel.
     * @returns An array of folders.
     */
    getStructure(): Promise<IStructureNode[]>;
    /**
     * Retrieves categories from the SharePoint list filtered by the current channel's folder path.
     *
     * @returns A promise that resolves to an array of category items, or an empty array if none found.
     */
    private getCategories;
    /**
 * Retrieves phase and stage information from the QA SharePoint list filtered by the current channel's folder path.
 *
 * @returns A promise that resolves to an array of phase and stage items, or an empty array if none found.
 */
    private getPhasesFromQnAList;
}

declare class UserApiClient implements IUserApiClient {
    private readonly sp;
    readonly eupNoReplyMail: string;
    constructor(sp: SPApiClient, eupNoReplyMail: string);
    /**
         * @param groupId
         * @returns teamMembers of current team mapped to array of {@link IPerson}
         */
    getTeamMembers(): Promise<IPerson[]>;
    getExistingUser(mail: string): Promise<IPerson | undefined>;
    getExistingUsers(mail: string): Promise<IPerson[]>;
    /**
     * Retrieves the photo of a user based on their ID.
     * @param userIdOrEmail The ID or email of the user whose photo is to be fetched.
     * @returns A promise that resolves to a URL representing the user's photo.
     */
    getUserPhoto(userIdOrEmail: string): Promise<string>;
    /**
     * Retrieves the current user's SharePoint groups and filters them to only include
     * groups defined in internal and external enums.
     *
     * @returns Array of group titles from SPGroupInternal or SPGroupExternal enums the user belongs to.
     */
    getCurrentUserGroups(): Promise<(SPGroupInternal | SPGroupExternal)[]>;
    getSiteGroupUsersBatch(groupIds: number[]): Promise<Map<number, ISiteGroupUser[]>>;
    ensureRole(): Promise<void>;
    private getRoleDefinitionId;
    /**
     * Creates multiple site groups based on the provided group names.
     * @param groupNames An array of group names to be created (either internal or external).
     * @param increaseProgressCount A callback function to track progress during the group creation process.
     * @returns A promise that resolves when all site groups are created.
     */
    createSiteGroups(groupNames: (SPGroupExternal | SPGroupInternal | SPGroupRO)[]): Promise<void>;
    /**
     * Creates a single site group based on the provided group name.
     * @param groupName The name of the group to be created (either internal, external, or read-only).
     * @param increaseProgressCount A callback function to track progress after each group creation.
     * @returns A promise that resolves when the site group is created.
     */
    private createSiteGroup;
    private createSiteROGroup;
    private assingMissingRoles;
    private assignRoleToGroup;
    private getMissingRoleAssignments;
    /**
     * Retrieves all invitations from the Invitations list.
     * @returns A promise that resolves to an array of invitation objects.
     */
    getInvitations(teamMembers: IPerson[]): Promise<IInvitation[]>;
    /**
     * Adds a user to a specific SharePoint site group.
     * @param id The ID of the site group.
     * @param userEmail The email address of the user to be added.
     * @returns A promise that resolves when the user is added to the group.
     */
    addUserToSiteGroup(id: number, userEmail: string): Promise<void>;
    /**
     * Removes a user from a specific SharePoint site group.
     * @param groupId The ID of the site group.
     * @param userId The ID of the user to be removed.
     * @returns A promise that resolves when the user is removed from the group.
     */
    removeUserFromsiteGroup(groupId: number, userId: number): Promise<void>;
    private removeUserFromsiteGroupByEmail;
    /**
     * Creates a new invitation in the Invitations list.
     * @param email The email address of the person being invited.
     * @param name The name of the person being invited.
     * @param roles The roles assigned to the person being invited.
     */
    createInvitation(email: string, name: string, roles: string, language: string, existingUser?: boolean): Promise<boolean>;
    /**
     * Checks if an Active Directory user exists by their email address.
     * @param email The email address of the user.
     * @returns A Promise that resolves to true if the user exists, otherwise false.
     */
    private checkADUserExistence;
    private checkADUsersExistence;
    /**
        * checks if user has access to EUP-Sablony site
        */
    checkTemplatesAccess(): Promise<boolean>;
    /**
     * Checks if a user exists in Active Directory by email,
     * creates an invitation if the user does not exist,
     * or adds the existing user to the team.
     *
     * @param email - The email address of the user to check or invite.
     * @param name - The display name of the user for the invitation.
     */
    checkUserAndCreateInvitation(email: string, name: string, roles: string, language: string): Promise<boolean>;
    /**
     * Adds a member to a team by their email address.
     * @param email The email address of the user to add to the team.
     * @returns A promise indicating the completion of the operation.
     */
    addUserByEmailToTeamAndChannel(email: string, role?: TeamRole): Promise<boolean>;
    addEUPAdminsAsOwners(): Promise<boolean>;
    addEUPMembersAsMembers(): Promise<boolean>;
    private getGroupIdByName;
    private addUserOrGroupByIdToTeamAndChannel;
    removeUserFromTeamOrChannel(person: IExtendedPerson): Promise<boolean>;
    private addMemberToTeam;
    private addMemberToPrivateChannel;
    private fixGuestGroup;
    getUserName(userEmail: string): Promise<string>;
    getCurrentUser(): Promise<IPerson | undefined>;
}

declare class ApiClient implements IApiClient {
    readonly users: UserApiClient;
    readonly structure: StructureApiClient;
    readonly requests: RequestApiClient;
    readonly qnA: QnAApiClient;
    readonly notifications: NotificationApiClient;
    readonly install: InstallApiClient;
    readonly forms: FormApiClient;
    readonly files: FileApiClient;
    readonly detail: DetailApiClient;
    readonly calendar: CalendarApiClient;
    readonly actions: ActionApiClient;
    readonly chat: ChatApiClient;
    readonly sp: SPApiClient;
    constructor(context: WebPartContext, internalSiteName: string, templateSiteName: string, eupNoReplyMail: string);
}

declare const createApiClient: (context: WebPartContext, internalSiteName: string, templateSiteName: string, eupNoReplyMail: string, ctx?: string, queryTeamId?: string, queryChannelId?: string) => Promise<ApiClient>;

declare class InstallService {
    private apiClient;
    constructor(apiClient: ApiClient);
    private getMissingSiteGroups;
    install(increaseProgressCount: () => void, locale: Language, setProgressText: (text: string) => void, setMaximum: (max: number) => void): Promise<void>;
}

export { AppTab, ChannelStatus, ChannelType, ColumnsActions, ColumnsCategories, ColumnsDetails, ColumnsDocTypes, ColumnsDuties, ColumnsForms, ColumnsQA, ColumnsReqTypes, ColumnsRequests, ColumnsRequired, ColumnsTrigger, ContactDesc, ContextMenuAction, DataDispatchAction, DocumentTypeName, EventPeriodicity, EventType, FileActionOptions, FileStatus, FileValidityAction, Filter, FinancialDoc, FolderStatus, FormStatus, FormType, type IActionApiClient, type IApiClient, type IAttachment, type ICalendarApiClient, type ICalendarEvent, type IChannelDetail, type IChatApiClient, type ICommonProps, type IContextMenuProps, type IDetailApiClient, type IExtendedFile, type IExtendedPerson, type IFile, type IFileApiClient, type IFileInput, type IFileType, type IFileUploadResult, type IFileVersion, type IFolder, type IFolderTemplate, type IFormApiClient, type IFormField, type IFormRecord, type IInstallApiClient, type IInvitation, type IListNotificationCount, type ILoanDetail, type ILocalizedStrings, type IModalState, type INewNotification, type INotification, type INotificationApiClient, type IPerson, type IQnAApiClient, type IQnAAttachment, type IQnAFilesData, type IQnAItem, type IRequest, type IRequestApiClient, type IRequestData, type IRequestOption, type IRequiredFile, type ISPApiClient, type ISPColumn, type ISPComment, type IScheduledAction, type ISiteGroup, type ISiteGroupUser, type IStructureApiClient, type IStructureNode, type ITeamDetail, type ITemplateContent, type ITemplateItemRelation, type ITemplateReference, type IToast, type IUserApiClient, type IWorkflowAction, InstallService, Language, LibraryName, ListName, LoadingStatus, NavDispatchAction, Permissions, QACategory, QAItemType, QAPriority, QASortType, QAStatus, RequestAction, RequestFields, RequestFormFields, RequestStatus, RequestType, RoleAlias, RoleDefinition, SPGroupExternal, SPGroupInternal, SPGroupRO, SettingsMenuKey, SharedDocuments, SystemAccount, TeamName, TeamRole, TemplatesListName, ToastType, UserType, addClassToComboBoxItem, addClassToDropdownItem, classifyUploader, cleanFolderName, convertUIVersionToLabel, copyLink, createApiClient, createCategoryOptions, createUniqueFileName, delay, downloadFile, downloadFiles, expandRoleAliases, expandRoles, findNodeById, formatDateForSharePoint, getCleanBaseName, getContextMenuHeight, getFileExtension, getFileName, getFileNameWOExt, getFileTypeTitle, getInitials, getLatestDate, getLocalizationPathById, getPhaseNodeByCategoryId, getStageByCategoryId, getStageNodeByCategoryId, isFormRecord, normalizeDocType, openFileInDesktopApp, openFileInTab, parseFormData, parseNameAndDate, parseSharePointDate, splitFolderPath, transformPath };
