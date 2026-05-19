import * as React from 'react';
import { FieldValues } from 'react-hook-form';
import { FileInfo } from '@syncfusion/ej2-react-inputs';
import { MenuItemModel, MenuEventArgs } from '@syncfusion/ej2-react-navigations';
import { FileInfo as FileInfo$1 } from '@syncfusion/ej2-inputs';
import { WebPartContext } from '@microsoft/sp-webpart-base';

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

declare const createApiClient: (context: WebPartContext, internalSiteName: string, templateSiteName: string, eupNoReplyMail: string, ctx?: string, queryTeamId?: string, queryChannelId?: string) => Promise<IApiClient>;

export { AppTab, ChannelStatus, ChannelType, ColumnsActions, ColumnsCategories, ColumnsDetails, ColumnsDocTypes, ColumnsDuties, ColumnsForms, ColumnsQA, ColumnsReqTypes, ColumnsRequests, ColumnsRequired, ColumnsTrigger, ContactDesc, ContextMenuAction, DataDispatchAction, DocumentTypeName, EventPeriodicity, EventType, FileActionOptions, FileStatus, FileValidityAction, Filter, FinancialDoc, FolderStatus, FormStatus, FormType, type IActionApiClient, type IApiClient, type IAttachment, type ICalendarApiClient, type ICalendarEvent, type IChannelDetail, type IChatApiClient, type ICommonProps, type IContextMenuProps, type IDetailApiClient, type IExtendedFile, type IExtendedPerson, type IFile, type IFileApiClient, type IFileInput, type IFileType, type IFileUploadResult, type IFileVersion, type IFolder, type IFolderTemplate, type IFormApiClient, type IFormField, type IFormRecord, type IInstallApiClient, type IInvitation, type IListNotificationCount, type ILoanDetail, type ILocalizedStrings, type IModalState, type INewNotification, type INotification, type INotificationApiClient, type IPerson, type IQnAApiClient, type IQnAAttachment, type IQnAFilesData, type IQnAItem, type IRequest, type IRequestApiClient, type IRequestData, type IRequestOption, type IRequiredFile, type ISPApiClient, type ISPColumn, type ISPComment, type IScheduledAction, type ISiteGroup, type ISiteGroupUser, type IStructureApiClient, type IStructureNode, type ITeamDetail, type ITemplateContent, type ITemplateItemRelation, type ITemplateReference, type IToast, type IUserApiClient, type IWorkflowAction, Language, LibraryName, ListName, LoadingStatus, NavDispatchAction, Permissions, QACategory, QAItemType, QAPriority, QASortType, QAStatus, RequestAction, RequestFields, RequestFormFields, RequestStatus, RequestType, RoleAlias, RoleDefinition, SPGroupExternal, SPGroupInternal, SPGroupRO, SettingsMenuKey, SharedDocuments, SystemAccount, TeamName, TeamRole, TemplatesListName, ToastType, UserType, addClassToComboBoxItem, addClassToDropdownItem, classifyUploader, cleanFolderName, convertUIVersionToLabel, copyLink, createApiClient, createCategoryOptions, createUniqueFileName, delay, downloadFile, downloadFiles, expandRoleAliases, expandRoles, findNodeById, formatDateForSharePoint, getCleanBaseName, getContextMenuHeight, getFileExtension, getFileName, getFileNameWOExt, getFileTypeTitle, getInitials, getLatestDate, getLocalizationPathById, getPhaseNodeByCategoryId, getStageByCategoryId, getStageNodeByCategoryId, isFormRecord, normalizeDocType, openFileInDesktopApp, openFileInTab, parseFormData, parseNameAndDate, parseSharePointDate, splitFolderPath, transformPath };
