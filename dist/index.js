import JSZip from 'jszip';
import { z } from 'zod';
import '@pnp/sp/files';
import '@pnp/sp/webs';
import { spfi, SPFx } from '@pnp/sp';
import LZString from 'lz-string';
import '@pnp/sp/lists';
import { SPHttpClient } from '@microsoft/sp-http';

// src/enums/AppTab.ts
var AppTab = /* @__PURE__ */ ((AppTab2) => {
  AppTab2["Requests"] = "requests";
  AppTab2["Forms"] = "forms";
  AppTab2["Calendar"] = "calendar";
  AppTab2["Documentation"] = "documentation";
  AppTab2["QnA"] = "qnA";
  AppTab2["AllDocuments"] = "allDocuments";
  AppTab2["Chat"] = "chat";
  return AppTab2;
})(AppTab || {});

// src/enums/ColumnsDetails.ts
var ColumnsDetails = /* @__PURE__ */ ((ColumnsDetails2) => {
  ColumnsDetails2["Contact"] = "contact";
  ColumnsDetails2["ChannelId"] = "channelId";
  ColumnsDetails2["CurrentStatus"] = "currentStatus";
  ColumnsDetails2["LoanText"] = "loanText";
  ColumnsDetails2["SecondaryContact"] = "secondaryContact";
  ColumnsDetails2["ChannelStatus"] = "ChannelStatus";
  ColumnsDetails2["ContactData"] = "contactData";
  ColumnsDetails2["SecondaryData"] = "secondaryData";
  ColumnsDetails2["Company"] = "company";
  ColumnsDetails2["Version"] = "version";
  return ColumnsDetails2;
})(ColumnsDetails || {});

// src/enums/ColumnsForms.ts
var ColumnsForms = /* @__PURE__ */ ((ColumnsForms2) => {
  ColumnsForms2["FormData"] = "FormData";
  ColumnsForms2["TitleEn"] = "Title_en";
  ColumnsForms2["Status"] = "Status";
  ColumnsForms2["FormType"] = "FormType";
  ColumnsForms2["Category"] = "Category";
  ColumnsForms2["Language"] = "Language";
  ColumnsForms2["Completed"] = "Completed";
  ColumnsForms2["FileId"] = "FileId";
  return ColumnsForms2;
})(ColumnsForms || {});

// src/enums/ColumnsQA.ts
var ColumnsQA = /* @__PURE__ */ ((ColumnsQA2) => {
  ColumnsQA2["Answer"] = "answer";
  ColumnsQA2["Order"] = "QAOrder";
  ColumnsQA2["Answered"] = "Answered";
  ColumnsQA2["Priority"] = "Priority";
  ColumnsQA2["Category"] = "Category";
  ColumnsQA2["FilesData"] = "FilesData";
  ColumnsQA2["SubCategory"] = "SubCategory";
  return ColumnsQA2;
})(ColumnsQA || {});

// src/enums/ColumnsRequests.ts
var ColumnsRequests = /* @__PURE__ */ ((ColumnsRequests2) => {
  ColumnsRequests2["FormData"] = "FormData";
  ColumnsRequests2["RequestType"] = "RequestType";
  ColumnsRequests2["RequestStatus"] = "RequestStatus";
  ColumnsRequests2["FileTemplate"] = "FileTemplate";
  ColumnsRequests2["CategoryId"] = "categoryId";
  ColumnsRequests2["FileIds"] = "FileIds";
  ColumnsRequests2["ExportData"] = "ExportData";
  return ColumnsRequests2;
})(ColumnsRequests || {});

// src/enums/ColumnsCategories.ts
var ColumnsCategories = /* @__PURE__ */ ((ColumnsCategories2) => {
  ColumnsCategories2["Stage"] = "Stage";
  ColumnsCategories2["Phase"] = "Phase";
  ColumnsCategories2["Description"] = "FolderDescription";
  ColumnsCategories2["DescriptionEn"] = "FolderDescription_en";
  ColumnsCategories2["StageEn"] = "Stage_en";
  ColumnsCategories2["PhaseEn"] = "Phase_en";
  ColumnsCategories2["CategoryEn"] = "Category_en";
  ColumnsCategories2["RequiredDocuments"] = "RequiredDocuments";
  ColumnsCategories2["Order"] = "CategoryOrder";
  ColumnsCategories2["CategoryStatus"] = "CategoryStatus";
  return ColumnsCategories2;
})(ColumnsCategories || {});

// src/enums/ContextMenuAction.ts
var ContextMenuAction = /* @__PURE__ */ ((ContextMenuAction2) => {
  ContextMenuAction2["NewRequests"] = "newRequests";
  ContextMenuAction2["CompletedRequest"] = "completedRequests";
  ContextMenuAction2["ActualDocuments"] = "actualDocuments";
  ContextMenuAction2["AcceptedDocuments"] = "acceptedDocuments";
  return ContextMenuAction2;
})(ContextMenuAction || {});

// src/enums/FileActionOptions.ts
var FileActionOptions = /* @__PURE__ */ ((FileActionOptions2) => {
  FileActionOptions2["Delete"] = "delete";
  FileActionOptions2["Download"] = "download";
  FileActionOptions2["Open"] = "open";
  FileActionOptions2["Accept"] = "accept";
  FileActionOptions2["Return"] = "return";
  FileActionOptions2["Process"] = "process";
  FileActionOptions2["Publish"] = "publish";
  FileActionOptions2["Decline"] = "decline";
  FileActionOptions2["AddToDoc"] = "addToDoc";
  FileActionOptions2["RemoveFromDoc"] = "removeFromDoc";
  FileActionOptions2["Archive"] = "archive";
  FileActionOptions2["CopyLink"] = "copyLink";
  FileActionOptions2["Extract"] = "extract";
  return FileActionOptions2;
})(FileActionOptions || {});

// src/enums/FileStatus.ts
var FileStatus = /* @__PURE__ */ ((FileStatus2) => {
  FileStatus2["Processing"] = "processing";
  FileStatus2["Accepted"] = "accepted";
  FileStatus2["Declined"] = "declined";
  FileStatus2["Waiting"] = "waiting";
  FileStatus2["Requested"] = "requested";
  FileStatus2["Deleted"] = "deleted";
  FileStatus2["Published"] = "published";
  FileStatus2["AmlReturned"] = "aml-returned";
  FileStatus2["Provided"] = "provided";
  FileStatus2["PreparedForClient"] = "prepared-for-client";
  FileStatus2["Verification"] = "progress-verification";
  FileStatus2["Returned"] = "returned";
  FileStatus2["NotValid"] = "not-valid";
  FileStatus2["BoVerified"] = "bo-verified";
  FileStatus2["Prepared"] = "prepared";
  FileStatus2["BankRevision"] = "bank-revision";
  FileStatus2["ClientRevision"] = "client-revision";
  FileStatus2["FinalVersion"] = "final-version";
  FileStatus2["AcceptedConditional"] = "accepted-conditional";
  FileStatus2["NoType"] = "no-type";
  FileStatus2["Exported"] = "exported";
  FileStatus2["Signed"] = "signed";
  FileStatus2["ReturnedToSign"] = "returned-to-sign";
  FileStatus2["NotProvided"] = "not-provided";
  return FileStatus2;
})(FileStatus || {});

// src/enums/Filter.ts
var Filter = /* @__PURE__ */ ((Filter2) => {
  Filter2["Actual"] = "actual";
  Filter2["Accepted"] = "accepted";
  return Filter2;
})(Filter || {});

// src/enums/FolderStatus.ts
var FolderStatus = /* @__PURE__ */ ((FolderStatus2) => {
  FolderStatus2["Accepted"] = "accepted";
  FolderStatus2["Default"] = "default";
  return FolderStatus2;
})(FolderStatus || {});

// src/enums/FormStatus.ts
var FormStatus = /* @__PURE__ */ ((FormStatus2) => {
  FormStatus2["New"] = "new";
  FormStatus2["Submitted"] = "submitted";
  FormStatus2["ReturnedToClient"] = "returnedToClient";
  FormStatus2["Processing"] = "processing";
  FormStatus2["Completed"] = "completed";
  FormStatus2["Exported"] = "exported";
  FormStatus2["Signed"] = "signed";
  FormStatus2["ReturnedToSign"] = "returned-to-sign";
  return FormStatus2;
})(FormStatus || {});

// src/enums/ListName.ts
var ListName = /* @__PURE__ */ ((ListName2) => {
  ListName2["Details"] = "Detaily";
  ListName2["Requests"] = "Zadosti";
  ListName2["QA"] = "QA";
  ListName2["Categories"] = "Kategorie";
  ListName2["Duties"] = "Povinnosti";
  ListName2["Forms"] = "Formulare";
  ListName2["RequiredDocuments"] = "PozadovaneDokumenty";
  ListName2["DocumentTypes"] = "TypyDokumentu";
  ListName2["RequestTypes"] = "TypyZadosti";
  ListName2["Notifications"] = "Notification";
  ListName2["WorkflowActions"] = "WorkflowActions";
  ListName2["ActionToTrigger"] = "ActionToTrigger";
  ListName2["Logs"] = "Logs";
  return ListName2;
})(ListName || {});

// src/enums/LoadingStatus.ts
var LoadingStatus = /* @__PURE__ */ ((LoadingStatus2) => {
  LoadingStatus2["Loading"] = "loading";
  LoadingStatus2["Loaded"] = "loaded";
  LoadingStatus2["NoTeam"] = "NoTeam";
  LoadingStatus2["Preparing"] = "preparing";
  LoadingStatus2["Unauthorized"] = "unauthorized";
  LoadingStatus2["WrongVersion"] = "wrong-version";
  return LoadingStatus2;
})(LoadingStatus || {});

// src/enums/Permissions.ts
var Permissions = /* @__PURE__ */ ((Permissions2) => {
  Permissions2["Owner"] = "owner";
  Permissions2["Member"] = "member";
  Permissions2["External"] = "external";
  Permissions2["Admin"] = "admin";
  return Permissions2;
})(Permissions || {});

// src/enums/RequestStatus.ts
var RequestStatus = /* @__PURE__ */ ((RequestStatus2) => {
  RequestStatus2["New"] = "new";
  RequestStatus2["Processing"] = "processing";
  RequestStatus2["Completed"] = "completed";
  RequestStatus2["Provided"] = "provided";
  return RequestStatus2;
})(RequestStatus || {});

// src/enums/RequestType.ts
var RequestType = /* @__PURE__ */ ((RequestType2) => {
  RequestType2["Activation"] = "activation";
  RequestType2["Repayment"] = "repayment";
  RequestType2["Statement"] = "statement";
  RequestType2["Template"] = "template";
  return RequestType2;
})(RequestType || {});

// src/enums/SettingsMenuKey.ts
var SettingsMenuKey = /* @__PURE__ */ ((SettingsMenuKey2) => {
  SettingsMenuKey2["Folders"] = "folders";
  SettingsMenuKey2["Contact"] = "contact";
  SettingsMenuKey2["Templates"] = "templates";
  SettingsMenuKey2["Status"] = "status";
  SettingsMenuKey2["Install"] = "install";
  SettingsMenuKey2["Users"] = "users";
  SettingsMenuKey2["Sharepoint"] = "sharepoint";
  SettingsMenuKey2["Internal"] = "internal";
  SettingsMenuKey2["Migration"] = "migration";
  return SettingsMenuKey2;
})(SettingsMenuKey || {});

// src/enums/TeamName.ts
var TeamName = /* @__PURE__ */ ((TeamName2) => {
  TeamName2["JTFG"] = "EUP-sablony";
  TeamName2["IUP"] = "IUP";
  TeamName2["JTFGIUP"] = "Uvery-Dev";
  return TeamName2;
})(TeamName || {});

// src/enums/ToastType.ts
var ToastType = /* @__PURE__ */ ((ToastType2) => {
  ToastType2["Error"] = "error";
  ToastType2["Success"] = "success";
  return ToastType2;
})(ToastType || {});

// src/enums/ColumnsDuties.ts
var ColumnsDuties = /* @__PURE__ */ ((ColumnsDuties2) => {
  ColumnsDuties2["Date"] = "Date";
  ColumnsDuties2["EventType"] = "EventType";
  ColumnsDuties2["Amount"] = "Amount";
  ColumnsDuties2["Currency"] = "Currency";
  ColumnsDuties2["Instruction"] = "Instruction";
  ColumnsDuties2["Category"] = "Category";
  ColumnsDuties2["EventDescription"] = "EventDescription";
  ColumnsDuties2["Note"] = "Note";
  ColumnsDuties2["Period"] = "Period";
  ColumnsDuties2["EndDate"] = "EndDate";
  return ColumnsDuties2;
})(ColumnsDuties || {});

// src/enums/UserType.ts
var UserType = /* @__PURE__ */ ((UserType2) => {
  UserType2["External"] = "03external";
  UserType2["Internal"] = "02internal";
  UserType2["Invited"] = "01invited";
  return UserType2;
})(UserType || {});

// src/enums/Language.ts
var Language = /* @__PURE__ */ ((Language4) => {
  Language4["CS"] = "cs-CZ";
  Language4["EN"] = "en-US";
  return Language4;
})(Language || {});

// src/enums/SPGroupInternal.ts
var SPGroupInternal = /* @__PURE__ */ ((SPGroupInternal2) => {
  SPGroupInternal2["UUO_RM"] = "Relationship manager";
  SPGroupInternal2["UUO_RSMA"] = "RSM Analyst";
  SPGroupInternal2["UUO_BO"] = "Back office";
  SPGroupInternal2["UUO_DEP"] = "UUO Deputy";
  SPGroupInternal2["UUO_DIR"] = "UUO Director";
  SPGroupInternal2["ORKR_RA"] = "Risk analyst";
  SPGroupInternal2["ORKR_GCCA"] = "GCC analyst";
  SPGroupInternal2["ORKR_RD"] = "Risk deputy";
  SPGroupInternal2["ORKR_RAPP"] = "Risk director";
  SPGroupInternal2["OAO_CC"] = "Credit controller";
  SPGroupInternal2["OAO_LAW"] = "Lawyer";
  SPGroupInternal2["OAO_DEP"] = "OAO Deputy";
  SPGroupInternal2["OAO_DIR"] = "OAO Director";
  SPGroupInternal2["WO_S"] = "WorkOut Specialist";
  SPGroupInternal2["WO_DEP"] = "WorkOut Deputy";
  SPGroupInternal2["WO_DIR"] = "WorkOut Director";
  SPGroupInternal2["AML_S"] = "AML Specialist";
  SPGroupInternal2["AML_DIR"] = "AML Director";
  SPGroupInternal2["APPR"] = "Approver";
  SPGroupInternal2["INT_AUD"] = "Intern\xED Auditor";
  SPGroupInternal2["BUS_ADMIN"] = "Business Admin";
  SPGroupInternal2["ADMIN"] = "Admin";
  return SPGroupInternal2;
})(SPGroupInternal || {});

// src/enums/SPGroupExternal.ts
var SPGroupExternal = /* @__PURE__ */ ((SPGroupExternal2) => {
  SPGroupExternal2["EXT_AUD"] = "External Auditor";
  SPGroupExternal2["EXT_ADVISER"] = "External Adviser";
  SPGroupExternal2["EXT_CLILOW"] = "Client - lower access";
  SPGroupExternal2["EXT_CLIFULL"] = "Client - full access";
  return SPGroupExternal2;
})(SPGroupExternal || {});

// src/enums/TemplatesListName.ts
var TemplatesListName = /* @__PURE__ */ ((TemplatesListName2) => {
  TemplatesListName2["Invitations"] = "Pozvanky";
  TemplatesListName2["DocumentStructure"] = "Kategorie";
  TemplatesListName2["Forms"] = "Forms";
  TemplatesListName2["DocumentTypes"] = "TypyDokumentu";
  TemplatesListName2["RequiredDocuments"] = "PozadovaneDokumenty";
  TemplatesListName2["RequestTypes"] = "TypyZadosti";
  TemplatesListName2["WorkflowActions"] = "WorkflowActions";
  TemplatesListName2["QA"] = "QA";
  TemplatesListName2["Duties"] = "Povinnosti";
  return TemplatesListName2;
})(TemplatesListName || {});

// src/enums/LibraryName.ts
var LibraryName = /* @__PURE__ */ ((LibraryName2) => {
  LibraryName2["Documents"] = "Dokumenty";
  LibraryName2["DocumentsAppGestion"] = "Documents";
  LibraryName2["Requests"] = "Requests";
  return LibraryName2;
})(LibraryName || {});

// src/enums/ColumnsRequired.ts
var ColumnsRequired = /* @__PURE__ */ ((ColumnsRequired2) => {
  ColumnsRequired2["TitleEn"] = "Title_en";
  ColumnsRequired2["FileDesc"] = "FileDescription";
  ColumnsRequired2["DocumentTypeId"] = "DocumentTypeId";
  ColumnsRequired2["DocumentType"] = "DocumentType";
  ColumnsRequired2["PermittedFormats"] = "PermittedFormats";
  ColumnsRequired2["UploadedFile"] = "uploadedFile";
  ColumnsRequired2["Deadline"] = "Deadline";
  ColumnsRequired2["Status"] = "Status";
  return ColumnsRequired2;
})(ColumnsRequired || {});

// src/enums/ColumnsDocTypes.ts
var ColumnsDocTypes = /* @__PURE__ */ ((ColumnsDocTypes2) => {
  ColumnsDocTypes2["TitleEn"] = "Title_en";
  ColumnsDocTypes2["UploadStatus"] = "UploadStatus";
  ColumnsDocTypes2["Permissions"] = "Permissions";
  ColumnsDocTypes2["FolderPath"] = "FolderPath";
  ColumnsDocTypes2["FolderPathAfter"] = "FolderPathAfter";
  ColumnsDocTypes2["ValidityStatuses"] = "ValidityStatuses";
  ColumnsDocTypes2["FolderPathInternal"] = "FolderPathInternal";
  ColumnsDocTypes2["FolderPathInternalAfter"] = "FolderPathInternalAfter";
  return ColumnsDocTypes2;
})(ColumnsDocTypes || {});

// src/enums/EventType.ts
var EventType = /* @__PURE__ */ ((EventType3) => {
  EventType3["Covenant"] = "covenant";
  EventType3["Repayment"] = "repayment";
  EventType3["Meeting"] = "meeting";
  return EventType3;
})(EventType || {});

// src/enums/RequestAction.ts
var RequestAction = /* @__PURE__ */ ((RequestAction2) => {
  RequestAction2["Accept"] = "accept";
  RequestAction2["Download"] = "download";
  RequestAction2["Open"] = "open";
  RequestAction2["DownloadAttachments"] = "DownloadAttachments";
  return RequestAction2;
})(RequestAction || {});

// src/enums/ColumnsReqTypes.ts
var ColumnsReqTypes = /* @__PURE__ */ ((ColumnsReqTypes2) => {
  ColumnsReqTypes2["TitleEn"] = "Title_en";
  ColumnsReqTypes2["FileTemplate"] = "FileTemplate";
  ColumnsReqTypes2["FileTemplateEn"] = "FileTemplateEn";
  ColumnsReqTypes2["Description"] = "Description";
  ColumnsReqTypes2["DescriptionEn"] = "Description_en";
  ColumnsReqTypes2["Stage"] = "Etapa";
  ColumnsReqTypes2["Fields"] = "Fields";
  ColumnsReqTypes2["FileAttachments"] = "FileAttachments";
  ColumnsReqTypes2["Notification"] = "Notification";
  ColumnsReqTypes2["FileReply"] = "FileReply";
  return ColumnsReqTypes2;
})(ColumnsReqTypes || {});

// src/enums/RequestFields.ts
var RequestFields = /* @__PURE__ */ ((RequestFields2) => {
  RequestFields2["Active"] = "Aktivn\xED \xFA\u010Det";
  RequestFields2["Amount"] = "\u010C\xE1stka";
  RequestFields2["Blocked"] = "Blokovan\xFD \xFA\u010Det";
  RequestFields2["ReqSpec"] = "\u017D\xE1dost";
  RequestFields2["Period"] = "Obdob\xED";
  RequestFields2["LoanSpec"] = "\xDAv\u011Br";
  RequestFields2["FeeSpec"] = "Poplatek";
  RequestFields2["Date"] = "Datum";
  RequestFields2["Currency"] = "M\u011Bna";
  RequestFields2["ActiveLoan"] = "Aktivn\xED \xFAv\u011Br";
  return RequestFields2;
})(RequestFields || {});

// src/enums/RequestFormFields.ts
var RequestFormFields = /* @__PURE__ */ ((RequestFormFields2) => {
  RequestFormFields2["Active"] = "active";
  RequestFormFields2["Amount"] = "amount";
  RequestFormFields2["Blocked"] = "blocked";
  RequestFormFields2["ReqSpec"] = "reqSpec";
  RequestFormFields2["Period"] = "period";
  RequestFormFields2["LoanSpec"] = "loanSpec";
  RequestFormFields2["FeeSpec"] = "feeSpec";
  RequestFormFields2["Date"] = "date";
  RequestFormFields2["Currency"] = "currency";
  RequestFormFields2["ActiveLoan"] = "activeLoan";
  RequestFormFields2["Comment"] = "comment";
  return RequestFormFields2;
})(RequestFormFields || {});

// src/enums/DataDispatchAction.ts
var DataDispatchAction = /* @__PURE__ */ ((DataDispatchAction2) => {
  DataDispatchAction2["SET_DETAIL"] = "SET_DETAIL";
  DataDispatchAction2["SET_PERMISSIONS"] = "SET_PERMISSIONS";
  DataDispatchAction2["SET_STRUCTURE"] = "SET_STRUCTURE";
  DataDispatchAction2["SET_TEAM_MEMBERS"] = "SET_TEAM_MEMBERS";
  DataDispatchAction2["SET_FILE_TYPES"] = "SET_FILE_TYPES";
  DataDispatchAction2["SET_SITE_GROUPS"] = "SET_SITE_GROUPS";
  DataDispatchAction2["SET_LIST_EXISTS"] = "SET_LISTS_EXISTS";
  DataDispatchAction2["SET_NOTIFICATIONS"] = "SET_NOTIFICATIONS";
  DataDispatchAction2["SET_WORKFLOW_ACTIONS"] = "SET_WORKFLOW_ACTIONS";
  DataDispatchAction2["SET_CURRENT_USER_GROUPS"] = "SET_CURRENT_USER_GROUPS";
  DataDispatchAction2["UPDATE_STRUCTURE_NODE"] = "UPDATE_STRUCTURE_NODE";
  DataDispatchAction2["ADD_FILES_TO_NODE"] = "ADD_FILES_TO_NODE";
  DataDispatchAction2["SET_FORM_RECORDS"] = "SET_FORM_RECORDS";
  DataDispatchAction2["SET_SELECTED_FORM"] = "SET_SELECTED_FORM";
  DataDispatchAction2["SET_SCHEDULED_ACTIONS"] = "SET_SCHEDULED_ACTIONS";
  DataDispatchAction2["SET_QA_ITEMS"] = "SET_QA_ITEMS";
  DataDispatchAction2["SET_REQUESTS"] = "SET_REQUESTS";
  DataDispatchAction2["ADD_FILES_TO_NODES"] = "ADD_FILES_TO_NODES";
  return DataDispatchAction2;
})(DataDispatchAction || {});

// src/enums/ColumnsActions.ts
var ColumnsActions = /* @__PURE__ */ ((ColumnsActions2) => {
  ColumnsActions2["DocumentType"] = "DocumentType";
  ColumnsActions2["DocumentStatus"] = "DocumentStatus";
  ColumnsActions2["Notification"] = "Notification";
  ColumnsActions2["ResultStatus"] = "ResultStatus";
  ColumnsActions2["Publish"] = "Publish";
  ColumnsActions2["AllowComment"] = "AllowComment";
  ColumnsActions2["ActionGroups"] = "ActionGroups";
  ColumnsActions2["Phase"] = "Phase";
  ColumnsActions2["NotificationGroups"] = "NotificationGroups";
  ColumnsActions2["Task"] = "Task";
  ColumnsActions2["TaskGroups"] = "TaskGroups";
  ColumnsActions2["TitleEn"] = "Title_en";
  ColumnsActions2["NotificationEn"] = "Notification_en";
  ColumnsActions2["TaskEn"] = "Task_en";
  ColumnsActions2["ReadPermissions"] = "ReadPermissions";
  ColumnsActions2["WritePermissions"] = "WritePermissions";
  ColumnsActions2["Validity"] = "validity";
  ColumnsActions2["Reupload"] = "Reupload";
  ColumnsActions2["FileLink"] = "FileLink";
  ColumnsActions2["OnMajorVersion"] = "onMajorVersion";
  return ColumnsActions2;
})(ColumnsActions || {});

// src/enums/FormType.ts
var FormType = /* @__PURE__ */ ((FormType2) => {
  FormType2["KYC"] = "KYC";
  FormType2["AML"] = "AML";
  FormType2["Default"] = "Default";
  return FormType2;
})(FormType || {});

// src/enums/SharedDocuments.ts
var SharedDocuments = /* @__PURE__ */ ((SharedDocuments2) => {
  SharedDocuments2["AppGestio"] = "Shared Documents";
  SharedDocuments2["JT"] = "Sdilene dokumenty";
  return SharedDocuments2;
})(SharedDocuments || {});

// src/enums/RoleDefinition.ts
var RoleDefinition = /* @__PURE__ */ ((RoleDefinition2) => {
  RoleDefinition2["Edit"] = "\xDApravy";
  RoleDefinition2["CustomEdit"] = "\xDApravy bez maz\xE1n\xED";
  RoleDefinition2["Read"] = "\u010Cten\xED";
  return RoleDefinition2;
})(RoleDefinition || {});

// src/enums/ChannelStatus.ts
var ChannelStatus = /* @__PURE__ */ ((ChannelStatus2) => {
  ChannelStatus2["Active"] = "Active";
  ChannelStatus2["Inactive"] = "Inactive";
  return ChannelStatus2;
})(ChannelStatus || {});

// src/enums/FileValidityAction.ts
var FileValidityAction = /* @__PURE__ */ ((FileValidityAction2) => {
  FileValidityAction2["Start"] = "start";
  FileValidityAction2["Update"] = "update";
  FileValidityAction2["End"] = "end";
  FileValidityAction2["StartFinal"] = "start+final";
  return FileValidityAction2;
})(FileValidityAction || {});

// src/enums/ColumnsTrigger.ts
var ColumnsTrigger = /* @__PURE__ */ ((ColumnsTrigger2) => {
  ColumnsTrigger2["FileId"] = "FileId";
  ColumnsTrigger2["ActionId"] = "ActionId";
  ColumnsTrigger2["Processed"] = "Processed";
  ColumnsTrigger2["Result"] = "Result";
  ColumnsTrigger2["Comment"] = "Comment";
  return ColumnsTrigger2;
})(ColumnsTrigger || {});

// src/enums/EventPeriodicity.ts
var EventPeriodicity = /* @__PURE__ */ ((EventPeriodicity3) => {
  EventPeriodicity3["Monthly"] = "Monthly";
  EventPeriodicity3["Quarterly"] = "Quarterly";
  EventPeriodicity3["Semiannually"] = "Semiannually";
  EventPeriodicity3["Annually"] = "Annually";
  return EventPeriodicity3;
})(EventPeriodicity || {});

// src/enums/TeamRole.ts
var TeamRole = /* @__PURE__ */ ((TeamRole2) => {
  TeamRole2["Member"] = "member";
  TeamRole2["Owner"] = "owner";
  TeamRole2["Guest"] = "guest";
  return TeamRole2;
})(TeamRole || {});

// src/enums/SystemAccount.ts
var SystemAccount = /* @__PURE__ */ ((SystemAccount2) => {
  SystemAccount2["EUPNoReply"] = "eup_noreply@jtfg.com";
  SystemAccount2["EUPNoReplyDev"] = "eup_noreply_dev@jtfg.com";
  SystemAccount2["EUPAdminsId"] = "7397659e-48e7-48fa-84f0-bd1d6f290538";
  SystemAccount2["EUPMembersId"] = "6834f589-81b2-4e47-9ec0-a41243a85aab";
  return SystemAccount2;
})(SystemAccount || {});

// src/enums/NavDispatchAction.ts
var NavDispatchAction = /* @__PURE__ */ ((NavDispatchAction2) => {
  NavDispatchAction2["SET_CURRENT_TAB"] = "SET_CURRENT_TAB";
  NavDispatchAction2["SET_ACTIVE_STEPS"] = "SET_ACTIVE_STEPS";
  NavDispatchAction2["SET_SCROLL_CATEGORY_ID"] = "SET_SCROLL_CATEGORY_ID";
  NavDispatchAction2["SET_NOTIFICATION_TARGET_ID"] = "SET_NOTIFICATION_TARGET_ID";
  NavDispatchAction2["SET_ACTIVE_STEP"] = "SET_ACTIVE_STEP";
  return NavDispatchAction2;
})(NavDispatchAction || {});

// src/enums/DocumentTypeName.ts
var DocumentTypeName = /* @__PURE__ */ ((DocumentTypeName2) => {
  DocumentTypeName2["AdministrativeExtract"] = "Administrativn\xED - V\xFDpis";
  DocumentTypeName2["AdministrativeRequest"] = "Administrativn\xED - \u017D\xE1dost";
  DocumentTypeName2["QA"] = "Q&A";
  return DocumentTypeName2;
})(DocumentTypeName || {});

// src/enums/QAPriority.ts
var QAPriority = /* @__PURE__ */ ((QAPriority2) => {
  QAPriority2["Min"] = "min";
  QAPriority2["Max"] = "max";
  QAPriority2["Normal"] = "normal";
  return QAPriority2;
})(QAPriority || {});

// src/enums/QASortType.ts
var QASortType = /* @__PURE__ */ ((QASortType2) => {
  QASortType2["Newest"] = "newest";
  QASortType2["Oldest"] = "oldest";
  QASortType2["Priority"] = "priority";
  return QASortType2;
})(QASortType || {});

// src/enums/QAItemType.ts
var QAItemType = /* @__PURE__ */ ((QAItemType2) => {
  QAItemType2["Question"] = "Question";
  QAItemType2["Answer"] = "Answer";
  QAItemType2["Comment"] = "Comment";
  return QAItemType2;
})(QAItemType || {});

// src/enums/QACategory.ts
var QACategory = /* @__PURE__ */ ((QACategory2) => {
  QACategory2["Basic"] = "Z\xE1kladn\xED data o spole\u010Dnosti";
  QACategory2["Project"] = "Projektov\xE1 dokumentace";
  QACategory2["Financial"] = "Finan\u010Dn\xED data";
  QACategory2["Legal"] = "Pr\xE1vn\xED dokumentace";
  QACategory2["Information"] = "Informa\u010Dn\xED povinnost";
  QACategory2["Request"] = "\u017D\xE1dosti na banku";
  QACategory2["Form"] = "Formul\xE1\u0159e";
  QACategory2["Other"] = "Ostatn\xED";
  return QACategory2;
})(QACategory || {});

// src/enums/QAStatus.ts
var QAStatus = /* @__PURE__ */ ((QAStatus2) => {
  QAStatus2["Completed"] = "completed";
  QAStatus2["New"] = "new";
  return QAStatus2;
})(QAStatus || {});

// src/enums/ContactDesc.ts
var ContactDesc = /* @__PURE__ */ ((ContactDesc2) => {
  ContactDesc2["Sales"] = "Sales";
  ContactDesc2["BackOffice"] = "BackOffice";
  ContactDesc2["LoanAdministration"] = "LoanAdministration";
  return ContactDesc2;
})(ContactDesc || {});

// src/enums/FinancialDoc.ts
var FinancialDoc = /* @__PURE__ */ ((FinancialDoc3) => {
  FinancialDoc3["PnL"] = "VZZ";
  FinancialDoc3["BalanceSheet"] = "Rozvaha";
  return FinancialDoc3;
})(FinancialDoc || {});

// src/enums/RoleAlias.ts
var RoleAlias = /* @__PURE__ */ ((RoleAlias2) => {
  RoleAlias2["UUO"] = "UUO";
  RoleAlias2["OAO"] = "OAO";
  RoleAlias2["AML"] = "AML";
  RoleAlias2["WO"] = "WO";
  RoleAlias2["ORKR"] = "ORKR";
  RoleAlias2["Admin"] = "admin";
  RoleAlias2["Internal"] = "internal";
  RoleAlias2["Client"] = "client";
  RoleAlias2["External"] = "external";
  return RoleAlias2;
})(RoleAlias || {});

// src/enums/SPGroupRO.ts
var SPGroupRO = /* @__PURE__ */ ((SPGroupRO2) => {
  SPGroupRO2["RORole"] = "Read Only role";
  return SPGroupRO2;
})(SPGroupRO || {});

// src/enums/ChannelType.ts
var ChannelType = /* @__PURE__ */ ((ChannelType2) => {
  ChannelType2["Standard"] = "Standard";
  ChannelType2["Private"] = "Private";
  return ChannelType2;
})(ChannelType || {});
var addClassToDropdownItem = (element, className) => {
  const arr = element.children[0] && element.children[0].children ? element.children[0].children : [];
  for (let i = 0; i < arr.length; i++) {
    const element2 = arr[i];
    element2.classList.add(className);
  }
};
var addClassToComboBoxItem = (element, className, width) => {
  element.classList.add(className);
  const content = element.querySelector(".e-content");
  if (!content) return;
  content.classList.add(className);
  const applyWidth = () => {
    if (content.children[0]) {
      const contentChildren = content.children[0].children;
      const updatedWidth = contentChildren.length <= 8 ? width + 14 : width;
      for (let i = 0; i < contentChildren.length; i++) {
        contentChildren[i].style.setProperty("width", `${updatedWidth}px`, "important");
      }
    }
  };
  requestAnimationFrame(applyWidth);
  const observer = new MutationObserver(() => {
    requestAnimationFrame(applyWidth);
  });
  observer.observe(content, { childList: true, subtree: true });
};
var createUniqueFileName = (fileName, uploadedFiles) => {
  const extensionIndex = fileName.lastIndexOf(".");
  const baseName = fileName.substring(0, extensionIndex);
  const extension = fileName.substring(extensionIndex);
  let counter = 1;
  let newFileName = fileName;
  const fileExists = (name) => uploadedFiles.some((file) => file.name === name);
  while (fileExists(newFileName)) {
    newFileName = `${baseName} (${counter})${extension}`;
    counter++;
  }
  return newFileName;
};
var getInitials = (name) => {
  return name.split(" ").map((part) => part[0]).join("");
};
var convertUIVersionToLabel = (uiVersion) => {
  const SP_UI_VERSION_MAJOR_DIVISOR = 512;
  const majorVersion = Math.floor(uiVersion / SP_UI_VERSION_MAJOR_DIVISOR);
  const minorVersion = uiVersion % SP_UI_VERSION_MAJOR_DIVISOR;
  return `${majorVersion}.${minorVersion}`;
};
var getFileExtension = (fileName) => {
  if (!fileName) {
    return "";
  }
  const parts = fileName.split(".");
  return parts && parts.length > 1 ? parts.pop() ?? "" : "";
};
var getFileNameWOExt = (fileName) => {
  const ext = getFileExtension(fileName);
  return ext ? fileName.replace(`.${ext}`, "") : fileName;
};
var getFileName = (relativeUrl) => {
  if (!relativeUrl) {
    return "";
  }
  return relativeUrl.split("/").pop() ?? "";
};
var getCleanBaseName = (fileName) => {
  const ext = getFileExtension(fileName);
  let nameWithoutExt = fileName;
  if (ext && fileName.toLowerCase().endsWith(`.${ext.toLowerCase()}`)) {
    nameWithoutExt = fileName.slice(0, -1 * (ext.length + 1));
  }
  const cleanName = nameWithoutExt.replace(/\s\(\d+\)$/, "");
  return cleanName;
};
var cleanFolderName = (folderName) => {
  if (!folderName) {
    return "";
  }
  const pattern = /^\d{2} /;
  const cleanedFolderName = folderName.replace(/^\n+/, "");
  if (pattern.test(cleanedFolderName)) {
    return cleanedFolderName.replace(pattern, "");
  }
  return cleanedFolderName;
};
var getContextMenuHeight = (items) => {
  return items.filter((item) => !item.separator).length * 36;
};
var findNodeById = (nodes, id) => {
  for (const node of nodes) {
    if (node.id === id) {
      return node;
    }
    const found = findNodeById(node.nodes, id);
    if (found) {
      return found;
    }
  }
  return void 0;
};
var getNodesByCategoryId = (nodes, id) => {
  for (const level1 of nodes) {
    for (const level2 of level1.nodes) {
      for (const level3 of level2.nodes) {
        if (level3.id === id) {
          return {
            stageNode: level1,
            phaseNode: level2,
            categoryNode: level3
          };
        }
      }
    }
  }
  return void 0;
};
var getLocalizationPathById = (nodes, id, locale) => {
  const nodesById = getNodesByCategoryId(nodes, id);
  if (nodesById) {
    return {
      stage: nodesById.stageNode.localization?.[locale] ?? "",
      phase: nodesById.phaseNode.localization?.[locale] ?? "",
      category: nodesById.categoryNode.localization?.[locale] ?? ""
    };
  }
  return void 0;
};
var getStageByCategoryId = (nodes, id) => {
  if (id) {
    const nodesById = getNodesByCategoryId(nodes, id);
    if (nodesById) {
      return nodesById.stageNode.name ?? "";
    }
  }
  return void 0;
};
var getStageNodeByCategoryId = (nodes, id) => {
  if (id) {
    const nodesById = getNodesByCategoryId(nodes, id);
    if (nodesById) {
      return nodesById.stageNode;
    }
  }
  return void 0;
};
var getPhaseNodeByCategoryId = (nodes, id) => {
  if (id) {
    const nodesById = getNodesByCategoryId(nodes, id);
    if (nodesById) {
      return nodesById.phaseNode;
    }
  }
  return void 0;
};
var parseFormData = (formData) => {
  try {
    return JSON.parse(formData);
  } catch (error) {
    console.error("Failed to parse formData:", formData, error);
    return {};
  }
};
var parseSharePointDate = (value) => {
  if (!value) {
    return void 0;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? void 0 : date;
};
var isFormRecord = (item) => {
  return "formData" in item;
};
var openFileInTab = (url) => {
  window.open(`${url}?web=1`, "_blank");
};
var getAppToOpen = (url) => {
  const ext = getFileExtension(url);
  switch (ext) {
    case "xlsx":
      return "excel";
    case "pptx":
      return "powerpoint";
    default:
      return "word";
  }
};
var openFileInDesktopApp = (url) => {
  const protocolMap = {
    word: "ms-word:ofe|u|",
    excel: "ms-excel:ofe|u|",
    powerpoint: "ms-powerpoint:ofe|u|"
  };
  const app = getAppToOpen(url);
  const protocol = protocolMap[app];
  const uri = `${protocol}${encodeURI(url)}`;
  try {
    window.location.href = uri;
  } catch (e) {
    const a = document.createElement("a");
    a.href = uri;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }
};
var getFileTypeTitle = (title) => {
  if (!title) return "";
  const parts = title.split(">").map((p) => p.trim());
  return parts.length > 1 ? parts[1] : parts[0];
};
var formatDateForSharePoint = (date) => {
  const month = date.getMonth() + 1;
  const day = date.getDate();
  const year = date.getFullYear();
  let hours = date.getHours();
  const minutes = date.getMinutes().toString().padStart(2, "0");
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12;
  if (hours === 0) hours = 12;
  return `${month}/${day}/${year} ${hours}:${minutes} ${ampm}`;
};
var GROUP_ALIAS_MAP = {
  ["UUO" /* UUO */]: ["Back office" /* UUO_BO */, "UUO Deputy" /* UUO_DEP */, "UUO Director" /* UUO_DIR */, "Relationship manager" /* UUO_RM */, "RSM Analyst" /* UUO_RSMA */],
  ["OAO" /* OAO */]: ["Credit controller" /* OAO_CC */, "OAO Deputy" /* OAO_DEP */, "OAO Director" /* OAO_DIR */, "Lawyer" /* OAO_LAW */],
  ["admin" /* Admin */]: ["Admin" /* ADMIN */, "Business Admin" /* BUS_ADMIN */],
  ["internal" /* Internal */]: Object.values(SPGroupInternal),
  ["client" /* Client */]: ["Client - full access" /* EXT_CLIFULL */, "Client - lower access" /* EXT_CLILOW */],
  ["external" /* External */]: Object.values(SPGroupExternal),
  ["AML" /* AML */]: ["AML Director" /* AML_DIR */, "AML Specialist" /* AML_S */],
  ["WO" /* WO */]: ["WorkOut Deputy" /* WO_DEP */, "WorkOut Director" /* WO_DIR */, "WorkOut Specialist" /* WO_S */],
  ["ORKR" /* ORKR */]: ["GCC analyst" /* ORKR_GCCA */, "Risk analyst" /* ORKR_RA */, "Risk director" /* ORKR_RAPP */, "Risk deputy" /* ORKR_RD */]
};
var resolveGroupAlias = (alias) => {
  return GROUP_ALIAS_MAP[alias] ?? [alias];
};
var expandRoles = (roles) => {
  const roleList = roles.split(";").map((r) => r.trim());
  const groupSet = /* @__PURE__ */ new Set();
  roleList.forEach((role) => {
    resolveGroupAlias(role).forEach((g) => groupSet.add(g));
  });
  return Array.from(groupSet);
};
var expandRoleAliases = (aliases) => {
  const groupSet = /* @__PURE__ */ new Set();
  aliases.forEach((alias) => {
    resolveGroupAlias(alias).forEach((g) => groupSet.add(g));
  });
  return Array.from(groupSet);
};
var delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
var triggerBlobDownload = (blob, name) => {
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
};
var downloadFile = async (name, serverRelativeUrl, apiClient) => {
  const arrayBuffer = await apiClient.files.getFileToDownload(serverRelativeUrl);
  if (!arrayBuffer) {
    return;
  }
  const blob = new Blob([arrayBuffer]);
  triggerBlobDownload(blob, name);
};
var classifyUploader = (currentUserGroups) => {
  if (currentUserGroups.some((g) => g === "Client - full access" /* EXT_CLIFULL */ || g === "Client - lower access" /* EXT_CLILOW */)) {
    return "client";
  } else if (currentUserGroups.some((g) => g === "External Adviser" /* EXT_ADVISER */ || g === "External Auditor" /* EXT_AUD */)) {
    return "external";
  } else if (currentUserGroups.some((g) => Object.values(SPGroupInternal).includes(g))) {
    return "bank";
  } else {
    return "";
  }
};
var splitFolderPath = (fullPath) => {
  const normalized = fullPath.replace(/\\/g, "/").replace(/\/+/g, "/");
  const lastSlash = normalized.lastIndexOf("/");
  if (lastSlash === -1) {
    return { path: "", name: normalized };
  }
  return {
    path: normalized.substring(0, lastSlash),
    name: normalized.substring(lastSlash + 1)
  };
};
var createCategoryOptions = (nodes, locale) => {
  return nodes.map((node) => ({
    value: node.id ? node.id.toString() : node.name,
    name: node.localization ? node.localization[locale] : node.name,
    expanded: true,
    selectable: !!node.id,
    subChild: createCategoryOptions(node.nodes, locale)
  }));
};
var downloadFilesAsZip = async (files, apiClient, zipName = "files.zip") => {
  const zip = new JSZip();
  const nameMap = {};
  for (const file of files) {
    try {
      const arrayBuffer = await apiClient.files.getFileToDownload(file.serverRelativeUrl);
      let finalName = file.name;
      if (nameMap[file.name]) {
        const baseName = getFileNameWOExt(file.name);
        const ext = getFileExtension(file.name);
        finalName = ext ? `${baseName} (${nameMap[file.name]}).${ext}` : `${baseName} (${nameMap[file.name]})`;
        nameMap[file.name]++;
      } else {
        nameMap[file.name] = 1;
      }
      if (!arrayBuffer) {
        continue;
      }
      zip.file(finalName, arrayBuffer);
    } catch (err) {
      console.error(`Failed to download file ${file.name}:`, err);
    }
  }
  const blob = await zip.generateAsync({ type: "blob" });
  triggerBlobDownload(blob, zipName);
};
var downloadFiles = async (files, apiClient, zipName = "files.zip") => {
  if (files.length === 1) {
    await downloadFile(files[0].name, files[0].serverRelativeUrl, apiClient);
  } else {
    await downloadFilesAsZip(files, apiClient, zipName);
  }
};
var parseNameAndDate = (input) => {
  if (!input || typeof input !== "string") {
    return { name: "", date: "" };
  }
  const match = input.match(
    /^(.*?)\s(\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\s+\d{1,2}:\d{2})$/
  );
  if (!match) {
    return { name: "", date: "" };
  }
  return {
    name: match[1].trim(),
    date: match[2].trim()
  };
};
var getLatestDate = (dates) => {
  return dates.reduce((latest, dateStr) => {
    const date = new Date(dateStr);
    return date > new Date(latest) ? dateStr : latest;
  }, dates[0]);
};
var copyLink = async (fileLink) => {
  await navigator.clipboard.writeText(fileLink);
};
var normalizeDocType = (s) => (s ?? "").normalize("NFKD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim().replace(/\u00A0/g, " ").replace(/\s*-\s*/g, "-").replace(/\s+/g, "").replace(/[^a-z0-9-]/g, "");
var transformPath = (path) => {
  return path.split("/").map((segment) => {
    return encodeURIComponent(segment).replace(/%20/g, "+");
  }).join("%252f");
};

// src/api/mappers/workflowActionMapper.ts
var mapWorkflowActions = (data) => {
  return data.map((item) => ({
    id: item.Id.toString(),
    title: {
      ["cs-CZ" /* CS */]: item.Title,
      ["en-US" /* EN */]: item["Title_en" /* TitleEn */]
    },
    fileTypes: item["DocumentType" /* DocumentType */] ? item["DocumentType" /* DocumentType */].split("\n") : [],
    statuses: item["DocumentStatus" /* DocumentStatus */] ? item["DocumentStatus" /* DocumentStatus */].split("\n") : [],
    notification: item["Notification" /* Notification */] ? {
      ["cs-CZ" /* CS */]: item["Notification" /* Notification */],
      ["en-US" /* EN */]: item["Notification_en" /* NotificationEn */]
    } : void 0,
    resultStatus: item["ResultStatus" /* ResultStatus */],
    publish: item["Publish" /* Publish */],
    allowComment: item["AllowComment" /* AllowComment */],
    actionGroups: item["ActionGroups" /* ActionGroups */] ? expandRoles(item["ActionGroups" /* ActionGroups */]) : [],
    notificationGroups: item["NotificationGroups" /* NotificationGroups */] ? expandRoles(item["NotificationGroups" /* NotificationGroups */]) : [],
    task: item["Task" /* Task */] ? {
      ["cs-CZ" /* CS */]: item["Task" /* Task */],
      ["en-US" /* EN */]: item["Task_en" /* TaskEn */]
    } : void 0,
    taskGroups: item["TaskGroups" /* TaskGroups */] ? expandRoles(item["TaskGroups" /* TaskGroups */]) : [],
    phase: item["Phase" /* Phase */],
    readPermissions: item["ReadPermissions" /* ReadPermissions */],
    writePermissions: item["WritePermissions" /* WritePermissions */],
    validity: Object.values(FileValidityAction).find((val) => val === item["validity" /* Validity */]),
    reupload: item["Reupload" /* Reupload */],
    fileLink: item["FileLink" /* FileLink */],
    type: item.FileDirRef.split("/").pop(),
    onMajorVersion: item["onMajorVersion" /* OnMajorVersion */]
  }));
};
var mapScheduledAction = (item) => {
  return {
    actionId: item["ActionId" /* ActionId */] ? Number(item["ActionId" /* ActionId */]) : 0,
    fileId: item["FileId" /* FileId */] ? Number(item["FileId" /* FileId */]) : 0,
    processed: item["Processed" /* Processed */] ? new Date(item["Processed" /* Processed */]) : void 0,
    id: item.Id
  };
};
var mapScheduledActions = (data) => {
  const actionMap = /* @__PURE__ */ new Map();
  data.forEach((item) => {
    const fileId = item["FileId" /* FileId */] ? Number(item["FileId" /* FileId */]) : 0;
    if (fileId) {
      actionMap.set(fileId, mapScheduledAction(item));
    }
  });
  return actionMap;
};

// src/api/mappers/notificationMapper.ts
var mapNotification = (item, id, attachmentsMap) => {
  return {
    id: item.Id,
    title: {
      ["cs-CZ" /* CS */]: item.Title,
      ["en-US" /* EN */]: item.Title_en
    },
    task: item.Task,
    completed: item.Completed,
    category: {
      ["cs-CZ" /* CS */]: item.Category ? item.Category.Title : "",
      ["en-US" /* EN */]: item.Category ? item.Category["Category_en" /* CategoryEn */] : ""
    },
    stage: {
      ["cs-CZ" /* CS */]: item.Category ? item.Category["Stage" /* Stage */] : "",
      ["en-US" /* EN */]: item.Category ? item.Category["Stage_en" /* StageEn */] : ""
    },
    phase: {
      ["cs-CZ" /* CS */]: item.Category ? item.Category["Phase" /* Phase */] : "",
      ["en-US" /* EN */]: item.Category ? item.Category["Phase_en" /* PhaseEn */] : ""
    },
    created: new Date(item.Created).toLocaleDateString(),
    author: item.Author.Title,
    unread: !item.Task && (item.ReadId && !item.ReadId.includes(id) || !item.ReadId),
    read: item.ReadId ? item.ReadId : [],
    targetId: item.TargetId,
    categoryId: item.Category && item.Category.Id ? item.Category.Id : 0,
    status: item.Status,
    comment: item.Comment,
    modified: item.Modified,
    editor: item.Editor ? item.Editor.Title : "",
    attachments: attachmentsMap?.get(item.Id) || [],
    link: item.Link,
    uniqueId: item.UniqueId,
    fileLink: item.FileLink
    // assingedTo: item.AssingedTo,
  };
};
var mapNotifications = (data, id, attachmentsMap) => {
  return data.map((item) => mapNotification(item, id, attachmentsMap));
};

// src/api/mappers/formMapper.ts
var mapFormRecords = (data) => {
  return data.map((item) => {
    return {
      id: item.Id,
      title: item.Title,
      formData: parseFormData(item.FormData),
      status: item.Status ?? "new" /* New */,
      localization: {
        ["cs-CZ" /* CS */]: item.Title,
        ["en-US" /* EN */]: item.Title_en ?? item.Title
      },
      categoryId: item[`${"Category" /* Category */}Id`],
      type: Object.values(FormType).includes(item["FormType" /* FormType */]) ? item["FormType" /* FormType */] : "Default" /* Default */,
      language: Object.values(Language).includes(item["Language" /* Language */]) ? item["Language" /* Language */] : void 0,
      completed: item["Completed" /* Completed */],
      fileId: item["FileId" /* FileId */]
    };
  });
};

// src/api/mappers/calendarMapper.ts
var mapCalendarEvents = (data) => {
  return data.map((item) => {
    return {
      id: item.Id,
      subject: item.Title,
      uniqueId: item.UniqueId,
      time: parseSharePointDate(item["Date" /* Date */]),
      startDate: parseSharePointDate(item["Date" /* Date */]),
      type: item["EventType" /* EventType */],
      amount: item["Amount" /* Amount */],
      currency: item["Currency" /* Currency */],
      instruction: item["Instruction" /* Instruction */],
      category: item["Category" /* Category */],
      description: item["EventDescription" /* EventDescription */],
      note: item["Note" /* Note */],
      period: item["Period" /* Period */] ? item["Period" /* Period */] : void 0,
      endDate: parseSharePointDate(item["EndDate" /* EndDate */])
    };
  });
};
var mapInvitations = (data, userExistenceMap) => {
  return data.map((item) => {
    return {
      name: item.Jmeno,
      email: item.Email,
      userExists: userExistenceMap.get(item.Email) ?? false
    };
  });
};
var mapPerson = (user) => {
  return {
    email: user.mail ?? "",
    name: user.displayName ?? "",
    phone: user.mobilePhone ?? "",
    photo: "",
    firstName: user.givenName ?? "",
    lastName: user.surname ?? "",
    id: user.id,
    userType: user.userPrincipalName && !user.userPrincipalName.includes("#EXT#") ? "02internal" /* Internal */ : "03external" /* External */,
    loadedPhoto: false,
    channelMemberId: user.channelMemberId
  };
};
var mapPersons = (users) => {
  return users.map((user) => mapPerson(user));
};
var UserSchema = z.array(z.object({
  Id: z.number(),
  Title: z.string(),
  Email: z.string()
}));
var parseBatchSiteGroupUsersResponse = (raw, groupIds) => {
  const result = /* @__PURE__ */ new Map();
  const parts = raw.split(/--batchresponse_[\s\S]*?Content-Type: application\/http/gi).filter((p) => p !== "");
  parts.forEach((part, index) => {
    const jsonMatch = part.match(/\{[\s\S]*\}/m);
    if (!jsonMatch) return;
    let json;
    try {
      json = JSON.parse(jsonMatch[0]);
    } catch {
      return;
    }
    const array = UserSchema.safeParse(json.value);
    if (!array.success) {
      return;
    }
    const users = array.data.map((u) => ({
      id: u.Id,
      name: u.Title,
      email: u.Email
    }));
    const groupId = groupIds[index];
    if (!groupId) return;
    result.set(groupId, users);
  });
  return result;
};

// src/api/mappers/templateMapper.ts
var mapTemplates = (data) => {
  return data.map((item) => {
    return {
      stage: item["Stage" /* Stage */],
      phase: item["Phase" /* Phase */],
      category: item.Title,
      description: item["FolderDescription" /* Description */],
      descriptionEn: item["FolderDescription_en" /* DescriptionEn */],
      stageEn: item["Stage_en" /* StageEn */] ?? item["Stage" /* Stage */],
      categoryEn: item["Category_en" /* CategoryEn */] ?? item.Title,
      phaseEn: item["Phase_en" /* PhaseEn */] ?? item["Phase" /* Phase */],
      //folder: item[ColumnsCategories.Folder],
      id: item.Id
    };
  });
};
var mapTemplateReferences = (data) => {
  return data.map((item) => ({
    fileRef: item.FileRef,
    fileLeafRef: item.FileLeafRef
  }));
};
var mapFolderTemplates = (data) => {
  return data.filter((item) => !data.some((it) => it.FileRef === item.FileDirRef)).map((item) => {
    return {
      name: item.FileLeafRef,
      nameEn: item.Name_en,
      templates: mapFolderTemplates(data.filter((it) => it.FileDirRef.startsWith(item.FileRef))),
      write: item.WritePermissions,
      read: item.ReadPermissions
    };
  });
};

// src/api/mappers/requestMapper.ts
var mapFileTemplates = (csTemplates, enTemplates) => {
  const cs = csTemplates ? csTemplates.split(";") : [];
  const en = enTemplates ? enTemplates.split(";") : [];
  const result = [];
  for (let i = 0; i < cs.length; i++) {
    result.push({ ["cs-CZ" /* CS */]: cs[i], ["en-US" /* EN */]: en[i] });
  }
  return result;
};
var mapRequestOption = (item) => {
  return {
    id: item.Id,
    type: item.Title,
    fileTemplates: mapFileTemplates(item["FileTemplate" /* FileTemplate */], item["FileTemplateEn" /* FileTemplateEn */]),
    loc: {
      ["cs-CZ" /* CS */]: item.Title,
      ["en-US" /* EN */]: item["Title_en" /* TitleEn */]
    },
    desc: {
      ["cs-CZ" /* CS */]: item["Description" /* Description */],
      ["en-US" /* EN */]: item["Description_en" /* DescriptionEn */]
    },
    stage: item["Etapa" /* Stage */],
    fields: item["Fields" /* Fields */],
    fileAttachments: item["FileAttachments" /* FileAttachments */],
    notifGroups: item["Notification" /* Notification */],
    fileReply: item["FileReply" /* FileReply */]
  };
};
var mapRequestOptions = (data) => {
  return data.map((item) => mapRequestOption(item));
};
var mapRequest = (item) => {
  const createdDate = new Date(item.Created) ?? /* @__PURE__ */ new Date();
  createdDate.setHours(0, 0, 0, 0);
  return {
    type: {
      id: item["RequestType" /* RequestType */].Id,
      type: item["RequestType" /* RequestType */].Title,
      loc: {
        ["cs-CZ" /* CS */]: item["RequestType" /* RequestType */].Title,
        ["en-US" /* EN */]: item["RequestType" /* RequestType */]["Title_en" /* TitleEn */]
      },
      fileTemplates: [],
      notifGroups: item["RequestType" /* RequestType */]["Notification" /* Notification */]
    },
    status: item["RequestStatus" /* RequestStatus */],
    created: createdDate,
    formData: parseFormData(item["FormData" /* FormData */] ?? "{}"),
    id: item.Id,
    //attachments: attachments,
    templateRef: item.FileTemplate,
    categoryId: item["categoryId" /* CategoryId */],
    fileIds: item["FileIds" /* FileIds */] ? item["FileIds" /* FileIds */].split(";").filter((id) => id !== "") : []
  };
};
var mapRequests = (data) => {
  return data.map((item) => {
    return mapRequest(
      item
      /* attachments*/
    );
  });
};

// src/api/mappers/fileFolderMapper.ts
var mapRequiredFiles = (data) => {
  return data.map((item) => {
    return {
      loc: {
        ["cs-CZ" /* CS */]: item.Title,
        ["en-US" /* EN */]: item["Title_en" /* TitleEn */] ? item["Title_en" /* TitleEn */] : item.Title
      },
      id: item.Id,
      //folderPath: item[ColumnsRequired.FolderPath],
      description: item["FileDescription" /* FileDesc */],
      type: item["DocumentType" /* DocumentType */] ? {
        loc: {
          ["en-US" /* EN */]: getFileTypeTitle(item["DocumentType" /* DocumentType */]["Title_en" /* TitleEn */]),
          ["cs-CZ" /* CS */]: getFileTypeTitle(item["DocumentType" /* DocumentType */].Title)
        },
        id: item.DocumentType.Id,
        uploadStatus: item["DocumentType" /* DocumentType */]["UploadStatus" /* UploadStatus */],
        permissions: item["DocumentType" /* DocumentType */]["Permissions" /* Permissions */] ? expandRoles(item["DocumentType" /* DocumentType */]["Permissions" /* Permissions */]) : [],
        folderPath: item["DocumentType" /* DocumentType */]["FolderPath" /* FolderPath */],
        folderPathAfter: item["DocumentType" /* DocumentType */]["FolderPathAfter" /* FolderPathAfter */],
        validityStatuses: []
      } : void 0,
      format: item["PermittedFormats" /* PermittedFormats */],
      uploadedFile: item["uploadedFile" /* UploadedFile */],
      deadline: item["Deadline" /* Deadline */] ? new Date(item["Deadline" /* Deadline */]) : void 0,
      status: Object.values(FileStatus).includes(item["Status" /* Status */]) ? item["Status" /* Status */] : void 0
    };
  });
};
var mapFolders = (data) => {
  return data.filter((item) => !data.find((it) => it.FileRef === item.FileDirRef)).map((item) => {
    return {
      name: item.FileLeafRef,
      serverRelativeUrl: item.FileRef,
      localization: {
        ["cs-CZ" /* CS */]: item.FileLeafRef,
        ["en-US" /* EN */]: item.Name_en
      },
      folders: mapFolders(data.filter((it) => it.FileDirRef.startsWith(item.FileRef)))
    };
  });
};
var mapItemAsFile = (item) => {
  return {
    name: item.FileLeafRef,
    serverRelativeUrl: item.FileRef,
    status: item.Status && item.Status !== "" ? item.Status : "processing" /* Processing */,
    uniqueId: item.UniqueId,
    version: item.OData__UIVersionString,
    type: item.DocumentType ? {
      loc: {
        ["en-US" /* EN */]: getFileTypeTitle(item.DocumentType.Title_en),
        ["cs-CZ" /* CS */]: getFileTypeTitle(item.DocumentType.Title)
      },
      id: item.DocumentType.Id,
      folderPath: item.DocumentType.FolderPath,
      folderPathAfter: item.DocumentType.FolderPathAfter,
      validityStatuses: item.DocumentType["ValidityStatuses" /* ValidityStatuses */] ? item.DocumentType["ValidityStatuses" /* ValidityStatuses */].split(";") : [],
      folderPathInternal: item.DocumentType["FolderPathInternal" /* FolderPathInternal */],
      folderPathInternalAfter: item.DocumentType["FolderPathInternalAfter" /* FolderPathInternalAfter */]
    } : void 0,
    created: item.Created ? new Date(item.Created) : void 0,
    id: item.Id,
    nodeId: item.StrukturaDokumentaceId,
    documentation: item.Documentation,
    provided: item.UploadedBy,
    contractValidity: item.ContractValidity,
    validTo: item.ValidTo ? new Date(item.ValidTo) : void 0,
    qaId: item.QAId ? item.QAId : "",
    author: item.Author && item.Author.Title ? item.Author.Title : "",
    modified: item.Modified ? new Date(item.Modified) : void 0
  };
};
var mapFile = (file) => {
  return {
    name: file.FileLeafRef,
    serverRelativeUrl: file.FileRef,
    status: file.Status,
    uniqueId: file.UniqueId,
    version: file.OData__UIVersionString,
    nodeId: file.StrukturaDokumentaceId,
    typeId: file.DocumentTypeId
  };
};
var mapItemsAsFiles = (items) => {
  return items.map(
    (item) => mapItemAsFile(item)
  );
};
var mapFileVersion = (item, teamName) => {
  const relativeUrl = item.FileRef.replace(`/sites/${teamName}`, "");
  return {
    editor: item.Editor ? item.Editor.LookupValue : "",
    modified: item.Modified,
    status: item.Status,
    versionLabel: item.VersionLabel,
    relativeUrl: item.IsCurrentVersion ? item.FileRef : `/sites/${teamName}/_vti_history/${item.VersionId}${relativeUrl}`,
    isCurrentVersion: item.IsCurrentVersion,
    id: item.VersionId
  };
};
var mapFileVersions = (data, teamName) => {
  return data.map((item) => mapFileVersion(item, teamName));
};
var mapFileTypes = (data) => {
  return data.filter((item) => !item.Title.includes(">")).map((item) => {
    const titleEn = typeof item["Title_en" /* TitleEn */] === "string" ? item["Title_en" /* TitleEn */] : "";
    return {
      loc: {
        ["cs-CZ" /* CS */]: item.Title,
        ["en-US" /* EN */]: titleEn
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
      uploadStatus: item["UploadStatus" /* UploadStatus */],
      permissions: item["Permissions" /* Permissions */] ? expandRoles(item["Permissions" /* Permissions */]) : [],
      folderPath: item["FolderPath" /* FolderPath */],
      folderPathAfter: item["FolderPathAfter" /* FolderPathAfter */],
      validityStatuses: item["ValidityStatuses" /* ValidityStatuses */] ? item["ValidityStatuses" /* ValidityStatuses */].split(";") : [],
      folderPathInternal: item["FolderPathInternal" /* FolderPathInternal */],
      folderPathInternalAfter: item["FolderPathInternalAfter" /* FolderPathInternalAfter */]
    };
  });
};

// src/api/mappers/structureMapper.ts
var sortStructureNodesByName = (nodes) => {
  const sorted = [...nodes].sort((a, b) => a.name.localeCompare(b.name));
  return sorted;
};
var mapCategoryNode = (item) => {
  return {
    name: item.Title,
    id: item.Id,
    uploadedFiles: [],
    nodes: [],
    status: item["CategoryStatus" /* CategoryStatus */] ? item["CategoryStatus" /* CategoryStatus */] : "default" /* Default */,
    description: {
      ["cs-CZ" /* CS */]: item["FolderDescription" /* Description */],
      ["en-US" /* EN */]: item["FolderDescription_en" /* DescriptionEn */]
    },
    isCategory: true,
    localization: {
      ["cs-CZ" /* CS */]: cleanFolderName(item.Title) ?? "",
      ["en-US" /* EN */]: cleanFolderName(item["Category_en" /* CategoryEn */]) ?? ""
    },
    requiredFiles: [],
    order: item["CategoryOrder" /* Order */]
  };
};
var mapStructure = (categories, phasesQnA) => {
  const nodes = [];
  for (let i = 0; i < categories.length; i++) {
    const item = categories[i];
    let stageNode = nodes.find((fol) => fol.name === item["Stage" /* Stage */]);
    if (!stageNode) {
      stageNode = {
        name: item["Stage" /* Stage */],
        uploadedFiles: [],
        nodes: [],
        status: "default" /* Default */,
        localization: {
          ["cs-CZ" /* CS */]: cleanFolderName(item["Stage" /* Stage */]) ?? "",
          ["en-US" /* EN */]: cleanFolderName(item["Stage_en" /* StageEn */]) ?? ""
        }
      };
      nodes.push(stageNode);
    }
    let phaseNode = stageNode.nodes.find((fol) => fol.name === item["Phase" /* Phase */]);
    if (!phaseNode) {
      phaseNode = {
        name: item["Phase" /* Phase */],
        uploadedFiles: [],
        nodes: [],
        status: "default" /* Default */,
        localization: {
          ["cs-CZ" /* CS */]: cleanFolderName(item["Phase" /* Phase */]) ?? "",
          ["en-US" /* EN */]: cleanFolderName(item["Phase_en" /* PhaseEn */]) ?? ""
        }
      };
      stageNode.nodes.push(phaseNode);
    }
    let categoryNode = phaseNode.nodes.find((fol) => fol.name === item.Title);
    if (!categoryNode) {
      categoryNode = mapCategoryNode(item);
      phaseNode.nodes.push(categoryNode);
    }
  }
  phasesQnA.forEach((phaseQnA) => {
    const stage = nodes.find((node) => node.name === phaseQnA.Stage);
    if (stage) {
      const phase = stage.nodes.find((node) => node.name === phaseQnA.Phase);
      if (!phase) {
        stage.nodes = [
          ...stage.nodes,
          {
            name: phaseQnA.Phase,
            nodes: [],
            uploadedFiles: [],
            status: "default" /* Default */,
            localization: {
              ["cs-CZ" /* CS */]: cleanFolderName(phaseQnA.Phase) ?? "",
              ["en-US" /* EN */]: cleanFolderName(phaseQnA.Phase_en) ?? ""
            }
          }
        ];
      }
    }
  });
  nodes.forEach((node) => {
    node.nodes = sortStructureNodesByName(node.nodes);
  });
  return sortStructureNodesByName(nodes);
};

// src/api/mappers/detailMapper.ts
var mapDetail = (item, iupLink, contactsData) => {
  const status = item["currentStatus" /* CurrentStatus */] ? item["currentStatus" /* CurrentStatus */].split("/") : "";
  const channelStatus = Object.values(ChannelStatus).find((v) => v === item["ChannelStatus" /* ChannelStatus */]);
  const contactData = item["contactData" /* ContactData */] ?? {};
  const secondaryData = item["secondaryData" /* SecondaryData */] ?? {};
  return {
    id: item.Id,
    currentStatus: status[0] ?? "",
    currentFolder: status[1] ?? "",
    text: item["loanText" /* LoanText */],
    title: item.Title,
    contactDesc: Object.values(ContactDesc).includes(contactData.desc) ? contactData.desc : void 0,
    secondaryDesc: Object.values(ContactDesc).includes(secondaryData.desc) ? secondaryData.desc : void 0,
    contact: contactsData && contactsData["contact" /* Contact */] && contactsData["contact" /* Contact */].Name ? {
      email: contactsData["contact" /* Contact */].EMail,
      firstName: contactsData["contact" /* Contact */].FirstName,
      lastName: contactsData["contact" /* Contact */].LastName,
      name: contactsData["contact" /* Contact */].Title,
      phone: contactData.phone ? contactData.phone : "",
      photo: "",
      userType: contactsData["contact" /* Contact */].Name.includes("#EXT#") ? "03external" /* External */ : "02internal" /* Internal */,
      id: contactData.id ? contactData.id : "",
      loadedPhoto: false
    } : void 0,
    secondaryContact: contactsData && contactsData["secondaryContact" /* SecondaryContact */] && contactsData["secondaryContact" /* SecondaryContact */].Name ? {
      email: contactsData["secondaryContact" /* SecondaryContact */].EMail,
      firstName: contactsData["secondaryContact" /* SecondaryContact */].FirstName,
      lastName: contactsData["secondaryContact" /* SecondaryContact */].LastName,
      name: contactsData["secondaryContact" /* SecondaryContact */].Title,
      phone: secondaryData.phone ? secondaryData.phone : "",
      photo: "",
      userType: contactsData["secondaryContact" /* SecondaryContact */].Name.includes("#EXT#") ? "03external" /* External */ : "02internal" /* Internal */,
      id: secondaryData.id ? secondaryData.id : "",
      loadedPhoto: false
    } : void 0,
    channelStatus: channelStatus ? channelStatus : "Active" /* Active */,
    iupLink,
    version: item["version" /* Version */] ? item["version" /* Version */] : 0
  };
};

// src/api/actions/ActionApiClient.ts
var ActionApiClient = class {
  /**
   * Constructs an ActionApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  /**
   * Retrieves workflow actions from the SharePoint WorkflowActions list.
   * Filters items by content type and maps them to structured workflow action objects.
   * @returns A Promise resolving to an array of IWorkflowAction objects, or an empty array if no data is found.
   */
  async getWorkflowActions(fromTemplates) {
    const absoluteUrl = fromTemplates ? this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), `/sites/${this.sp.templateSiteName}`) : this.sp.getAbsoluteUrl();
    const filter = `startswith(ContentTypeId,'0x0100')`;
    const apiUrl = `${absoluteUrl}/_api/web/lists/getbytitle('${"WorkflowActions" /* WorkflowActions */}')/items?$filter=${filter}&$select=*,FileDirRef`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) ? mapWorkflowActions(result.data) : [];
  }
  /**
   * Retrieves scheduled actions from the SharePoint ActionToTrigger list.
   * Filters items by folder and processed status, then maps them to structured scheduled action objects.
   * @returns A Promise resolving to a Map of scheduled actions, or an empty Map if no data is found.
   */
  async getScheduledActions() {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"ActionToTrigger" /* ActionToTrigger */}/${this.sp.getChannelName()}`;
    const filter = `startswith(FileDirRef,'${fileDirRef}') and ${"Processed" /* Processed */} eq null`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"ActionToTrigger" /* ActionToTrigger */}')/items?$filter=${filter}&$select=*`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) ? mapScheduledActions(result.data) : /* @__PURE__ */ new Map();
  }
  async getScheduledActionByFileId(fileId) {
    if (!fileId) {
      return void 0;
    }
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"ActionToTrigger" /* ActionToTrigger */}/${this.sp.getChannelName()}`;
    const filter = `startswith(FileDirRef,'${fileDirRef}') and ${"Processed" /* Processed */} eq null and ${"FileId" /* FileId */} eq ${fileId}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"ActionToTrigger" /* ActionToTrigger */}')/items?$filter=${filter}&$select=*`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) && result.data[0] ? mapScheduledAction(result.data[0]) : void 0;
  }
  /**
   * Creates a new scheduled action in the ActionToTrigger list.
   * @param action The scheduled action object to create.
   * @param title The title for the scheduled action.
   * @returns Promise resolving to true if creation succeeded, otherwise false.
   */
  async createScheduledAction(action, title, comment) {
    const body = {
      "formValues": [
        {
          "FieldName": "Title",
          "FieldValue": title
        },
        {
          "FieldName": "FileId" /* FileId */,
          "FieldValue": action.fileId.toString()
        },
        {
          "FieldName": "ActionId" /* ActionId */,
          "FieldValue": action.actionId.toString()
        },
        {
          "FieldName": "Comment" /* Comment */,
          "FieldValue": comment ?? ""
        }
      ]
    };
    const result = await this.sp.createListItemUsingPath(body, "ActionToTrigger" /* ActionToTrigger */);
    return !!result.ok;
  }
};

// src/api/calendar/CalendarApiClient.ts
var CalendarApiClient = class {
  /**
   * Constructs a CalendarApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  /**
   * Retrieves calendar events from the Duties SharePoint list.
   * Filters items by folder and content type, then maps them to structured calendar event objects.
   * @returns Promise resolving to an array of ICalendarEvent objects, or an empty array if no events are found.
   */
  async getCalendarEvents() {
    const folderPath = `${this.sp.getRelativeUrl()}/lists/${"Povinnosti" /* Duties */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
    const select = "*,UniqueId";
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Povinnosti" /* Duties */}')/items?$select=${select}&$filter=${filter}`;
    const response = await this.sp.spGet(apiUrl);
    return Array.isArray(response.data) ? mapCalendarEvents(response.data) : [];
  }
  /**
   * Creates a new calendar event item in the Duties SharePoint list.
   * Normalizes event dates to noon and formats as ISO strings before saving.
   * @param event The calendar event data to create.
   * @returns Promise resolving to true if creation succeeded, otherwise false.
   */
  async createCalendarEvent(event) {
    const date = event.startDate;
    date?.setHours(12);
    const formattedDate = date ? date.toISOString() : "";
    const endDate = event.endDate;
    endDate?.setHours(12);
    const formattedEndDate = endDate ? endDate.toISOString() : "";
    const body = {
      "formValues": [
        {
          "FieldName": "Title",
          "FieldValue": event.subject
        },
        {
          "FieldName": "Date" /* Date */,
          "FieldValue": formattedDate
        },
        {
          "FieldName": "EventType" /* EventType */,
          "FieldValue": event.type
        },
        {
          "FieldName": "Amount" /* Amount */,
          "FieldValue": event.amount?.toString()
        },
        {
          "FieldName": "Currency" /* Currency */,
          "FieldValue": event.currency
        },
        {
          "FieldName": "Instruction" /* Instruction */,
          "FieldValue": event.instruction
        },
        {
          "FieldName": "Category" /* Category */,
          "FieldValue": event.category
        },
        {
          "FieldName": "EventDescription" /* EventDescription */,
          "FieldValue": event.description
        },
        {
          "FieldName": "Note" /* Note */,
          "FieldValue": event.note
        },
        {
          "FieldName": "Period" /* Period */,
          "FieldValue": event.period
        },
        {
          "FieldName": "EndDate" /* EndDate */,
          "FieldValue": formattedEndDate
        }
      ]
    };
    const result = await this.sp.createListItemUsingPath(body, "Povinnosti" /* Duties */);
    return result.ok ?? false;
  }
};

// src/api/chat/ChatApiClient.ts
var ChatApiClient = class {
  /**
   * Constructs a ChatApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  async getMessages() {
    const apiUrl = `/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/messages`;
    const result = await this.sp.graphGet(apiUrl);
    console.log(result);
    return [];
  }
};
var ContactDataSchema = z.object({
  id: z.string().nullable().optional(),
  phone: z.string().nullable().optional(),
  desc: z.string().nullable().optional()
});
var parseContactData = (value) => {
  if (!value) {
    return {};
  }
  try {
    if (typeof value !== "string") {
      return {};
    }
    const parsedJson = JSON.parse(value);
    const parsedContactData = ContactDataSchema.safeParse(parsedJson);
    return parsedContactData.success ? parsedContactData.data : {};
  } catch {
    return {};
  }
};
var SPDetailItemSchema = z.object({
  Id: z.number(),
  Title: z.string().nullable().optional(),
  ["currentStatus" /* CurrentStatus */]: z.string().nullable().optional(),
  ["ChannelStatus" /* ChannelStatus */]: z.string().nullable().optional(),
  ["loanText" /* LoanText */]: z.string().nullable().optional(),
  ["contactData" /* ContactData */]: z.string().nullable().optional().transform(parseContactData),
  ["secondaryData" /* SecondaryData */]: z.string().nullable().optional().transform(parseContactData),
  ["version" /* Version */]: z.number().nullable().optional()
});
var DetailApiClient = class {
  /**
   * Constructs a DetailApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  documentFolder = "Sdilene dokumenty";
  /**
       * create detail of loan when installing
       * @param channelId 
       * @param title 
       * @param currentStatus 
       */
  async createDetail(currentStatus, version) {
    const body = {
      "Title": this.sp.getChannelName(),
      ["channelId" /* ChannelId */]: this.sp.getChannelId(),
      ["currentStatus" /* CurrentStatus */]: currentStatus,
      ["ChannelStatus" /* ChannelStatus */]: "Active" /* Active */,
      ["version" /* Version */]: version
    };
    const result = await this.sp.createListItem(body, "Detaily" /* Details */);
    return result.ok ?? false;
  }
  /**
   * @param channelId 
   * @returns detail {@link ILoanDetail}
   */
  async getLoanDetail() {
    const filter = `${"channelId" /* ChannelId */} eq '${this.sp.getChannelId()}'`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Detaily" /* Details */}')/items?$filter=${filter}&$select=*`;
    const response = await this.sp.spGet(apiUrl);
    const contactResponse = await this.getLoanDetailContacts(this.sp.getChannelId());
    const contactsData = Array.isArray(contactResponse.data) ? contactResponse.data[0] : void 0;
    const rawItem = Array.isArray(response.data) ? response.data[0] : void 0;
    if (!rawItem) {
      return void 0;
    }
    const parsed = SPDetailItemSchema.safeParse(rawItem);
    if (!parsed.success) {
      console.error(parsed.error);
      return void 0;
    }
    const item = parsed.data;
    const iupLink = await this.getIUPLink();
    return mapDetail(item, iupLink, contactsData);
  }
  /**
  * updates detail using {@link updateListItem}
  * @param itemId 
  * @param body 
  * @returns boolean
  */
  async updateDetail(itemId, body) {
    const result = await this.sp.updateListItem(itemId, body, "Detaily" /* Details */);
    return result;
  }
  /**
   * ensure user and update detail using {@link updateListItem}
   * @param itemId 
   * @param email 
   * @returns boolean
   */
  async updateContact(itemId, contactData, secondaryData, email, secondaryEmail) {
    const userId = email ? await this.sp.getUserSPId(email) : null;
    const secondaryId = secondaryEmail ? await this.sp.getUserSPId(secondaryEmail) : null;
    const body = {
      [`${"contact" /* Contact */}Id`]: userId,
      [`${"secondaryContact" /* SecondaryContact */}Id`]: secondaryId,
      ["contactData" /* ContactData */]: contactData ?? "",
      ["secondaryData" /* SecondaryData */]: secondaryData ?? ""
    };
    const result = await this.sp.updateListItem(itemId, body, "Detaily" /* Details */);
    return result;
  }
  async getLoanDetailContacts(channelId) {
    const filter = `${"channelId" /* ChannelId */} eq '${channelId}'`;
    const select = [
      "*",
      `${"contact" /* Contact */}/FirstName`,
      `${"contact" /* Contact */}/LastName`,
      `${"contact" /* Contact */}/Title`,
      `${"contact" /* Contact */}/Name`,
      `${"contact" /* Contact */}/UserName`,
      `${"contact" /* Contact */}/EMail`,
      `${"contact" /* Contact */}/Id`,
      `${"contact" /* Contact */}/MobilePhone`,
      `${"contact" /* Contact */}/WorkPhone`,
      `${"secondaryContact" /* SecondaryContact */}/FirstName`,
      `${"secondaryContact" /* SecondaryContact */}/LastName`,
      `${"secondaryContact" /* SecondaryContact */}/Title`,
      `${"secondaryContact" /* SecondaryContact */}/Name`,
      `${"secondaryContact" /* SecondaryContact */}/UserName`,
      `${"secondaryContact" /* SecondaryContact */}/EMail`,
      `${"secondaryContact" /* SecondaryContact */}/Id`,
      `${"secondaryContact" /* SecondaryContact */}/MobilePhone`,
      `${"secondaryContact" /* SecondaryContact */}/WorkPhone`
    ].join(",");
    const expand = `${"contact" /* Contact */},${"secondaryContact" /* SecondaryContact */}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Detaily" /* Details */}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
    const response = await this.sp.spGet(apiUrl);
    return response;
  }
  async getIUPLink() {
    const teamName = await this.sp.getTeamName();
    if (!teamName) {
      return "";
    }
    const updatedTeamName = teamName.replace("_ext", "");
    const id = `/sites/${this.sp.internalSiteName}/${this.documentFolder}/${updatedTeamName}/${this.sp.getChannelName()}`;
    const url = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}/${this.documentFolder}/Forms/AllItems.aspx?id=${id}`;
    return url;
  }
};
var FileItemSchema = z.object({
  FileLeafRef: z.string(),
  FileRef: z.string(),
  Status: z.nativeEnum(FileStatus).nullable().optional(),
  UniqueId: z.string(),
  OData__UIVersionString: z.string(),
  DocumentType: z.object({
    Title: z.string(),
    ["Title_en" /* TitleEn */]: z.string(),
    Id: z.number(),
    ["ValidityStatuses" /* ValidityStatuses */]: z.string().nullable().optional(),
    ["FolderPathInternal" /* FolderPathInternal */]: z.string().nullable().optional(),
    ["FolderPathInternalAfter" /* FolderPathInternalAfter */]: z.string().nullable().optional(),
    ["FolderPath" /* FolderPath */]: z.string().nullable().optional(),
    ["FolderPathAfter" /* FolderPathAfter */]: z.string().nullable().optional()
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
    Title: z.string().nullable().optional()
  }).nullable().optional()
});
z.array(FileItemSchema);
var parseFileItem = (value) => {
  if (!value) {
    return void 0;
  }
  const parsedFileItem = FileItemSchema.safeParse(value);
  return parsedFileItem.success ? parsedFileItem.data : void 0;
};
var parseFileItems = (value) => {
  if (!value || !Array.isArray(value)) {
    return [];
  }
  const items = value.map((item) => parseFileItem(item)).filter((item) => item !== void 0);
  return items;
};
var parseAndMapFileItems = (value) => {
  const data = parseFileItems(value);
  return mapItemsAsFiles(data);
};
var FileApiClient = class {
  /**
   * Constructs a FileApiClient instance.
   * @param spClient The SPApiClient instance for SharePoint operations.
   */
  constructor(spClient) {
    this.spClient = spClient;
  }
  spClient;
  documentFolder = "Sdilene dokumenty";
  trashPath = "KLIENT - EXT/KO\u0160";
  /**
   * Moves a file to the bin (trash) folder.
   * @param fileRef The file reference (server-relative URL).
   * @param fileName The file name.
   * @param uniqueId Optional unique ID for the file.
   * @returns Promise resolving to true if moved successfully.
   */
  async moveFileToBin(fileRef, fileName) {
    const newFileDirRef = this.trashPath;
    return await this.moveFile(fileRef, fileName, newFileDirRef);
  }
  /**
  * Loads files from a SharePoint document library filtered by a structure ID,
  * excluding deleted files and those in the trash folder
  * @param id - The structure ID used to filter files
  * @returns A promise resolving to an array of IFile objects matching the filter criteria
  */
  async loadFiles(id) {
    const select = `*,FileRef,FileLeafRef,UniqueId,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${"ValidityStatuses" /* ValidityStatuses */},DocumentType/${"FolderPathInternal" /* FolderPathInternal */},DocumentType/${"FolderPath" /* FolderPath */},DocumentType/${"FolderPathAfter" /* FolderPathAfter */},DocumentType/${"FolderPathInternalAfter" /* FolderPathInternalAfter */}`;
    const filter = `StrukturaDokumentaceId eq ${id} and Status ne '${"deleted" /* Deleted */}'`;
    const expand = "DocumentType,Author";
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/lists/getbytitle('${"Dokumenty" /* Documents */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const result = await this.spClient.spGet(apiUrl);
    return parseAndMapFileItems(result.data);
  }
  /**
   * Loads files from SharePoint by their item IDs.
   * @param ids Array of item IDs.
   * @returns Promise resolving to an array of IFile objects.
   */
  async loadFilesByIds(ids) {
    if (ids.length === 0) {
      return [];
    }
    const select = `*,FileRef,FileLeafRef,UniqueId,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${"ValidityStatuses" /* ValidityStatuses */},DocumentType/${"FolderPathInternal" /* FolderPathInternal */},DocumentType/${"FolderPath" /* FolderPath */},DocumentType/${"FolderPathAfter" /* FolderPathAfter */},DocumentType/${"FolderPathInternalAfter" /* FolderPathInternalAfter */}`;
    const filter = `${ids.map((id) => `Id eq ${id}`).join(" or ")} and Status ne '${"deleted" /* Deleted */}'`;
    const expand = "DocumentType,Author";
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/lists/getbytitle('${"Dokumenty" /* Documents */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const result = await this.spClient.spGet(apiUrl);
    return parseAndMapFileItems(result.data);
  }
  /**
   * Loads files from SharePoint by their unique IDs.
   * @param uniqueIds Array of unique IDs.
   * @returns Promise resolving to an array of IFile objects.
   */
  async loadFilesByUniqueIds(uniqueIds) {
    const result = await Promise.all(
      uniqueIds.map(
        (id) => this.loadFileByUniqueId(id)
      )
    );
    return result.filter((file) => file !== void 0);
  }
  /**
   * Loads a file from SharePoint by its server-relative URL.
   * @param serverRelativeUrl The server-relative URL of the file.
   * @returns Promise resolving to an IFile object or undefined.
   */
  async loadFileByRelativeUrl(serverRelativeUrl) {
    const baseUrl = this.spClient.getAbsoluteUrl();
    const result = await this.spClient.spGet(
      `${baseUrl}/_api/web/getfilebyserverrelativeurl('${serverRelativeUrl}')/ListItemAllFields?$select=*,StrukturaDokumentaceId,FileRef,FileLeafRef,UIVersionLabel,UniqueId`
    );
    return result.data ? mapFile(result.data) : void 0;
  }
  /**
   * Loads a file from SharePoint by its unique ID (private).
   * @param uniqueId The unique ID of the file.
   * @returns Promise resolving to an IFile object or undefined.
   */
  async loadFileByUniqueId(uniqueId) {
    const baseUrl = this.spClient.getAbsoluteUrl();
    const select = `*,StrukturaDokumentaceId,FileRef,FileLeafRef,UIVersionLabel,UniqueId`;
    const result = await this.spClient.spGet(`${baseUrl}/_api/web/getfilebyid('${uniqueId}')/ListItemAllFields?$select=${select}`);
    return result.data ? mapFile(result.data) : void 0;
  }
  /**
  * Loads all files from a channel folder path within the SharePoint document library,
  * optionally filtering by documentation flag, excluding deleted files and those in the trash folder
  * @param documentation - If true, filters files marked as documentation
  * @returns A promise resolving to an array of IFile objects matching the criteria
  */
  async loadAllFiles(documentation) {
    const folderPath = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}`;
    const filter = `startswith(ContentTypeId,'0x0101') and startswith(FileDirRef,'${folderPath}') and Status ne '${"deleted" /* Deleted */}' ${documentation ? `and Documentation eq 1` : ""}`;
    const select = `*,UniqueId,FileRef,FileLeafRef,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id,Author/Title,DocumentType/${"ValidityStatuses" /* ValidityStatuses */},DocumentType/${"FolderPathInternal" /* FolderPathInternal */},DocumentType/${"FolderPathAfter" /* FolderPathAfter */},DocumentType/${"FolderPath" /* FolderPath */},DocumentType/${"FolderPathInternalAfter" /* FolderPathInternalAfter */}`;
    const expand = "DocumentType,Author";
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/GetListByTitle('${"Dokumenty" /* Documents */}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
    const result = await this.spClient.spGet(apiUrl);
    return parseAndMapFileItems(result.data);
  }
  /**
   * publishs file by using its uniqueId
   * @param uniqueId 
   * @returns boolean
   */
  async publishFile(uniqueId) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/GetItemByUniqueId('${uniqueId}')/File/publish('')`;
    const result = await this.spClient.spPost(apiUrl);
    return result.ok ?? false;
  }
  async assignFilePermissions(uniqueId, readGroupsIds, writeGroupIds) {
    const boundary = "batch_" + Date.now();
    const webUrl = this.spClient.getAbsoluteUrl();
    const apiUrlBase = `${webUrl}/_api/web/GetFileById('${uniqueId}')/ListItemAllFields/roleassignments/addroleassignment`;
    const readPart = readGroupsIds.map(
      (id) => this.spClient.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
    ).join("");
    const writePart = writeGroupIds.map(
      (id) => this.spClient.getPermissionBatchString(boundary, id, 1073741827, apiUrlBase)
    ).join("");
    const batchBody = `${readPart}${writePart}--${boundary}--\r
`;
    await this.spClient.spPost(
      `${webUrl}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`
      }
    );
  }
  async setFilePermissions(action, siteGroups, uniqueId) {
    const writeGroups = action.writePermissions ? expandRoles(action.writePermissions) : [];
    const readGroupsAll = action.readPermissions ? expandRoles(action.readPermissions) : [];
    const writeGroupIds = siteGroups.filter((g) => writeGroups.includes(g.title)).map((g) => g.id);
    await this.resetAndBreakInheritance(uniqueId);
    const readGroups = [...readGroupsAll].filter((g) => !writeGroups.includes(g));
    const readGroupIds = siteGroups.filter((g) => readGroups.includes(g.title)).map((g) => g.id);
    await this.assignFilePermissions(uniqueId, readGroupIds, writeGroupIds);
  }
  /**
   * upload array of files using serverRelativeUrl of folder
   * @param folderPath 
   * @param files 
   * @returns array of fileNames whose upload failed
   */
  async uploadFiles(files, secondStage, uploadedBy, folderId) {
    const result = await Promise.all(
      files.map(async (file) => {
        const folderPath = !file.type ? "/Temporary" : secondStage && file.type.folderPathAfter ? file.type.folderPathAfter : file.type.folderPath;
        return await this.uploadFile(folderPath, file.file, uploadedBy, folderId, file.type?.id, file.type?.uploadStatus);
      })
    );
    return result;
  }
  /**
  * Retrieves the list of available file types from the SharePoint 'DocumentTypes' list
  * @returns A promise resolving to an array of IFileType objects
  */
  async getFileTypes(webUrl = this.spClient.getAbsoluteUrl()) {
    const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${"TypyDokumentu" /* DocumentTypes */}')/items`;
    const result = await this.spClient.spGet(apiUrl);
    return Array.isArray(result.data) ? mapFileTypes(result.data) : [];
  }
  /**
   * Undoes the check-out of a file by unique ID.
   * @param uniqueId The unique ID of the file.
   * @returns Promise resolving to true if successful.
   */
  async undoFileCheckOut(uniqueId) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyid(guid'${uniqueId}')/UndoCheckOut()`;
    const result = await this.spClient.spPost(apiUrl);
    return !!result.ok;
  }
  /**
  * Retrieves a single file by its unique ID, including its document type details if available
  * @param uniqueId - The GUID of the file to retrieve
  * @returns A promise resolving to the file object or undefined if not found
  */
  async getFileByUniqueId(uniqueId) {
    const select = "*,FileRef,FileLeafRef,UniqueId";
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyid(guid'${uniqueId}')/ListItemAllFields?$select=${select}`;
    const result = await this.spClient.spGet(apiUrl);
    const item = result.data;
    if (!item) return void 0;
    const documentTypeId = item.DocumentTypeId;
    const file = mapItemAsFile(item);
    if (documentTypeId) {
      const lookupApiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"TypyDokumentu" /* DocumentTypes */}')/items(${documentTypeId})?$select=Id,Title,Title_en`;
      const lookupResult = await this.spClient.spGet(lookupApiUrl);
      const documentType = lookupResult.data;
      file.type = documentType ? {
        loc: {
          ["en-US" /* EN */]: getFileTypeTitle(documentType.Title_en),
          ["cs-CZ" /* CS */]: getFileTypeTitle(documentType.Title)
        },
        id: documentType.Id,
        folderPath: documentType["FolderPath" /* FolderPath */],
        folderPathAfter: documentType["FolderPathAfter" /* FolderPathAfter */],
        validityStatuses: documentType["ValidityStatuses" /* ValidityStatuses */] ? documentType["ValidityStatuses" /* ValidityStatuses */].split(";") : []
      } : void 0;
    }
    return file;
  }
  /**
  * Retrieves the version history for a file based on its list item ID.
  * @param id - The list item ID of the file.
  * @returns A promise resolving to an array of file version objects.
  */
  async getFileVersionHistory(id) {
    const expand = `Version`;
    const select = `*,Version/Id`;
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/items(${id})/versions?$expand=${expand}&$select=${select}`;
    const result = await this.spClient.spGet(apiUrl);
    const teamName = this.spClient.getAbsoluteUrl().split("/").pop();
    return Array.isArray(result.data) ? mapFileVersions(result.data, teamName ?? "") : [];
  }
  /**
   * Retrieves a specific version of a file as an ArrayBuffer.
   * @param relativeUrl The server-relative URL of the file.
   * @param versionId The version ID.
   * @returns Promise resolving to the file content as ArrayBuffer.
   */
  async getFileVersion(relativeUrl, versionId) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${relativeUrl}')/versions(${versionId})/$value`;
    const result = await this.spClient.spGetBinary(apiUrl);
    return result.ok && result.data instanceof ArrayBuffer ? result.data : void 0;
  }
  /**
   * Retrieves a file for download as an ArrayBuffer.
   * @param relativeUrl The server-relative URL of the file.
   * @returns Promise resolving to the file content as ArrayBuffer.
   */
  async getFileToDownload(relativeUrl) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${relativeUrl}')/$value`;
    const result = await this.spClient.spGetBinary(apiUrl);
    return result.ok && result.data instanceof ArrayBuffer ? result.data : void 0;
  }
  /**
   * Archives a file to a designated location, updating metadata.
   * @param file The file object to archive.
   * @param stage Optional stage node.
   * @param phase Optional phase node.
   * @returns Promise resolving to true if archived successfully.
   */
  async archiveFile(file, stage, phase) {
    const sourceSiteUrl = new URL(this.spClient.getAbsoluteUrl()).origin;
    const endpoint = `${sourceSiteUrl}/_api/SP.MoveCopyUtil.CopyFileByPath`;
    const destinationUrl = `${this.spClient.getAbsoluteUrl().replace(this.spClient.getRelativeUrl(), "")}/sites/${this.spClient.internalSiteName}`;
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
      const updateUrl2 = `${destinationUrl}/_api/web/getfilebyserverrelativeurl('${archivedFileRef}')/ListItemAllFields`;
      await this.spClient.spPatch(updateUrl2, { externalFileId: file.uniqueId });
      return true;
    }
    const teamName = await this.spClient.getTeamName();
    const updatedTeamName = teamName.replace("_ext", "");
    const folderPath = !stage?.name.startsWith("02") ? file.type?.folderPathInternal : file.type?.loc["en-US" /* EN */].startsWith("Legal - Fulfillment of Contractual") && (phase?.name?.startsWith("01") || phase?.name?.startsWith("02")) ? file.type?.folderPathInternal : file.type?.folderPathInternalAfter;
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
  async getArchivedFile(uniqueId) {
    const apiUrl = `${this.spClient.getAbsoluteUrl().replace(this.spClient.getRelativeUrl(), "")}/sites/${this.spClient.internalSiteName}/_api/web/lists/getbytitle('Dokumenty')/items?$filter=externalFileId eq '${uniqueId}'&$select=FileRef&$top=1`;
    const result = await this.spClient.spGet(apiUrl);
    return result.ok && Array.isArray(result.data) ? result.data[0]?.FileRef : void 0;
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
  async moveFile(fileRef, fileName, newFileDirRef) {
    const basePath = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}`;
    const path = `${basePath}/${newFileDirRef}`;
    const uniqueName = await this.getUniqueFileName(path, fileName);
    try {
      await this.spClient.sp.web.getFileByServerRelativePath(fileRef).moveByPath(`${path}/${uniqueName}`, false);
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
  async updateFileByUniqueId(uniqueId, body = {}) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/GetItemByUniqueId('${uniqueId}')`;
    const result = await this.spClient.spPatch(apiUrl, body);
    return result;
  }
  /**
   * Updates a file's metadata by server-relative URL.
   * @param fileRef The server-relative URL of the file.
   * @param body Metadata update payload.
   * @returns API result object.
   */
  async updateFileByUrl(fileRef, body = {}) {
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/getfilebyserverrelativeurl('${fileRef}')/ListItemAllFields`;
    const result = await this.spClient.spPatch(apiUrl, body);
    return result;
  }
  async resetAndBreakInheritance(uniqueId) {
    const baseUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/GetItemByUniqueId('${uniqueId}')`;
    await this.spClient.spPost(`${baseUrl}/resetroleinheritance`);
    await this.spClient.spPost(`${baseUrl}/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`);
    await this.clearRoleAssignments(uniqueId);
  }
  async clearRoleAssignments(uniqueId) {
    const webUrl = this.spClient.getAbsoluteUrl();
    const apiRoot = `${webUrl}/_api/web/GetFileById('${uniqueId}')/ListItemAllFields/RoleAssignments`;
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
  async uploadFile(folderPath, file, uploadedBy, folderId, documentTypeId, status) {
    const path = `${this.spClient.getRelativeUrl()}/${this.documentFolder}/${this.spClient.getChannelName()}/${folderPath}`;
    const body = await file.rawFile.arrayBuffer();
    const maxAttempts = 3;
    let attempt = 0;
    let result = { };
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
      const uniqueId = result.data.UniqueId;
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
  async resetFileMetadata(uniqueId, uploadedBy, folderId, documentTypeId, status) {
    const body = {
      Status: status ? status : "",
      StrukturaDokumentaceId: folderId ?? null,
      DocumentTypeId: documentTypeId,
      ContractValidity: null,
      ValidTo: null,
      ScheduledNotificationDate: null,
      ScheduledNotificationMessage: null,
      ScheduledNotificationMessageEn: null,
      ScheduledNotificationRecipients: null,
      Documentation: null,
      ActionToTrigger: null
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
  async reuploadFile(folderPath, file, fileName, uploadedBy, folderId, documentTypeId, status) {
    const url = `${this.spClient.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${folderPath}')/Files/add(url='${fileName}',overwrite=true)`;
    const body = await file.rawFile.arrayBuffer();
    const result = await this.spClient.spPost(url, body);
    if (result.data) {
      const uniqueId = result.data.UniqueId;
      await this.resetFileMetadata(uniqueId, uploadedBy, folderId, documentTypeId, status);
      const uploadedFile = await this.getFileByUniqueId(uniqueId);
      return {
        fileName: file.name,
        success: true,
        file: uploadedFile
      };
    }
    return { fileName: file.name, success: false };
  }
  /**
   * Retrieves required files for a folder by ID.
   * @param id The folder ID.
   * @returns Array of required files.
   */
  async getRequiredFiles(id) {
    const select = `*,FileDirRef,DocumentType/Title,DocumentType/${"Title_en" /* TitleEn */},DocumentType/Id,DocumentType/${"UploadStatus" /* UploadStatus */},DocumentType/${"FolderPath" /* FolderPath */},DocumentType/${"FolderPathAfter" /* FolderPathAfter */}`;
    const expand = "DocumentType";
    const fileDirRef = `${this.spClient.getRelativeUrl()}/Lists/${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.spClient.getChannelName()}/${id}`;
    const filter = `FileDirRef eq '${fileDirRef}'`;
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"PozadovaneDokumenty" /* RequiredDocuments */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const result = await this.spClient.spGet(apiUrl);
    return Array.isArray(result.data) ? mapRequiredFiles(result.data) : [];
  }
  /**
   * Retrieves all required files for the current channel, grouped by folder number.
   * @returns Map of folder number to array of required files.
   */
  async getAllRequiredFiles() {
    const select = `*,FileDirRef`;
    const fileDirRef = `${this.spClient.getRelativeUrl()}/Lists/${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.spClient.getChannelName()}`;
    const filter = `startswith(FileDirRef,'${fileDirRef}')`;
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"PozadovaneDokumenty" /* RequiredDocuments */}')/items?$select=${select}&$filter=${filter}`;
    const result = await this.spClient.spGet(apiUrl);
    const data = Array.isArray(result.data) ? result.data.filter((item) => item.ContentTypeId.startsWith("0x010")) : [];
    const map = /* @__PURE__ */ new Map();
    for (const item of data) {
      const parts = item.FileDirRef.split("/");
      const folderName = parts[parts.length - 1];
      const folderNumber = parseInt(folderName, 10);
      if (isNaN(folderNumber)) continue;
      if (!map.has(folderNumber)) {
        map.set(folderNumber, []);
      }
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
  async getUniqueFileName(folderPath, fileName, countStart = 0, siteUrl = this.spClient.getAbsoluteUrl()) {
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
  async checkFileExists(path, siteUrl = this.spClient.getAbsoluteUrl()) {
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
  async getFolders() {
    const folderName = this.spClient.getAbsoluteUrl().includes("appgestio") ? "Shared documents" : "Sdilene dokumenty";
    const fileDirRef = `${this.spClient.getRelativeUrl()}/${folderName}/${this.spClient.getChannelName()}`;
    const filter = `startswith(FileDirRef,'${fileDirRef}') and startswith(ContentTypeId,'0x0120')`;
    const select = "*,FileDirRef,FileLeafRef,FileRef";
    const apiUrl = `${this.spClient.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/items?$filter=${filter}&$select=${select}`;
    const result = await this.spClient.spGet(apiUrl);
    return Array.isArray(result.data) ? mapFolders(result.data) : [];
  }
  async sendFiletoExtraction(file, type) {
    const transformedPath = transformPath(file.serverRelativeUrl.replace(this.spClient.getRelativeUrl(), ""));
    const origin = new URL(this.spClient.getAbsoluteUrl()).origin;
    const apiUrl = `${origin}/sites/${this.spClient.internalSiteName}/_api/web/lists/getbytitle('Extrakce')/items`;
    const body = {
      Title: transformedPath,
      SiteUrl: this.spClient.getAbsoluteUrl(),
      DocumentType: type,
      FolderPath: file.serverRelativeUrl.split("/").slice(0, -1).join("/").replace(this.spClient.getRelativeUrl(), ""),
      FileId: file.id
    };
    const result = await this.spClient.spPost(apiUrl, body);
    return !!result.ok;
  }
};

// src/api/forms/FormApiClient.ts
var FormApiClient = class {
  /**
   * Constructs a FormApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  /**
   * Retrieves form records from the SharePoint list.
   * @returns Promise<IFormRecord[]> - A promise that resolves to an array of form records
   */
  async getFormRecords() {
    const folderPath = `${this.sp.getRelativeUrl()}/lists/${"Formulare" /* Forms */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Formulare" /* Forms */}')/items?$filter=${filter}`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) ? mapFormRecords(result.data) : [];
  }
  /**
   * Saves the form data to a SharePoint list and updates its status.
   * @param id - The ID of the form record to update
   * @param formData - The form data to be saved
   * @param asDraft - A flag indicating whether the form data should be saved as a draft (optional, defaults to false)
   * @returns Promise<boolean> - A promise that resolves to true if the data is successfully saved, otherwise false
   */
  async saveFormData(id, formData, language, asDraft = false) {
    const body = {
      ["FormData" /* FormData */]: JSON.stringify(formData),
      ["Status" /* Status */]: asDraft ? "new" /* New */ : "submitted" /* Submitted */,
      ["Language" /* Language */]: language
    };
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Formulare" /* Forms */}')/items(${id})`;
    const result = await this.sp.spPatch(apiUrl, body);
    return result.ok ?? false;
  }
};

// src/api/listColumns.ts
var listColumns = {
  ["Detaily" /* Details */]: [
    { FieldTypeKind: "20", Title: "contact" /* Contact */, displayName: "Kontakt" },
    // { FieldTypeKind: '9', Title: ColumnsDetails.Limit, displayName: 'Limit' },
    // { FieldTypeKind: '9', Title: ColumnsDetails.Withdrawn, displayName: 'Čerpání' },
    { FieldTypeKind: "2", Title: "channelId" /* ChannelId */, displayName: "Id kan\xE1lu" },
    { FieldTypeKind: "2", Title: "currentStatus" /* CurrentStatus */, displayName: "Stav" },
    { FieldTypeKind: "3", Title: "loanText" /* LoanText */, displayName: "Text" },
    { FieldTypeKind: "20", Title: "secondaryContact" /* SecondaryContact */, displayName: "Sekund\xE1rn\xED kontakt" },
    // { FieldTypeKind: '2', Title: ColumnsDetails.ContactDesc, displayName: 'Kontakt popis' },
    // { FieldTypeKind: '2', Title: ColumnsDetails.SecondaryDesc, displayName: 'Sekundární popis' },
    { FieldTypeKind: "2", Title: "ChannelStatus" /* ChannelStatus */, displayName: "Stav kan\xE1lu" },
    { FieldTypeKind: "2", Title: "contactData" /* ContactData */, displayName: "Data kontaktu" },
    { FieldTypeKind: "2", Title: "secondaryData" /* SecondaryData */, displayName: "Data sekund\xE1rn\xEDho kontaktu" },
    { FieldTypeKind: "9", Title: "version" /* Version */, displayName: "Verze" }
  ],
  ["Zadosti" /* Requests */]: [
    { FieldTypeKind: "3", Title: "FormData" /* FormData */, displayName: "Data formul\xE1\u0159e" },
    { FieldTypeKind: "7", Title: "RequestType" /* RequestType */, lookUpListTitle: "TypyZadosti" /* RequestTypes */, displayName: "Typ \u017E\xE1dosti" },
    { FieldTypeKind: "6", Title: "RequestStatus" /* RequestStatus */, Choices: Object.values(RequestStatus), displayName: "Stav" },
    { FieldTypeKind: "2", Title: "FileTemplate" /* FileTemplate */, displayName: "\u0160ablona" },
    { FieldTypeKind: "2", Title: "categoryId" /* CategoryId */, displayName: "Id kategorie" },
    { FieldTypeKind: "2", Title: "FileIds" /* FileIds */, displayName: "Id soubor\u016F" },
    { FieldTypeKind: "3", Title: "ExportData" /* ExportData */, displayName: "Export data" }
  ],
  ["QA" /* QA */]: [
    //{ FieldTypeKind: '2', Title: ColumnsQA.Question, displayName: 'Otázka' },
    { FieldTypeKind: "2", Title: "answer" /* Answer */, displayName: "Odpov\u011B\u010F" },
    { FieldTypeKind: "2", Title: "Status", displayName: "Stav" },
    { FieldTypeKind: "6", Title: "Category" /* Category */, Choices: Object.values(QACategory), displayName: "Kategorie" },
    { FieldTypeKind: "2", Title: "FileLink", displayName: "Soubor" },
    { FieldTypeKind: "2", Title: "FolderLink", displayName: "Slo\u017Eka" },
    { FieldTypeKind: "2", Title: "QuestionFileLink", displayName: "Soubor ot\xE1zka" },
    { FieldTypeKind: "2", Title: "QuestionFolderLink", displayName: "Slo\u017Eka ot\xE1zka" },
    { FieldTypeKind: "9", Title: "QAOrder" /* Order */, displayName: "Po\u0159ad\xED" },
    { FieldTypeKind: "2", Title: "Stage", displayName: "Etapa" },
    { FieldTypeKind: "2", Title: "Stage_en", displayName: "Etapa_en" },
    { FieldTypeKind: "2", Title: "Phase", displayName: "F\xE1ze" },
    { FieldTypeKind: "2", Title: "Phase_en", displayName: "F\xE1ze_en" },
    { FieldTypeKind: "2", Title: "Answered" /* Answered */, displayName: "Odpov\u011Bd\u011Bl" },
    { FieldTypeKind: "2", Title: "Priority" /* Priority */, displayName: "Priorita" },
    { FieldTypeKind: "3", Title: "FilesData" /* FilesData */, displayName: "FilesData" },
    { FieldTypeKind: "2", Title: "SubCategory" /* SubCategory */, displayName: "Podkategorie" }
  ],
  ["Kategorie" /* Categories */]: [
    //{ FieldTypeKind: '2', Title: ColumnsCategories.Folder, displayName: 'Složka' },
    { FieldTypeKind: "2", Title: "Stage" /* Stage */, displayName: "Etapa" },
    { FieldTypeKind: "2", Title: "Phase" /* Phase */, displayName: "F\xE1ze" },
    { FieldTypeKind: "2", Title: "Stage_en" /* StageEn */, displayName: "Etapa en" },
    { FieldTypeKind: "2", Title: "Phase_en" /* PhaseEn */, displayName: "F\xE1ze en" },
    { FieldTypeKind: "2", Title: "Category_en" /* CategoryEn */, displayName: "Nadpis en" },
    { FieldTypeKind: "2", Title: "FolderDescription" /* Description */, displayName: "Popis" },
    { FieldTypeKind: "2", Title: "FolderDescription_en" /* DescriptionEn */, displayName: "Popis en" },
    { FieldTypeKind: "11", Title: "RequiredDocuments" /* RequiredDocuments */, displayName: "Po\u017Eadovan\xE9 dokumenty" },
    //{ FieldTypeKind: '4', Title: ColumnsCategories.Deadline, displayName: 'Deadline' },
    { FieldTypeKind: "9", Title: "CategoryOrder" /* Order */, displayName: "Po\u0159ad\xED" },
    { FieldTypeKind: "2", Title: "CategoryStatus" /* CategoryStatus */, displayName: "Stav" }
  ],
  ["Povinnosti" /* Duties */]: [
    { FieldTypeKind: "2", Title: "Date" /* Date */, displayName: "Datum" },
    { FieldTypeKind: "9", Title: "Amount" /* Amount */, displayName: "\u010C\xE1stka" },
    { FieldTypeKind: "6", Title: "EventType" /* EventType */, Choices: Object.values(EventType), displayName: "Typ" },
    { FieldTypeKind: "2", Title: "Currency" /* Currency */, displayName: "M\u011Bna" },
    { FieldTypeKind: "2", Title: "Instruction" /* Instruction */, displayName: "Instrukce" },
    { FieldTypeKind: "6", Title: "Category" /* Category */, displayName: "Kategorie", Choices: ["covenant", "information"] },
    { FieldTypeKind: "2", Title: "EventDescription" /* EventDescription */, displayName: "Popis" },
    { FieldTypeKind: "2", Title: "Note" /* Note */, displayName: "Pozn\xE1mka" },
    { FieldTypeKind: "6", Title: "Period" /* Period */, Choices: Object.values(EventPeriodicity), displayName: "Perioda" },
    { FieldTypeKind: "2", Title: "EndDate" /* EndDate */, displayName: "Konec" }
  ],
  ["Formulare" /* Forms */]: [
    { FieldTypeKind: "3", Title: "FormData" /* FormData */, displayName: "Data formul\xE1\u0159e" },
    { FieldTypeKind: "2", Title: "Title_en" /* TitleEn */, displayName: "Nadpis en" },
    { FieldTypeKind: "6", Title: "Status" /* Status */, Choices: Object.values(FormStatus), displayName: "Stav" },
    { FieldTypeKind: "2", Title: "FormType" /* FormType */, displayName: "Typ formul\xE1\u0159e" },
    { FieldTypeKind: "7", Title: "Category" /* Category */, lookUpListTitle: "Kategorie" /* Categories */, displayName: "Kategorie" },
    { FieldTypeKind: "6", Title: "Language" /* Language */, Choices: Object.values(Language), displayName: "Jazyk" },
    { FieldTypeKind: "2", Title: "Completed" /* Completed */, displayName: "Dokon\u010Deno" },
    { FieldTypeKind: "9", Title: "FileId" /* FileId */, displayName: "File Id" }
  ],
  ["Notification" /* Notifications */]: [
    { FieldTypeKind: "2", Title: "Title_en", displayName: "Nadpis_en" },
    { FieldTypeKind: "8", Title: "Task", displayName: "Task" },
    { FieldTypeKind: "8", Title: "Completed", displayName: "Completed" },
    { FieldTypeKind: "7", Title: "Category", lookUpListTitle: "Kategorie" /* Categories */, displayName: "Category" },
    { FieldTypeKind: "20", Title: "AssignedTo", displayName: "AssignedTo", allowMultiple: true },
    //TODO add multichoice
    { FieldTypeKind: "20", Title: "Read", displayName: "Read", allowMultiple: true },
    { FieldTypeKind: "2", Title: "Comment", displayName: "Comment" },
    { FieldTypeKind: "6", Title: "Status", displayName: "Status", Choices: ["Akceptovat", "Dokon\u010Dit", "Poslat d\xE1l"] },
    { FieldTypeKind: "3", Title: "Assigned", displayName: "Assigned" },
    { FieldTypeKind: "2", Title: "TargetId", displayName: "TargetId" },
    { FieldTypeKind: "2", Title: "Link", displayName: "Link" },
    { FieldTypeKind: "2", Title: "FileLink", displayName: "FileLink" }
  ],
  ["Requests" /* Requests */]: [
    { FieldTypeKind: "2", Title: "RequestType" /* RequestType */, displayName: "Typ \u017E\xE1dosti" },
    { FieldTypeKind: "2", Title: "RequestStatus" /* RequestStatus */, displayName: "Stav" }
  ],
  ["PozadovaneDokumenty" /* RequiredDocuments */]: [
    { FieldTypeKind: "2", Title: "Title_en" /* TitleEn */, displayName: "Nadpis en" },
    //{ FieldTypeKind: '2', Title: ColumnsRequired.FolderPath, displayName: 'Umístění souboru' },
    { FieldTypeKind: "2", Title: "FileDescription" /* FileDesc */, displayName: "Popis souboru" },
    { FieldTypeKind: "7", Title: "DocumentType" /* DocumentType */, lookUpListTitle: "TypyDokumentu" /* DocumentTypes */, displayName: "Typ souboru" },
    { FieldTypeKind: "2", Title: "PermittedFormats" /* PermittedFormats */, displayName: "P\u0159\xEDpustn\xE9 form\xE1ty" },
    { FieldTypeKind: "2", Title: "uploadedFile" /* UploadedFile */, displayName: "ID souboru" },
    { FieldTypeKind: "4", Title: "Deadline" /* Deadline */, displayName: "Deadline" },
    { FieldTypeKind: "2", Title: "Status" /* Status */, displayName: "Stav" }
  ],
  ["TypyDokumentu" /* DocumentTypes */]: [
    { FieldTypeKind: "2", Title: "Title_en" /* TitleEn */, displayName: "Nadpis en" },
    { FieldTypeKind: "2", Title: "UploadStatus" /* UploadStatus */, displayName: "Stav po nahr\xE1n\xED" },
    { FieldTypeKind: "2", Title: "Permissions" /* Permissions */, displayName: "Opr\xE1vn\u011Bn\xED" },
    { FieldTypeKind: "2", Title: "FolderPath" /* FolderPath */, displayName: "Slo\u017Eka" },
    { FieldTypeKind: "2", Title: "FolderPathAfter" /* FolderPathAfter */, displayName: "Slo\u017Eka po schv\xE1len\xED" },
    { FieldTypeKind: "2", Title: "ValidityStatuses" /* ValidityStatuses */, displayName: "Stavy platnosti" },
    { FieldTypeKind: "2", Title: "FolderPathInternal" /* FolderPathInternal */, displayName: "Intern\xED slo\u017Eka" },
    { FieldTypeKind: "2", Title: "FolderPathInternalAfter" /* FolderPathInternalAfter */, displayName: "Intern\xED slo\u017Eka po schv\xE1len\xED" }
  ],
  ["TypyZadosti" /* RequestTypes */]: [
    { FieldTypeKind: "2", Title: "Title_en" /* TitleEn */, displayName: "Nadpis en" },
    { FieldTypeKind: "2", Title: "FileTemplate" /* FileTemplate */, displayName: "\u0160ablona" },
    { FieldTypeKind: "2", Title: "FileTemplateEn" /* FileTemplateEn */, displayName: "\u0160ablona en" },
    { FieldTypeKind: "2", Title: "Description" /* Description */, displayName: "Popis" },
    { FieldTypeKind: "2", Title: "Description_en" /* DescriptionEn */, displayName: "Popis en" },
    { FieldTypeKind: "15", Title: "Etapa" /* Stage */, displayName: "Etapa" },
    { FieldTypeKind: "15", Title: "Fields" /* Fields */, displayName: "Pole" },
    { FieldTypeKind: "8", Title: "FileAttachments" /* FileAttachments */, displayName: "P\u0159\xEDloha" },
    { FieldTypeKind: "2", Title: "Notification" /* Notification */, displayName: "Notifikace" },
    { FieldTypeKind: "8", Title: "FileReply" /* FileReply */, displayName: "P\u0159\xEDloha - odpov\u011B\u010F" }
  ],
  ["WorkflowActions" /* WorkflowActions */]: [
    { FieldTypeKind: "2", Title: "Title_en" /* TitleEn */, displayName: "Nadpis_en" },
    { FieldTypeKind: "3", Title: "DocumentStatus" /* DocumentStatus */, displayName: "Stav dokumentu" },
    { FieldTypeKind: "2", Title: "Notification" /* Notification */, displayName: "Notifikace" },
    { FieldTypeKind: "2", Title: "Notification_en" /* NotificationEn */, displayName: "Notifikace_en" },
    { FieldTypeKind: "2", Title: "ResultStatus" /* ResultStatus */, displayName: "C\xEDlov\xFD stav" },
    { FieldTypeKind: "3", Title: "DocumentType" /* DocumentType */, displayName: "Typ souboru" },
    { FieldTypeKind: "8", Title: "Publish" /* Publish */, displayName: "Publikovat" },
    { FieldTypeKind: "8", Title: "AllowComment" /* AllowComment */, displayName: "Povolit koment\xE1\u0159" },
    { FieldTypeKind: "2", Title: "ActionGroups" /* ActionGroups */, displayName: "Skupiny pro akci" },
    { FieldTypeKind: "2", Title: "Phase" /* Phase */, displayName: "F\xE1ze" },
    { FieldTypeKind: "2", Title: "NotificationGroups" /* NotificationGroups */, displayName: "Skupiny pro notifikaci" },
    { FieldTypeKind: "2", Title: "Task" /* Task */, displayName: "\xDAkol" },
    { FieldTypeKind: "2", Title: "Task_en" /* TaskEn */, displayName: "\xDAkol_en" },
    { FieldTypeKind: "2", Title: "TaskGroups" /* TaskGroups */, displayName: "Skupiny pro \xFAkol" },
    { FieldTypeKind: "3", Title: "ReadPermissions" /* ReadPermissions */, displayName: "Opr\xE1vn\u011Bn\xED pro \u010Dten\xED" },
    { FieldTypeKind: "3", Title: "WritePermissions" /* WritePermissions */, displayName: "Opr\xE1vn\u011Bn\xED upravit" },
    { FieldTypeKind: "6", Title: "validity" /* Validity */, displayName: "Nastaven\xED platnosti", Choices: Object.values(FileValidityAction) },
    { FieldTypeKind: "8", Title: "Reupload" /* Reupload */, displayName: "P\u0159epsat soubor" },
    { FieldTypeKind: "8", Title: "FileLink" /* FileLink */, displayName: "P\u0159idat odkaz" },
    { FieldTypeKind: "8", Title: "onMajorVersion" /* OnMajorVersion */, displayName: "Na hlavn\xED verzi" }
  ],
  ["ActionToTrigger" /* ActionToTrigger */]: [
    { FieldTypeKind: "9", Title: "FileId" /* FileId */, displayName: "FileId" },
    { FieldTypeKind: "9", Title: "ActionId" /* ActionId */, displayName: "ActionId" },
    { FieldTypeKind: "4", Title: "Processed" /* Processed */, displayName: "Processed" },
    { FieldTypeKind: "9", Title: "Result" /* Result */, displayName: "Result" },
    { FieldTypeKind: "2", Title: "Comment" /* Comment */, displayName: "Comment" }
  ],
  ["Logs" /* Logs */]: [
    { FieldTypeKind: "3", Title: "Data", displayName: "Data" }
  ]
};

// src/api/install/InstallApiClient.ts
var InstallApiClient = class {
  /**
   * Constructs an InstallApiClient instance.
   * @param sp The SPApiClient instance for SharePoint operations.
   */
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  documentFolder = "Sdilene dokumenty";
  /**
  * Initializes request-related folders and copies request templates.
  * 
  * Creates a request folder for the current channel and an 'options' folder,
  * calling the provided progress callback after each step.
  * Finally, copies the request templates.
  * 
  * @param increaseProgressCount - A callback function to indicate progress after each step.
  * @returns Promise<void> - A promise that resolves when all initialization steps are complete.
  */
  async initRequests(increaseProgressCount) {
    await this.createRequestFolder(this.sp.getChannelName());
    increaseProgressCount();
    await this.createRequestFolder("options");
    increaseProgressCount();
    await this.copyRequestTemplates();
    increaseProgressCount();
  }
  async ensureIUPData() {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
    await this.ensureIUPChannel(webUrl);
    await this.ensureIUPFolders(webUrl, this.sp.internalSiteName);
  }
  async updateIUPChannel(body) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
    const baseApiUrl = `${webUrl}/_api/web/lists/getbytitle('Channels')/items`;
    const channelRes = await this.sp.spGet(`${baseApiUrl}?$filter=ChannelId eq '${this.sp.getChannelId()}'`);
    const id = Array.isArray(channelRes.data) && typeof channelRes.data[0]?.Id === "number" ? channelRes.data[0].Id : void 0;
    if (id === void 0) {
      return;
    }
    await this.sp.spPatch(`${baseApiUrl}(${id})`, body);
  }
  async updateIUPChannelContacts(email, secondaryEmail) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
    const userId = email ? await this.sp.getUserSPId(email, webUrl) : null;
    const secondaryId = secondaryEmail ? await this.sp.getUserSPId(secondaryEmail, webUrl) : null;
    const body = {
      ContactId: userId,
      SecondaryContactId: secondaryId
    };
    await this.updateIUPChannel(body);
  }
  async updateIUPGroupColumns(oaoEmails, uuoEmails, amlEmails) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
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
  async moveFolderAndChangeTitle(path, newPath, newTitle, listName) {
    const moved = await this.sp.moveFolder(path, newPath);
    if (!moved) {
      return;
    }
    const getItemUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${newPath}')`;
    const itemResponse = await this.sp.spGet(getItemUrl);
    const uniqueId = itemResponse.data ? itemResponse.data.UniqueId : null;
    const body = {
      Title: newTitle
    };
    await this.sp.updateItemInList(uniqueId, body, listName);
  }
  async ensureFolderForAllChannels(basePath, name) {
    const topLevelFolders = await this.getTopLevelFolders();
    for (const channelName of topLevelFolders) {
      const channelFolderPath = `${basePath}/${channelName}`;
      await this.sp.ensureFolder(channelFolderPath, name);
    }
  }
  /**
   * ensures document library and then creates folders
   * @param folderTemplates 
   * @param increaseProgressCount 
   */
  async createFoldersFromTemplates(templateItems, increaseProgressCount) {
    await this.createMainFolderInDocumentStructure();
    const ids = await this.createCategories(templateItems, increaseProgressCount);
    return ids;
  }
  /**
       * Updates the 'RequiredDocuments' links for all category items in the current channel,
       * ensuring the description contains the correct relative path.
       *
       * @returns A promise that resolves when all category items have been updated.
       */
  async updateLinksInCategories() {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"Kategorie" /* Categories */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${fileDirRef}'`;
    const select = `Id,${"RequiredDocuments" /* RequiredDocuments */}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${"Kategorie" /* Categories */}')/items?$select=${select}&$filter=${filter}`;
    const response = await this.sp.spGet(apiUrl);
    const items = response.data;
    if (Array.isArray(items)) {
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const body = {
          ["RequiredDocuments" /* RequiredDocuments */]: {
            Url: item["RequiredDocuments" /* RequiredDocuments */].Url,
            Description: `${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.sp.getChannelName()}/${item.Id}`
            // Optional description
          }
        };
        await this.sp.updateListItem(item.Id, body, "Kategorie" /* Categories */);
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
  async createCategoryItem(templateItem) {
    const path = `/${this.sp.getRelativeUrl()}/Lists/${"Kategorie" /* Categories */}/${this.sp.getChannelName()}`;
    const { stage, phase, category, description, stageEn, categoryEn, phaseEn, descriptionEn, order } = templateItem;
    const body = {
      ["Stage" /* Stage */]: stage,
      ["Phase" /* Phase */]: phase,
      "Title": category,
      //[ColumnsCategories.Folder]: folder,
      ["FolderDescription_en" /* DescriptionEn */]: descriptionEn,
      ["FolderDescription" /* Description */]: description,
      ["Stage_en" /* StageEn */]: stageEn,
      ["Phase_en" /* PhaseEn */]: phaseEn,
      ["Category_en" /* CategoryEn */]: categoryEn,
      //[ColumnsCategories.Deadline]: deadline ? deadline : null,
      ["CategoryOrder" /* Order */]: order ? order : null
    };
    const result = await this.sp.createListItem(body, "Kategorie" /* Categories */);
    if (result.data) {
      const itemId = result.data.Id;
      await this.sp.moveListItem(itemId, path, "Kategorie" /* Categories */);
      const body2 = {
        ["RequiredDocuments" /* RequiredDocuments */]: {
          Url: `${this.sp.getAbsoluteUrl()}/Lists/${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.sp.getChannelName()}/${itemId}`,
          Description: `${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.sp.getChannelName()}/${itemId}`
          // Optional description
        }
      };
      await this.sp.updateListItem(itemId, body2, "Kategorie" /* Categories */);
      return itemId;
    }
  }
  async updateLoanDetailsVersion(version) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Detaily" /* Details */}')/items?$select=*`;
    const response = await this.sp.spGet(apiUrl);
    const data = Array.isArray(response.data) ? response.data : [];
    const details = data.filter((item) => item["version" /* Version */] < version);
    for (let i = 0; i < details.length; i++) {
      const detail = details[i];
      const updated = await this.sp.updateListItem(detail.Id, { ["version" /* Version */]: version }, "Detaily" /* Details */);
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
  async getTemplates(fileRef) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const apiUrl = `${webUrl}/_api/web/GetListByTitle('${"Kategorie" /* DocumentStructure */}')/items?$select=*,FileRef,FileLeafRef,FileDirRef&$filter=FileDirRef eq '${fileRef}'`;
    const response = await this.sp.spGet(apiUrl);
    return Array.isArray(response.data) ? mapTemplates(response.data) : [];
  }
  /**
   * Retrieves the names of templates from the Templates list.
   * @returns An array of template reference objects containing file reference and leaf file reference.
   */
  async getTemplatesNames() {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const apiUrl = `${webUrl}/_api/web/GetListByTitle('${"Kategorie" /* DocumentStructure */}')/items?$select=FileRef,FileLeafRef&$filter=FSObjType eq 1`;
    const response = await this.sp.spGet(apiUrl);
    return Array.isArray(response.data) ? mapTemplateReferences(response.data) : [];
  }
  /**
       * Ensures the existence of necessary folders for different lists under the current channel,
       * invoking a progress increment callback after each folder is created or confirmed.
       *
       * @param increaseProgressCount - Callback function to increment progress tracking.
       * @returns Promise that resolves when all folders have been ensured.
       */
  async createListFolders(increaseProgressCount) {
    await this.ensureListFolder(this.sp.getChannelName(), "Zadosti" /* Requests */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "Formulare" /* Forms */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "Povinnosti" /* Duties */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "PozadovaneDokumenty" /* RequiredDocuments */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "Notification" /* Notifications */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "ActionToTrigger" /* ActionToTrigger */);
    increaseProgressCount();
    await this.ensureListFolder(this.sp.getChannelName(), "Logs" /* Logs */);
    increaseProgressCount();
  }
  async setListReadOnly(listName) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web?$select=AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id&$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup`;
    const groupsResult = await this.sp.spGet(apiUrl);
    if (!groupsResult.data) {
      return;
    }
    const data = groupsResult.data;
    const ownersGroupId = data.AssociatedOwnerGroup && data.AssociatedOwnerGroup.Id ? data.AssociatedOwnerGroup.Id : void 0;
    const membersGroupId = data.AssociatedMemberGroup && data.AssociatedMemberGroup.Id ? data.AssociatedMemberGroup.Id : void 0;
    const visitorsGroupId = data.AssociatedVisitorGroup && data.AssociatedVisitorGroup.Id ? data.AssociatedVisitorGroup.Id : void 0;
    const breakApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
    await this.sp.spPost(breakApiUrl);
    const boundary = "batch_" + Date.now();
    const apiUrlBase = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/roleassignments/addroleassignment`;
    const readPart = [membersGroupId, visitorsGroupId].filter((id) => id !== void 0).map(
      (id) => this.sp.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
    ).join("");
    const writePart = ownersGroupId ? [ownersGroupId].map(
      (id) => this.sp.getPermissionBatchString(boundary, id, 1073741829, apiUrlBase)
    ).join("") : "";
    const batchBody = `${readPart}${writePart}--${boundary}--\r
`;
    await this.sp.spPost(
      `${this.sp.getAbsoluteUrl()}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`
      }
    );
  }
  async ensureListFolder(name, listName, path) {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${listName}${path ? `/${path}` : ""}`;
    const fileRef = `${fileDirRef}/${name}`;
    const checkApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${fileRef}')`;
    const folderCheck = await this.sp.spGet(checkApiUrl, {}, false);
    if (!folderCheck.ok) {
      const createBody = {
        "ContentTypeId": "0x0120",
        "Title": name
      };
      const result = await this.sp.createListItem(createBody, listName);
      const renameBody = {
        "Title": name,
        "FileLeafRef": name
      };
      await this.sp.updateListItem(result.data.Id, renameBody, listName);
      if (path) {
        await this.sp.moveFolder(`${this.sp.getRelativeUrl()}/Lists/${listName}/${name}`, fileRef);
      }
    }
  }
  /**
       * check existence of all necessary columns in document library and create them if needed
       */
  async ensureDocumentLibrary(selectedTemplate, increaseProgressCount) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')`;
    const fieldsApiUrl = `${apiUrl}/fields`;
    const result = await this.sp.spGet(`${fieldsApiUrl}`);
    const fields = result.data;
    const requiredFields = [
      { title: "Status", type: "2" },
      { title: "FileType", type: "2" },
      { title: "StrukturaDokumentaceId", type: "9" },
      { title: "Name_en", type: "2" },
      { title: "Documentation", type: "8" },
      { title: "UploadedBy", type: "2" },
      { type: "4", title: "ValidTo" },
      { title: "ContractValidity", type: "2" },
      { title: "ScheduledNotificationDate", type: "4" },
      { title: "ScheduledNotificationMessage", type: "3" },
      { title: "ScheduledNotificationMessageEn", type: "3" },
      { title: "ScheduledNotificationRecipients", type: "3" },
      { title: "ActionToTrigger", type: "9" },
      { title: "QAId", type: "2" }
    ];
    for (let i = 0; i < requiredFields.length; i++) {
      const reqField = requiredFields[i];
      const field = fields?.find((f) => f.Title === reqField.title);
      if (!field) {
        const body2 = { "FieldTypeKind": reqField.type, "InternalName": reqField.title, "Title": reqField.title };
        await this.sp.spPost(fieldsApiUrl, body2);
      }
    }
    const lookupFieldTitle = "DocumentType";
    const lookupField = fields.find((f) => f.Title === lookupFieldTitle);
    const lookupListUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"TypyDokumentu" /* DocumentTypes */}')`;
    const lookupListResult = await this.sp.spGet(lookupListUrl);
    const lookupListId = lookupListResult.data.Id;
    if (!lookupField) {
      if (lookupListId) {
        const lookupFieldXml = `<Field Type='Lookup' DisplayName='${lookupFieldTitle}' Name='${lookupFieldTitle}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
        const lookupBody = { "parameters": { "SchemaXml": lookupFieldXml } };
        const xmlApiUrl = `${fieldsApiUrl}/CreateFieldAsXml`;
        await this.sp.spPost(xmlApiUrl, lookupBody);
      }
    } else {
      const existingLookupListId = lookupField.LookupList;
      if (existingLookupListId !== lookupListId) {
        const updateFieldXml = `<Field Type='Lookup' DisplayName='${lookupFieldTitle}' Name='${lookupFieldTitle}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
        const updateBody = { "SchemaXml": updateFieldXml };
        const updateApiUrl = `${fieldsApiUrl}/getbytitle('${lookupFieldTitle}')`;
        await this.sp.spPatch(updateApiUrl, updateBody);
      }
    }
    const headers = {
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE"
    };
    const body = { "EnableVersioning": true, "EnableMinorVersions": true, "DraftVersionVisibility": 1 };
    await this.sp.spPost(apiUrl, body, headers);
    await this.ensureFoldersInDocumentLibrary(selectedTemplate, increaseProgressCount);
  }
  /**
       * Adds required files for categories under the current channel and copies template-related files.
       * Ensures each category folder exists in the RequiredDocuments list and copies files based on relations.
       * Calls the provided progress increment callback once the process completes.
       *
       * @param increaseProgressCount - Callback function to increment progress tracking.
       * @param relations - Array of template item relations to copy required files for.
       * @param templateName - Name of the template used for copying files.
       * @returns Promise that resolves when all required files are processed.
       */
  async addRequiredFiles(increaseProgressCount, relations, templateName) {
    const folderPath = `${this.sp.getRelativeUrl()}/lists/${"Kategorie" /* Categories */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${folderPath}'`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Kategorie" /* Categories */}')/items?$select=Id&$filter=${filter}`;
    const result = await this.sp.spGet(apiUrl);
    const idArray = Array.isArray(result.data) ? result.data.map((item) => item.Id) : [];
    for (let i = 0; i < idArray.length; i++) {
      const id = idArray[i];
      await this.ensureListFolder(id.toString(), "PozadovaneDokumenty" /* RequiredDocuments */, this.sp.getChannelName());
    }
    for (let i = 0; i < relations.length; i++) {
      const relation = relations[i];
      await this.copyRequiredfiles(relation, templateName);
    }
    increaseProgressCount();
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
  async ensureListAndCopyItems(listName, templatesListName) {
    const listExists = await this.sp.checkListExistence(listName);
    if (!listExists) {
      const result = await this.sp.createList(listName);
      if (!result) {
        return;
      }
    }
    const missingColumns = await this.sp.ensureColumns(listName);
    if (!listExists) {
      const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
      const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items`;
      const result = await this.sp.spGet(apiUrl);
      if (Array.isArray(result.data)) {
        const itemBodies = [];
        for (let i = 0; i < result.data.length; i++) {
          const item = result.data[i];
          const body = {
            Title: item.Title
          };
          listColumns[listName].forEach((col) => {
            if ((col.Title === "FileTemplate" /* FileTemplate */ || col.Title === "FileTemplateEn" /* FileTemplateEn */) && listName === "TypyZadosti" /* RequestTypes */) {
              const fileTemplate = item[col.Title] ? item[col.Title].split(";").map(
                (part) => part.replace(
                  `/sites/${this.sp.templateSiteName}/Requests`,
                  `${this.sp.getRelativeUrl()}/Requests/options`
                )
              ).join(";") : "";
              body[col.Title] = fileTemplate;
            } else if (col.Title === "Fields" /* Fields */ && listName === "TypyZadosti" /* RequestTypes */) {
              body[col.Title] = item[col.Title] ? item[col.Title] : [];
            } else {
              body[col.Title] = item[col.Title];
            }
          });
          itemBodies.push(body);
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
  async copyFoldersAndItemsWithStructure(sourceListName, targetListName, fields) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const sourceApiUrl = `${webUrl}/_api/web/lists/getbytitle('${sourceListName}')/items?$select=*,Id,Title,FileSystemObjectType,FileDirRef,FileLeafRef,Folder&$orderby=ID&$top=5000`;
    const result = await this.sp.spGet(sourceApiUrl);
    if (!Array.isArray(result.data)) return;
    const baseSourceFolder = `/sites/${this.sp.templateSiteName}/Lists/${sourceListName}`;
    for (const item of result.data.filter((i) => i.FileSystemObjectType === 1)) {
      const folderName = item.FileLeafRef;
      await this.ensureListFolder(folderName, targetListName);
    }
    for (const item of result.data.filter((i) => i.FileSystemObjectType !== 1)) {
      const sourcePath = item.FileDirRef;
      const relativePath = sourcePath.toLowerCase().startsWith(baseSourceFolder.toLowerCase()) ? sourcePath.slice(baseSourceFolder.length) : sourcePath;
      const formValues = [{
        "FieldName": "Title",
        "FieldValue": item.Title
      }];
      fields.forEach((field) => {
        let value = item[field];
        if (typeof value === "boolean") {
          value = value ? "true" : "false";
        } else if (value === null || value === void 0) {
          value = "";
        } else {
          value = String(value);
        }
        formValues.push({
          "FieldName": field,
          "FieldValue": value
        });
      });
      const body = { "formValues": formValues };
      await this.sp.createListItemUsingPath(body, targetListName, relativePath, false);
    }
  }
  async ensureListAndCopyItemsIntoChannelFolder(listName, templatesListName, templateName) {
    const listExists = await this.sp.checkListExistence(listName);
    if (!listExists) {
      const result = await this.sp.createList(listName);
      if (!result) {
        return;
      }
    }
    const missingColumns = await this.sp.ensureColumns(listName);
    await this.ensureListFolder(this.sp.getChannelName(), "QA" /* QA */);
    const itemsExists = await this.checkCopiedItemsExistence(listName);
    if (!itemsExists) {
      const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
      const folderPath = `/sites/${this.sp.templateSiteName}/lists/${templatesListName}/${templateName}`;
      const filter = `FileDirRef eq '${folderPath}'`;
      const select = "*,FileDirRef";
      const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items?$select=${select}&$filter=${filter}`;
      const result = await this.sp.spGet(apiUrl);
      if (Array.isArray(result.data)) {
        const itemBodies = [];
        for (let i = 0; i < result.data.length; i++) {
          const item = result.data[i];
          const formValues = [{
            "FieldName": "Title",
            "FieldValue": item.Title
          }];
          listColumns[listName].forEach((col) => {
            let value = item[col.Title];
            if (typeof value === "boolean") {
              value = value ? "true" : "false";
            } else if (value === null || value === void 0) {
              value = "";
            } else {
              value = String(value);
            }
            formValues.push({
              "FieldName": col.Title,
              "FieldValue": value
            });
          });
          const body = { "formValues": formValues };
          itemBodies.push(body);
        }
        await this.createListItemsBatchUsingPath(listName, itemBodies);
      }
      await this.sp.ensureListView(listName, missingColumns);
    }
  }
  async copyTemplateAttachments(templatesListName, listName) {
    const itemTitle = "\u0161ablony";
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${templatesListName}')/items?$filter=Title eq '${itemTitle}'`;
    const result = await this.sp.spGet(apiUrl);
    const id = Array.isArray(result.data) && result.data[0] ? result.data[0].Id : void 0;
    if (!id) {
      return;
    }
    const attachUrl = `${webUrl}/_api/lists/getbytitle('${templatesListName}')/items(${id})/AttachmentFiles`;
    const attachResult = await this.sp.spGet(attachUrl);
    const files = Array.isArray(attachResult.data) ? await Promise.all(attachResult.data.map(async (item) => {
      const result2 = await this.sp.spGetBinary(`${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}${item.ServerRelativeUrl}`);
      await this.sp.spGetBinary(`${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}${item.ServerRelativeUrl}`);
      const arrayBuffer = result2.ok && result2.data instanceof ArrayBuffer ? result2.data : void 0;
      return {
        fileName: item.FileName,
        arrayBuffer
      };
    })) : [];
    const filter = `Title eq '${itemTitle}' and FileDirRef eq '${this.sp.getRelativeUrl()}/Lists/${listName}'`;
    const checkApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/items?$filter=${filter}&$select=*,FileDirRef`;
    const checkResult = await this.sp.spGet(checkApiUrl);
    let itemid = Array.isArray(checkResult.data) && checkResult.data[0] ? checkResult.data[0].Id : void 0;
    if (!itemid) {
      const body = { Title: itemTitle };
      const createResult = await this.sp.createListItem(body, listName);
      itemid = createResult.data ? createResult.data.Id : void 0;
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
  async createRequestFolder(name) {
    const path = `${this.sp.getRelativeUrl()}${"Requests" /* Requests */}/`;
    return await this.sp.ensureFolder(path, name) !== void 0;
  }
  async createListItemsBatchUsingPath(listName, itemBodies, path, channelSpecific = true) {
    const chunkSize = 50;
    const chunks = this.sp.chunkArray(itemBodies, chunkSize);
    const results = [];
    for (const chunk of chunks) {
      const result = await this.createListItemsBatchChunkUsingPath(listName, chunk, path, channelSpecific);
      results.push(result);
      await delay(200);
    }
    return {
      ok: results.every((r) => r.ok),
      data: results.flatMap((r) => r.data ?? []),
      error: results.find((r) => !r.ok)?.error ?? void 0
    };
  }
  async createListItemsBatchChunkUsingPath(listName, itemBodies, path, channelSpecific = true) {
    const boundary = "batch_" + Date.now();
    let folderPath = `${this.sp.getAbsoluteUrl()}/lists/${listName}`;
    if (channelSpecific) {
      folderPath += `/${this.sp.getChannelName()}`;
    }
    if (path) {
      folderPath += `${path}`;
    }
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/AddValidateUpdateItemUsingPath`;
    const batchParts = itemBodies.map((body) => {
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
        "Content-Type: application/http",
        "Content-Transfer-Encoding: binary",
        "",
        `POST ${apiUrl} HTTP/1.1`,
        "Content-Type: application/json",
        "Accept: application/json",
        "",
        JSON.stringify(bodyWithPath),
        ""
      ].join("\r\n");
    });
    const batchBody = [...batchParts, `--${boundary}--\r
`].join("\r\n");
    const response = await this.sp.spPost(
      `${this.sp.getAbsoluteUrl()}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`
      }
    );
    return response;
  }
  async checkCopiedItemsExistence(listName) {
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
  async copyRequiredfiles(relation, templateName) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const fileDirRef = `/sites/${this.sp.templateSiteName}/Lists/${"PozadovaneDokumenty" /* RequiredDocuments */}/${templateName}/${relation.templateId}`;
    const filter = `FileDirRef eq '${fileDirRef}'`;
    const select = "*,FileDirRef,DocumentType/Title,DocumentType/Title_en,DocumentType/Id";
    const expand = "DocumentType";
    const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${"PozadovaneDokumenty" /* RequiredDocuments */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const templates = await this.sp.spGet(apiUrl);
    const path = `${this.sp.getRelativeUrl()}/Lists/${"PozadovaneDokumenty" /* RequiredDocuments */}/${this.sp.getChannelName()}/${relation.itemId}`;
    const docTypesUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"TypyDokumentu" /* DocumentTypes */}')/items`;
    const docTypes = await this.sp.spGet(docTypesUrl);
    const docTypeMap = new Map(
      docTypes.data.map((t) => [
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
          ["Title_en" /* TitleEn */]: item["Title_en" /* TitleEn */],
          ["FileDescription" /* FileDesc */]: item["FileDescription" /* FileDesc */],
          ["DocumentTypeId" /* DocumentTypeId */]: typeId
        };
        const result = await this.sp.createListItem(body, "PozadovaneDokumenty" /* RequiredDocuments */);
        if (result.data) {
          await this.sp.moveListItem(result.data.Id, path, "PozadovaneDokumenty" /* RequiredDocuments */);
        }
      }
    }
  }
  /**
  * Ensures that all necessary folders based on the selected template exist in the document library.
  * Calls folder creation for each folder template and updates progress via callback.
  * 
  * @param selectedTemplate - The name or identifier of the selected folder template.
  * @param increaseProgressCount - Callback function to update progress after each folder is ensured.
  */
  async ensureFoldersInDocumentLibrary(selectedTemplate, increaseProgressCount) {
    const folderTemplates = await this.getFolderTemplates(selectedTemplate);
    const siteGroups = await this.sp.getSiteGroups(this.sp.getAbsoluteUrl());
    await this.generateFoldersFromTemplates(folderTemplates, "", increaseProgressCount, siteGroups);
  }
  /**
   * Retrieves folder templates for the given selected template from a specific team's document library.
   *
   * @param selectedTemplate - The name of the selected folder template to retrieve.
   * @returns Promise resolving to an array of folder templates ({@link IFolderTemplate}).
   */
  async getFolderTemplates(selectedTemplate) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
    const fileDirRef = `/sites/${this.sp.templateSiteName}/${this.documentFolder}/${selectedTemplate}`;
    const apiUrl = `${webUrl}/_api/web/GetListByTitle('${"Dokumenty" /* Documents */}')/items?$select=*,FileRef,FileLeafRef,FileDirRef&$filter=startswith(FileDirRef,'${fileDirRef}')`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) ? mapFolderTemplates(result.data) : [];
  }
  /**
       * recursively generate folders from templates loaded from site EUP-Sablony
       * @param folderTemplates 
       * @param parentFolderPath 
       * @param increaseProgressCount 
       */
  async generateFoldersFromTemplates(folderTemplates, parentFolderPath, increaseProgressCount, siteGroups) {
    for (let i = 0; i < folderTemplates.length; i++) {
      const folderTemplate = folderTemplates[i];
      const uniqueId = await this.createFolder(folderTemplate, parentFolderPath);
      if (folderTemplate.nameEn && uniqueId) {
        const body = {
          Name_en: folderTemplate.nameEn
        };
        await this.sp.updateItemInList(uniqueId, body, "Dokumenty" /* Documents */);
      }
      if (uniqueId && (folderTemplate.write || folderTemplate.read)) {
        const folderRelativeUrl = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}${parentFolderPath}/${folderTemplate.name}`;
        await this.setFolderPermissions(folderTemplate, folderRelativeUrl, siteGroups, uniqueId, this.sp.getAbsoluteUrl());
      }
      if (folderTemplate.templates) {
        const path = `${parentFolderPath}/${folderTemplate.name}`;
        await this.generateFoldersFromTemplates(folderTemplate.templates, path, increaseProgressCount, siteGroups);
      } else {
        increaseProgressCount();
      }
    }
  }
  async setFolderPermissions(folderTemplate, folderRelativeUrl, siteGroups, uniqueId, webUrl) {
    const writeGroups = folderTemplate.write ? expandRoles(folderTemplate.write) : [];
    const readGroupsAll = folderTemplate.read ? expandRoles(folderTemplate.read) : [];
    const writeGroupIds = siteGroups.filter((g) => writeGroups.includes(g.title)).map((g) => g.id);
    const apiUrl = `${webUrl}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/GetItemByUniqueId('${uniqueId}')/breakroleinheritance(copyRoleAssignments=false, clearSubscopes=true)`;
    await this.sp.spPost(apiUrl);
    const readGroups = [...readGroupsAll].filter((g) => !writeGroups.includes(g));
    const readGroupIds = siteGroups.filter((g) => readGroups.includes(g.title)).map((g) => g.id);
    await this.assignFolderPermissions(folderRelativeUrl, readGroupIds, writeGroupIds, webUrl);
  }
  async assignFolderPermissions(folderRelativeUrl, readGroupsIds, writeGroupIds, webUrl) {
    const boundary = "batch_" + Date.now();
    const apiUrlBase = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${folderRelativeUrl}')/ListItemAllFields/roleassignments/addroleassignment`;
    const readPart = readGroupsIds.map(
      (id) => this.sp.getPermissionBatchString(boundary, id, 1073741826, apiUrlBase)
    ).join("");
    const writePart = writeGroupIds.map(
      (id) => this.sp.getPermissionBatchString(boundary, id, 1073741827, apiUrlBase)
    ).join("");
    const batchBody = `${readPart}${writePart}--${boundary}--\r
`;
    await this.sp.spPost(
      `${webUrl}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`
      }
    );
  }
  /**
  * creates new subfolder with custom attributes in folder using serverRelativeUrl
  * @param folder 
  * @param parentFolderPath 
  * @returns boolean
  */
  async createFolder(folder, parentFolderPath) {
    const path = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}/${parentFolderPath}/`;
    return await this.sp.ensureFolder(path, folder.name);
  }
  /**
  * Creates a folder structure based on the provided template items and updates progress.
  * @param templateItems The template content items used to create the folder structure.
  * @param increaseProgressCount A callback function to update the progress count.
  * @returns A promise indicating the completion of the folder structure creation.
  */
  async createCategories(templateItems, increaseProgressCount) {
    const ids = [];
    for (let i = 0; i < templateItems.length; i++) {
      const itemId = await this.createCategoryItem(templateItems[i]);
      if (itemId) {
        ids.push({ itemId, templateId: templateItems[i].id });
      }
      increaseProgressCount();
    }
    return ids;
  }
  async getTopLevelFolders() {
    const path = `${this.sp.getRelativeUrl()}/${this.documentFolder}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativeUrl('${path}')/Folders?$select=Name,ServerRelativeUrl`;
    const result = await this.sp.spGet(apiUrl);
    if (!result.ok || !Array.isArray(result.data)) {
      return [];
    }
    return result.data.filter((f) => !f.Name.startsWith("_")).map((f) => f.ServerRelativeUrl);
  }
  /**
   * Creates the main folder in the document structure for the current channel.
   * @returns A promise indicating the completion of the folder creation.
   */
  async createMainFolderInDocumentStructure() {
    const folderPath = encodeURIComponent(`${this.sp.getRelativeUrl()}/Lists/${"Kategorie" /* Categories */}`);
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a1)/AddSubFolderUsingPath(DecodedUrl=@a2)?@a1='${folderPath}'&@a2='${this.sp.getChannelName()}'&siteUrl=${this.sp.getAbsoluteUrl()}`;
    await this.sp.spPost(apiUrl);
  }
  async ensureIUPFolders(webUrl, iupTeamName) {
    const siteGroups = await this.sp.getSiteGroups(webUrl);
    const teamName = await this.sp.getTeamName();
    const updatedTeamName = teamName.replace("_ext", "");
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
        const templates = await this.getFolderTemplates("Intern\xED ulo\u017Ei\u0161t\u011B");
        await this.createIUPFolders(`${path}/${this.sp.getChannelName()}`, templates, webUrl, iupTeamName, siteGroups);
      }
    }
  }
  async createIUPFolders(parentFolderPath, templates, webUrl, iupTeamName, siteGroups) {
    for (let i = 0; i < templates.length; i++) {
      const template = templates[i];
      const apiUrl = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${parentFolderPath}')/folders/add('${template.name}')`;
      await this.sp.spPost(apiUrl);
      if (template.templates) {
        await this.createIUPFolders(`${parentFolderPath}/${template.name}`, template.templates, webUrl, iupTeamName, siteGroups);
      }
    }
  }
  async ensureIUPChannel(webUrl) {
    const apiUrl = `${webUrl}/_api/web/lists/getbytitle('Channels')/items`;
    const channelDataExists = await this.sp.spGet(`${apiUrl}?$filter=ChannelId eq '${this.sp.getChannelId()}'`);
    const detail = await new DetailApiClient(this.sp).getLoanDetail();
    if (!channelDataExists.ok || channelDataExists.data.length === 0) {
      const body = {
        Title: this.sp.getChannelName(),
        TeamId: this.sp.getTeamId(),
        ChannelId: this.sp.getChannelId(),
        TeamSiteUrl: this.sp.getAbsoluteUrl(),
        ChannelStatus: "Active",
        Stage: detail?.currentStatus || "",
        Phase: detail?.currentFolder || ""
        //ChannelType: isPrivate ? 'Private' : 'Standart'
      };
      await this.sp.spPost(apiUrl, body);
      await this.updateIUPChannelContacts(detail?.contact?.email, detail?.secondaryContact?.email);
    } else if (channelDataExists.ok && Array.isArray(channelDataExists.data) && channelDataExists.data[0]) {
      const body = {
        ChannelStatus: "Active",
        Stage: detail?.currentStatus || "",
        Phase: detail?.currentFolder || "",
        TeamSiteUrl: this.sp.getAbsoluteUrl()
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
  async copyRequestTemplates() {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}`;
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
  async copyRequestTemplate(fileRef, fileLeafRef, webUrl) {
    const downloadUrl = `${webUrl}/_api/web/getfilebyserverrelativeurl('${fileRef}')/$value`;
    const result = await this.sp.spGetBinary(downloadUrl);
    const fileContent = result.ok && result.data instanceof ArrayBuffer ? result.data : void 0;
    const uploadUrl = `${this.sp.getAbsoluteUrl()}/_api/web/getfolderbyserverrelativeurl('${this.sp.getRelativeUrl()}/Requests/options')/files/add(overwrite=true, url='${fileLeafRef}')`;
    if (fileContent) {
      await this.sp.spPostBinary(uploadUrl, fileContent);
    }
  }
  async updateDateColumns() {
    const docApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Dokumenty" /* Documents */}')/fields/GetByTitle('ScheduledNotificationDate')`;
    await this.sp.spPatch(docApiUrl, { "DisplayFormat": 0 });
    const reqApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"PozadovaneDokumenty" /* RequiredDocuments */}')/fields/GetByTitle('${"Deadline" /* Deadline */}')`;
    await this.sp.spPatch(reqApiUrl, { "DisplayFormat": 0 });
  }
};
var NotificationItemSchema = z.object({
  Id: z.number(),
  Title: z.string().nullable().optional(),
  Title_en: z.string().nullable().optional(),
  Task: z.boolean().nullable().optional(),
  Completed: z.boolean().nullable().optional(),
  Created: z.string(),
  Modified: z.string().nullable().optional(),
  Assigned: z.string().nullable().optional(),
  ReadId: z.string().nullable().optional(),
  TargetId: z.string().nullable().optional(),
  Status: z.string().nullable().optional(),
  Comment: z.string().nullable().optional(),
  Link: z.string().nullable().optional(),
  UniqueId: z.string().nullable().optional(),
  FileLink: z.string().nullable().optional(),
  Attachments: z.boolean().nullable().optional(),
  Author: z.object({
    Title: z.string().nullable().optional()
  }).nullish().transform((author) => ({
    Title: author?.Title ?? ""
  })),
  Editor: z.object({
    Title: z.string().nullable().optional()
  }).nullable().optional(),
  Category: z.object({
    Id: z.number().nullable().optional(),
    Title: z.string().nullable().optional(),
    ["Category_en" /* CategoryEn */]: z.string().nullable().optional(),
    ["Stage" /* Stage */]: z.string().nullable().optional(),
    ["Stage_en" /* StageEn */]: z.string().nullable().optional(),
    ["Phase" /* Phase */]: z.string().nullable().optional(),
    ["Phase_en" /* PhaseEn */]: z.string().nullable().optional()
  }).nullable().optional()
});
var parseNotificationItem = (value) => {
  if (!value) {
    return void 0;
  }
  const parsed = NotificationItemSchema.safeParse(value);
  return parsed.success ? parsed.data : void 0;
};
var parseNotificationItems = (value) => {
  if (!value || !Array.isArray(value)) {
    return [];
  }
  return value.map((item) => parseNotificationItem(item)).filter((item) => item !== void 0);
};
var NotificationApiClient = class {
  constructor(sp, webPartContext) {
    this.sp = sp;
    this.webPartContext = webPartContext;
  }
  sp;
  webPartContext;
  /**
       * Retrieves and filters notifications assigned to the specified user groups from the SharePoint Notifications list.
       * @param groups - An array of group identifiers to filter notifications by assignment.
       * @returns A Promise resolving to a sorted array of notifications relevant to the user, mapped with user-specific data.
       */
  async getAllNotifications(groups) {
    const folderPath = `${this.sp.getRelativeUrl()}/lists/${"Notification" /* Notifications */}/${this.sp.getChannelName()}`;
    const id = await this.sp.getUserSPId(this.webPartContext.pageContext.user.email);
    const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
    const expand = "Category,Author";
    const select = `*,Category/Title,Category/${"Phase" /* Phase */},Category/${"Stage" /* Stage */},Author/Title,Category/Id,Category/${"Category_en" /* CategoryEn */},Category/${"Phase_en" /* PhaseEn */},Category/${"Stage_en" /* StageEn */},UniqueId`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Notification" /* Notifications */}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
    const result = await this.sp.spGet(apiUrl);
    const parsedItems = parseNotificationItems(result.data);
    const data = parsedItems.filter((item) => item.Assigned && item.Assigned.split(";").some((e) => groups.includes(e)));
    data.sort((a, b) => {
      return new Date(b.Created).getTime() - new Date(a.Created).getTime();
    });
    const returnData = mapNotifications(data, id);
    return returnData;
  }
  /**
   * Creates a new notification item in the SharePoint Notifications list using provided form values.
   * @param notification - The notification data including title, task, category, assigned groups, and file reference.
   * @returns A Promise resolving to true if the item was successfully created, or false otherwise.
   */
  async createNotification(notification) {
    const body = {
      "formValues": [
        {
          "FieldName": "Title",
          "FieldValue": notification.title
        },
        {
          "FieldName": "Title_en",
          "FieldValue": notification.titleEn
        },
        {
          "FieldName": "Task",
          "FieldValue": `${notification.task}`
        },
        {
          "FieldName": "Category",
          "FieldValue": notification.categoryId.toString()
        },
        {
          "FieldName": "Assigned",
          "FieldValue": notification.userGroups
        },
        {
          "FieldName": "TargetId",
          "FieldValue": notification.targetId
        },
        {
          "FieldName": "Link",
          "FieldValue": notification.link
        },
        {
          "FieldName": "Comment",
          "FieldValue": notification.comment
        },
        {
          "FieldName": "FileLink",
          "FieldValue": notification.fileLink
        }
      ]
    };
    const result = await this.sp.createListItemUsingPath(body, "Notification" /* Notifications */);
    return !!result.ok;
  }
  /**
   * Marks a task as completed in the SharePoint Notifications list and optionally attaches a file.
   * @param id - The ID of the notification item to update.
   * @param body - An object containing the updated field values for the task.
   * @param file - (Optional) A file to be attached to the notification item.
   * @returns A Promise resolving to true if the update (and optional file attachment) was successful, or false otherwise.
   */
  async completeTask(id, body, files) {
    const updated = await this.sp.updateListItem(id, body, "Notification" /* Notifications */);
    if (!updated || files.length === 0) {
      return updated;
    }
    let result = true;
    for (const file of files) {
      const addedFile = await this.sp.addAttachmentFile(id, file, "Notification" /* Notifications */);
      if (!addedFile.ok) {
        result = false;
      }
    }
    return result;
  }
  async deleteNotifications(notifications) {
    for (let i = 0; i < notifications.length; i++) {
      const notification = notifications[i];
      await this.sp.deleteListItem(notification.id, "Notification" /* Notifications */);
    }
  }
  async deleteNotificationsByTargetId(targetId, link) {
    const filter = `TargetId eq '${targetId}' and Task ne 1${link ? ` and Link eq '${link}'` : ""}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"Notification" /* Notifications */}')/items?$filter=${filter}`;
    const result = await this.sp.spGet(apiUrl);
    const data = Array.isArray(result.data) ? result.data : [];
    for (let i = 0; i < data.length; i++) {
      const id = data[i].Id;
      await this.sp.deleteListItem(id, "Notification" /* Notifications */);
    }
  }
  async completeTasksByTargetId(targetId, link, resultStatus) {
    const filter = `TargetId eq '${targetId}' and Task eq 1${link ? ` and Link eq '${link}'` : ""}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"Notification" /* Notifications */}')/items?$filter=${filter}`;
    const result = await this.sp.spGet(apiUrl);
    const data = Array.isArray(result.data) ? result.data : [];
    for (let i = 0; i < data.length; i++) {
      const id = data[i].Id;
      const body = { Completed: true, Status: resultStatus ? resultStatus : null };
      await this.sp.updateListItem(id, body, "Notification" /* Notifications */);
    }
  }
  async getTaskByTargetId(targetId) {
    const filter = `TargetId eq '${targetId}' and Task eq 1 and Completed ne 1`;
    const expand = "Category,Author,Editor";
    const select = `*,UniqueId,Category/Title,Category/${"Phase" /* Phase */},Category/${"Stage" /* Stage */},Author/Title,Category/Id,Editor/Title`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"Notification" /* Notifications */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const result = await this.sp.spGet(apiUrl);
    const parsedItems = parseNotificationItems(result.data);
    const mapped = mapNotifications(parsedItems);
    return mapped[0];
  }
  /**
   * Retrieves the history of completed tasks associated with a specific file ID from the SharePoint Notifications list.
   * @param fileId - The ID of the file to filter task history by.
   * @returns A Promise resolving to a chronologically sorted array of notifications with mapped attachments and metadata.
   */
  async getTaskHistory(fileId) {
    const filter = `TargetId eq '${fileId}' and Task eq 1`;
    const expand = "Category,Author,Editor";
    const select = `*,UniqueId,Category/Title,Category/${"Phase" /* Phase */},Category/${"Stage" /* Stage */},Author/Title,Category/Id,Editor/Title`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"Notification" /* Notifications */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const result = await this.sp.spGet(apiUrl);
    const parsedItems = parseNotificationItems(result.data);
    const sorted = parsedItems.sort((a, b) => {
      return new Date(a.Created).getTime() - new Date(b.Created).getTime();
    });
    const attachmentsMap = await this.getAttachmentsMap(sorted);
    return mapNotifications(sorted, void 0, attachmentsMap);
  }
  /**
  * Builds a map of attachment metadata for each item in the provided list.
  * @param items - An array of SharePoint list items to retrieve attachments for.
  * @returns A Promise resolving to a map where each key is an item ID and the value is an array of attachment objects containing file names and URLs.
  */
  async getAttachmentsMap(items) {
    const attachmentsMap = /* @__PURE__ */ new Map();
    await Promise.all(items.map(async (item) => {
      if (item.Attachments) {
        const attachUrl = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"Notification" /* Notifications */}')/items(${item.Id})/AttachmentFiles`;
        const attachResult = await this.sp.spGet(attachUrl);
        const attachments = (Array.isArray(attachResult.data) ? attachResult.data : []).map((att) => ({
          fileName: att.FileName,
          url: att.ServerRelativeUrl
        }));
        attachmentsMap.set(item.Id, attachments);
      }
    }));
    return attachmentsMap;
  }
};

// src/api/mappers/parser.ts
function isQnAFilesData(obj) {
  if (!obj || typeof obj !== "object") return false;
  return Object.values(obj).every(isFileArray);
}
function isFileArray(val) {
  if (!Array.isArray(val)) return false;
  return val.every(isValidFileItem);
}
function isValidFileItem(item) {
  if (!item || typeof item !== "object") return false;
  const f = item;
  return typeof f.name === "string" && typeof f.serverRelativeUrl === "string";
}
var parseQnAFilesData = (value) => {
  if (!value) return void 0;
  try {
    const parsed = JSON.parse(value);
    if (!isQnAFilesData(parsed)) {
      return void 0;
    }
    return parsed;
  } catch (error) {
    console.warn("Failed to parse FilesData:", error);
    return void 0;
  }
};

// src/api/qnA/QnAApiClient.ts
var QnAApiClient = class {
  constructor(files, sp) {
    this.files = files;
    this.sp = sp;
  }
  files;
  sp;
  documentFolder = "Sdilene dokumenty";
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
  async createQnAItem(question, category, stage, phase, priority, uploadedBy, fileType, order, files, subCategory) {
    const body = {
      "formValues": [
        {
          "FieldName": "Title",
          "FieldValue": question
        },
        {
          "FieldName": "Category",
          "FieldValue": category
        },
        {
          "FieldName": "Status",
          "FieldValue": "new"
        },
        {
          "FieldName": "Stage",
          "FieldValue": stage
        },
        {
          "FieldName": "Phase",
          "FieldValue": phase
        },
        {
          "FieldName": "QAOrder",
          "FieldValue": order ? order.toString() : ""
        },
        {
          "FieldName": "Priority" /* Priority */,
          "FieldValue": priority
        },
        {
          "FieldName": "SubCategory" /* SubCategory */,
          "FieldValue": subCategory
        }
      ]
    };
    const result = await this.sp.createListItemUsingPath(body, "QA" /* QA */);
    const value = result.data && Array.isArray(result.data.value) ? result.data.value : [];
    const idField = value.find((item) => item.FieldName === "Id");
    if (files && files.length > 0) {
      const addedFiles = await this.uploadQAFiles(files, uploadedBy, Number(idField.FieldValue), fileType);
      const filesData = { ["Question" /* Question */]: [...addedFiles] };
      if (result.ok) {
        const body2 = { ["FilesData" /* FilesData */]: JSON.stringify(filesData) };
        await this.sp.updateListItem(idField.FieldValue, body2, "QA" /* QA */);
      }
    }
    return idField ? idField.FieldValue : void 0;
  }
  /**
   * Adds a comment to a list item based on its unique ID.
   * @param uniqueId The unique identifier of the list item.
   * @param listName The name of the list where the item exists.
   * @param text The comment text to add.
   * @returns A boolean indicating whether the comment was successfully added.
   */
  async addComment(qnAItem, listName, text, files, uploadedBy) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${qnAItem.uniqueId}')/Comments()`;
    const body = {
      "text": text
    };
    const response = await this.sp.spPost(apiUrl, body);
    let filesData = qnAItem.filesData;
    const fileType = files[0] ? files[0].type : void 0;
    if (files.length > 0 && fileType) {
      const addedFiles = await this.uploadQAFiles(files, uploadedBy, qnAItem.id, fileType, Number(response.data.id));
      filesData = {
        ...qnAItem.filesData ?? {},
        ["Comment" /* Comment */]: [
          ...qnAItem.filesData?.["Comment" /* Comment */] ?? [],
          ...addedFiles
        ]
      };
      await this.sp.updateItemInList(
        qnAItem.uniqueId,
        { ["FilesData" /* FilesData */]: JSON.stringify(filesData) },
        "QA" /* QA */
        /*, file ?? [] */
      );
    }
    return response.ok && response.data ? { newComment: this.mapItemComments([response.data])[0], filesData } : { newComment: void 0, filesData: void 0 };
  }
  /**
   * Uploads Q&A files to the specified folder and updates file metadata.
   * @param files Array of files to upload.
   * @param uploadedBy User uploading the files.
   * @param qnAItemId Q&A item ID for association.
   * @param fileType File type information (optional).
   * @param commentId Comment ID for association (optional).
   * @returns Promise resolving to an array of IQnAAttachment objects.
   */
  async uploadQAFiles(files, uploadedBy, qnAItemId, fileType, commentId) {
    if (!fileType || !files?.length) {
      return [];
    }
    const { name, path } = splitFolderPath(fileType.folderPath);
    const parentFolderPath = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}/${path}`;
    await this.sp.ensureFolder(parentFolderPath, name);
    const uploadResult = await this.files.uploadFiles(files, false, uploadedBy);
    const addedFiles = uploadResult.filter((res) => res.success).map((res) => ({
      name: res.file ? res.file.name : res.fileName,
      serverRelativeUrl: res.file?.serverRelativeUrl ?? "",
      commentId: commentId ? commentId : void 0
    }));
    for (const file of addedFiles) {
      if (file.serverRelativeUrl) {
        await this.files.updateFileByUrl(file.serverRelativeUrl, { QAId: qnAItemId.toString().padStart(4, "0") });
      }
    }
    return addedFiles;
  }
  /**
   * Deletes a comment from a list item by its unique ID and comment ID.
   * @param listName The name of the list.
   * @param uniqueId The unique identifier of the list item.
   * @param commentId The ID of the comment to delete.
   * @returns Promise resolving to true if deleted successfully, otherwise false.
   */
  async deleteComment(listName, uniqueId, commentId) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${uniqueId}')/Comments('${commentId}')`;
    const response = await this.sp.spDelete(apiUrl);
    return !!response.ok;
  }
  /**
  * Retrieves Q&A items from the specified folder in the QA list.
  * @param folderPAth The relative path of the folder within the QA list.
  * @returns An array of Q&A items.
  */
  async getQnAItems(stage, phase) {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"QA" /* QA */}/${this.sp.getChannelName()}`;
    const expand = "Author";
    const select = "*,FileDirRef,UniqueId,Author/Title";
    const filter = `startswith(FileDirRef,'${fileDirRef}') and startswith(ContentTypeId,'0x0100') ${stage && phase ? `and Stage eq '${stage}' and Phase eq '${phase}'` : ``}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${"QA" /* QA */}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
    const response = await this.sp.spGet(apiUrl);
    return Array.isArray(response.data) ? this.mapQnAItems(response.data) : [];
  }
  /**
   * Retrieves the choice options for the 'Category' field from the QnA list.
   * 
   * @returns An array of category choice strings available in the QnA list.
   */
  async getQnACategories() {
    const columnsApi = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${"QA" /* QA */}')/fields/getbyinternalnameortitle('Category')`;
    const col = await this.sp.spGet(columnsApi);
    return col.data.Choices ?? [];
  }
  /**
  * Maps raw comment data to a structured array of comment objects.
  * @param data The raw data array of comments.
  * @returns An array of mapped comment objects.
  */
  mapItemComments(data) {
    data.sort((a, b) => a.createdDate.localeCompare(b.createdDate));
    return data.map((item) => {
      return {
        text: item.text,
        author: {
          name: item.author.name,
          email: item.author.email,
          isExternal: item.author.isExternal
        },
        createdDate: item.createdDate,
        id: Number(item.id)
      };
    });
  }
  /**
       * Maps raw Q&A data to a structured array of Q&A items, including comments for each item.
       * @param data The raw data array of Q&A items.
       * @returns A promise that resolves to an array of mapped Q&A item objects.
       */
  async mapQnAItems(data) {
    return await Promise.all(
      data.map(async (item) => await this.mapQnAItem(item))
    );
  }
  /**
   * Maps a single raw Q&A item to a structured IQnAItem object, including comments and metadata.
   * @param item The raw Q&A item data.
   * @returns Promise resolving to a mapped IQnAItem object.
   */
  async mapQnAItem(item) {
    const comments = await this.getItemComments(item.Id);
    const { name, date } = parseNameAndDate(item["Answered" /* Answered */]);
    const modified = getLatestDate([...comments.map((c) => c.createdDate), item.Modified]);
    return {
      question: item.Title,
      answer: item["answer" /* Answer */],
      comments,
      uniqueId: item.UniqueId,
      id: item.Id,
      status: item.Status,
      category: item.Category,
      phase: item.Phase,
      order: item["QAOrder" /* Order */],
      stage: item.Stage,
      createdText: `${item.Author.Title} ${new Date(item.Created).toLocaleDateString()}`,
      answeredText: item["Answered" /* Answered */],
      priority: item["Priority" /* Priority */],
      created: new Date(item.Created),
      createdUser: item.Author.Title,
      createdDate: new Date(item.Created).toLocaleString(void 0, {
        year: "numeric",
        month: "numeric",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit"
      }),
      answeredDate: date,
      answeredUser: name,
      filesData: parseQnAFilesData(item["FilesData" /* FilesData */]),
      subCategory: item["SubCategory" /* SubCategory */],
      modified: new Date(modified).toLocaleString(void 0, {
        year: "numeric",
        month: "numeric",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit"
      })
    };
  }
  /**
  * Retrieves comments for a specific item in the QA list based on its ID.
  * @param id The ID of the item to retrieve comments for.
  * @returns An array of comments associated with the item.
  */
  async getItemComments(id) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${"QA" /* QA */}')/items(${id})/comments`;
    const response = await this.sp.spGet(apiUrl);
    return Array.isArray(response.data) ? this.mapItemComments(response.data) : [];
  }
};
var ParsedFormDataSchema = z.record(z.unknown());
var normalizeRequestFormData = (value) => {
  if (!value || typeof value !== "string") {
    return "{}";
  }
  try {
    const parsed = JSON.parse(value);
    const parsedFormData = ParsedFormDataSchema.safeParse(parsed);
    return parsedFormData.success ? value : "{}";
  } catch {
    return "{}";
  }
};
var RequestSchema = z.object({
  Id: z.number(),
  Created: z.string(),
  ["RequestType" /* RequestType */]: z.object({
    Id: z.number(),
    Title: z.string(),
    ["Title_en" /* TitleEn */]: z.string(),
    ["Notification" /* Notification */]: z.string()
  }),
  ["RequestStatus" /* RequestStatus */]: z.nativeEnum(RequestStatus),
  ["FormData" /* FormData */]: z.string().nullable().optional().transform(normalizeRequestFormData),
  ["FileTemplate" /* FileTemplate */]: z.string().nullable().optional(),
  ["FileIds" /* FileIds */]: z.string().nullable().optional(),
  ["categoryId" /* CategoryId */]: z.string().nullable().optional()
});
z.array(RequestSchema);
var parseRequest = (value) => {
  if (!value) {
    return void 0;
  }
  const parsedRequest = RequestSchema.safeParse(value);
  return parsedRequest.success ? parsedRequest.data : void 0;
};
var parseRequests = (value) => {
  if (!value || !Array.isArray(value)) {
    return [];
  }
  const items = value.map((item) => parseRequest(item)).filter((item) => item !== void 0);
  return items;
};
var RequestApiClient = class {
  constructor(files, sp) {
    this.files = files;
    this.sp = sp;
  }
  files;
  sp;
  /**
       * create request in specific folder using /AddValidateUpdateItemUsingPath endpoint and attach files to it
       * @param data 
       * @param files 
       * @returns boolean
       */
  async createRequest(data, files, fileType, secondStage, uploadedBy) {
    const folderPath = secondStage && fileType.folderPathAfter ? fileType.folderPathAfter : fileType.folderPath;
    let fileTemplate = "";
    const fileIds = [];
    const failedUploads = [];
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const ext = getFileExtension(file.name);
      const today = /* @__PURE__ */ new Date();
      const day = String(today.getDate()).padStart(2, "0");
      const month = String(today.getMonth() + 1).padStart(2, "0");
      const year = String(today.getFullYear()).slice(-2);
      const dateSuffix = `${day}${month}${year}`;
      const fileName = data.templateRef && i === 0 ? `${data.templateName}_${dateSuffix}.${ext}` : file.name;
      const uploadResult = await this.files.uploadFile(folderPath, { ...file, name: fileName }, uploadedBy, void 0, fileType.id, fileType.uploadStatus);
      if (uploadResult.file) {
        if (data.templateRef && i === 0) {
          fileTemplate = uploadResult.file.serverRelativeUrl;
        } else {
          fileIds.push(uploadResult.file.uniqueId);
        }
      }
      if (!uploadResult.success) {
        failedUploads.push(file.name);
      }
    }
    const body = {
      "formValues": [
        {
          "FieldName": "Title",
          "FieldValue": data.templateName
        },
        {
          "FieldName": "FormData" /* FormData */,
          "FieldValue": JSON.stringify(data.formData)
        },
        {
          "FieldName": "RequestType" /* RequestType */,
          "FieldValue": data.typeId?.toString()
        },
        {
          "FieldName": "RequestStatus" /* RequestStatus */,
          "FieldValue": "new" /* New */
        },
        {
          "FieldName": "FileTemplate" /* FileTemplate */,
          "FieldValue": fileTemplate
        },
        {
          "FieldName": "FileIds" /* FileIds */,
          "FieldValue": fileIds.join(";")
        },
        {
          "FieldName": "ExportData" /* ExportData */,
          "FieldValue": data.exportData ? data.exportData : ""
        }
      ]
    };
    const result = await this.sp.createListItemUsingPath(body, "Zadosti" /* Requests */);
    return { result, failedUploads };
  }
  /**
   * loads request from SP list from folder with name of channel
   * @returns array of {@link IRequest}
   */
  async getRequests() {
    const folderPath = `${this.sp.getRelativeUrl()}/lists/${"Zadosti" /* Requests */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${folderPath}'`;
    const select = `*,${"RequestType" /* RequestType */}/Id,${"RequestType" /* RequestType */}/Title,${"RequestType" /* RequestType */}/${"Title_en" /* TitleEn */},${"RequestType" /* RequestType */}/${"Notification" /* Notification */}`;
    const expand = `${"RequestType" /* RequestType */}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"Zadosti" /* Requests */}')/items?&$select=${select}&$filter=${filter}&$expand=${expand}`;
    const response = await this.sp.spGet(apiUrl);
    const data = parseRequests(response.data);
    return mapRequests(data);
  }
  /**
   * Loads all request options from the SharePoint list specified by {@link ListName.RequestTypes}.
   * @returns A promise that resolves to an array of {@link IRequestOption} objects.
   */
  async getRequestsOptions() {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${"TypyZadosti" /* RequestTypes */}')/items`;
    const result = await this.sp.spGet(apiUrl);
    return Array.isArray(result.data) ? mapRequestOptions(result.data) : [];
  }
  /**
       * update request and attach new files to it
       * @param data 
       * @param itemId 
       * @param files 
       * @returns boolean
       */
  async updateRequest(body, itemId, uploadedBy, secondStage, files, category) {
    const uploadedFiles = [];
    if (files) {
      const result2 = await this.files.uploadFiles(files, !!secondStage, uploadedBy, category?.id);
      const fileIds = result2.map((f) => f.file?.uniqueId);
      for (const item of result2) {
        if (item.file) {
          uploadedFiles.push(item.file);
        }
      }
      const fileIdsString = fileIds.join(";");
      body["FileIds" /* FileIds */] = body["FileIds" /* FileIds */] ? body["FileIds" /* FileIds */] + ";" + fileIdsString : fileIdsString;
    }
    const result = await this.sp.updateListItem(itemId, body, "Zadosti" /* Requests */);
    return { files: uploadedFiles, result };
  }
};
var BaseApiClient = class {
  constructor(webPartContext, getUrl) {
    this.webPartContext = webPartContext;
    this.getUrl = getUrl;
  }
  webPartContext;
  getUrl;
  /**
   * Retrieves a fresh request digest value from SharePoint for authentication.
   * @returns Promise resolving to the FormDigestValue string.
   */
  async getFreshDigest() {
    const baseUrl = this.getUrl();
    const apiUrl = `${baseUrl}/_api/contextinfo`;
    const response = await this.spPost(apiUrl);
    return response.data && response.data.FormDigestValue ? response.data.FormDigestValue : "";
  }
  /**
   * Sends a POST request to a specified SharePoint API endpoint with the given request body and headers.
   * @param url The URL of the SharePoint API endpoint.
   * @param body The body of the request, which can be an object, string, or ArrayBuffer (default is an empty object).
   * @param headers Additional headers to include in the request (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
   */
  async spPost(url, body = {}, headers = {}) {
    try {
      const isBatch = headers["Content-Type"]?.startsWith("multipart/mixed");
      const hasCustomContentType = Object.keys(headers).some(
        (h) => h.toLowerCase() === "content-type"
      );
      const digest = isBatch ? await this.getFreshDigest() : void 0;
      const spOpts = {
        headers: {
          ...headers,
          ...!isBatch && !hasCustomContentType && {
            "Accept": "application/json",
            "Content-type": "application/json"
          },
          ...isBatch && {
            "X-RequestDigest": digest
          }
        },
        body: typeof body === "object" && !(body instanceof ArrayBuffer) ? JSON.stringify(body) : body
      };
      const response = await this.webPartContext.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts);
      const responseJson = response.status === 204 ? true : response.status === 200 && isBatch ? await response.text() : await response.json();
      const result = {
        ok: response.ok,
        code: response.status
      };
      if (response.ok) {
        result.data = responseJson;
      } else {
        console.error(responseJson);
        result.exception = responseJson;
      }
      return result;
    } catch (exception) {
      console.error("Error in POST request: ", url, body);
      console.error(exception.error, exception);
      return {
        ok: false,
        code: 500,
        error: exception.message,
        exception
      };
    }
  }
  /**
   * Sends a DELETE request to a specified SharePoint API endpoint with the given request body and headers.
   * @param url The URL of the SharePoint API endpoint.
   * @param body The body of the request, which can be an object (default is an empty object).
   * @param headers Additional headers to include in the request (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
   */
  // REFACTOR: `return await` is redundant in all three delegate methods below (spDelete, spPatch, spPut).
  // `return this.spPost(...)` is sufficient and avoids an extra microtask tick.
  async spDelete(url, body = {}, headers = {}) {
    return this.spPost(
      url,
      body,
      {
        ...headers,
        "X-Http-Method": "DELETE",
        "IF-MATCH": "*"
      }
    );
  }
  /**
   * Sends a PATCH request (MERGE method) to a specified SharePoint API endpoint with the given request body and headers.
   * @param url The URL of the SharePoint API endpoint.
   * @param body The body of the request, which can be an object (default is an empty object).
   * @param headers Additional headers to include in the request (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
   */
  async spPatch(url, body = {}, headers = {}) {
    return this.spPost(
      url,
      body,
      {
        ...headers,
        "IF-MATCH": "*",
        "X-HTTP-Method": "MERGE"
      }
    );
  }
  /**
   * Sends a PUT request to a specified SharePoint API endpoint with the given request body and headers.
   * @param url The URL of the SharePoint API endpoint.
   * @param body The body of the request, which can be an object or string (default is an empty object).
   * @param headers Additional headers to include in the request (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data or errors.
   */
  async spPut(url, body = {}, headers = {}) {
    return this.spPost(
      url,
      body,
      {
        ...headers,
        "IF-MATCH": "*",
        "X-HTTP-Method": "PUT"
      }
    );
  }
  /**
   * Sends a GET request to a specified SharePoint API endpoint with the given headers.
   * @param url The URL of the SharePoint API endpoint.
   * @param headers Additional headers to include in the request (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request, any response data, or errors.
   */
  async spGet(url, headers = {}, logError = true, responseType = "json") {
    try {
      const spOpts = {
        headers: {
          "Accept": "application/json",
          ...responseType === "json" && { "Content-type": "application/json" },
          ...headers
        }
      };
      let results = [];
      let nextUrl = url;
      let lastResponseJson = null;
      let isPaginated = false;
      while (nextUrl) {
        const response = await this.webPartContext.spHttpClient.get(nextUrl, SPHttpClient.configurations.v1, spOpts);
        let responseJson;
        if (response.ok && responseType === "text") {
          const text = await response.text();
          return {
            ok: true,
            code: 200,
            data: text
          };
        }
        try {
          responseJson = await response.clone().json();
        } catch (_) {
          if (!response.ok && logError) {
            const error = await response.text();
            console.error(error);
          }
        }
        if (response.ok) {
          lastResponseJson = responseJson;
          if (responseJson.value) {
            results = results.concat(responseJson.value);
            isPaginated = true;
          } else {
            results.push(responseJson);
          }
          nextUrl = responseJson["@odata.nextLink"];
        } else {
          if (logError) {
            console.error(responseJson || `${response.status} ${response.statusText}`.trim());
          }
          return {
            ok: false,
            code: response.status,
            exception: JSON.stringify(responseJson)
          };
        }
      }
      return {
        ok: true,
        code: 200,
        data: isPaginated ? results : lastResponseJson
        // Keep original format if no paging
      };
    } catch (exception) {
      if (logError) {
        console.error("Error in GET request: ", url);
        console.error(exception.error, exception);
      }
      return {
        ok: false,
        code: 500,
        error: exception.message,
        exception
      };
    }
  }
  /**
   * Sends a GET request to retrieve binary data (as an ArrayBuffer) from a specified SharePoint API URL.
   * @param url The API endpoint to fetch the binary content from.
   * @param headers (Optional) Additional headers to include in the request.
   * @returns A Promise resolving to an ArrayBuffer containing the binary data, or throws an error if the request fails.
   */
  // REFACTOR: `spGetBinary` and `spPostBinary` break the uniform `IApiResult` return contract
  // used by all other methods. Consider returning `IApiResult<ArrayBuffer>` (generic) so callers
  // can handle failures consistently without try/catch at every call site.
  async spGetBinary(url, headers = {}) {
    try {
      const spOpts = {
        headers: {
          "Accept": "application/octet-stream",
          ...headers
        }
      };
      const response = await this.webPartContext.spHttpClient.get(url, SPHttpClient.configurations.v1, spOpts);
      if (response.ok) {
        const arrayBuffer = await response.arrayBuffer();
        return {
          ok: true,
          data: arrayBuffer
        };
      } else {
        const error = await response.text();
        console.error("Error fetching binary data: ", error);
        return {
          ok: false,
          error
        };
      }
    } catch (exception) {
      console.error("Error in binary GET request: ", url);
      console.error(exception);
      return {
        ok: false,
        error: exception.message,
        exception
      };
    }
  }
  /**
   * Sends a POST request to upload binary data (as an ArrayBuffer) to a specified SharePoint API URL.
   * @param url The API endpoint to which the binary content should be uploaded.
   * @param fileContent The binary content to be uploaded, provided as an ArrayBuffer.
   * @returns A Promise that resolves when the upload is successful, or throws an error if the request fails.
   */
  async spPostBinary(url, fileContent) {
    try {
      const spOpts = {
        body: fileContent,
        // Attach the binary content to the body
        headers: {
          "Accept": "application/json",
          "Content-Type": "application/octet-stream"
          // Set the content type as binary
        }
      };
      const response = await this.webPartContext.spHttpClient.post(url, SPHttpClient.configurations.v1, spOpts);
      if (!response.ok) {
        const error = await response.text();
        console.error("Error uploading binary data: ", error);
        throw new Error(error);
      }
      return {
        ok: true
      };
    } catch (exception) {
      console.error("Error in binary POST request: ", url);
      console.error(exception);
      return {
        ok: false
      };
    }
  }
  async getGraphClient(url) {
    const GRAPH_API_VERSION = "3";
    const client = await this.webPartContext.msGraphClientFactory.getClient(GRAPH_API_VERSION);
    const graphClient = client.api(url);
    return graphClient;
  }
  /**
   * Sends a GET request to the Microsoft Graph API and retrieves data.
   * @param url The URL of the Microsoft Graph API endpoint.
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
   */
  // REFACTOR: `graphGet` and `graphPost` both repeat the same two-line client initialization.
  // Extract a private `getGraphClient(url: string)` helper to eliminate the duplication.
  // Also: the magic string "3" is the MSGraph API version — extract it to a named constant, e.g.:
  //   private static readonly GRAPH_API_VERSION = "3";
  async graphGet(url, attempt = 1) {
    try {
      const graphClient = await this.getGraphClient(url);
      const response = await graphClient.get();
      if (!response) {
        return {
          ok: false
        };
      }
      return {
        ok: true,
        data: response.value ?? response
      };
    } catch (exception) {
      if (exception?.statusCode === 429 && attempt < 3) {
        await delay(1e3);
        return this.graphGet(url, attempt + 1);
      }
      return {
        ok: false
      };
    }
  }
  /**
   * Sends a POST request to the Microsoft Graph API with a specified body.
   * @param url The URL of the Microsoft Graph API endpoint.
   * @param body The body of the POST request, which is an object (default is an empty object).
   * @returns A Promise that resolves to an IApiResult, containing the status of the request and any response data.
   */
  async graphPost(url, body = {}, attempt = 1) {
    try {
      const graphClient = await this.getGraphClient(url);
      const response = await graphClient.post(body);
      if (!response) {
        return {
          ok: false
        };
      }
      return {
        ok: true,
        data: response.value ?? response
      };
    } catch (exception) {
      if (exception?.statusCode === 429 && attempt < 3) {
        await delay(1e3);
        return this.graphPost(url, body, attempt + 1);
      }
      return {
        ok: false
      };
    }
  }
  async graphDelete(url, attempt = 1) {
    try {
      const graphClient = await this.getGraphClient(url);
      const response = await graphClient.delete();
      if (!response) {
        return {
          ok: false
        };
      }
      return {
        ok: true,
        data: response.value ?? response
      };
    } catch (exception) {
      if (exception?.statusCode === 429 && attempt < 3) {
        await delay(1e3);
        return this.graphDelete(url, attempt + 1);
      }
      return {
        ok: false
      };
    }
  }
};

// src/api/SPApiClient.ts
var ContextSchema = z.object({
  c: z.string(),
  t: z.string(),
  n: z.string(),
  s: z.string().url()
});
var parseContext = (value) => {
  if (!value) {
    return void 0;
  }
  try {
    if (typeof value !== "string") {
      return void 0;
    }
    const parsedJson = JSON.parse(LZString.decompressFromEncodedURIComponent(value));
    const parsedContactData = ContextSchema.safeParse(parsedJson);
    if (!parsedContactData.success) {
      return void 0;
    }
    return {
      teamId: parsedContactData.data.t,
      channelId: parsedContactData.data.c,
      channelName: parsedContactData.data.n,
      siteUrl: parsedContactData.data.s
    };
  } catch {
    return void 0;
  }
};
var SPApiClient = class extends BaseApiClient {
  constructor(webPartContext, internalSiteName, templateSiteName) {
    super(webPartContext, () => this.getAbsoluteUrl());
    this.internalSiteName = internalSiteName;
    this.templateSiteName = templateSiteName;
  }
  internalSiteName;
  templateSiteName;
  channelId;
  teamId;
  absoluteUrl;
  relativeUrl;
  userRole;
  channelName;
  teamName;
  channelType;
  // REFACTOR: Hardcoded Czech string — move to a named constant or config value
  // so it can be changed without hunting through implementation files.
  documentFolder = "Sdilene dokumenty";
  listMappings = {
    ["TypyZadosti" /* RequestTypes */]: "TypyZadosti" /* RequestTypes */,
    ["Kategorie" /* Categories */]: "Kategorie" /* DocumentStructure */
  };
  libraryName = "Dokumenty" /* Documents */;
  sp;
  /**
   * Checks existence of all required SharePoint lists and libraries.
   * @returns Object mapping each list/library name to a boolean indicating existence.
   */
  async checkAllListExistence() {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists?$select=Title`;
    const response = await this.spGet(apiUrl);
    const listItems = Array.isArray(response.data) ? response.data : [];
    const titles = listItems.map((l) => l.Title);
    const allNames = [
      "Detaily" /* Details */,
      "Zadosti" /* Requests */,
      "QA" /* QA */,
      "Kategorie" /* Categories */,
      "Povinnosti" /* Duties */,
      "Formulare" /* Forms */,
      "Requests" /* Requests */,
      "PozadovaneDokumenty" /* RequiredDocuments */,
      "TypyDokumentu" /* DocumentTypes */,
      "TypyZadosti" /* RequestTypes */,
      "Notification" /* Notifications */,
      "WorkflowActions" /* WorkflowActions */,
      "ActionToTrigger" /* ActionToTrigger */,
      "Logs" /* Logs */
    ];
    const result = {};
    allNames.forEach((name) => {
      result[name] = titles.includes(name);
    });
    return result;
  }
  /**
   * Ensures a SharePoint list exists, creates it if missing, and ensures required columns and views.
   * @param listName The name of the list or library.
   */
  async ensureList(listName) {
    const listExists = await this.checkListExistence(listName);
    if (!listExists) {
      const result = await this.createList(listName);
      if (!result) {
        return;
      }
    }
    const missingColumns = await this.ensureColumns(listName);
    if (listName !== "Requests" /* Requests */ && !listExists) {
      await this.ensureListView(listName, missingColumns);
    }
  }
  /**
   * Deletes a SharePoint list and all its items and folders.
   * @param listName The name of the list to delete.
   * @returns Promise resolving to true if deleted successfully.
   */
  async deleteList(listName) {
    const itemsFilter = `startswith(ContentTypeId,'0x0100')`;
    const itemsUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    const itemsResponse = await this.spGet(`${itemsUrl}?$filter=${itemsFilter}`);
    const items = Array.isArray(itemsResponse.data) ? itemsResponse.data : [];
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      await this.spDelete(`${itemsUrl}(${item.Id})`);
    }
    const foldersResponse = await this.spGet(itemsUrl);
    const folders = Array.isArray(foldersResponse.data) ? foldersResponse.data : [];
    for (let i = 0; i < folders.length; i++) {
      const item = folders[i];
      await this.spDelete(`${itemsUrl}(${item.Id})`);
    }
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
    const result = await this.spDelete(apiUrl);
    return !!result.ok;
  }
  /**
   * Updates a SharePoint list with the provided body.
   * @param body Update payload.
   * @param listName The name of the list (defaults to libraryName).
   */
  async updateList(body, listName = this.libraryName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
    await this.spPatch(apiUrl, body);
  }
  /**
   * Updates a column in a SharePoint list.
   * @param body Update payload.
   * @param listName The name of the list.
   * @param columnName The column to update.
   */
  async updateColumn(body, listName, columnName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getByTitle('${listName}')/fields/getByInternalNameOrTitle('${columnName}')`;
    const headers = {
      "Accept": "application/json;odata=nometadata",
      "Content-type": "application/json;odata=nometadata",
      "OData-Version": ""
    };
    await this.spPatch(apiUrl, body, headers);
  }
  /**
   * Updates an item in a SharePoint list by unique ID.
   * @param uniqueId The unique identifier of the item.
   * @param body Update payload.
   * @param listName The name of the list (defaults to libraryName).
   * @returns Promise resolving to true if updated successfully.
   */
  async updateItemInList(uniqueId, body = {}, listName = this.libraryName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${uniqueId}')`;
    const result = await this.spPatch(apiUrl, body);
    return result.ok ?? false;
  }
  /**
   * Deletes an item from the specified SharePoint list by its ID.
   * 
   * @param id - The ID of the item to delete.
   * @param listName - The name of the SharePoint list.
   * @returns A promise resolving to true if the item was successfully deleted, otherwise false.
   */
  async deleteListItem(id, listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})`;
    const result = await this.spDelete(apiUrl);
    return result.ok === true;
  }
  /**
   * Updates a list item by item ID.
   * @param itemId The item ID.
   * @param body Update payload.
   * @param listName The name of the list or library.
   * @returns Promise resolving to true if updated successfully.
   */
  async updateListItem(itemId, body = {}, listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})`;
    const result = await this.spPatch(apiUrl, body);
    return result.ok ?? false;
  }
  /**
   * Retrieves choice options for a field in a SharePoint list.
   * @param listName The name of the list.
   * @param columnTitle The column title.
   * @returns Array of choice strings.
   */
  async getFieldChoices(listName, columnTitle) {
    const apiUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/fields/getbyinternalnameortitle('${columnTitle}')`;
    const result = await this.spGet(apiUrl);
    return result.data && result.data.Choices ? result.data.Choices : [];
  }
  /**
   * Creates a list item using a specified path in SharePoint.
   * @param body Item creation payload.
   * @param listName The name of the list.
   * @param path Optional path for item placement.
   * @param channelSpecific Whether to use channel-specific path.
   * @returns API result of the operation.
   */
  async createListItemUsingPath(body = {}, listName, path, channelSpecific = true) {
    let folderPath = `${this.absoluteUrl}/lists/${listName}`;
    if (channelSpecific) {
      folderPath += `/${this.channelName}`;
    }
    if (path) {
      folderPath += `${path}`;
    }
    const bodyWithPath = {
      "listItemCreateInfo": {
        "FolderPath": {
          "DecodedUrl": folderPath
        },
        "UnderlyingObjectType": 0
      },
      "bNewDocumentUpdate": false,
      ...body
    };
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/AddValidateUpdateItemUsingPath`;
    const result = await this.spPost(apiUrl, bodyWithPath);
    return result;
  }
  /**
   * Updates multiple list items in batch.
   * @param updates Array of update objects with uniqueId and body.
   * @param listName The name of the list.
   */
  async updateListItemsBatch(updates, listName) {
    const chunkSize = 50;
    const chunks = this.chunkArray(updates, chunkSize);
    for (const chunk of chunks) {
      await this.sendItemsUpdateBatchChunk(listName, chunk);
    }
  }
  /**
   * Splits an array into chunks of specified size.
   * @param arr The array to chunk.
   * @param size The chunk size.
   * @returns Array of chunked arrays.
   */
  chunkArray(arr, size) {
    const chunks = [];
    for (let i = 0; i < arr.length; i += size) {
      chunks.push(arr.slice(i, i + size));
    }
    return chunks;
  }
  /**
   * Sends a batch update chunk for list items.
   * @param listName The name of the list.
   * @param updates Array of update objects.
   * @returns API result of the batch operation.
   */
  async sendItemsUpdateBatchChunk(listName, updates) {
    const boundary = "batch_" + Date.now();
    const changeBoundary = "changeset_" + Date.now();
    const batchParts = updates.map((u, index) => {
      const itemUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${u.uniqueId}')`;
      return [
        `--${changeBoundary}`,
        `Content-Type: application/http`,
        `Content-Transfer-Encoding: binary`,
        `Content-ID: ${index}`,
        // <-- first empty line
        "",
        `PATCH ${itemUrl} HTTP/1.1`,
        `Content-Type: application/json;odata=nometadata`,
        `IF-MATCH: *`,
        //`X-HTTP-Method: MERGE`,             // <-- second empty line
        "",
        JSON.stringify(u.body),
        "",
        ""
      ].join("\r\n");
    });
    const batchBody = `--${boundary}\r
Content-Type: multipart/mixed; boundary=${changeBoundary}\r
\r
` + batchParts.join("\r\n") + `--${changeBoundary}--\r
--${boundary}--\r
`;
    const response = await this.spPost(
      `${this.absoluteUrl}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`,
        "Accept": "application/json;odata=nometadata",
        "OData-Version": "4.0",
        "OData-MaxVersion": "4.0"
      }
    );
    return response;
  }
  /**
   * Checks if a SharePoint list exists.
   * @param listName The name of the list or library.
   * @returns Promise resolving to true if exists.
   */
  async checkListExistence(listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')`;
    const response = await this.spGet(apiUrl, {}, false);
    return response.code !== 404;
  }
  /**
   * Creates a SharePoint list or library.
   * @param listName The name of the list or library.
   * @returns Promise resolving to true if created successfully.
   */
  async createList(listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists`;
    const body = {
      Title: listName,
      BaseTemplate: listName === "Requests" /* Requests */ ? 101 : 100,
      EnableFolderCreation: true
    };
    const response = await this.spPost(apiUrl, body);
    return response.ok === true;
  }
  /**
   * Ensures required columns exist in a SharePoint list.
   * @param listName The name of the list or library.
   * @returns Array of missing columns.
   */
  async ensureColumns(listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields`;
    const result = await this.spGet(apiUrl);
    const fields = result.data;
    const columns = listColumns[listName];
    const missingColumns = [];
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i];
      const field = fields.find((f) => f.InternalName === column.Title);
      if (!field) {
        missingColumns.push(column);
      }
      if (!field) {
        await this.createColumn(listName, column);
      } else if (column.FieldTypeKind === "7") {
        const lookupListUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${column.lookUpListTitle}')`;
        const lookupListResult = await this.spGet(lookupListUrl);
        const lookupListId = lookupListResult.data.Id;
        const fieldType = column.allowMultiple ? "LookupMulti" : "Lookup";
        const multAttr = column.allowMultiple ? "Mult='TRUE'" : "";
        const lookupFieldXml = `<Field Type='${fieldType}' DisplayName='${column.Title}' ${multAttr} Name='${column.Title}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
        const existingLookupListId = field.LookupList;
        if (existingLookupListId !== lookupListId) {
          const body = { "SchemaXml": lookupFieldXml };
          const updateApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
          await this.spPatch(updateApiUrl, body);
        }
      }
    }
    return missingColumns;
  }
  /**
   * Ensures the default list view contains required columns.
   * @param listName The name of the list.
   * @param columns Array of columns to ensure in the view.
   */
  async ensureListView(listName, columns) {
    const viewId = await this.getDefaultListViewId(listName);
    for (let i = 0; i < columns.length; i++) {
      const column = columns[i];
      const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/views('${viewId}')/ViewFields/AddViewField('${column.Title}')`;
      await this.spPost(apiUrl);
    }
  }
  /**
   * Retrieves the default view ID for a SharePoint list.
   * @param listName The name of the list.
   * @returns The default view ID as a string.
   */
  async getDefaultListViewId(listName) {
    const apiUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/DefaultView`;
    const result = await this.spGet(apiUrl);
    return result.data ? result.data.Id : "";
  }
  /**
   * Creates a column in a SharePoint list.
   * @param listName The name of the list or library.
   * @param column Column definition object.
   * @returns Promise resolving to true if created successfully.
   */
  async createColumn(listName, column) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields`;
    if (column.FieldTypeKind !== "7") {
      const { displayName, allowMultiple, ...baseColumnProps } = column;
      const body = { ...baseColumnProps };
      if (column.FieldTypeKind === "6" || column.FieldTypeKind === "15") {
        body["@odata.type"] = column.FieldTypeKind === "15" ? "SP.FieldMultiChoice" : "SP.FieldChoice";
        const templatesListName = listName !== "Requests" /* Requests */ ? this.listMappings[listName] : void 0;
        if (!column.Choices && templatesListName) {
          const webUrl = `${this.absoluteUrl.replace(this.relativeUrl, "")}/sites/${this.templateSiteName}`;
          const columnsApi = `${webUrl}/_api/lists/getbytitle('${templatesListName}')/fields/getbyinternalnameortitle('${column.Title}')`;
          const col = await this.spGet(columnsApi);
          body.Choices = col.data.Choices ?? [];
        }
      }
      if (column.FieldTypeKind === "20") {
        body["@odata.type"] = "SP.FieldUser";
        body.AllowMultipleValues = !!allowMultiple;
      }
      const result = await this.spPost(apiUrl, body);
      if (result.ok) {
        const fieldApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
        const body2 = { Title: displayName };
        await this.spPatch(fieldApiUrl, body2);
      }
      return !!result.ok;
    } else {
      const lookupListUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${column.lookUpListTitle}')`;
      const lookupListResult = await this.spGet(lookupListUrl);
      const lookupListId = lookupListResult.data.Id;
      const fieldType = column.allowMultiple ? "LookupMulti" : "Lookup";
      const multAttr = column.allowMultiple ? "Mult='TRUE'" : "";
      const lookupFieldXml = `<Field Type='${fieldType}' DisplayName='${column.Title}' ${multAttr} Name='${column.Title}' Required='FALSE' List='${lookupListId}' ShowField='Title' />`;
      const xmlApiUrl = `${apiUrl}/CreateFieldAsXml`;
      if (lookupListId) {
        const body = { "parameters": { "SchemaXml": lookupFieldXml } };
        const result = await this.spPost(xmlApiUrl, body);
        if (result.ok) {
          const fieldApiUrl = `${apiUrl}/getbytitle('${column.Title}')`;
          const body2 = { Title: column.displayName };
          await this.spPatch(fieldApiUrl, body2);
        }
        return !!result.ok;
      }
    }
    return false;
  }
  /**
   * Creates an item in a SharePoint list.
   * @param body Item creation payload.
   * @param listName The name of the list or library.
   * @returns API result of the operation.
   */
  async createListItem(body = {}, listName) {
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    const result = await this.spPost(apiUrl, body);
    return result;
  }
  /**
   * Ensures a user exists in SharePoint and returns their ID.
   * @param userEmail The user's email address.
   * @param webUrl Optional web URL (defaults to absoluteUrl).
   * @returns The SharePoint user ID.
   */
  async getUserSPId(userEmail, webUrl = this.absoluteUrl) {
    const result = await this.spPost(`${webUrl}/_api/web/ensureUser`, {
      logonName: userEmail
    });
    return result.data && result.data.Id ? result.data.Id : 0;
  }
  async init(ctx, teamId, channelId) {
    if (ctx) {
      await this.initWithCtx(ctx);
    } else if (teamId && channelId) {
      await this.initWithIds(teamId, channelId);
    }
  }
  async initWithCtx(ctx) {
    const parsed = parseContext(ctx);
    if (!parsed) {
      throw new Error("Invalid ctx format");
    }
    const { teamId, channelId, channelName, siteUrl } = parsed;
    const teamSiteUrl = siteUrl;
    this.absoluteUrl = teamSiteUrl;
    this.relativeUrl = new URL(teamSiteUrl).pathname;
    this.teamId = teamId;
    this.channelId = channelId;
    this.userRole = await this.getUserChannelRole(teamId, channelId);
    this.channelName = channelName;
    this.sp = spfi(teamSiteUrl).using(SPFx(this.webPartContext));
  }
  async initWithIds(teamId, channelId) {
    const channelDetail = await this.getChannelDetail(teamId, channelId);
    const teamSiteUrl = channelDetail.absoluteUrl ? channelDetail.absoluteUrl : await this.getTeamSiteUrl(teamId);
    this.absoluteUrl = teamSiteUrl;
    this.relativeUrl = new URL(teamSiteUrl).pathname;
    this.teamId = teamId;
    this.channelId = channelId;
    this.userRole = channelDetail.userRole;
    this.channelName = channelDetail.displayName;
    this.sp = spfi(teamSiteUrl).using(SPFx(this.webPartContext));
    this.upgradeLegacyUrl();
  }
  upgradeLegacyUrl() {
    const newCtx = LZString.compressToEncodedURIComponent(JSON.stringify({
      t: this.teamId,
      c: this.channelId,
      n: this.channelName,
      s: this.absoluteUrl
    }));
    const newUrl = `${window.location.pathname}?ctx=${newCtx}`;
    window.history.replaceState(null, "", newUrl);
  }
  async getChannelDetail(teamId, channelId) {
    const apiUrl = `teams/${teamId}/channels/${channelId}/filesFolder`;
    const result = await this.graphGet(apiUrl);
    const displayName = await this.getChannelDisplayName(teamId, channelId);
    const relativeUrl = result.data ? new URL(result.data.webUrl).pathname : "";
    const origin = result.data ? new URL(result.data.webUrl).origin : "";
    const parts = relativeUrl ? relativeUrl.replace("/sites/", "").split("/") : "";
    const absoluteUrl = parts && parts[0] && origin ? `${origin}/sites/${parts[0]}` : "";
    const userRole = await this.getUserChannelRole(teamId, channelId);
    return { displayName, absoluteUrl, userRole };
  }
  async getChannelDisplayName(teamId, channelId) {
    const apiUrl = `teams/${teamId}/channels/${channelId}`;
    const result = await this.graphGet(apiUrl);
    return result.data ? result.data.displayName : "";
  }
  async getTeamSiteUrl(id) {
    const teamSite = await this.graphGet(`groups/${id}/sites/root`);
    return teamSite.data ? teamSite.data.webUrl : "";
  }
  /**
   * Gets the current user's role in the channel.
   * @returns The user role as a number.
   */
  getUserRole() {
    return this.userRole;
  }
  /**
   * Gets the current team ID.
   * @returns The team ID as a string.
   */
  getTeamId() {
    return this.teamId;
  }
  /**
   * Gets the current channel ID.
   * @returns The channel ID as a string.
   */
  getChannelId() {
    return this.channelId;
  }
  /**
   * Gets the absolute URL for the current site.
   * @returns The absolute URL as a string.
   */
  getAbsoluteUrl() {
    return this.absoluteUrl;
  }
  /**
   * Gets the current channel name.
   * @returns The channel name as a string.
   */
  getChannelName() {
    return this.channelName;
  }
  /**
   * Gets the relative URL for the current site.
   * @returns The relative URL as a string.
   */
  getRelativeUrl() {
    return this.relativeUrl;
  }
  async getChannelType() {
    if (this.channelType) {
      return this.channelType;
    }
    const channelDetails = await this.graphGet(`/teams/${this.getTeamId()}/channels/${this.getChannelId()}`);
    const channelType = channelDetails.data?.membershipType === "private" ? "Private" /* Private */ : "Standard" /* Standard */;
    this.channelType = channelType;
    return channelType;
  }
  async isPrivateChannel() {
    const channelType = await this.getChannelType();
    return channelType === "Private" /* Private */;
  }
  /**
   * Gets the current team name, loading it if not already set.
   * @returns Promise resolving to the team name.
   */
  async getTeamName() {
    if (this.teamName) {
      return this.teamName;
    }
    const teamName = await this.loadTeamName();
    return teamName;
  }
  /**
   * Gets the user's role in a Teams channel.
   * @param teamId The team ID.
   * @param channelId The channel ID.
   * @returns Role as a number.
   */
  async getUserChannelRole(teamId, channelId) {
    const memberApiUrl = `/teams/${teamId}/channels/${channelId}/members`;
    const members = await this.graphGet(memberApiUrl);
    const userRole = Array.isArray(members.data) && members.data.some((item) => item.email === this.webPartContext.pageContext.user.email && item.roles.some((role) => role.toLowerCase() === "owner")) ? 0 : 1;
    return userRole;
  }
  /**
   * Loads the team name from Microsoft Graph if not already set.
   * @returns Promise resolving to the team name.
   */
  async loadTeamName() {
    if (this.teamName) {
      return this.teamName;
    }
    const teamSite = await this.graphGet(`groups/${this.teamId}/sites/root`);
    const teamName = teamSite.data ? teamSite.data.displayName : "";
    this.teamName = teamName;
    return teamName;
  }
  /**
   * Ensures a folder exists at the specified path, creating it if necessary.
   * @param path The server-relative path.
   * @param name The folder name.
   * @returns UniqueId of the folder or undefined.
   */
  async ensureFolder(path, name) {
    const checkApiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}/${name}')`;
    const check = await this.spGet(checkApiUrl, {}, false);
    if (check.ok && check.data) {
      return check.data.UniqueId;
    }
    const apiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')/folders/add('${name}')`;
    const result = await this.spPost(apiUrl);
    return result.data.UniqueId;
  }
  /**
   * Creates multiple list items in batch.
   * @param itemBodies Array of item creation payloads.
   * @param listName The name of the list or library.
   */
  async createListItemsBatch(itemBodies, listName) {
    const chunkSize = 50;
    const chunks = this.chunkArray(itemBodies, chunkSize);
    for (const chunk of chunks) {
      await this.sendItemsPostBatchChunk(listName, chunk);
    }
  }
  /**
   * Sends a batch chunk for creating list items.
   * @param listName The name of the list or library.
   * @param itemBodies Array of item creation payloads.
   * @returns API result of the batch operation.
   */
  async sendItemsPostBatchChunk(listName, itemBodies) {
    const boundary = "batch_" + Date.now();
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    const batchParts = itemBodies.map((body) => {
      return [
        `--${boundary}`,
        "Content-Type: application/http",
        "Content-Transfer-Encoding: binary",
        "",
        `POST ${apiUrl} HTTP/1.1`,
        "Content-Type: application/json",
        "Accept: application/json",
        "",
        JSON.stringify(body),
        ""
      ].join("\r\n");
    });
    const batchBody = [...batchParts, `--${boundary}--\r
`].join("\r\n");
    const response = await this.spPost(
      `${this.absoluteUrl}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`
      }
    );
    return response;
  }
  /**
   * Moves a list item to a new folder within the document structure.
   * @param id The item ID.
   * @param fileDirRef The target folder's file directory reference.
   * @param listName The name of the list.
   */
  async moveListItem(id, fileDirRef, listName) {
    const getUrl = `${this.absoluteUrl}/_api/web/GetListByTitle('${listName}')/items(${id})?$select=FileRef,FileDirRef,FileLeafRef`;
    const item = await this.spGet(getUrl);
    if (item.data) {
      const newFileDirRef = encodeURIComponent(item.data.FileRef.replace(item.data.FileDirRef, fileDirRef));
      const moveUrl = `${this.absoluteUrl}/_api/web/getfilebyserverrelativeurl('${item.data.FileRef}')/moveto(newurl='${newFileDirRef}',flags=1)`;
      await this.spPost(moveUrl);
    }
  }
  /**
   * Generates a batch string for setting permissions.
   * @param boundary Batch boundary string.
   * @param groupId Group ID.
   * @param roledefid Role definition ID.
   * @param apiUrlBase Base API URL.
   * @returns Batch string for permissions.
   */
  getPermissionBatchString(boundary, groupId, roledefid, apiUrlBase) {
    return [
      `--${boundary}`,
      "Content-Type: application/http",
      "Content-Transfer-Encoding: binary",
      "",
      `POST ${apiUrlBase}(principalid=${groupId},roledefid=${roledefid}) HTTP/1.1`,
      "Accept: application/json",
      "",
      ""
    ].join("\r\n");
  }
  /**
   * Retrieves site groups from the current SharePoint site.
   * @param webUrl Optional web URL.
   * @returns Array of site group objects.
   */
  async getSiteGroups(webUrl) {
    const apiUrl = `${webUrl ? webUrl : this.absoluteUrl}/_api/web/sitegroups`;
    const result = await this.spGet(apiUrl);
    const sitegroups = [];
    const groups = Array.isArray(result.data) ? result.data : [];
    for (let i = 0; i < groups.length; i++) {
      const group = groups[i];
      if (Object.values(SPGroupInternal).includes(group.Title) || Object.values(SPGroupExternal).includes(group.Title) || Object.values(SPGroupRO).includes(group.Title)) {
        sitegroups.push({
          id: group.Id,
          title: group.Title,
          users: []
        });
      }
    }
    return sitegroups;
  }
  async getSiteGroupByName(groupName, webUrl) {
    const apiUrl = `${webUrl ? webUrl : this.absoluteUrl}/_api/web/sitegroups/getByName('${encodeURIComponent(groupName)}')`;
    const result = await this.spGet(apiUrl);
    const group = result.data;
    if (!group) {
      return void 0;
    }
    return {
      id: group.Id,
      title: group.Title,
      users: []
    };
  }
  /**
   * Attaches a file to a list item.
   * @param id The item ID.
   * @param file The file to attach.
   * @param listName The name of the list.
   * @returns API result of the operation.
   */
  async addAttachmentFile(id, file, listName) {
    const body = await file.rawFile.arrayBuffer();
    return await this.addAttachmentArrayBuffer(id, file.name, body, listName);
  }
  /**
   * Attaches an array buffer as a file to a list item, optionally skipping if file exists.
   * @param id The item ID.
   * @param fileName The file name.
   * @param buffer The file content as ArrayBuffer.
   * @param listName The name of the list.
   * @param skipIfExists Whether to skip if file exists.
   * @returns API result of the operation.
   */
  async addAttachmentArrayBuffer(id, fileName, buffer, listName, skipIfExists = false) {
    const ext = getFileExtension(fileName);
    const name = getFileNameWOExt(fileName);
    let finalName = fileName;
    let count = 1;
    while (await this.checkAttachmentExists(id, finalName, listName)) {
      if (skipIfExists) {
        return { ok: true };
      }
      finalName = `${name} (${count}).${ext}`;
      count++;
    }
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${id})/AttachmentFiles/add(FileName='${finalName}')`;
    const response = await this.spPost(apiUrl, buffer);
    return response;
  }
  /**
   * Checks if an attachment exists for a list item.
   * @param itemId The item ID.
   * @param fileName The file name.
   * @param listName The name of the list.
   * @returns Promise resolving to true if exists.
   */
  async checkAttachmentExists(itemId, fileName, listName) {
    const apiUrl = `${this.relativeUrl}/_api/web/lists/getByTitle('${listName}')/items(${itemId})/AttachmentFiles?$filter=FileName eq '${fileName}'`;
    const result = await this.spGet(apiUrl);
    return !!result.ok && Array.isArray(result.data) && result.data.length > 0;
  }
  /**
   * Retrieves template attachments from a SharePoint list.
   * @param listName The name of the list.
   * @returns Array of attachment objects.
   */
  async getTemplatesAttachments(listName) {
    const folderPath = `${this.relativeUrl}/lists/${listName}`;
    const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100') and Title eq '\u0161ablony'`;
    const apiUrl = `${this.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=${filter}`;
    const result = await this.spGet(apiUrl);
    const id = Array.isArray(result.data) ? result.data[0]?.Id : void 0;
    if (!id) {
      return [];
    }
    const attachUrl = `${this.absoluteUrl}/_api/lists/getbytitle('${listName}')/items(${id})/AttachmentFiles`;
    const attachResult = await this.spGet(attachUrl);
    const attachments = (Array.isArray(attachResult.data) ? attachResult.data : []).map((att) => ({
      name: att.FileName,
      serverRelativeUrl: att.ServerRelativeUrl
    }));
    return attachments;
  }
  /**
   * Ensures a folder exists in a SharePoint list, creating and moving it if necessary.
   * @param name The folder name.
   * @param listName The name of the list or library.
   * @param path Optional path for folder placement.
   */
  async ensureListFolder(name, listName, path) {
    const fileDirRef = `${this.relativeUrl}/Lists/${listName}${path ? `/${path}` : ""}`;
    const fileRef = `${fileDirRef}/${name}`;
    const checkApiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${fileRef}')`;
    const folderCheck = await this.spGet(checkApiUrl, {}, false);
    if (!folderCheck.ok) {
      const createBody = {
        "ContentTypeId": "0x0120",
        "Title": name
      };
      const result = await this.createListItem(createBody, listName);
      const renameBody = {
        "Title": name,
        "FileLeafRef": name
      };
      await this.updateListItem(result.data.Id, renameBody, listName);
      if (path) {
        await this.moveFolder(`${this.relativeUrl}/Lists/${listName}/${name}`, fileRef);
      }
    }
  }
  /**
   * Moves a folder to a new path in SharePoint.
   * @param path The current folder path.
   * @param newPath The new folder path.
   * @returns Promise resolving to true if moved successfully.
   */
  async moveFolder(path, newPath) {
    const apiUrl = `${this.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${path}')/moveTo(newurl='${newPath}')`;
    const result = await this.spPost(apiUrl);
    return result.ok ?? false;
  }
  /**
   * Logs a message and data to the SharePoint Logs list.
   * @param message The log message.
   * @param data Optional log data.
   */
  async log(message, data) {
    const body = {
      formValues: [
        { FieldName: "Title", FieldValue: message },
        { FieldName: "Data", FieldValue: JSON.stringify(data) }
      ]
    };
    await this.ensureListFolder(this.channelName, "Logs" /* Logs */);
    await this.createListItemUsingPath(body, "Logs" /* Logs */, "", true);
  }
  /**
   * Hides specified SharePoint lists.
   * @param listNames Array of list names to hide.
   */
  async hideLists(listNames) {
    for (const listName of listNames) {
      await this.sp.web.lists.getByTitle(listName).update({ Hidden: true });
    }
  }
  // public async test(): Promise<void> {
  //     await this.graphGet(`groups/${SystemAccount.EUPAdminsId}/members`);
  //     await this.graphGet(`groups/${SystemAccount.EUPMembersId}/members`);
  // }
};
var StructureCategoryItemSchema = z.object({
  Id: z.number(),
  Title: z.string(),
  ["Stage" /* Stage */]: z.string(),
  ["Phase" /* Phase */]: z.string(),
  ["Stage_en" /* StageEn */]: z.string().nullable().optional(),
  ["Phase_en" /* PhaseEn */]: z.string().nullable().optional(),
  ["Category_en" /* CategoryEn */]: z.string().nullable().optional(),
  ["FolderDescription" /* Description */]: z.string().nullable().optional(),
  ["FolderDescription_en" /* DescriptionEn */]: z.string().nullable().optional(),
  ["CategoryStatus" /* CategoryStatus */]: z.nativeEnum(FolderStatus).nullable().optional(),
  ["CategoryOrder" /* Order */]: z.number().nullable().optional()
});
var StructureQnAPhaseItemSchema = z.object({
  Stage: z.string(),
  Phase: z.string(),
  Stage_en: z.string().optional(),
  Phase_en: z.string().optional()
});
var parseStructureCategoryItems = (value) => {
  if (!value || !Array.isArray(value)) {
    return [];
  }
  return value.map((item) => StructureCategoryItemSchema.safeParse(item)).filter((item) => item.success).map((item) => item.data);
};
var parseStructureQnAPhaseItems = (value) => {
  if (!value || !Array.isArray(value)) {
    return [];
  }
  return value.map((item) => StructureQnAPhaseItemSchema.safeParse(item)).filter((item) => item.success).map((item) => item.data);
};
var StructureApiClient = class {
  constructor(sp) {
    this.sp = sp;
  }
  sp;
  /**
   * Retrieves all folders within the document structure for the current channel.
   * @returns An array of folders.
   */
  async getStructure() {
    const categories = await this.getCategories();
    const phases = await this.getPhasesFromQnAList();
    return mapStructure(categories, phases);
  }
  /**
   * Retrieves categories from the SharePoint list filtered by the current channel's folder path.
   *
   * @returns A promise that resolves to an array of category items, or an empty array if none found.
   */
  async getCategories() {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"Kategorie" /* Categories */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${fileDirRef}'`;
    const select = "*,FileDirRef,UniqueId";
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${"Kategorie" /* Categories */}')/items?$select=${select}&$filter=${filter}`;
    const response = await this.sp.spGet(apiUrl);
    return parseStructureCategoryItems(response.data);
  }
  /**
  * Retrieves phase and stage information from the QA SharePoint list filtered by the current channel's folder path.
  *
  * @returns A promise that resolves to an array of phase and stage items, or an empty array if none found.
  */
  async getPhasesFromQnAList() {
    const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${"QA" /* QA */}/${this.sp.getChannelName()}`;
    const filter = `FileDirRef eq '${fileDirRef}'`;
    const select = "Stage,Phase,Stage_en,Phase_en";
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${"QA" /* QA */}')/items?$select=${select}&$filter=${filter}`;
    const response = await this.sp.spGet(apiUrl);
    return parseStructureQnAPhaseItems(response.data);
  }
};
var RoleAssignmentSchema = z.array(z.object({
  Member: z.object({
    LoginName: z.string(),
    Id: z.number()
  }),
  RoleDefinitionBindings: z.array(z.object({
    Id: z.number()
  }))
}));
var parseRoleAssignments = (value) => {
  if (!value) {
    return [];
  }
  const parsedRoleAssignments = RoleAssignmentSchema.safeParse(value);
  return parsedRoleAssignments.success ? parsedRoleAssignments.data : [];
};
var UserApiClient = class {
  constructor(sp, eupNoReplyMail) {
    this.sp = sp;
    this.eupNoReplyMail = eupNoReplyMail;
  }
  sp;
  eupNoReplyMail;
  /**
       * @param groupId 
       * @returns teamMembers of current team mapped to array of {@link IPerson}
       */
  async getTeamMembers() {
    const channelMembersApiUrl = `/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/members`;
    const channelMembersResponse = await this.sp.graphGet(channelMembersApiUrl);
    const channelMembers = Array.isArray(channelMembersResponse.data) ? channelMembersResponse.data : [];
    const apiUrl = `groups/${this.sp.getTeamId()}/members`;
    const response = await this.sp.graphGet(apiUrl);
    const teamMembersData = Array.isArray(response.data) ? response.data : [];
    const filteredData = teamMembersData.filter(
      (teamMember) => channelMembers.some(
        (channelMember) => channelMember.email === teamMember.mail
      ) && !Object.values(SystemAccount).includes(teamMember.mail)
    ).map((teamMember) => {
      const match = channelMembers.find(
        (channelMember) => channelMember.email === teamMember.mail
      );
      return {
        ...teamMember,
        channelMemberId: match?.id
        // 👈 THIS is what you were missing
      };
    });
    return mapPersons(filteredData);
  }
  async getExistingUser(mail) {
    const encodedMail = encodeURIComponent(mail);
    const apiUrl = `/users?$filter=mail eq '${encodedMail}'`;
    const result = await this.sp.graphGet(apiUrl);
    return Array.isArray(result.data) && result.data.length > 0 ? mapPerson(result.data[0]) : void 0;
  }
  async getExistingUsers(mail) {
    const encodedMail = encodeURIComponent(mail);
    const apiUrl = `/users?$filter=startswith(mail,'${encodedMail}') or startswith(displayName,'${encodedMail}') or startswith(userPrincipalName,'${encodedMail}')&$select=id,displayName,mail,userPrincipalName&$top=10`;
    const result = await this.sp.graphGet(apiUrl);
    const data = Array.isArray(result.data) ? result.data : [];
    return mapPersons(data);
  }
  /**
   * Retrieves the photo of a user based on their ID.
   * @param userIdOrEmail The ID or email of the user whose photo is to be fetched.
   * @returns A promise that resolves to a URL representing the user's photo.
   */
  async getUserPhoto(userIdOrEmail) {
    const photoResponse = await this.sp.graphGet(`/users/${userIdOrEmail}/photo/$value`);
    return photoResponse.data && photoResponse.data instanceof Blob ? URL.createObjectURL(photoResponse.data) : "";
  }
  /**
   * Retrieves the current user's SharePoint groups and filters them to only include
   * groups defined in internal and external enums.
   * 
   * @returns Array of group titles from SPGroupInternal or SPGroupExternal enums the user belongs to.
   */
  async getCurrentUserGroups() {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/currentuser?$expand=Groups`;
    const result = await this.sp.spGet(apiUrl);
    const groups = result.data && Array.isArray(result.data.Groups) ? result.data.Groups : [];
    const allEnumGroupTitles = [
      ...Object.values(SPGroupInternal),
      ...Object.values(SPGroupExternal)
    ];
    const relevantGroupTitles = groups.map((g) => {
      const parsed = z.object({
        Title: z.string()
      }).safeParse(g);
      return parsed.success ? parsed.data.Title : void 0;
    }).filter((title) => title && allEnumGroupTitles.includes(title));
    return relevantGroupTitles;
  }
  async getSiteGroupUsersBatch(groupIds) {
    const boundary = "batch_" + Date.now();
    const batchParts = groupIds.map((groupId) => {
      const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users`;
      return [
        `--${boundary}`,
        "Content-Type: application/http",
        "Content-Transfer-Encoding: binary",
        `Content-ID: ${groupId}`,
        "",
        `GET ${apiUrl} HTTP/1.1`,
        "Accept: application/json",
        "",
        ""
      ].join("\r\n");
    });
    const batchBody = batchParts.join("\r\n") + `\r
--${boundary}--\r
`;
    const response = await this.sp.spPost(
      `${this.sp.getAbsoluteUrl()}/_api/$batch`,
      batchBody,
      {
        "Content-Type": `multipart/mixed; boundary=${boundary}`,
        "Accept": "application/json",
        "OData-Version": "4.0",
        "OData-MaxVersion": "4.0"
      }
    );
    const groupMap = response.data && typeof response.data === "string" ? parseBatchSiteGroupUsersResponse(response.data, groupIds) : /* @__PURE__ */ new Map();
    return groupMap;
  }
  async ensureRole() {
    const roleId = await this.getRoleDefinitionId("\xDApravy bez maz\xE1n\xED" /* CustomEdit */);
    if (!roleId) {
      const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roledefinitions`;
      const body = {
        BasePermissions: { High: 432, Low: 1011028583 },
        Description: "Read, Add, Edit without Delete",
        Name: "\xDApravy bez maz\xE1n\xED" /* CustomEdit */,
        Order: 2147483647
      };
      await this.sp.spPost(apiUrl, body);
    }
    const missingRoles = await this.getMissingRoleAssignments();
    await this.assingMissingRoles(missingRoles);
  }
  async getRoleDefinitionId(roleName) {
    const filter = `Name eq '${roleName}'`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roledefinitions`;
    const result = await this.sp.spGet(`${apiUrl}?$filter=${filter}`);
    return Array.isArray(result.data) && result.data[0] ? result.data[0].Id : void 0;
  }
  /**
   * Creates multiple site groups based on the provided group names.
   * @param groupNames An array of group names to be created (either internal or external).
   * @param increaseProgressCount A callback function to track progress during the group creation process.
   * @returns A promise that resolves when all site groups are created.
   */
  async createSiteGroups(groupNames, increaseProgressCount) {
    for (let i = 0; i < groupNames.length; i++) {
      if (Object.values(SPGroupRO).includes(groupNames[i])) {
        await this.createSiteROGroup(groupNames[i]);
      } else {
        await this.createSiteGroup(groupNames[i], increaseProgressCount);
      }
    }
  }
  /**
   * Creates a single site group based on the provided group name.
   * @param groupName The name of the group to be created (either internal, external, or read-only).
   * @param increaseProgressCount A callback function to track progress after each group creation.
   * @returns A promise that resolves when the site group is created.
   */
  async createSiteGroup(groupName, increaseProgressCount) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups`;
    const body = { Title: groupName };
    await this.sp.spPost(apiUrl, body);
    increaseProgressCount();
  }
  async createSiteROGroup(groupName) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups`;
    const body = { Title: groupName };
    const result = await this.sp.spPost(apiUrl, body);
    const groupId = result.data ? (result?.data).Id : void 0;
    const roleId = await this.getRoleDefinitionId("\u010Cten\xED" /* Read */);
    if (groupId && roleId) {
      await this.assignRoleToGroup(groupId, roleId);
    }
  }
  async assingMissingRoles(missingRoles) {
    for (let i = 0; i < missingRoles.length; i++) {
      const { groupId, roleId } = missingRoles[i];
      if (roleId) {
        await this.assignRoleToGroup(groupId, roleId);
      }
    }
  }
  async assignRoleToGroup(groupId, roleId) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
    await this.sp.spPost(apiUrl);
  }
  async getMissingRoleAssignments() {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings`;
    const result = await this.sp.spGet(apiUrl);
    const customEditId = await this.getRoleDefinitionId("\xDApravy bez maz\xE1n\xED" /* CustomEdit */);
    const editId = await this.getRoleDefinitionId("\xDApravy" /* Edit */);
    const missingAssignments = [];
    const data = Array.isArray(result.data) ? parseRoleAssignments(result.data) : [];
    for (let i = 0; i < data.length; i++) {
      const item = data[i];
      const name = item.Member.LoginName;
      if (!Object.values(SPGroupInternal).includes(name) && !Object.values(SPGroupExternal).includes(name)) {
        continue;
      }
      const roleId = Object.values(SPGroupInternal).includes(name) ? editId : customEditId;
      const role = item.RoleDefinitionBindings.find((r) => r.Id === roleId);
      if (!role) {
        missingAssignments.push({ groupId: item.Member.Id, roleId });
      }
    }
    return missingAssignments;
  }
  /**
   * Retrieves all invitations from the Invitations list.
   * @returns A promise that resolves to an array of invitation objects.
   */
  async getInvitations(teamMembers) {
    const filter = `ChannelId eq '${this.sp.getChannelId()}'`;
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
    const apiUrl = `${webUrl}/_api/web/GetListByTitle('${"Pozvanky" /* Invitations */}')/items?$filter=${filter}`;
    const response = await this.sp.spGet(apiUrl);
    const data = Array.isArray(response.data) ? response.data : [];
    const filteredData = data.filter((invitation) => !teamMembers.some((teamMember) => teamMember.email === invitation.Email));
    const userExistenceMap = await this.checkADUsersExistence(filteredData.map((item) => item.Email));
    return Array.isArray(response.data) ? mapInvitations(filteredData, userExistenceMap) : [];
  }
  /**
   * Adds a user to a specific SharePoint site group.
   * @param id The ID of the site group.
   * @param userEmail The email address of the user to be added.
   * @returns A promise that resolves when the user is added to the group.
   */
  async addUserToSiteGroup(id, userEmail) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
    const body = { LoginName: `i:0#.f|membership|${userEmail}` };
    await this.sp.spPost(apiUrl, body);
  }
  /**
   * Removes a user from a specific SharePoint site group.
   * @param groupId The ID of the site group.
   * @param userId The ID of the user to be removed.
   * @returns A promise that resolves when the user is removed from the group.
   */
  async removeUserFromsiteGroup(groupId, userId) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users/removeById(${userId})`;
    await this.sp.spPost(apiUrl);
  }
  async removeUserFromsiteGroupByEmail(groupId, email) {
    const loginName = `i:0#.f|membership|${email}`;
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users/removeByLoginName('${encodeURIComponent(loginName)}')`;
    await this.sp.spPost(apiUrl);
  }
  /**
   * Creates a new invitation in the Invitations list.
   * @param email The email address of the person being invited.
   * @param name The name of the person being invited.
   * @param roles The roles assigned to the person being invited.
   */
  async createInvitation(email, name, roles, language, existingUser = false) {
    const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.internalSiteName}`;
    const apiUrl = `${webUrl}/_api/web/GetListByTitle('${"Pozvanky" /* Invitations */}')/items`;
    const body = {
      "Title": name,
      "Jmeno": name,
      "Email": email,
      "ChannelId": this.sp.getChannelId(),
      "Roles": roles,
      "Language": language,
      "User": existingUser ? email : ""
    };
    const result = await this.sp.spPost(apiUrl, body);
    return !!result.ok;
  }
  /**
   * Checks if an Active Directory user exists by their email address.
   * @param email The email address of the user.
   * @returns A Promise that resolves to true if the user exists, otherwise false.
   */
  async checkADUserExistence(email) {
    const result = await this.sp.graphGet(`/users/${email}`);
    return result.ok ?? false;
  }
  async checkADUsersExistence(emails) {
    const resultMap = /* @__PURE__ */ new Map();
    const chunkSize = 20;
    for (let i = 0; i < emails.length; i += chunkSize) {
      const chunk = emails.slice(i, i + chunkSize);
      const batchBody = {
        requests: chunk.map((email, index) => ({
          id: `${index + 1}`,
          method: "GET",
          url: `/users/?$filter=mail eq '${encodeURIComponent(email)}'`
        }))
      };
      const result = await this.sp.graphPost(`/$batch`, batchBody);
      const responses = result.data?.responses || [];
      for (const response of responses) {
        const email = chunk[parseInt(response.id, 10) - 1];
        resultMap.set(email, response.status === 200 && Array.isArray(response.body.value) && response.body.value.length > 0);
      }
    }
    return resultMap;
  }
  /**
      * checks if user has access to EUP-Sablony site
      */
  async checkTemplatesAccess() {
    const apiUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), "")}/sites/${this.sp.templateSiteName}/_api/web?$select=Id`;
    const result = await this.sp.spGet(apiUrl);
    return !!result.ok;
  }
  /**
   * Checks if a user exists in Active Directory by email, 
   * creates an invitation if the user does not exist, 
   * or adds the existing user to the team.
   * 
   * @param email - The email address of the user to check or invite.
   * @param name - The display name of the user for the invitation.
   */
  async checkUserAndCreateInvitation(email, name, roles, language) {
    const userExists = await this.checkADUserExistence(email);
    if (!userExists) {
      return await this.createInvitation(email, name, roles, language);
    } else {
      return await this.addUserByEmailToTeamAndChannel(email);
    }
  }
  // public async checkUserAndAddToTeam(email:string):Promise<
  /**
   * Adds a member to a team by their email address.
   * @param email The email address of the user to add to the team.
   * @returns A promise indicating the completion of the operation.
   */
  async addUserByEmailToTeamAndChannel(email, role = "member" /* Member */) {
    const userData = await this.sp.graphGet(`/users/?$filter=mail eq '${encodeURIComponent(email)}'`);
    const userId = Array.isArray(userData.data) ? userData.data[0]?.id : void 0;
    if (!userId) {
      return false;
    }
    const result = await this.addUserOrGroupByIdToTeamAndChannel(userId, role, email);
    return result;
  }
  async addEUPAdminsAsOwners() {
    const ownersApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id,Title,LoginName`;
    const ownersResult = await this.sp.spGet(ownersApiUrl);
    if (!ownersResult.ok) {
      return false;
    }
    const ownersId = ownersResult.data.Id;
    const adminId = await this.getGroupIdByName("Admin" /* ADMIN */);
    const ids = [ownersId, adminId];
    const aadGroupClaim = `c:0t.c|tenant|${"7397659e-48e7-48fa-84f0-bd1d6f290538" /* EUPAdminsId */}`;
    const emailClaim = `i:0#.f|membership|${this.eupNoReplyMail}`;
    let result = true;
    for (const id of ids) {
      const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
      for (const login of [aadGroupClaim, emailClaim]) {
        const response = await this.sp.spPost(apiUrl, { LoginName: login });
        if (!response.ok) {
          result = false;
        }
      }
    }
    return result;
  }
  async addEUPMembersAsMembers() {
    const memberssApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/associatedmembergroup?$select=Id,Title,LoginName`;
    const membersResult = await this.sp.spGet(memberssApiUrl);
    if (!membersResult.ok) {
      return false;
    }
    const membersId = membersResult.data.Id;
    const roRoleId = await this.getGroupIdByName("Read Only role" /* RORole */);
    const ids = [membersId, ...roRoleId !== void 0 ? [roRoleId] : []];
    const aadGroupClaim = `c:0t.c|tenant|${"6834f589-81b2-4e47-9ec0-a41243a85aab" /* EUPMembersId */}`;
    let result = true;
    for (const id of ids) {
      const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
      const response = await this.sp.spPost(apiUrl, { LoginName: aadGroupClaim });
      if (!response.ok) {
        result = false;
      }
    }
    return result;
  }
  async getGroupIdByName(groupName) {
    const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups?$filter=Title eq '${groupName}'`;
    const result = await this.sp.spGet(apiUrl);
    if (!result.ok) {
      return void 0;
    }
    const groups = Array.isArray(result.data) ? result.data : [];
    return groups.length > 0 ? groups[0].Id : void 0;
  }
  async addUserOrGroupByIdToTeamAndChannel(userId, role = "member" /* Member */, email) {
    const addedToTeam = await this.addMemberToTeam(userId, this.sp.getTeamId(), role);
    const isPrivateChannel = await this.sp.isPrivateChannel();
    if (isPrivateChannel) {
      const addedToChannel = await this.addMemberToPrivateChannel(this.sp.getTeamId(), this.sp.getChannelId(), userId, role, email);
      const owners = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id`);
      await this.addUserToSiteGroup(owners.data?.Id, email);
      return addedToChannel;
    }
    if (role === "guest" /* Guest */) {
      await this.fixGuestGroup(email);
    }
    return addedToTeam;
  }
  async removeUserFromTeamOrChannel(person) {
    const isPrivateChannel = await this.sp.isPrivateChannel();
    if (isPrivateChannel) {
      const result = await this.sp.graphDelete(`/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/members/${person.channelMemberId}`);
      return !!result.ok;
    } else {
      const result = await this.sp.graphDelete(`/groups/${this.sp.getTeamId()}/members/${person.id}/$ref`);
      return !!result.ok;
    }
  }
  async addMemberToTeam(userId, teamId, role) {
    const body = {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": [role],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
    };
    const apiUrl = `/teams/${teamId}/members`;
    const result = await this.sp.graphPost(apiUrl, body);
    return !!result.ok;
  }
  async addMemberToPrivateChannel(teamId, channelId, userId, role, email) {
    const body = {
      "@odata.type": "#microsoft.graph.aadUserConversationMember",
      "roles": [role],
      "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
    };
    const apiUrl = `/teams/${teamId}/channels/${channelId}/members`;
    const result = await this.sp.graphPost(apiUrl, body);
    const owners = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id`);
    await this.addUserToSiteGroup(owners.data?.Id, email);
    if (role === "guest" /* Guest */) {
      await this.fixGuestGroup(email);
    }
    return !!result.ok;
  }
  async fixGuestGroup(email) {
    const visitors = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedvisitorgroup?$select=Id`);
    const members = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedmembergroup?$select=Id`);
    await this.addUserToSiteGroup(visitors.data.Id, email);
    await this.removeUserFromsiteGroupByEmail(members.data.Id, email);
  }
  async getUserName(userEmail) {
    const loginName = encodeURIComponent(`i:0#.f|membership|${userEmail}`);
    const result = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/siteusers/getByLoginName('${loginName}')`);
    return result.data && result.data.Title ? result.data.Title : "";
  }
  async getCurrentUser() {
    const result = await this.sp.graphGet("/me");
    return result.data ? mapPerson(result.data) : void 0;
  }
};

// src/api/ApiClient.ts
var ApiClient = class {
  users;
  structure;
  requests;
  qnA;
  notifications;
  install;
  forms;
  files;
  detail;
  calendar;
  actions;
  chat;
  sp;
  constructor(context, internalSiteName, templateSiteName, eupNoReplyMail) {
    this.sp = new SPApiClient(context, internalSiteName, templateSiteName);
    this.users = new UserApiClient(this.sp, eupNoReplyMail);
    this.files = new FileApiClient(this.sp);
    this.structure = new StructureApiClient(this.sp);
    this.requests = new RequestApiClient(this.files, this.sp);
    this.qnA = new QnAApiClient(this.files, this.sp);
    this.notifications = new NotificationApiClient(this.sp, context);
    this.install = new InstallApiClient(this.sp);
    this.forms = new FormApiClient(this.sp);
    this.detail = new DetailApiClient(this.sp);
    this.calendar = new CalendarApiClient(this.sp);
    this.actions = new ActionApiClient(this.sp);
    this.chat = new ChatApiClient(this.sp);
  }
};

// src/api/createApiClient.ts
var createApiClient = async (context, internalSiteName, templateSiteName, eupNoReplyMail, ctx, queryTeamId, queryChannelId) => {
  const client = new ApiClient(context, internalSiteName, templateSiteName, eupNoReplyMail);
  if (ctx || queryTeamId && queryChannelId) {
    await client.sp.init(ctx, queryTeamId, queryChannelId);
  }
  return client;
};

export { AppTab, ChannelStatus, ChannelType, ColumnsActions, ColumnsCategories, ColumnsDetails, ColumnsDocTypes, ColumnsDuties, ColumnsForms, ColumnsQA, ColumnsReqTypes, ColumnsRequests, ColumnsRequired, ColumnsTrigger, ContactDesc, ContextMenuAction, DataDispatchAction, DocumentTypeName, EventPeriodicity, EventType, FileActionOptions, FileStatus, FileValidityAction, Filter, FinancialDoc, FolderStatus, FormStatus, FormType, Language, LibraryName, ListName, LoadingStatus, NavDispatchAction, Permissions, QACategory, QAItemType, QAPriority, QASortType, QAStatus, RequestAction, RequestFields, RequestFormFields, RequestStatus, RequestType, RoleAlias, RoleDefinition, SPGroupExternal, SPGroupInternal, SPGroupRO, SettingsMenuKey, SharedDocuments, SystemAccount, TeamName, TeamRole, TemplatesListName, ToastType, UserType, addClassToComboBoxItem, addClassToDropdownItem, classifyUploader, cleanFolderName, convertUIVersionToLabel, copyLink, createApiClient, createCategoryOptions, createUniqueFileName, delay, downloadFile, downloadFiles, expandRoleAliases, expandRoles, findNodeById, formatDateForSharePoint, getCleanBaseName, getContextMenuHeight, getFileExtension, getFileName, getFileNameWOExt, getFileTypeTitle, getInitials, getLatestDate, getLocalizationPathById, getPhaseNodeByCategoryId, getStageByCategoryId, getStageNodeByCategoryId, isFormRecord, normalizeDocType, openFileInDesktopApp, openFileInTab, parseFormData, parseNameAndDate, parseSharePointDate, splitFolderPath, transformPath };
//# sourceMappingURL=index.js.map
//# sourceMappingURL=index.js.map