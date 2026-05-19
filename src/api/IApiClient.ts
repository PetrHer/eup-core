import { IActionApiClient } from "./actions/IActionApiClient";
import { ICalendarApiClient } from "./calendar/ICalendarApiClient";
import { IChatApiClient } from "./chat/IChatApiClient";
import { IDetailApiClient } from "./detail/IDetailApiClient";
import { IFileApiClient } from "./files/IFileApiClient";
import { IFormApiClient } from "./forms/IFormApiClient";
import { IInstallApiClient } from "./install/IInstallApiClient";
import { ISPApiClient } from "./ISPApiClient";
import { INotificationApiClient } from "./notifications/INotificationApiClient";
import { IQnAApiClient } from "./qnA/IQnAApiClient";
import { IRequestApiClient } from "./requests/IRequestApiClient";
import { IStructureApiClient } from "./structure/IStructureApiClient";
import { IUserApiClient } from "./users/IUserApiClient";

export interface IApiClient {
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