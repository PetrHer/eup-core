import { WebPartContext } from "@microsoft/sp-webpart-base";

import ActionApiClient from "./actions/ActionApiClient";
import CalendarApiClient from "./calendar/CalendarApiClient";
import ChatApiClient from "./chat/ChatApiClient";
import DetailApiClient from "./detail/DetailApiClient";
import FileApiClient from "./files/FileApiClient";
import FormApiClient from "./forms/FormApiClient";
import { IApiClient } from "./IApiClient";
import InstallApiClient from "./install/InstallApiClient";
import NotificationApiClient from "./notifications/NotificationApiClient";
import QnAApiClient from "./qnA/QnAApiClient";
import RequestApiClient from "./requests/RequestApiClient";
import SPApiClient from "./SPApiClient";
import StructureApiClient from "./structure/StructureApiClient";
import UserApiClient from "./users/UserApiClient";

export default class ApiClient implements IApiClient {
    public readonly users: UserApiClient;
    public readonly structure: StructureApiClient;
    public readonly requests: RequestApiClient;
    public readonly qnA: QnAApiClient;
    public readonly notifications: NotificationApiClient;
    public readonly install: InstallApiClient;
    public readonly forms: FormApiClient;
    public readonly files: FileApiClient;
    public readonly detail: DetailApiClient;
    public readonly calendar: CalendarApiClient;
    public readonly actions: ActionApiClient;
    public readonly chat: ChatApiClient;

    public readonly sp: SPApiClient;

    constructor(context: WebPartContext, internalSiteName: string, templateSiteName: string, eupNoReplyMail: string) {
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
}