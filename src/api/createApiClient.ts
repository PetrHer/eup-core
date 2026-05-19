import { WebPartContext } from "@microsoft/sp-webpart-base";
import ApiClient from "./ApiClient";

export const createApiClient = async (
    context: WebPartContext,
    internalSiteName: string,
    templateSiteName: string,
    eupNoReplyMail: string,
    ctx?: string,
    queryTeamId?: string,
    queryChannelId?: string
): Promise<ApiClient> => {
    const client = new ApiClient(context, internalSiteName, templateSiteName, eupNoReplyMail);
    if (ctx || (queryTeamId && queryChannelId)) {
        await client.sp.init(ctx, queryTeamId, queryChannelId);
    }
    return client;
}