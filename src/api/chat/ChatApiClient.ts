import SPApiClient from "../SPApiClient";
import { IChatApiClient } from "./IChatApiClient";

export default class ChatApiClient implements IChatApiClient {
    /**
     * Constructs a ChatApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    public async getMessages(): Promise<any[]> {
        const apiUrl = `/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/messages`;
        const result = await this.sp.graphGet(apiUrl);
        console.log(result);
        return [];
    }
}