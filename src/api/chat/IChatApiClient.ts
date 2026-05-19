export interface IChatApiClient {
    getMessages: () => Promise<any[]>;
}