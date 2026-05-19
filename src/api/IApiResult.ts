interface IApiResult {
    ok: boolean;
    code?: number;
    data?: unknown;
    error?: string;
    exception?: unknown;
}

export default IApiResult;