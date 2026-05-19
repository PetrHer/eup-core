/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import { BaseApiClient } from './BaseApiClient';

jest.mock('@microsoft/sp-http', () => ({
    SPHttpClient: {
        configurations: {
            v1: {}
        }
    }
}));

// Concrete subclass to test the abstract class
class TestApiClient extends BaseApiClient {
    constructor(ctx: any) {
        super(ctx, () => 'https://example.sharepoint.com/sites/team');
    }
}

function makeResponse(status: number, ok: boolean, json?: any, text?: string) {
    return {
        status,
        ok,
        json: jest.fn().mockResolvedValue(json),
        text: jest.fn().mockResolvedValue(text ?? ''),
        clone: jest.fn().mockReturnValue({
            json: jest.fn().mockResolvedValue(json)
        }),
        arrayBuffer: jest.fn().mockResolvedValue(new ArrayBuffer(8)),
    };
}

function createMockContext() {
    const spHttpClientGet = jest.fn();
    const spHttpClientPost = jest.fn();
    const graphClientGet = jest.fn();
    const graphClientPost = jest.fn();
    const graphApi = jest.fn().mockReturnValue({ get: graphClientGet, post: graphClientPost });
    const getClient = jest.fn().mockResolvedValue({ api: graphApi });

    return {
        ctx: {
            spHttpClient: { get: spHttpClientGet, post: spHttpClientPost },
            msGraphClientFactory: { getClient },
            pageContext: { user: { email: 'user@test.com' } },
        } as any,
        spGet: spHttpClientGet,
        spPost: spHttpClientPost,
        graphGet: graphClientGet,
        graphPost: graphClientPost,
        graphApi,
    };
}

describe('BaseApiClient', () => {
    let client: TestApiClient;
    let mocks: ReturnType<typeof createMockContext>;

    beforeEach(() => {
        jest.clearAllMocks();
        mocks = createMockContext();
        client = new TestApiClient(mocks.ctx);
    });

    describe('spGet', () => {
        it('returns ok result with data for a single-object response', async () => {
            mocks.spGet.mockResolvedValue(makeResponse(200, true, { Id: 1, Title: 'Test' }));
            const result = await client.spGet('https://example.com/_api/test');
            expect(result.ok).toBe(true);
            expect(result.code).toBe(200);
            expect((result.data as any).Id).toBe(1);
        });

        it('returns aggregated array when response has a value array', async () => {
            mocks.spGet.mockResolvedValue(makeResponse(200, true, { value: [{ Id: 1 }, { Id: 2 }] }));
            const result = await client.spGet('https://example.com/_api/test');
            expect(result.ok).toBe(true);
            expect(Array.isArray(result.data)).toBe(true);
            expect((result.data as any[]).length).toBe(2);
        });

        it('paginates when @odata.nextLink is present', async () => {
            mocks.spGet
                .mockResolvedValueOnce(makeResponse(200, true, { value: [{ Id: 1 }], '@odata.nextLink': 'https://example.com/next' }))
                .mockResolvedValueOnce(makeResponse(200, true, { value: [{ Id: 2 }] }));
            const result = await client.spGet('https://example.com/_api/test');
            expect(mocks.spGet).toHaveBeenCalledTimes(2);
            expect((result.data as any[]).length).toBe(2);
        });

        it('returns ok:false with code on non-ok response', async () => {
            mocks.spGet.mockResolvedValue(makeResponse(404, false, { error: 'not found' }));
            const result = await client.spGet('https://example.com/_api/test', {}, false);
            expect(result.ok).toBe(false);
            expect(result.code).toBe(404);
        });

        it('returns ok:false with code 500 when an exception is thrown', async () => {
            mocks.spGet.mockRejectedValue(new Error('network failure'));
            const result = await client.spGet('https://example.com/_api/test', {}, false);
            expect(result.ok).toBe(false);
            expect(result.code).toBe(500);
        });

        it('returns text data when responseType is text', async () => {
            mocks.spGet.mockResolvedValue(makeResponse(200, true, undefined, 'raw text'));
            const result = await client.spGet('https://example.com/_api/test', {}, true, 'text');
            expect(result.ok).toBe(true);
            expect(result.data).toBe('raw text');
        });
    });

    describe('spPost', () => {
        it('returns ok result with parsed json body on success', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, { Id: 5 }));
            const result = await client.spPost('https://example.com/_api/test', { Title: 'New' });
            expect(result.ok).toBe(true);
            expect((result.data as any).Id).toBe(5);
        });

        it('returns data:true for 204 No Content', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(204, true));
            const result = await client.spPost('https://example.com/_api/test');
            expect(result.ok).toBe(true);
            expect(result.data).toBe(true);
        });

        it('returns ok:false with exception on non-ok response', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(400, false, { error: 'bad request' }));
            const result = await client.spPost('https://example.com/_api/test');
            expect(result.ok).toBe(false);
        });

        it('returns ok:false with code 500 when an exception is thrown', async () => {
            mocks.spPost.mockRejectedValue(new Error('network failure'));
            const result = await client.spPost('https://example.com/_api/test');
            expect(result.ok).toBe(false);
            expect(result.code).toBe(500);
        });

        it('serializes object body to JSON string', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, {}));
            await client.spPost('https://example.com/_api/test', { key: 'val' });
            const opts = mocks.spPost.mock.calls[0][2];
            expect(opts.body).toBe(JSON.stringify({ key: 'val' }));
        });

        it('passes ArrayBuffer body as-is without serialization', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, {}));
            const buf = new ArrayBuffer(4);
            await client.spPost('https://example.com/_api/test', buf);
            const opts = mocks.spPost.mock.calls[0][2];
            expect(opts.body).toBe(buf);
        });
    });

    describe('spPatch', () => {
        it('delegates to spPost with X-HTTP-Method MERGE and IF-MATCH headers', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, {}));
            await client.spPatch('https://example.com/_api/test', { field: 'val' });
            const opts = mocks.spPost.mock.calls[0][2];
            expect(opts.headers['X-HTTP-Method']).toBe('MERGE');
            expect(opts.headers['IF-MATCH']).toBe('*');
        });
    });

    describe('spDelete', () => {
        it('delegates to spPost with X-Http-Method DELETE', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, {}));
            await client.spDelete('https://example.com/_api/test');
            const opts = mocks.spPost.mock.calls[0][2];
            expect(opts.headers['X-Http-Method']).toBe('DELETE');
            expect(opts.headers['IF-MATCH']).toBe('*');
        });
    });

    describe('spPut', () => {
        it('delegates to spPost with X-HTTP-Method PUT', async () => {
            mocks.spPost.mockResolvedValue(makeResponse(200, true, {}));
            await client.spPut('https://example.com/_api/test');
            const opts = mocks.spPost.mock.calls[0][2];
            expect(opts.headers['X-HTTP-Method']).toBe('PUT');
        });
    });

    describe('spGetBinary', () => {
        it('returns IApiResult with ok:true and ArrayBuffer data on success', async () => {
            const buf = new ArrayBuffer(16);
            mocks.spGet.mockResolvedValue({ ok: true, arrayBuffer: jest.fn().mockResolvedValue(buf) });
            const result = await client.spGetBinary('https://example.com/_api/file');
            expect(result.ok).toBe(true);
            expect(result.data).toBe(buf);
        });

        it('returns IApiResult with ok:false and error text when response is not ok', async () => {
            mocks.spGet.mockResolvedValue({ ok: false, text: jest.fn().mockResolvedValue('error text') });
            const result = await client.spGetBinary('https://example.com/_api/file');
            expect(result.ok).toBe(false);
            expect(result.error).toBe('error text');
        });

        it('returns IApiResult with ok:false when a network exception is thrown', async () => {
            mocks.spGet.mockRejectedValue(new Error('network error'));
            const result = await client.spGetBinary('https://example.com/_api/file');
            expect(result.ok).toBe(false);
            expect(result.error).toBe('network error');
        });
    });

    describe('spPostBinary', () => {
        it('returns IApiResult with ok:true on success', async () => {
            mocks.spPost.mockResolvedValue({ ok: true });
            const result = await client.spPostBinary('https://example.com/_api/upload', new ArrayBuffer(4));
            expect(result.ok).toBe(true);
        });

        it('returns IApiResult with ok:false when response is not ok', async () => {
            mocks.spPost.mockResolvedValue({ ok: false, text: jest.fn().mockResolvedValue('upload failed') });
            const result = await client.spPostBinary('https://example.com/_api/upload', new ArrayBuffer(4));
            expect(result.ok).toBe(false);
        });
    });

    describe('graphGet', () => {
        it('returns ok result with data on success', async () => {
            mocks.graphGet.mockResolvedValue({ value: [{ id: '1' }] });
            const result = await client.graphGet('/teams/abc/channels');
            expect(result.ok).toBe(true);
            expect(Array.isArray(result.data)).toBe(true);
        });

        it('returns data directly when no value wrapper', async () => {
            mocks.graphGet.mockResolvedValue({ displayName: 'Test Team' });
            const result = await client.graphGet('/groups/abc/sites/root');
            expect(result.ok).toBe(true);
            expect((result.data as any).displayName).toBe('Test Team');
        });

        it('returns ok:false when graph response is falsy', async () => {
            mocks.graphGet.mockResolvedValue(null);
            const result = await client.graphGet('/teams/abc');
            expect(result.ok).toBe(false);
        });

        it('returns ok:false when exception is thrown', async () => {
            mocks.graphGet.mockRejectedValue(new Error('graph error'));
            const result = await client.graphGet('/teams/abc');
            expect(result.ok).toBe(false);
        });
    });

    describe('graphPost', () => {
        it('returns ok result with data on success', async () => {
            mocks.graphPost.mockResolvedValue({ id: 'msg-1' });
            const result = await client.graphPost('/teams/abc/channels/xyz/messages', { body: 'hi' });
            expect(result.ok).toBe(true);
        });

        it('returns ok:false when graph response is falsy', async () => {
            mocks.graphPost.mockResolvedValue(null);
            const result = await client.graphPost('/teams/abc/channels/xyz/messages');
            expect(result.ok).toBe(false);
        });

        it('returns ok:false when exception is thrown', async () => {
            mocks.graphPost.mockRejectedValue(new Error('graph error'));
            const result = await client.graphPost('/teams/abc/channels/xyz/messages');
            expect(result.ok).toBe(false);
        });
    });
});
