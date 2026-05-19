/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import StructureApiClient from './StructureApiClient';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
    } as unknown as SPApiClient;
}

describe('StructureApiClient', () => {
    let client: StructureApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new StructureApiClient(mockSp);
    });

    describe('getStructure', () => {
        it('makes two spGet calls (categories and phases)', async () => {
            await client.getStructure();
            expect(mockSp.spGet).toHaveBeenCalledTimes(2);
        });

        it('queries categories with channel name in URL', async () => {
            await client.getStructure();
            const firstCallUrl: string = (mockSp.spGet as jest.Mock).mock.calls[0][0];
            expect(firstCallUrl).toContain('TestChannel');
        });

        it('queries QA list with channel name in URL', async () => {
            await client.getStructure();
            const secondCallUrl: string = (mockSp.spGet as jest.Mock).mock.calls[1][0];
            expect(secondCallUrl).toContain('TestChannel');
        });

        it('uses absolute URL for both queries', async () => {
            await client.getStructure();
            const calls = (mockSp.spGet as jest.Mock).mock.calls;
            calls.forEach(([url]: [string]) => {
                expect(url).toContain('https://example.sharepoint.com/sites/team');
            });
        });

        it('returns an array', async () => {
            const result = await client.getStructure();
            expect(Array.isArray(result)).toBe(true);
        });
    });
});
