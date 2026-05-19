/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import ActionApiClient from './ActionApiClient';
import { IScheduledAction } from '../../interfaces';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelId: jest.fn().mockReturnValue('channel-123'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        getTeamId: jest.fn().mockReturnValue('team-id-123'),
        templateSiteName: 'template-site',
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        createListItemUsingPath: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as SPApiClient;
}

describe('ActionApiClient', () => {
    let client: ActionApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new ActionApiClient(mockSp);
    });

    describe('getWorkflowActions', () => {
        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getWorkflowActions();
            expect(result).toEqual([]);
        });

        it('calls spGet with site absolute URL', async () => {
            await client.getWorkflowActions();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('uses template site URL when fromTemplates is true', async () => {
            await client.getWorkflowActions(true);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('/sites/template-site')
            );
        });

        it('returns mapped array from API data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getWorkflowActions();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('getScheduledActions', () => {
        it('calls spGet with channel name in URL', async () => {
            await client.getScheduledActions();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('returns empty Map when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getScheduledActions();
            expect(result).toBeInstanceOf(Map);
            expect(result.size).toBe(0);
        });

        it('returns a Map for array data', async () => {
            const result = await client.getScheduledActions();
            expect(result).toBeInstanceOf(Map);
        });
    });

    describe('getScheduledActionByFileId', () => {
        it('returns undefined without calling API when fileId is not provided', async () => {
            const result = await client.getScheduledActionByFileId(undefined);
            expect(result).toBeUndefined();
            expect(mockSp.spGet).not.toHaveBeenCalled();
        });

        it('calls spGet with fileId in URL filter', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            await client.getScheduledActionByFileId(42);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('42')
            );
        });

        it('returns undefined when no matching item found', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getScheduledActionByFileId(42);
            expect(result).toBeUndefined();
        });
    });

    describe('createScheduledAction', () => {
        const action = { fileId: 10, actionId: 5 } as unknown as IScheduledAction;

        it('calls createListItemUsingPath', async () => {
            await client.createScheduledAction(action, 'Test Action');
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });

        it('returns true when createListItemUsingPath succeeds', async () => {
            const result = await client.createScheduledAction(action, 'Test Action');
            expect(result).toBe(true);
        });

        it('returns false when createListItemUsingPath fails', async () => {
            (mockSp.createListItemUsingPath as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.createScheduledAction(action, 'Test Action');
            expect(result).toBe(false);
        });

        it('includes fileId and actionId in request body', async () => {
            await client.createScheduledAction(action, 'My Title');
            const callArg = (mockSp.createListItemUsingPath as jest.Mock).mock.calls[0][0];
            const formValues: any[] = callArg.formValues;
            const fileIdField = formValues.find((f: any) => f.FieldValue === '10');
            const actionIdField = formValues.find((f: any) => f.FieldValue === '5');
            expect(fileIdField).toBeDefined();
            expect(actionIdField).toBeDefined();
        });
    });
});
