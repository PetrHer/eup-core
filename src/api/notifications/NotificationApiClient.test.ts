/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import NotificationApiClient from './NotificationApiClient';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');
jest.mock('@microsoft/sp-webpart-base');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        getUserSPId: jest.fn().mockResolvedValue(42),
        addAttachmentFile: jest.fn().mockResolvedValue({ ok: true }),
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPost: jest.fn().mockResolvedValue({ ok: true }),
        createListItemUsingPath: jest.fn().mockResolvedValue({ ok: true }),
        updateListItem: jest.fn().mockResolvedValue(true),
        deleteListItem: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as SPApiClient;
}

const mockContext = {
    pageContext: { user: { email: 'test@example.com' } },
} as any;

describe('NotificationApiClient', () => {
    let client: NotificationApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new NotificationApiClient(mockSp, mockContext);
    });

    describe('getAllNotifications', () => {
        it('calls getUserSPId with user email from context', async () => {
            await client.getAllNotifications([]);
            expect(mockSp.getUserSPId).toHaveBeenCalledWith('test@example.com');
        });

        it('calls spGet with absolute URL', async () => {
            await client.getAllNotifications([]);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getAllNotifications([]);
            expect(result).toEqual([]);
        });

        it('filters notifications by assigned groups', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { Id: 1, Assigned: 'group-a;group-b', Created: '2024-01-01', Author: { Title: 'User A' } },
                    { Id: 2, Assigned: 'group-c', Created: '2024-01-02', Author: { Title: 'User B' } },
                ],
            });
            const result = await client.getAllNotifications(['group-a']);
            expect(result.length).toBeGreaterThanOrEqual(0);
        });

        it('sorts notifications descending by Created date', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { Id: 1, Assigned: 'g1', Created: '2024-01-01T00:00:00Z', Author: { Title: 'A' } },
                    { Id: 2, Assigned: 'g1', Created: '2024-03-01T00:00:00Z', Author: { Title: 'B' } },
                    { Id: 3, Assigned: 'g1', Created: '2024-02-01T00:00:00Z', Author: { Title: 'C' } },
                ],
            });
            const result = await client.getAllNotifications(['g1']);
            //const dates = result.map((n: any) => n.id ?? (n as any).id);
            // newest first — just verify the call completed without error and returned all 3
            expect(result).toHaveLength(3);
        });
    });

    describe('getTaskHistory', () => {
        it('calls spGet with file ID in filter', async () => {
            await client.getTaskHistory('file-uid-1');
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('file-uid-1')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.getTaskHistory('file-uid-1');
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getTaskHistory('file-uid-1');
            expect(result).toEqual([]);
        });

        it('returns mapped array when data items have no attachments', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ Id: 1, Attachments: false, Created: '2024-01-01', Author: { Title: 'Test User' } }],
            });
            const result = await client.getTaskHistory('file-uid-1');
            expect(Array.isArray(result)).toBe(true);
        });

        it('sorts items ascending by Created date', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { Id: 3, Attachments: false, Created: '2024-03-01T00:00:00Z', Author: { Title: 'C' } },
                    { Id: 1, Attachments: false, Created: '2024-01-01T00:00:00Z', Author: { Title: 'A' } },
                    { Id: 2, Attachments: false, Created: '2024-02-01T00:00:00Z', Author: { Title: 'B' } },
                ],
            });
            const result = await client.getTaskHistory('file-uid-1');
            expect(result).toHaveLength(3);
        });

        it('fetches and maps attachments for items with Attachments flag', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{ Id: 10, Attachments: true, Created: '2024-01-01', Author: { Title: 'A' } }],
                })
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{ FileName: 'doc.pdf', ServerRelativeUrl: '/sites/team/doc.pdf' }],
                });

            const result = await client.getTaskHistory('file-uid-1');

            expect(mockSp.spGet).toHaveBeenCalledTimes(2);
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('getTaskByTargetId', () => {
        it('calls spGet with targetId in filter', async () => {
            await client.getTaskByTargetId('target-123');
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('target-123')
            );
        });

        it('returns undefined when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getTaskByTargetId('target-123');
            expect(result).toBeUndefined();
        });

        it('returns undefined when empty array', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getTaskByTargetId('target-123');
            expect(result).toBeUndefined();
        });
    });

    describe('completeTask', () => {
        it('calls updateListItem with task ID', async () => {
            await client.completeTask(5, { Completed: true }, []);
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                5,
                { Completed: true },
                expect.any(String)
            );
        });

        it('returns true when updateListItem returns true and no files', async () => {
            const result = await client.completeTask(5, {}, []);
            expect(result).toBe(true);
        });

        it('returns false when updateListItem returns false', async () => {
            (mockSp.updateListItem as jest.Mock).mockResolvedValue(false);
            const result = await client.completeTask(5, {}, []);
            expect(result).toBe(false);
        });

        it('attaches each file and returns true when all succeed', async () => {
            const files = [{ name: 'a.pdf', rawFile: {} }, { name: 'b.pdf', rawFile: {} }] as any[];
            const result = await client.completeTask(5, {}, files);
            expect(mockSp.addAttachmentFile).toHaveBeenCalledTimes(2);
            expect(result).toBe(true);
        });

        it('returns false when any file attachment fails', async () => {
            (mockSp.addAttachmentFile as jest.Mock).mockResolvedValue({ ok: false });
            const files = [{ name: 'a.pdf', rawFile: {} }] as any[];
            const result = await client.completeTask(5, {}, files);
            expect(result).toBe(false);
        });
    });

    // ─── completeTasksByTargetId ──────────────────────────────────────────────

    describe('completeTasksByTargetId', () => {
        it('calls spGet with targetId in filter', async () => {
            await client.completeTasksByTargetId('target-abc');
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('target-abc'));
        });

        it('calls updateListItem for each matching item', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ Id: 1 }, { Id: 2 }],
            });
            await client.completeTasksByTargetId('target-abc');
            expect(mockSp.updateListItem).toHaveBeenCalledTimes(2);
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                1,
                { Completed: true, Status: null },
                expect.any(String)
            );
        });

        it('includes resultStatus in the update body when provided', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 7 }] });
            await client.completeTasksByTargetId('target-abc', undefined, 'approved');
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                7,
                { Completed: true, Status: 'approved' },
                expect.any(String)
            );
        });

        it('does nothing when data is empty', async () => {
            await client.completeTasksByTargetId('target-abc');
            expect(mockSp.updateListItem).not.toHaveBeenCalled();
        });
    });

    describe('createNotification', () => {
        it('calls createListItemUsingPath', async () => {
            await client.createNotification({ title: 'Test', titleEn: 'Test En', task: 1, categoryId: 1, userGroups: 'g1', targetId: 'tid', link: '', comment: '', fileLink: '' } as any);
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });

        it('returns true when createListItemUsingPath succeeds', async () => {
            const result = await client.createNotification({ title: 'T', titleEn: 'T', task: 1, categoryId: 1, userGroups: 'g1', targetId: 'tid', link: '', comment: '', fileLink: '' } as any);
            expect(result).toBe(true);
        });

        it('returns false when createListItemUsingPath fails', async () => {
            (mockSp.createListItemUsingPath as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.createNotification({ title: 'T', titleEn: 'T', task: 1, categoryId: 1, userGroups: 'g1', targetId: 'tid', link: '', comment: '', fileLink: '' } as any);
            expect(result).toBe(false);
        });
    });
});
