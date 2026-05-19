/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import QnAApiClient from './QnAApiClient';
import { ColumnsQA, QAItemType } from '../../enums';
import FileApiClient from '../files/FileApiClient';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');
jest.mock('../files/FileApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPost: jest.fn().mockResolvedValue({ ok: true, data: { id: '7', text: 'hi', author: { name: 'U', email: 'u@x.com', isExternal: false }, createdDate: '2024-01-01T00:00:00Z' } }),
        spDelete: jest.fn().mockResolvedValue({ ok: true }),
        createListItemUsingPath: jest.fn().mockResolvedValue({
            ok: true,
            data: { value: [{ FieldName: 'Id', FieldValue: '123' }] },
        }),
        updateListItem: jest.fn().mockResolvedValue(true),
        updateItemInList: jest.fn().mockResolvedValue({ ok: true }),
        ensureFolder: jest.fn().mockResolvedValue(undefined),
    } as unknown as SPApiClient;
}

function createMockFiles() {
    return {
        loadFiles: jest.fn().mockResolvedValue([]),
        uploadFile: jest.fn().mockResolvedValue({ success: true, file: null }),
        uploadFiles: jest.fn().mockResolvedValue([
            { success: true, file: { name: 'doc.pdf', serverRelativeUrl: '/sites/team/Shared/doc.pdf' }, fileName: 'doc.pdf' },
        ]),
        updateFileByUrl: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as FileApiClient;
}

function makeRawItem(overrides: Record<string, any> = {}) {
    return {
        Id: 1,
        Title: 'Q1',
        UniqueId: 'uuid-1',
        Status: 'new',
        Category: 'cat',
        Phase: 'p1',
        Stage: 's1',
        Created: '2024-01-01T00:00:00Z',
        Modified: '2024-01-02T00:00:00Z',
        Author: { Title: 'User1' },
        [ColumnsQA.Answer]: 'answer text',
        [ColumnsQA.Answered]: '',
        [ColumnsQA.Order]: 1,
        [ColumnsQA.Priority]: 'high',
        [ColumnsQA.FilesData]: null,
        [ColumnsQA.SubCategory]: 'sub',
        ...overrides,
    };
}

describe('QnAApiClient', () => {
    let client: QnAApiClient;
    let mockSp: ReturnType<typeof createMockSp>;
    let mockFiles: ReturnType<typeof createMockFiles>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        mockFiles = createMockFiles();
        client = new QnAApiClient(mockFiles, mockSp);
    });

    describe('getQnAItems', () => {
        it('calls spGet with channel name in URL', async () => {
            await client.getQnAItems();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.getQnAItems();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getQnAItems();
            expect(result).toEqual([]);
        });

        it('returns mapped array from API data', async () => {
            const result = await client.getQnAItems();
            expect(Array.isArray(result)).toBe(true);
        });

        it('passes stage and phase filter when provided', async () => {
            await client.getQnAItems('stage1', 'phase1');
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('stage1')
            );
        });
    });

    describe('createQnAItem', () => {
        it('calls createListItemUsingPath', async () => {
            await client.createQnAItem('My question', 'cat1', 'stage1', 'phase1', 'high', 'user@example.com');
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });

        it('returns new item ID when createListItemUsingPath succeeds', async () => {
            const result = await client.createQnAItem('My question', 'cat1', 'stage1', 'phase1', 'high', 'user@example.com');
            expect(result).toBe('123');
        });

        it('returns undefined when no Id field returned', async () => {
            (mockSp.createListItemUsingPath as jest.Mock).mockResolvedValue({
                ok: true,
                data: { value: [] },
            });
            const result = await client.createQnAItem('Q', 'cat', 'stage', 'phase', 'low', 'user@example.com');
            expect(result).toBeUndefined();
        });

        it('includes question text in formValues', async () => {
            await client.createQnAItem('Test question', 'cat1', 'stage1', 'phase1', 'high', 'user@example.com');
            const [body] = (mockSp.createListItemUsingPath as jest.Mock).mock.calls[0];
            const titleField = body.formValues.find((f: any) => f.FieldName === 'Title');
            expect(titleField?.FieldValue).toBe('Test question');
        });

        it('uploads files and calls updateListItem when files are provided', async () => {
            const fileType = { folderPath: 'docs/sub' } as any;
            const files = [{ name: 'a.pdf', rawFile: new ArrayBuffer(0) }] as any[];
            await client.createQnAItem('Q', 'cat', 'stage', 'phase', 'high', 'user@example.com', fileType, 1, files);
            expect(mockFiles.uploadFiles).toHaveBeenCalled();
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                '123',
                expect.objectContaining({ [ColumnsQA.FilesData]: expect.any(String) }),
                expect.any(String)
            );
        });
    });

    describe('getQnACategories', () => {
        it('calls spGet with absolute URL', async () => {
            await client.getQnACategories();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns API data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getQnACategories();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('deleteComment', () => {
        it('calls spDelete with comment list name in URL', async () => {
            await client.deleteComment('QAList' as any, 'unique-id-abc', 1);
            expect(mockSp.spDelete).toHaveBeenCalledWith(
                expect.stringContaining('unique-id-abc')
            );
        });

        it('returns true when spDelete succeeds', async () => {
            const result = await client.deleteComment('QAList' as any, 'unique-id-abc', 1);
            expect(result).toBe(true);
        });

        it('returns false when spDelete fails', async () => {
            (mockSp.spDelete as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.deleteComment('QAList' as any, 'unique-id-abc', 1);
            expect(result).toBe(false);
        });
    });

    // ─── uploadQAFiles ────────────────────────────────────────────────────────

    describe('uploadQAFiles', () => {
        const fileType = { folderPath: 'docs/subfolder' } as any;
        const files = [{ name: 'a.pdf', rawFile: new ArrayBuffer(0) }] as any[];

        it('returns empty array when fileType is undefined', async () => {
            const result = await client.uploadQAFiles(files, 'user', 1, undefined);
            expect(result).toEqual([]);
        });

        it('returns empty array when files array is empty', async () => {
            const result = await client.uploadQAFiles([], 'user', 1, fileType);
            expect(result).toEqual([]);
        });

        it('calls ensureFolder and uploadFiles', async () => {
            await client.uploadQAFiles(files, 'user', 1, fileType);
            expect(mockSp.ensureFolder).toHaveBeenCalled();
            expect(mockFiles.uploadFiles).toHaveBeenCalled();
        });

        it('calls updateFileByUrl for each uploaded file with a serverRelativeUrl', async () => {
            await client.uploadQAFiles(files, 'user', 42, fileType);
            expect(mockFiles.updateFileByUrl).toHaveBeenCalledWith(
                '/sites/team/Shared/doc.pdf',
                { QAId: '0042' }
            );
        });

        it('skips updateFileByUrl when serverRelativeUrl is empty', async () => {
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([
                { success: true, file: { name: 'doc.pdf', serverRelativeUrl: '' }, fileName: 'doc.pdf' },
            ]);
            await client.uploadQAFiles(files, 'user', 1, fileType);
            expect(mockFiles.updateFileByUrl).not.toHaveBeenCalled();
        });

        it('skips failed upload entries', async () => {
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([
                { success: false, file: null, fileName: 'bad.pdf' },
            ]);
            const result = await client.uploadQAFiles(files, 'user', 1, fileType);
            expect(result).toHaveLength(0);
        });

        it('includes commentId in result when provided', async () => {
            const result = await client.uploadQAFiles(files, 'user', 1, fileType, 99);
            expect(result[0].commentId).toBe(99);
        });

        it('sets commentId to undefined when not provided', async () => {
            const result = await client.uploadQAFiles(files, 'user', 1, fileType);
            expect(result[0].commentId).toBeUndefined();
        });
    });

    // ─── addComment ───────────────────────────────────────────────────────────

    describe('addComment', () => {
        const qnAItem: any = {
            id: 1,
            uniqueId: 'uuid-1',
            filesData: undefined,
        };

        it('calls spPost with comment URL', async () => {
            await client.addComment(qnAItem, 'QA', 'hello', [], 'user');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('uuid-1'),
                expect.objectContaining({ text: 'hello' })
            );
        });

        it('returns newComment mapped from response when no files', async () => {
            const result = await client.addComment(qnAItem, 'QA', 'hello', [], 'user');
            expect(result.newComment).toBeDefined();
            expect(result.newComment?.text).toBe('hi');
        });

        it('returns { newComment: undefined, filesData: undefined } when spPost fails', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false, data: null });
            const result = await client.addComment(qnAItem, 'QA', 'hello', [], 'user');
            expect(result.newComment).toBeUndefined();
            expect(result.filesData).toBeUndefined();
        });

        it('uploads files and calls updateItemInList when files are provided', async () => {
            const files = [{ name: 'f.pdf', rawFile: new ArrayBuffer(0), type: { folderPath: 'docs/sub' } }] as any[];
            await client.addComment(qnAItem, 'QA', 'hello', files, 'user');
            expect(mockFiles.uploadFiles).toHaveBeenCalled();
            expect(mockSp.updateItemInList).toHaveBeenCalled();
        });

        it('merges existing filesData.Comment entries when files are provided', async () => {
            const itemWithFiles: any = {
                id: 1,
                uniqueId: 'uuid-1',
                filesData: { [QAItemType.Comment]: [{ name: 'old.pdf', serverRelativeUrl: '/old', commentId: 3 }] },
            };
            const files = [{ name: 'new.pdf', rawFile: new ArrayBuffer(0), type: { folderPath: 'docs/sub' } }] as any[];
            const result = await client.addComment(itemWithFiles, 'QA', 'hi', files, 'user');
            const commentFiles = result.filesData?.[QAItemType.Comment];
            expect(commentFiles?.length).toBeGreaterThan(1);
        });
    });

    // ─── getQnAItems with data (covers mapQnAItem / mapItemComments) ──────────

    describe('getQnAItems with items', () => {
        it('maps returned items into IQnAItem objects', async () => {
            const raw = makeRawItem();
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [raw] })
                .mockResolvedValueOnce({ ok: true, data: [] }); // getItemComments
            const result = await client.getQnAItems();
            expect(result).toHaveLength(1);
            expect(result[0].question).toBe('Q1');
            expect(result[0].id).toBe(1);
        });

        it('maps comments and sorts them by createdDate', async () => {
            const raw = makeRawItem();
            const comments = [
                { id: '2', text: 'second', author: { name: 'B', email: 'b@x.com', isExternal: false }, createdDate: '2024-02-01T00:00:00Z' },
                { id: '1', text: 'first', author: { name: 'A', email: 'a@x.com', isExternal: false }, createdDate: '2024-01-01T00:00:00Z' },
            ];
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [raw] })
                .mockResolvedValueOnce({ ok: true, data: comments });
            const result = await client.getQnAItems();
            expect(result[0].comments[0].text).toBe('first');
            expect(result[0].comments[1].text).toBe('second');
        });

        it('includes filesData from item when FilesData field is present', async () => {
            const filesData = JSON.stringify({ [QAItemType.Question]: [{ name: 'f.pdf', serverRelativeUrl: '/f.pdf' }] });
            const raw = makeRawItem({ [ColumnsQA.FilesData]: filesData });
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [raw] })
                .mockResolvedValueOnce({ ok: true, data: [] });
            const result = await client.getQnAItems();
            expect(result[0].filesData).toBeDefined();
        });
    });
});
