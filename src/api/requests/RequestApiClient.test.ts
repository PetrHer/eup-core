/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import RequestApiClient from './RequestApiClient';
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
        createListItemUsingPath: jest.fn().mockResolvedValue({ ok: true }),
        updateListItem: jest.fn().mockResolvedValue(true),
    } as unknown as SPApiClient;
}

function createMockFiles() {
    return {
        loadFiles: jest.fn().mockResolvedValue([]),
        uploadFile: jest.fn().mockResolvedValue({ success: true, file: null }),
        uploadFiles: jest.fn().mockResolvedValue([]),
    } as unknown as FileApiClient;
}

describe('RequestApiClient', () => {
    let client: RequestApiClient;
    let mockSp: ReturnType<typeof createMockSp>;
    let mockFiles: ReturnType<typeof createMockFiles>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        mockFiles = createMockFiles();
        client = new RequestApiClient(mockFiles, mockSp);
    });

    describe('getRequests', () => {
        it('calls spGet with channel name in URL', async () => {
            await client.getRequests();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.getRequests();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getRequests();
            expect(result).toEqual([]);
        });

        it('returns mapped array from API data', async () => {
            const result = await client.getRequests();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('getRequestsOptions', () => {
        it('calls spGet with absolute URL', async () => {
            await client.getRequestsOptions();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getRequestsOptions();
            expect(result).toEqual([]);
        });

        it('returns mapped array from API data', async () => {
            const result = await client.getRequestsOptions();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('updateRequest', () => {
        it('calls updateListItem with item ID and body', async () => {
            await client.updateRequest({ Status: 'done' }, 11);
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                11,
                expect.any(Object),
                expect.any(String)
            );
        });

        it('returns result true when updateListItem returns true', async () => {
            const { result } = await client.updateRequest({}, 11);
            expect(result).toBe(true);
        });

        it('returns result false when updateListItem returns false', async () => {
            (mockSp.updateListItem as jest.Mock).mockResolvedValue(false);
            const { result } = await client.updateRequest({}, 11);
            expect(result).toBe(false);
        });

        it('returns files array in result', async () => {
            const { files } = await client.updateRequest({}, 11);
            expect(Array.isArray(files)).toBe(true);
        });

        it('does not call uploadFiles when no files provided', async () => {
            await client.updateRequest({}, 11);
            expect(mockFiles.uploadFiles).not.toHaveBeenCalled();
        });
    });

    // ─── createRequest ────────────────────────────────────────────────────────

    describe('createRequest', () => {
        const fileType: any = { folderPath: '/sites/team/docs', id: 'ft1', uploadStatus: 'pending' };
        const baseData: any = {
            templateRef: undefined,
            templateName: 'My Template',
            formData: { field: 'value' },
            typeId: 5,
            exportData: '',
        };

        it('calls createListItemUsingPath after uploading files', async () => {
            const files: any[] = [{ name: 'file.pdf' }];
            await client.createRequest(baseData, files, fileType, false, 'user@test.com');
            expect(mockFiles.uploadFile).toHaveBeenCalled();
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });

        it('returns the result from createListItemUsingPath and empty failedUploads on success', async () => {
            (mockFiles.uploadFile as jest.Mock).mockResolvedValue({ success: true, file: { serverRelativeUrl: '/f.pdf', uniqueId: 'uid-1' } });
            const files: any[] = [{ name: 'file.pdf' }];
            const { result, failedUploads } = await client.createRequest(baseData, files, fileType, false, 'user@test.com');
            expect(result).toEqual({ ok: true });
            expect(failedUploads).toEqual([]);
        });

        it('adds failed file name to failedUploads when upload fails', async () => {
            (mockFiles.uploadFile as jest.Mock).mockResolvedValue({ success: false, file: null });
            const files: any[] = [{ name: 'bad.pdf' }];
            const { failedUploads } = await client.createRequest(baseData, files, fileType, false, 'user@test.com');
            expect(failedUploads).toContain('bad.pdf');
        });

        it('uses folderPathAfter when secondStage is true and fileType has folderPathAfter', async () => {
            const ftWithAfter: any = { ...fileType, folderPathAfter: '/sites/team/after' };
            const files: any[] = [{ name: 'doc.pdf' }];
            await client.createRequest(baseData, files, ftWithAfter, true, 'user@test.com');
            expect(mockFiles.uploadFile).toHaveBeenCalledWith(
                '/sites/team/after',
                expect.any(Object),
                'user@test.com',
                undefined,
                'ft1',
                'pending'
            );
        });

        it('uses folderPath when secondStage is false', async () => {
            const files: any[] = [{ name: 'doc.pdf' }];
            await client.createRequest(baseData, files, fileType, false, 'user@test.com');
            expect(mockFiles.uploadFile).toHaveBeenCalledWith(
                '/sites/team/docs',
                expect.any(Object),
                'user@test.com',
                undefined,
                'ft1',
                'pending'
            );
        });

        it('sets fileTemplate from serverRelativeUrl for first file when templateRef is set', async () => {
            (mockFiles.uploadFile as jest.Mock).mockResolvedValue({ success: true, file: { serverRelativeUrl: '/tmpl.pdf', uniqueId: 'uid-tmpl' } });
            const dataWithRef: any = { ...baseData, templateRef: 'ref123', templateName: 'TemplateName' };
            const files: any[] = [{ name: 'orig.pdf' }];
            await client.createRequest(dataWithRef, files, fileType, false, 'user@test.com');
            const [body] = (mockSp.createListItemUsingPath as jest.Mock).mock.calls[0];
            const templateField = body.formValues.find((f: any) => f.FieldName === 'FileTemplate');
            expect(templateField?.FieldValue).toBe('/tmpl.pdf');
        });

        it('pushes uniqueId into FileIds for non-template files', async () => {
            (mockFiles.uploadFile as jest.Mock).mockResolvedValue({ success: true, file: { serverRelativeUrl: '/f.pdf', uniqueId: 'uid-99' } });
            const files: any[] = [{ name: 'a.pdf' }];
            await client.createRequest(baseData, files, fileType, false, 'user@test.com');
            const [body] = (mockSp.createListItemUsingPath as jest.Mock).mock.calls[0];
            const fileIdsField = body.formValues.find((f: any) => f.FieldName === 'FileIds');
            expect(fileIdsField?.FieldValue).toContain('uid-99');
        });

        it('works with no files', async () => {
            const { result, failedUploads } = await client.createRequest(baseData, [], fileType, false, 'user@test.com');
            expect(mockFiles.uploadFile).not.toHaveBeenCalled();
            expect(result).toEqual({ ok: true });
            expect(failedUploads).toEqual([]);
        });
    });

    // ─── updateRequest with files ─────────────────────────────────────────────

    describe('updateRequest with files', () => {
        it('calls uploadFiles and adds uniqueIds to body FileIds', async () => {
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([
                { file: { uniqueId: 'uid-a', serverRelativeUrl: '/a.pdf' } },
                { file: { uniqueId: 'uid-b', serverRelativeUrl: '/b.pdf' } },
            ]);
            const body: any = {};
            await client.updateRequest(body, 1, 'user', false, [{ name: 'a.pdf' }] as any[]);
            expect(mockFiles.uploadFiles).toHaveBeenCalled();
            expect(body.FileIds).toContain('uid-a');
            expect(body.FileIds).toContain('uid-b');
        });

        it('appends to existing FileIds when body already has them', async () => {
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([
                { file: { uniqueId: 'uid-new', serverRelativeUrl: '/new.pdf' } },
            ]);
            const body: any = { FileIds: 'uid-old' };
            await client.updateRequest(body, 1, 'user', false, [{ name: 'f.pdf' }] as any[]);
            expect(body.FileIds).toContain('uid-old');
            expect(body.FileIds).toContain('uid-new');
        });

        it('pushes uploaded files into result files array', async () => {
            const uploadedFile = { uniqueId: 'uid-z', serverRelativeUrl: '/z.pdf' };
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([{ file: uploadedFile }]);
            const { files } = await client.updateRequest({}, 1, 'user', false, [{ name: 'z.pdf' }] as any[]);
            expect(files).toContain(uploadedFile);
        });

        it('ignores entries with no file in upload result', async () => {
            (mockFiles.uploadFiles as jest.Mock).mockResolvedValue([{ file: null }]);
            const { files } = await client.updateRequest({}, 1, 'user', false, [{ name: 'x.pdf' }] as any[]);
            expect(files).toHaveLength(0);
        });
    });
});
