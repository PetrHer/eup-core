/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import FileApiClient from './FileApiClient';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        getChannelId: jest.fn().mockReturnValue('channel-123'),
        getTeamName: jest.fn().mockResolvedValue('TestTeam'),
        internalSiteName: 'internal-site',
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPost: jest.fn().mockResolvedValue({ ok: true }),
        spPatch: jest.fn().mockResolvedValue({ ok: true }),
        spDelete: jest.fn().mockResolvedValue({ ok: true }),
        spPut: jest.fn().mockResolvedValue({ ok: true }),
        spGetBinary: jest.fn().mockResolvedValue({ ok: true, data: new ArrayBuffer(0) }),
        spPostBinary: jest.fn().mockResolvedValue({ ok: true }),
        graphGet: jest.fn().mockResolvedValue({ ok: true, data: {} }),
        graphPost: jest.fn().mockResolvedValue({ ok: true }),
        updateItemInList: jest.fn().mockResolvedValue({ ok: true }),
        getPermissionBatchString: jest.fn().mockReturnValue(''),
        sp: {
            web: {
                getFileByServerRelativePath: jest.fn().mockReturnValue({
                    moveByPath: jest.fn().mockResolvedValue(undefined),
                }),
            },
        },
    } as unknown as SPApiClient;
}

describe('FileApiClient', () => {
    let client: FileApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new FileApiClient(mockSp);
    });

    describe('loadFiles', () => {
        it('calls spGet with structure ID in URL', async () => {
            await client.loadFiles(42);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('42')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.loadFiles(42);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.loadFiles(42);
            expect(result).toEqual([]);
        });

        it('returns mapped array from API data', async () => {
            const result = await client.loadFiles(42);
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('loadFilesByIds', () => {
        it('returns empty array when no IDs provided', async () => {
            const result = await client.loadFilesByIds([]);
            expect(result).toEqual([]);
        });

        it('calls spGet for given IDs', async () => {
            await client.loadFilesByIds([1, 2, 3]);
            expect(mockSp.spGet).toHaveBeenCalled();
        });

        it('includes IDs in URL filter', async () => {
            await client.loadFilesByIds([101]);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('101')
            );
        });
    });

    describe('publishFile', () => {
        it('calls spPost with unique ID in URL', async () => {
            await client.publishFile('unique-abc');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('unique-abc')
            );
        });

        it('returns true when spPost succeeds', async () => {
            const result = await client.publishFile('unique-abc');
            expect(result).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.publishFile('unique-abc');
            expect(result).toBe(false);
        });
    });

    describe('getFileTypes', () => {
        it('calls spGet with absolute URL', async () => {
            await client.getFileTypes();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getFileTypes();
            expect(result).toEqual([]);
        });
    });

    describe('undoFileCheckOut', () => {
        it('calls spPost with file unique ID in URL', async () => {
            await client.undoFileCheckOut('unique-id-abc');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('unique-id-abc')
            );
        });

        it('returns true when spPost succeeds', async () => {
            const result = await client.undoFileCheckOut('uid-1');
            expect(result).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.undoFileCheckOut('uid-1');
            expect(result).toBe(false);
        });
    });

    describe('updateFileByUniqueId', () => {
        it('calls spPatch with unique ID in URL', async () => {
            await client.updateFileByUniqueId('unique-id-abc', {});
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('unique-id-abc'),
                expect.any(Object)
            );
        });

        it('returns ok:true result when spPatch succeeds', async () => {
            const result = await client.updateFileByUniqueId('uid-1', {});
            expect(result).toMatchObject({ ok: true });
        });

        it('returns ok:false result when spPatch fails', async () => {
            (mockSp.spPatch as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.updateFileByUniqueId('uid-1', {});
            expect(result).toMatchObject({ ok: false });
        });
    });

    describe('getFileVersionHistory', () => {
        it('calls spGet with list item ID in URL', async () => {
            await client.getFileVersionHistory(55);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('55')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getFileVersionHistory(55);
            expect(result).toEqual([]);
        });
    });

    describe('getFileToDownload', () => {
        it('calls spGetBinary with relative URL in path', async () => {
            await client.getFileToDownload('myfile.pdf');
            expect(mockSp.spGetBinary).toHaveBeenCalledWith(
                expect.stringContaining('myfile.pdf')
            );
        });

        it('returns ArrayBuffer from spGetBinary', async () => {
            const buffer = new ArrayBuffer(8);
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: true, data: buffer });
            const result = await client.getFileToDownload('myfile.pdf');
            expect(result).toBe(buffer);
        });

        it('returns undefined when spGetBinary fails', async () => {
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.getFileToDownload('myfile.pdf');
            expect(result).toBeUndefined();
        });
    });

    // ─── moveFileToBin ────────────────────────────────────────────────────────

    describe('moveFileToBin', () => {
        it('calls moveFile with the internal trash path', async () => {
            const moveSpy = jest.spyOn(client, 'moveFile').mockResolvedValue(true);
            await client.moveFileToBin('/sites/team/file.pdf', 'file.pdf');
            expect(moveSpy).toHaveBeenCalledWith('/sites/team/file.pdf', 'file.pdf', 'KLIENT - EXT/KOŠ');
        });

        it('returns the result from moveFile', async () => {
            jest.spyOn(client, 'moveFile').mockResolvedValue(false);
            expect(await client.moveFileToBin('/f', 'f.pdf')).toBe(false);
        });
    });

    // ─── loadFilesByUniqueIds ─────────────────────────────────────────────────

    describe('loadFilesByUniqueIds', () => {
        it('returns empty array when no IDs provided', async () => {
            expect(await client.loadFilesByUniqueIds([])).toEqual([]);
        });

        it('calls spGet for each unique ID', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.loadFilesByUniqueIds(['uid-1', 'uid-2']);
            expect(mockSp.spGet).toHaveBeenCalledTimes(2);
        });

        it('filters out undefined results', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.loadFilesByUniqueIds(['uid-1']);
            expect(result).toEqual([]);
        });
    });

    // ─── loadFileByRelativeUrl ────────────────────────────────────────────────

    describe('loadFileByRelativeUrl', () => {
        it('calls spGet with server-relative URL', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.loadFileByRelativeUrl('/sites/team/doc.pdf');
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('doc.pdf'));
        });

        it('returns undefined when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            expect(await client.loadFileByRelativeUrl('/f')).toBeUndefined();
        });
    });

    // ─── loadAllFiles ─────────────────────────────────────────────────────────

    describe('loadAllFiles', () => {
        it('calls spGet with channel name in path', async () => {
            await client.loadAllFiles(false);
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('TestChannel'));
        });

        it('includes documentation filter when true', async () => {
            await client.loadAllFiles(true);
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('Documentation eq 1'));
        });

        it('returns empty array when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            expect(await client.loadAllFiles(false)).toEqual([]);
        });
    });

    // ─── getFileByUniqueId ────────────────────────────────────────────────────

    describe('getFileByUniqueId', () => {
        it('returns undefined when no data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            expect(await client.getFileByUniqueId('uid-1')).toBeUndefined();
        });

        it('does not call second spGet when DocumentTypeId is absent', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: { UniqueId: 'uid-1', FileRef: '/f', FileLeafRef: 'f.pdf' }
            });
            await client.getFileByUniqueId('uid-1');
            expect(mockSp.spGet).toHaveBeenCalledTimes(1);
        });

        it('fetches document type when DocumentTypeId is present', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { UniqueId: 'uid-1', FileRef: '/f', FileLeafRef: 'f.pdf', DocumentTypeId: 5 } })
                .mockResolvedValueOnce({ ok: true, data: { Id: 5, Title: 'Contract', Title_en: 'Contract' } });
            const result = await client.getFileByUniqueId('uid-1');
            expect(mockSp.spGet).toHaveBeenCalledTimes(2);
            expect(result?.type).toBeDefined();
        });
    });

    // ─── getFileVersion ───────────────────────────────────────────────────────

    describe('getFileVersion', () => {
        it('calls spGetBinary with relative URL and version ID', async () => {
            await client.getFileVersion('/sites/team/file.pdf', 512);
            expect(mockSp.spGetBinary).toHaveBeenCalledWith(
                expect.stringContaining('512')
            );
        });

        it('returns ArrayBuffer from spGetBinary', async () => {
            const buf = new ArrayBuffer(4);
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: true, data: buf });
            expect(await client.getFileVersion('/f', 1)).toBe(buf);
        });

        it('returns undefined when spGetBinary fails', async () => {
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: false });
            expect(await client.getFileVersion('/f', 1)).toBeUndefined();
        });
    });

    // ─── updateFileByUrl ──────────────────────────────────────────────────────

    describe('updateFileByUrl', () => {
        it('calls spPatch with file URL', async () => {
            await client.updateFileByUrl('/sites/team/file.pdf', { Status: 'active' });
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('file.pdf'),
                { Status: 'active' }
            );
        });

        it('returns the API result', async () => {
            const result = await client.updateFileByUrl('/f', {});
            expect(result).toMatchObject({ ok: true });
        });
    });

    // ─── moveFile ─────────────────────────────────────────────────────────────

    describe('moveFile', () => {
        it('calls moveByPath and returns true on success', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.moveFile('/sites/team/file.pdf', 'file.pdf', 'Archive');
            expect(mockSp.sp.web.getFileByServerRelativePath).toHaveBeenCalledWith('/sites/team/file.pdf');
            expect(result).toBe(true);
        });

        it('returns false when moveByPath throws', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            (mockSp.sp.web.getFileByServerRelativePath as jest.Mock).mockReturnValue({
                moveByPath: jest.fn().mockRejectedValue(new Error('Move failed')),
            });
            const result = await client.moveFile('/sites/team/file.pdf', 'file.pdf', 'Archive');
            expect(result).toBe(false);
        });
    });

    // ─── setFilePermissions ───────────────────────────────────────────────────

    describe('setFilePermissions', () => {
        it('calls resetAndBreakInheritance and assignFilePermissions', async () => {
            const resetSpy = jest.spyOn(client as any, 'resetAndBreakInheritance').mockResolvedValue(undefined);
            const assignSpy = jest.spyOn(client as any, 'assignFilePermissions').mockResolvedValue(undefined);
            const action = { writePermissions: 'UUO_BO', readPermissions: 'UUO_BO' } as any;

            await client.setFilePermissions(action, [], 'uid-1');

            expect(resetSpy).toHaveBeenCalledWith('uid-1');
            expect(assignSpy).toHaveBeenCalled();
        });

        it('handles missing permissions gracefully', async () => {
            jest.spyOn(client as any, 'resetAndBreakInheritance').mockResolvedValue(undefined);
            jest.spyOn(client as any, 'assignFilePermissions').mockResolvedValue(undefined);

            await expect(client.setFilePermissions({} as any, [], 'uid-1')).resolves.not.toThrow();
        });
    });

    // ─── uploadFiles ──────────────────────────────────────────────────────────

    describe('uploadFiles', () => {
        it('calls uploadFile for each file', async () => {
            const uploadSpy = jest.spyOn(client, 'uploadFile').mockResolvedValue({ fileName: 'a.pdf', success: true } as any);
            const mockType = { folderPath: '/Docs', folderPathAfter: '/DocsAfter', id: 1, uploadStatus: 'active' };
            const files = [
                { file: { name: 'a.pdf', rawFile: {} } as any, type: mockType } as any,
            ];

            await client.uploadFiles(files, false);

            expect(uploadSpy).toHaveBeenCalledTimes(1);
        });

        it('uses Temporary folder when file has no type', async () => {
            const uploadSpy = jest.spyOn(client, 'uploadFile').mockResolvedValue({ fileName: 'b.pdf', success: true } as any);
            const files = [{ file: { name: 'b.pdf' } as any, type: null } as any];

            await client.uploadFiles(files, false);

            expect(uploadSpy).toHaveBeenCalledWith('/Temporary', expect.anything(), undefined, undefined, undefined, undefined);
        });
    });

    // ─── uploadFile ───────────────────────────────────────────────────────────

    describe('uploadFile', () => {
        const makeMockFile = () => ({
            rawFile: { arrayBuffer: jest.fn().mockResolvedValue(new ArrayBuffer(4)) },
            name: 'report.pdf',
        } as any);

        it('returns success result with uploaded file', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: false })
                .mockResolvedValueOnce({ ok: true, data: { UniqueId: 'uid-1', FileRef: '/f', FileLeafRef: 'f.pdf' } })
                .mockResolvedValueOnce({ ok: true, data: null });
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: true, data: { UniqueId: 'uid-1' } });

            const result = await client.uploadFile('/FolderPath', makeMockFile());
            expect(result.success).toBe(true);
            expect(result.fileName).toBe('report.pdf');
        });

        it('returns failure when spPost returns no data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false, data: null });

            const result = await client.uploadFile('/FolderPath', makeMockFile());
            expect(result.success).toBe(false);
        });

        it('retries on 403 and gives up after max attempts', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false, code: 403, data: null });

            const result = await client.uploadFile('/FolderPath', makeMockFile());
            expect(mockSp.spPost).toHaveBeenCalledTimes(3);
            expect(result.success).toBe(false);
        });
    });

    // ─── reuploadFile ─────────────────────────────────────────────────────────

    describe('reuploadFile', () => {
        const makeMockFile = () => ({
            rawFile: { arrayBuffer: jest.fn().mockResolvedValue(new ArrayBuffer(4)) },
            name: 'report.pdf',
        } as any);

        it('returns success result when spPost returns data', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: true, data: { UniqueId: 'uid-2' } });
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { UniqueId: 'uid-2', FileRef: '/f', FileLeafRef: 'f.pdf' } })
                .mockResolvedValueOnce({ ok: true, data: null });

            const result = await client.reuploadFile('/folder', makeMockFile(), 'report.pdf');
            expect(result.success).toBe(true);
        });

        it('returns failure when spPost returns no data', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false, data: null });

            const result = await client.reuploadFile('/folder', makeMockFile(), 'report.pdf');
            expect(result.success).toBe(false);
        });
    });

    // ─── archiveFile ──────────────────────────────────────────────────────────

    describe('archiveFile', () => {
        const mockFile = {
            uniqueId: 'uid-1',
            name: 'doc.pdf',
            serverRelativeUrl: '/sites/team/Sdilene dokumenty/TestChannel/doc.pdf',
            type: { folderPathInternal: 'Legal', folderPathInternalAfter: 'LegalAfter', loc: { EN: 'Contract' } },
        } as any;

        it('updates existing archived file in place and returns true', async () => {
            const buf = new ArrayBuffer(8);
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ FileRef: '/sites/internal/doc.pdf' }] });
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: true, data: buf });

            const result = await client.archiveFile(mockFile);

            expect(mockSp.spGetBinary).toHaveBeenCalled();
            expect(mockSp.spPut).toHaveBeenCalledWith(expect.any(String), buf);
            expect(mockSp.spPatch).toHaveBeenCalled();
            expect(result).toBe(true);
        });

        it('returns false when spGetBinary fails on existing archived file', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ FileRef: '/sites/internal/doc.pdf' }] });
            (mockSp.spGetBinary as jest.Mock).mockResolvedValue({ ok: false });

            const result = await client.archiveFile(mockFile);

            expect(mockSp.spPut).not.toHaveBeenCalled();
            expect(result).toBe(false);
        });

        it('copies file to archive when no existing archived file', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [] })
                .mockResolvedValue({ ok: false });

            const result = await client.archiveFile(mockFile);

            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('CopyFileByPath'),
                expect.any(Object)
            );
            expect(result).toBe(true);
        });
    });

    // ─── getRequiredFiles ─────────────────────────────────────────────────────

    describe('getRequiredFiles', () => {
        it('calls spGet with folder ID in filter', async () => {
            await client.getRequiredFiles(7);
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('7'));
        });

        it('returns empty array when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            expect(await client.getRequiredFiles(7)).toEqual([]);
        });
    });

    // ─── getAllRequiredFiles ───────────────────────────────────────────────────

    describe('getAllRequiredFiles', () => {
        it('returns a Map grouped by folder number', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { ContentTypeId: '0x010123', FileDirRef: '/sites/team/Lists/RequiredDocuments/TestChannel/3' },
                    { ContentTypeId: '0x010123', FileDirRef: '/sites/team/Lists/RequiredDocuments/TestChannel/3' },
                ],
            });

            const result = await client.getAllRequiredFiles();
            expect(result).toBeInstanceOf(Map);
            expect(result.has(3)).toBe(true);
        });

        it('skips items where last path segment is not a number', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ ContentTypeId: '0x010123', FileDirRef: '/lists/channel/notanumber' }],
            });
            const result = await client.getAllRequiredFiles();
            expect(result.size).toBe(0);
        });

        it('returns empty map when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getAllRequiredFiles();
            expect(result.size).toBe(0);
        });
    });

    // ─── getFolders ───────────────────────────────────────────────────────────

    describe('getFolders', () => {
        it('uses Czech folder name by default', async () => {
            await client.getFolders();
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('Sdilene dokumenty'));
        });

        it('uses English folder name for appgestio URL', async () => {
            (mockSp.getAbsoluteUrl as jest.Mock).mockReturnValue('https://appgestio.sharepoint.com/sites/team');
            await client.getFolders();
            expect(mockSp.spGet).toHaveBeenCalledWith(expect.stringContaining('Shared documents'));
        });

        it('returns empty array when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            expect(await client.getFolders()).toEqual([]);
        });
    });

    // ─── sendFiletoExtraction ─────────────────────────────────────────────────

    describe('sendFiletoExtraction', () => {
        const mockFile = {
            serverRelativeUrl: '/sites/team/Sdilene dokumenty/TestChannel/doc.pdf',
            id: 99,
        } as any;

        it('calls spPost with Extrakce list URL and body', async () => {
            await client.sendFiletoExtraction(mockFile, 'Invoice' as any);
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('Extrakce'),
                expect.objectContaining({ FileId: 99 })
            );
        });

        it('returns true when spPost succeeds', async () => {
            expect(await client.sendFiletoExtraction(mockFile, 'Invoice' as any)).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false });
            expect(await client.sendFiletoExtraction(mockFile, 'Invoice' as any)).toBe(false);
        });
    });
});
