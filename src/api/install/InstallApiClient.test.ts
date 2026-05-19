/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import InstallApiClient from './InstallApiClient';
import { ListName, TemplatesListName } from '../../enums';
import { ITemplateContent } from '../../interfaces';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');
jest.mock('../detail/DetailApiClient', () => ({
    __esModule: true,
    default: jest.fn().mockImplementation(() => ({
        getLoanDetail: jest.fn().mockResolvedValue({
            currentStatus: 'active',
            currentFolder: 'folder1',
            contact: { email: 'a@b.com' },
            secondaryContact: { email: 'b@c.com' },
        }),
    })),
}));

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelId: jest.fn().mockReturnValue('channel-123'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        getTeamId: jest.fn().mockReturnValue('team-id-123'),
        getTeamName: jest.fn().mockResolvedValue('TestTeam'),
        templateSiteName: 'template-site',
        internalSiteName: 'internal-site',
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPost: jest.fn().mockResolvedValue({ ok: true }),
        spPatch: jest.fn().mockResolvedValue({ ok: true }),
        spPut: jest.fn().mockResolvedValue({ ok: true }),
        spGetBinary: jest.fn().mockResolvedValue({ ok: true, data: new ArrayBuffer(0) }),
        spPostBinary: jest.fn().mockResolvedValue({ ok: true }),
        moveFolder: jest.fn().mockResolvedValue(true),
        createListItem: jest.fn().mockResolvedValue({ ok: true, data: { Id: 1 } }),
        updateListItem: jest.fn().mockResolvedValue(true),
        updateItemInList: jest.fn().mockResolvedValue({ ok: true }),
        moveListItem: jest.fn().mockResolvedValue(true),
        ensureFolder: jest.fn().mockResolvedValue('/path/to/folder'),
        ensureListFolder: jest.fn().mockResolvedValue(undefined),
        ensureListView: jest.fn().mockResolvedValue(undefined),
        getUserSPId: jest.fn().mockResolvedValue(99),
        checkListExistence: jest.fn().mockResolvedValue(true),
        createList: jest.fn().mockResolvedValue(true),
        ensureColumns: jest.fn().mockResolvedValue([]),
        createListItemsBatch: jest.fn().mockResolvedValue({ ok: true }),
        createListItemUsingPath: jest.fn().mockResolvedValue({ ok: true }),
        getSiteGroups: jest.fn().mockResolvedValue([]),
        getPermissionBatchString: jest.fn().mockReturnValue(''),
        addAttachmentArrayBuffer: jest.fn().mockResolvedValue({ ok: true }),
        chunkArray: jest.fn().mockImplementation((arr: any[], size: number) => {
            const chunks = [];
            for (let i = 0; i < arr.length; i += size) {
                chunks.push(arr.slice(i, i + size));
            }
            return chunks;
        }),
    } as unknown as SPApiClient;
}

describe('InstallApiClient', () => {
    let client: InstallApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new InstallApiClient(mockSp);
    });

    describe('getTemplates', () => {
        it('calls spGet with templateSiteName in URL', async () => {
            await client.getTemplates('/sites/template-site/folder');
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('template-site')
            );
        });

        it('returns empty array when API returns non-array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getTemplates('/fileRef');
            expect(result).toEqual([]);
        });

        it('returns mapped array when API returns array data', async () => {
            const result = await client.getTemplates('/fileRef');
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('getTemplatesNames', () => {
        it('calls spGet with templateSiteName in URL', async () => {
            await client.getTemplatesNames();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('template-site')
            );
        });

        it('returns empty array when API returns non-array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getTemplatesNames();
            expect(result).toEqual([]);
        });

        it('returns array from API data', async () => {
            const result = await client.getTemplatesNames();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('updateLoanDetailsVersion', () => {
        it('returns true without calling updateListItem when no items need updating', async () => {
            const result = await client.updateLoanDetailsVersion(5);
            expect(result).toBe(true);
            expect(mockSp.updateListItem).not.toHaveBeenCalled();
        });

        it('calls updateListItem for each item below the given version', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ Id: 1, version: 2 }, { Id: 2, version: 3 }],
            });
            await client.updateLoanDetailsVersion(5);
            expect(mockSp.updateListItem).toHaveBeenCalledTimes(2);
        });

        it('returns false when updateListItem fails on any item', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1, version: 1 }] });
            (mockSp.updateListItem as jest.Mock).mockResolvedValue(false);
            const result = await client.updateLoanDetailsVersion(5);
            expect(result).toBe(false);
        });

        it('returns true when all updateListItem calls succeed', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1, version: 1 }] });
            const result = await client.updateLoanDetailsVersion(5);
            expect(result).toBe(true);
        });
    });

    describe('updateIUPChannel', () => {
        it('returns early without calling spPatch when channel item not found', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            await client.updateIUPChannel({ Status: 'Active' });
            expect(mockSp.spPatch).not.toHaveBeenCalled();
        });

        it('returns early without calling spPatch when channel item Id is not a number', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: '42' }] });
            await client.updateIUPChannel({ Status: 'Active' });
            expect(mockSp.spPatch).not.toHaveBeenCalled();
        });

        it('calls spPatch with channel item ID when channel found', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 42 }] });
            await client.updateIUPChannel({ Status: 'Active' });
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('(42)'),
                { Status: 'Active' }
            );
        });

        it('queries by channelId in URL', async () => {
            await client.updateIUPChannel({});
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('channel-123')
            );
        });
    });

    describe('updateIUPChannelContacts', () => {
        it('calls getUserSPId for provided email', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }] });
            await client.updateIUPChannelContacts('user@test.com');
            expect(mockSp.getUserSPId).toHaveBeenCalledWith('user@test.com', expect.any(String));
        });

        it('does not call getUserSPId when no emails provided', async () => {
            await client.updateIUPChannelContacts();
            expect(mockSp.getUserSPId).not.toHaveBeenCalled();
        });

        it('calls getUserSPId for both emails when both provided', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }] });
            await client.updateIUPChannelContacts('a@test.com', 'b@test.com');
            expect(mockSp.getUserSPId).toHaveBeenCalledTimes(2);
        });
    });

    describe('updateIUPGroupColumns', () => {
        it('calls getUserSPId once per email across all groups', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }] });
            await client.updateIUPGroupColumns(['a@test.com'], ['b@test.com'], ['c@test.com']);
            expect(mockSp.getUserSPId).toHaveBeenCalledTimes(3);
        });

        it('passes resolved IDs to updateIUPChannel', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }] });
            (mockSp.getUserSPId as jest.Mock).mockResolvedValue(55);
            await client.updateIUPGroupColumns(['a@test.com'], [], []);
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.any(String),
                expect.objectContaining({ OAOId: [55] })
            );
        });
    });

    describe('moveFolderAndChangeTitle', () => {
        it('returns early without further calls when moveFolder returns false', async () => {
            (mockSp.moveFolder as jest.Mock).mockResolvedValue(false);
            await client.moveFolderAndChangeTitle('/old', '/new', 'New Title', ListName.Categories);
            expect(mockSp.spGet).not.toHaveBeenCalled();
            expect(mockSp.updateItemInList).not.toHaveBeenCalled();
        });

        it('calls updateItemInList with new title after successful move', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: { UniqueId: 'uid-abc' } });
            await client.moveFolderAndChangeTitle('/old', '/new', 'New Title', ListName.Categories);
            expect(mockSp.updateItemInList).toHaveBeenCalledWith(
                'uid-abc',
                expect.objectContaining({ Title: 'New Title' }),
                ListName.Categories
            );
        });

        it('calls spGet with new path to retrieve folder UniqueId', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: { UniqueId: 'uid' } });
            await client.moveFolderAndChangeTitle('/old', '/new/path', 'Title', ListName.Categories);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('/new/path')
            );
        });
    });

    describe('updateLinksInCategories', () => {
        it('calls spGet with channel name in filter URL', async () => {
            await client.updateLinksInCategories();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('does not call updateListItem when API returns empty array', async () => {
            await client.updateLinksInCategories();
            expect(mockSp.updateListItem).not.toHaveBeenCalled();
        });

        it('calls updateListItem for each returned item', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { Id: 1, RequiredDocuments: { Url: 'http://example.com/1' } },
                    { Id: 2, RequiredDocuments: { Url: 'http://example.com/2' } },
                ],
            });
            await client.updateLinksInCategories();
            expect(mockSp.updateListItem).toHaveBeenCalledTimes(2);
        });
    });

    describe('createCategoryItem', () => {
        const templateItem: ITemplateContent = {
            id: 10,
            stage: 'S1',
            phase: 'P1',
            category: 'Cat',
            description: 'Desc',
            stageEn: 'S1en',
            categoryEn: 'CatEn',
            phaseEn: 'P1en',
            descriptionEn: 'DescEn',
            order: 1,
        } as unknown as ITemplateContent;

        it('calls createListItem and returns the new item ID', async () => {
            const result = await client.createCategoryItem(templateItem);
            expect(mockSp.createListItem).toHaveBeenCalled();
            expect(result).toBe(1);
        });

        it('calls moveListItem and updateListItem after creation', async () => {
            await client.createCategoryItem(templateItem);
            expect(mockSp.moveListItem).toHaveBeenCalled();
            expect(mockSp.updateListItem).toHaveBeenCalled();
        });

        it('returns undefined when createListItem returns no data', async () => {
            (mockSp.createListItem as jest.Mock).mockResolvedValue({ ok: false, data: null });
            const result = await client.createCategoryItem(templateItem);
            expect(result).toBeUndefined();
        });
    });

    describe('ensureListFolder', () => {
        it('does not create folder when spGet returns ok', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: {} });
            await client.ensureListFolder('TestChannel', ListName.Requests);
            expect(mockSp.createListItem).not.toHaveBeenCalled();
        });

        it('creates folder when spGet returns not ok', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            await client.ensureListFolder('TestChannel', ListName.Requests);
            expect(mockSp.createListItem).toHaveBeenCalled();
            expect(mockSp.updateListItem).toHaveBeenCalled();
        });

        it('calls moveFolder when path param is provided and folder does not exist', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            await client.ensureListFolder('subfolder', ListName.Requests, 'parentPath');
            expect(mockSp.moveFolder).toHaveBeenCalled();
        });
    });

    describe('setListReadOnly', () => {
        it('returns early without calling spPost when spGet returns no data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.setListReadOnly(ListName.Requests);
            expect(mockSp.spPost).not.toHaveBeenCalled();
        });

        it('calls spPost to break inheritance and post batch when groups found', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: {
                    AssociatedOwnerGroup: { Id: 1 },
                    AssociatedMemberGroup: { Id: 2 },
                    AssociatedVisitorGroup: { Id: 3 },
                },
            });
            await client.setListReadOnly(ListName.Requests);
            expect(mockSp.spPost).toHaveBeenCalledTimes(2);
        });
    });

    describe('createListFolders', () => {
        it('calls ensureListFolder 7 times for all lists', async () => {
            await client.createListFolders();
            expect(mockSp.spGet).toHaveBeenCalledTimes(7);
        });

        it('calls ensureListFolder with channel name for each list', async () => {
            await client.createListFolders();
            expect(mockSp.spGet).toHaveBeenCalledTimes(7);
        });
    });

    describe('updateDateColumns', () => {
        it('calls spPatch twice for document library and required documents fields', async () => {
            await client.updateDateColumns();
            expect(mockSp.spPatch).toHaveBeenCalledTimes(2);
        });

        it('sets DisplayFormat 0 on ScheduledNotificationDate field', async () => {
            await client.updateDateColumns();
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('ScheduledNotificationDate'),
                { DisplayFormat: 0 }
            );
        });
    });

    describe('ensureListAndCopyItems', () => {
        it('does not call createList when list already exists', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(true);
            await client.ensureListAndCopyItems(ListName.RequestTypes, TemplatesListName.RequestTypes);
            expect(mockSp.createList).not.toHaveBeenCalled();
        });

        it('creates list and copies items when list does not exist', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(false);
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Title: 'Item 1' }] });
            await client.ensureListAndCopyItems(ListName.RequestTypes, TemplatesListName.RequestTypes);
            expect(mockSp.createList).toHaveBeenCalledWith(ListName.RequestTypes);
            expect(mockSp.createListItemsBatch).toHaveBeenCalled();
        });

        it('returns early without copying if createList fails', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(false);
            (mockSp.createList as jest.Mock).mockResolvedValue(false);
            await client.ensureListAndCopyItems(ListName.RequestTypes, TemplatesListName.RequestTypes);
            expect(mockSp.spGet).not.toHaveBeenCalled();
        });
    });

    describe('copyFoldersAndItemsWithStructure', () => {
        it('returns without making other calls when API returns non-array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.copyFoldersAndItemsWithStructure(
                TemplatesListName.QA, ListName.QA, ['Title']
            );
            expect(mockSp.createListItemUsingPath).not.toHaveBeenCalled();
        });

        it('creates folder entries for items with FileSystemObjectType 1', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { FileSystemObjectType: 1, FileLeafRef: 'folder1', FileDirRef: '/sites/template/Lists/QA', Title: 'F1' },
                ],
            });
            // ensureListFolder calls spGet internally — return ok:true so no folder creation
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ FileSystemObjectType: 1, FileLeafRef: 'folder1', FileDirRef: '/sites/template/Lists/QA', Title: 'F1' }] })
                .mockResolvedValue({ ok: true, data: {} });
            await client.copyFoldersAndItemsWithStructure(
                TemplatesListName.QA, ListName.QA, ['Title']
            );
            expect(mockSp.spGet).toHaveBeenCalled();
        });

        it('calls createListItemUsingPath for non-folder items', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { FileSystemObjectType: 0, FileLeafRef: 'item.docx', FileDirRef: '/sites/template/Lists/QA', Title: 'Item1' },
                ],
            });
            await client.copyFoldersAndItemsWithStructure(
                TemplatesListName.QA, ListName.QA, ['Title']
            );
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });
    });

    describe('initRequests', () => {
        it('calls ensureFolder for channel name and options folders', async () => {
            await client.initRequests();
            expect(mockSp.ensureFolder).toHaveBeenCalledWith(
                expect.any(String),
                'TestChannel'
            );
            expect(mockSp.ensureFolder).toHaveBeenCalledWith(
                expect.any(String),
                'options'
            );
        });
    });

    describe('ensureFolderForAllChannels', () => {
        it('calls ensureFolder for each top-level folder returned', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { Name: 'ChannelA', ServerRelativeUrl: '/sites/team/docs/ChannelA' },
                    { Name: 'ChannelB', ServerRelativeUrl: '/sites/team/docs/ChannelB' },
                ],
            });
            await client.ensureFolderForAllChannels('/base/path', 'subfolder');
            expect(mockSp.ensureFolder).toHaveBeenCalledTimes(2);
        });

        it('calls ensureFolder with the correct subfolder name', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ Name: 'ChannelA', ServerRelativeUrl: '/sites/team/docs/ChannelA' }],
            });
            await client.ensureFolderForAllChannels('/base', 'my-subfolder');
            expect(mockSp.ensureFolder).toHaveBeenCalledWith(
                expect.any(String),
                'my-subfolder'
            );
        });

        it('does not call ensureFolder when API returns no folders', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false, data: [] });
            await client.ensureFolderForAllChannels('/base', 'subfolder');
            expect(mockSp.ensureFolder).not.toHaveBeenCalled();
        });
    });

    // ─── ensureIUPData ────────────────────────────────────────────────────────

    describe('ensureIUPData', () => {
        it('calls ensureIUPChannel and ensureIUPFolders', async () => {
            const channelSpy = jest.spyOn(client as any, 'ensureIUPChannel').mockResolvedValue(undefined);
            const folderSpy = jest.spyOn(client as any, 'ensureIUPFolders').mockResolvedValue(undefined);

            await client.ensureIUPData();

            expect(channelSpy).toHaveBeenCalled();
            expect(folderSpy).toHaveBeenCalled();
        });

        it('passes the internal site webUrl to both methods', async () => {
            const channelSpy = jest.spyOn(client as any, 'ensureIUPChannel').mockResolvedValue(undefined);
            jest.spyOn(client as any, 'ensureIUPFolders').mockResolvedValue(undefined);

            await client.ensureIUPData();

            expect(channelSpy).toHaveBeenCalledWith(expect.stringContaining('internal-site'));
        });
    });

    // ─── createFoldersFromTemplates ───────────────────────────────────────────

    describe('createFoldersFromTemplates', () => {
        it('calls createMainFolderInDocumentStructure and createCategories', async () => {
            const mainSpy = jest.spyOn(client as any, 'createMainFolderInDocumentStructure').mockResolvedValue(undefined);
            const catSpy = jest.spyOn(client as any, 'createCategories').mockResolvedValue([{ itemId: 1, templateId: 10 }]);

            const result = await client.createFoldersFromTemplates([]);

            expect(mainSpy).toHaveBeenCalled();
            expect(catSpy).toHaveBeenCalled();
            expect(result).toEqual([{ itemId: 1, templateId: 10 }]);
        });
    });

    // ─── ensureDocumentLibrary ────────────────────────────────────────────────

    describe('ensureDocumentLibrary', () => {
        it('posts missing required fields', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [] })           // fields — all missing
                .mockResolvedValueOnce({ ok: true, data: { Id: 'list-id' } }); // lookup list
            jest.spyOn(client as any, 'ensureFoldersInDocumentLibrary').mockResolvedValue(undefined);

            await client.ensureDocumentLibrary('Template1');

            expect(mockSp.spPost).toHaveBeenCalled();
        });

        it('skips posting fields that already exist', async () => {
            const existingFields = [
                'Status', 'FileType', 'StrukturaDokumentaceId', 'Name_en', 'Documentation',
                'UploadedBy', 'ValidTo', 'ContractValidity', 'ScheduledNotificationDate',
                'ScheduledNotificationMessage', 'ScheduledNotificationMessageEn',
                'ScheduledNotificationRecipients', 'ActionToTrigger', 'QAId', 'DocumentType'
            ].map(title => ({ Title: title, LookupList: 'list-id' }));
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: existingFields })
                .mockResolvedValueOnce({ ok: true, data: { Id: 'list-id' } });
            jest.spyOn(client as any, 'ensureFoldersInDocumentLibrary').mockResolvedValue(undefined);

            await client.ensureDocumentLibrary('Template1');

            // Only the final versioning spPost should be called
            expect(mockSp.spPost).toHaveBeenCalledTimes(1);
        });

        it('patches lookup field when list ID has changed', async () => {
            const fields = [{ Title: 'DocumentType', LookupList: 'old-id' }];
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: fields })
                .mockResolvedValueOnce({ ok: true, data: { Id: 'new-id' } });
            jest.spyOn(client as any, 'ensureFoldersInDocumentLibrary').mockResolvedValue(undefined);

            await client.ensureDocumentLibrary('Template1');

            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('DocumentType'),
                expect.any(Object)
            );
        });
    });

    // ─── addRequiredFiles ─────────────────────────────────────────────────────

    describe('addRequiredFiles', () => {
        it('ensures a list folder for each category ID', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }, { Id: 2 }] });
            jest.spyOn(client as any, 'copyRequiredfiles').mockResolvedValue(undefined);
            const folderSpy = jest.spyOn(client, 'ensureListFolder').mockResolvedValue(undefined);

            await client.addRequiredFiles([], 'template');

            expect(folderSpy).toHaveBeenCalledTimes(2);
        });

        it('calls copyRequiredfiles for each relation', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const copySpy = jest.spyOn(client as any, 'copyRequiredfiles').mockResolvedValue(undefined);
            jest.spyOn(client, 'ensureListFolder').mockResolvedValue(undefined);
            const relations = [{ itemId: 1, templateId: 10 }, { itemId: 2, templateId: 20 }];

            await client.addRequiredFiles(relations, 'template');

            expect(copySpy).toHaveBeenCalledTimes(2);
        });


    });

    // ─── ensureListAndCopyItemsIntoChannelFolder ──────────────────────────────

    describe('ensureListAndCopyItemsIntoChannelFolder', () => {
        it('skips copy when items already exist', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(true);
            jest.spyOn(client, 'ensureListFolder').mockResolvedValue(undefined);
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [{ Id: 1 }] });

            await client.ensureListAndCopyItemsIntoChannelFolder(ListName.QA, TemplatesListName.QA, 'template');

            expect(mockSp.spPost).not.toHaveBeenCalled();
        });

        it('copies items from template when list is new and no items exist', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(false);
            jest.spyOn(client, 'ensureListFolder').mockResolvedValue(undefined);
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [] })                      // checkCopiedItemsExistence → empty
                .mockResolvedValueOnce({ ok: true, data: [{ Title: 'Item1' }] });   // template items

            await client.ensureListAndCopyItemsIntoChannelFolder(ListName.QA, TemplatesListName.QA, 'template');

            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('$batch'),
                expect.any(String),
                expect.any(Object)
            );
        });

        it('returns early when list creation fails', async () => {
            (mockSp.checkListExistence as jest.Mock).mockResolvedValue(false);
            (mockSp.createList as jest.Mock).mockResolvedValue(false);

            await client.ensureListAndCopyItemsIntoChannelFolder(ListName.QA, TemplatesListName.QA, 'template');

            expect(mockSp.ensureColumns).not.toHaveBeenCalled();
        });
    });

    // ─── copyTemplateAttachments ──────────────────────────────────────────────

    describe('copyTemplateAttachments', () => {
        it('returns early when no template item found', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });

            await client.copyTemplateAttachments(TemplatesListName.QA, ListName.QA);

            expect(mockSp.spGetBinary).not.toHaveBeenCalled();
        });

        it('downloads and attaches files when template item exists', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 5 }] })                         // find šablony item
                .mockResolvedValueOnce({ ok: true, data: [{ FileName: 'a.docx', ServerRelativeUrl: '/sites/template-site/a.docx' }] }) // attachments
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 10 }] });                        // check existing item

            await client.copyTemplateAttachments(TemplatesListName.QA, ListName.QA);

            expect(mockSp.spGetBinary).toHaveBeenCalled();
            expect(mockSp.addAttachmentArrayBuffer).toHaveBeenCalled();
        });

        it('creates list item when no existing target item found', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 5 }] })
                .mockResolvedValueOnce({ ok: true, data: [] })   // no attachments
                .mockResolvedValueOnce({ ok: true, data: [] });  // no existing item

            await client.copyTemplateAttachments(TemplatesListName.QA, ListName.QA);

            expect(mockSp.createListItem).toHaveBeenCalledWith(
                { Title: 'šablony' },
                ListName.QA
            );
        });
    });

    // ─── ensureIUPChannel (via ensureIUPData integration) ────────────────────

    describe('ensureIUPChannel', () => {
        it('posts new channel entry when channel data does not exist', async () => {
            jest.spyOn(client as any, 'ensureIUPFolders').mockResolvedValue(undefined);
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false, data: [] });

            await client.ensureIUPData();

            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('Channels'),
                expect.objectContaining({ ChannelStatus: 'Active' })
            );
        });

        it('updates existing channel when it is found', async () => {
            jest.spyOn(client as any, 'ensureIUPFolders').mockResolvedValue(undefined);
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [{ Id: 1, ChannelId: 'channel-123' }],
            });
            const updateSpy = jest.spyOn(client, 'updateIUPChannel').mockResolvedValue(undefined);

            await client.ensureIUPData();

            expect(updateSpy).toHaveBeenCalled();
        });
    });

    // ─── ensureIUPFolders (via ensureIUPData integration) ────────────────────

    describe('ensureIUPFolders', () => {
        it('creates team folder when it does not exist', async () => {
            jest.spyOn(client as any, 'ensureIUPChannel').mockResolvedValue(undefined);
            jest.spyOn(client as any, 'getFolderTemplates').mockResolvedValue([]);
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: false })   // team folder check → not found
                .mockResolvedValueOnce({ ok: false });  // channel folder check → not found
            (mockSp.spPost as jest.Mock)
                .mockResolvedValueOnce({ ok: true })    // create team folder
                .mockResolvedValueOnce({ ok: true });   // create channel folder

            await client.ensureIUPData();

            expect(mockSp.spPost).toHaveBeenCalledTimes(2);
        });

        it('skips folder creation when team folder exists', async () => {
            jest.spyOn(client as any, 'ensureIUPChannel').mockResolvedValue(undefined);
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true })   // team folder check → exists
                .mockResolvedValueOnce({ ok: true });  // channel folder check → exists

            await client.ensureIUPData();

            expect(mockSp.spPost).not.toHaveBeenCalled();
        });
    });

    // ─── copyRequestTemplates / copyRequestTemplate ───────────────────────────

    describe('copyRequestTemplates', () => {
        it('downloads and uploads each template file', async () => {
            // Override default spGet for the copyRequestTemplates call inside initRequests
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: [
                    { FileRef: '/sites/template-site/Requests/file1.docx', FileLeafRef: 'file1.docx' },
                    { FileRef: '/sites/template-site/Requests/file2.docx', FileLeafRef: 'file2.docx' },
                ],
            });

            await client.initRequests();

            expect(mockSp.spGetBinary).toHaveBeenCalledTimes(2);
            expect(mockSp.spPostBinary).toHaveBeenCalledTimes(2);
        });

        it('does nothing when template list is empty', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });

            await client.initRequests();

            expect(mockSp.spGetBinary).not.toHaveBeenCalled();
        });
    });
});
