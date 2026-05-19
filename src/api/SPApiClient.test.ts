/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import LZString from 'lz-string';

import SPApiClient from './SPApiClient';
import { LibraryName, ListName, SPGroupInternal } from '../enums';

jest.mock('lz-string', () => ({
    decompressFromEncodedURIComponent: jest.fn(),
    compressToEncodedURIComponent: jest.fn(() => 'compressed-ctx'),
    default: {
        decompressFromEncodedURIComponent: jest.fn(),
        compressToEncodedURIComponent: jest.fn(() => 'compressed-ctx'),
    },
}));
jest.mock('@pnp/sp', () => ({
    spfi: jest.fn(() => ({ using: jest.fn().mockReturnValue({}) })),
    SPFx: jest.fn(),
}));
jest.mock('./SPApiClient');

// We need to test the real implementation, not the auto-mock.
// Re-import after clearing the automock by using the actual module.
const ActualSPApiClient = jest.requireActual('./SPApiClient').default;

function createMockContext() {
    return {
        spHttpClient: {
            get: jest.fn(),
            post: jest.fn(),
        },
        msGraphClientFactory: {
            getClient: jest.fn(),
        },
        pageContext: { user: { email: 'user@test.com' } },
    } as any;
}

function createClient() {
    const ctx = createMockContext();
    const client: SPApiClient = new ActualSPApiClient(ctx, 'internal-site', 'template-site');
    // Set protected fields directly so tests don't require init()
    (client as any).absoluteUrl = 'https://example.sharepoint.com/sites/team';
    (client as any).relativeUrl = '/sites/team';
    (client as any).channelId = 'ch-123';
    (client as any).teamId = 'team-abc';
    (client as any).channelName = 'TestChannel';
    (client as any).userRole = 1;
    (client as any).teamName = 'My Team';
    return { client, ctx };
}

describe('SPApiClient', () => {
    let client: SPApiClient;

    beforeEach(() => {
        jest.clearAllMocks();
        ({ client } = createClient());
    });

    // ─── Pure / synchronous methods ───────────────────────────────────────────

    describe('chunkArray', () => {
        it('splits array into chunks of given size', () => {
            const result = client.chunkArray([1, 2, 3, 4, 5], 2);
            expect(result).toEqual([[1, 2], [3, 4], [5]]);
        });

        it('returns single chunk when array is smaller than chunk size', () => {
            const result = client.chunkArray([1, 2], 10);
            expect(result).toEqual([[1, 2]]);
        });

        it('returns empty array for empty input', () => {
            expect(client.chunkArray([], 5)).toEqual([]);
        });
    });

    describe('getPermissionBatchString', () => {
        it('returns a string containing principalid and roledefid', () => {
            const result = client.getPermissionBatchString('batch_1', 42, 1073741826, 'https://example.com/api');
            expect(result).toContain('principalid=42');
            expect(result).toContain('roledefid=1073741826');
        });

        it('includes boundary markers', () => {
            const result = client.getPermissionBatchString('batch_1', 1, 1, 'https://example.com/api');
            expect(result).toContain('--batch_1');
        });
    });

    describe('getters', () => {
        it('getTeamId returns team id', () => {
            expect(client.getTeamId()).toBe('team-abc');
        });

        it('getChannelId returns channel id', () => {
            expect(client.getChannelId()).toBe('ch-123');
        });

        it('getAbsoluteUrl returns absolute url', () => {
            expect(client.getAbsoluteUrl()).toBe('https://example.sharepoint.com/sites/team');
        });

        it('getChannelName returns channel name', () => {
            expect(client.getChannelName()).toBe('TestChannel');
        });

        it('getRelativeUrl returns relative url', () => {
            expect(client.getRelativeUrl()).toBe('/sites/team');
        });

        it('getUserRole returns user role', () => {
            expect(client.getUserRole()).toBe(1);
        });
    });

    describe('getTeamName', () => {
        it('returns cached teamName without calling graph', async () => {
            const loadSpy = jest.spyOn(client as any, 'loadTeamName');
            const result = await client.getTeamName();
            expect(result).toBe('My Team');
            expect(loadSpy).not.toHaveBeenCalled();
        });

        it('loads from graph when teamName is not set', async () => {
            (client as any).teamName = undefined;
            jest.spyOn(client as any, 'loadTeamName').mockResolvedValue('Loaded Team');
            const result = await client.getTeamName();
            expect(result).toBe('Loaded Team');
        });
    });

    // ─── Async methods (spy on spGet/spPost/spPatch/spDelete) ─────────────────

    describe('checkListExistence', () => {
        it('returns true when response code is not 404', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, code: 200 });
            const result = await client.checkListExistence(ListName.Requests);
            expect(result).toBe(true);
        });

        it('returns false when response code is 404', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: false, code: 404 });
            const result = await client.checkListExistence(ListName.Requests);
            expect(result).toBe(false);
        });

        it('calls spGet with list name in URL', async () => {
            const spy = jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, code: 200 });
            await client.checkListExistence(ListName.Requests);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining(ListName.Requests),
                expect.anything(),
                false
            );
        });
    });

    describe('createList', () => {
        it('returns true when spPost succeeds', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            const result = await client.createList(ListName.QA);
            expect(result).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: false });
            const result = await client.createList(ListName.QA);
            expect(result).toBe(false);
        });

        it('uses BaseTemplate 101 for Requests library', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createList(LibraryName.Requests);
            const body = spy.mock.calls[0][1] as any;
            expect(body.BaseTemplate).toBe(101);
        });

        it('uses BaseTemplate 100 for regular lists', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createList(ListName.QA);
            const body = spy.mock.calls[0][1] as any;
            expect(body.BaseTemplate).toBe(100);
        });
    });

    describe('createListItem', () => {
        it('calls spPost with item body', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true, data: { Id: 1 } });
            await client.createListItem({ Title: 'Test' }, ListName.QA);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining(ListName.QA),
                { Title: 'Test' }
            );
        });

        it('returns the API result', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true, data: { Id: 7 } });
            const result = await client.createListItem({}, ListName.QA);
            expect(result.ok).toBe(true);
            expect((result.data as any).Id).toBe(7);
        });
    });

    describe('updateListItem', () => {
        it('calls spPatch with item id in URL', async () => {
            const spy = jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            await client.updateListItem(5, { Title: 'Updated' }, ListName.QA);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('items(5)'),
                expect.any(Object)
            );
        });

        it('returns true when spPatch succeeds', async () => {
            jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            expect(await client.updateListItem(1, {}, ListName.QA)).toBe(true);
        });

        it('returns false when spPatch fails', async () => {
            jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: false });
            expect(await client.updateListItem(1, {}, ListName.QA)).toBe(false);
        });
    });

    describe('updateItemInList', () => {
        it('calls spPatch with GetItemByUniqueId URL', async () => {
            const spy = jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            await client.updateItemInList('uid-abc', { Name_en: 'Test' });
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining("GetItemByUniqueId('uid-abc')"),
                expect.any(Object)
            );
        });

        it('returns true when spPatch succeeds', async () => {
            jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            expect(await client.updateItemInList('uid', {})).toBe(true);
        });
    });

    describe('deleteListItem', () => {
        it('calls spDelete with item id in URL', async () => {
            const spy = jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: true });
            await client.deleteListItem(3, ListName.QA);
            expect(spy).toHaveBeenCalledWith(expect.stringContaining('items(3)'));
        });

        it('returns true when spDelete succeeds', async () => {
            jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: true });
            expect(await client.deleteListItem(1, ListName.QA)).toBe(true);
        });

        it('returns false when spDelete fails', async () => {
            jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: false });
            expect(await client.deleteListItem(1, ListName.QA)).toBe(false);
        });
    });

    describe('getFieldChoices', () => {
        it('returns Choices array from API response', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: { Choices: ['A', 'B', 'C'] } });
            const result = await client.getFieldChoices(ListName.QA, 'Status');
            expect(result).toEqual(['A', 'B', 'C']);
        });

        it('returns empty array when Choices is absent', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: {} });
            const result = await client.getFieldChoices(ListName.QA, 'Status');
            expect(result).toEqual([]);
        });
    });

    describe('getUserSPId', () => {
        it('calls spPost with ensureUser endpoint and email', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true, data: { Id: 55 } });
            await client.getUserSPId('user@test.com');
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('ensureUser'),
                expect.objectContaining({ logonName: 'user@test.com' })
            );
        });

        it('returns user Id from response', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true, data: { Id: 55 } });
            const result = await client.getUserSPId('user@test.com');
            expect(result).toBe(55);
        });

        it('returns 0 when Id is absent from response', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: false, data: null });
            const result = await client.getUserSPId('user@test.com');
            expect(result).toBe(0);
        });
    });

    describe('ensureFolder', () => {
        it('returns UniqueId from existing folder when spGet succeeds', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: { UniqueId: 'existing-uid' } });
            const result = await client.ensureFolder('/path', 'folderName');
            expect(result).toBe('existing-uid');
        });

        it('calls spPost to create folder when it does not exist', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: false });
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true, data: { UniqueId: 'new-uid' } });
            await client.ensureFolder('/path', 'newFolder');
            expect(postSpy).toHaveBeenCalledWith(expect.stringContaining("folders/add('newFolder')"));
        });
    });

    describe('getSiteGroups', () => {
        it('returns only groups matching SPGroupInternal or SPGroupExternal', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({
                ok: true,
                data: [
                    { Id: 1, Title: SPGroupInternal.ADMIN, Users: [] },
                    { Id: 2, Title: 'Random Group', Users: [] },
                ],
            });
            const result = await client.getSiteGroups();
            expect(result).toHaveLength(1);
            expect(result[0].title).toBe(SPGroupInternal.ADMIN);
        });

        it('returns empty array when data is not an array', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: null });
            const result = await client.getSiteGroups();
            expect(result).toEqual([]);
        });
    });

    describe('createListItemUsingPath', () => {
        it('calls spPost with AddValidateUpdateItemUsingPath endpoint', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createListItemUsingPath({ formValues: [] }, ListName.QA);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('AddValidateUpdateItemUsingPath'),
                expect.any(Object)
            );
        });

        it('includes channel name in folder path when channelSpecific is true (default)', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createListItemUsingPath({}, ListName.QA);
            const body = spy.mock.calls[0][1] as any;
            expect(body.listItemCreateInfo.FolderPath.DecodedUrl).toContain('TestChannel');
        });

        it('excludes channel name from folder path when channelSpecific is false', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createListItemUsingPath({}, ListName.QA, undefined, false);
            const body = spy.mock.calls[0][1] as any;
            expect(body.listItemCreateInfo.FolderPath.DecodedUrl).not.toContain('TestChannel');
        });

        it('appends optional path to folder path', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createListItemUsingPath({}, ListName.QA, '/subpath');
            const body = spy.mock.calls[0][1] as any;
            expect(body.listItemCreateInfo.FolderPath.DecodedUrl).toContain('/subpath');
        });
    });

    describe('checkAllListExistence', () => {
        it('returns an object with list names mapped to booleans', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({
                ok: true,
                data: [{ Title: ListName.Requests }, { Title: ListName.QA }],
            });
            const result = await client.checkAllListExistence();
            expect(result[ListName.Requests]).toBe(true);
            expect(result[ListName.QA]).toBe(true);
            expect(result[ListName.Details]).toBe(false);
        });

        it('marks all as false when data is not an array', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: null });
            const result = await client.checkAllListExistence();
            expect(result[ListName.Requests]).toBe(false);
        });
    });

    // ─── ensureList ───────────────────────────────────────────────────────────

    describe('ensureList', () => {
        it('creates list and ensures columns when list does not exist', async () => {
            jest.spyOn(client, 'checkListExistence').mockResolvedValue(false);
            const createSpy = jest.spyOn(client, 'createList').mockResolvedValue(true);
            jest.spyOn(client, 'ensureColumns').mockResolvedValue([]);
            jest.spyOn(client, 'ensureListView').mockResolvedValue(undefined);

            await client.ensureList(ListName.QA);

            expect(createSpy).toHaveBeenCalledWith(ListName.QA);
        });

        it('skips creation when list already exists', async () => {
            jest.spyOn(client, 'checkListExistence').mockResolvedValue(true);
            const createSpy = jest.spyOn(client, 'createList').mockResolvedValue(true);
            jest.spyOn(client, 'ensureColumns').mockResolvedValue([]);

            await client.ensureList(ListName.QA);

            expect(createSpy).not.toHaveBeenCalled();
        });

        it('returns early without ensuring columns when creation fails', async () => {
            jest.spyOn(client, 'checkListExistence').mockResolvedValue(false);
            jest.spyOn(client, 'createList').mockResolvedValue(false);
            const columnsSpy = jest.spyOn(client, 'ensureColumns').mockResolvedValue([]);

            await client.ensureList(ListName.QA);

            expect(columnsSpy).not.toHaveBeenCalled();
        });
    });

    // ─── deleteList ───────────────────────────────────────────────────────────

    describe('deleteList', () => {
        it('deletes items, folders and list; returns true on success', async () => {
            jest.spyOn(client, 'spGet')
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 1 }] })
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 2 }] });
            const deleteSpy = jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: true });

            const result = await client.deleteList(ListName.QA);

            expect(deleteSpy).toHaveBeenCalledTimes(3);
            expect(result).toBe(true);
        });

        it('returns false when final delete fails', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: [] });
            jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: false });

            expect(await client.deleteList(ListName.QA)).toBe(false);
        });

        it('handles empty items and folders gracefully', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: null });
            jest.spyOn(client, 'spDelete').mockResolvedValue({ ok: true });

            expect(await client.deleteList(ListName.QA)).toBe(true);
        });
    });

    // ─── updateList ───────────────────────────────────────────────────────────

    describe('updateList', () => {
        it('calls spPatch with list URL and body', async () => {
            const spy = jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            await client.updateList({ EnableVersioning: true }, ListName.QA);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining(ListName.QA),
                { EnableVersioning: true }
            );
        });
    });

    // ─── updateColumn ─────────────────────────────────────────────────────────

    describe('updateColumn', () => {
        it('calls spPatch with column name in URL', async () => {
            const spy = jest.spyOn(client, 'spPatch').mockResolvedValue({ ok: true });
            await client.updateColumn({ Required: true }, ListName.QA, 'MyColumn');
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('MyColumn'),
                { Required: true },
                expect.any(Object)
            );
        });
    });

    // ─── updateListItemsBatch ─────────────────────────────────────────────────

    describe('updateListItemsBatch', () => {
        it('sends batch request to $batch endpoint', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.updateListItemsBatch(
                [{ uniqueId: 'uid-1', body: { Title: 'A' } }],
                ListName.QA
            );
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('$batch'),
                expect.any(String),
                expect.any(Object)
            );
        });

        it('chunks updates into groups of 50', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            const updates = Array.from({ length: 60 }, (_, i) => ({ uniqueId: `uid-${i}`, body: {} }));
            await client.updateListItemsBatch(updates, ListName.QA);
            expect(spy).toHaveBeenCalledTimes(2);
        });
    });

    // ─── ensureColumns ────────────────────────────────────────────────────────

    describe('ensureColumns', () => {
        it('calls createColumn for missing columns', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: [] });
            const createColumnSpy = jest.spyOn(client as any, 'createColumn').mockResolvedValue(true);

            await client.ensureColumns(ListName.QA);

            expect(createColumnSpy).toHaveBeenCalled();
        });

        it('returns missing columns array', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: [] });
            jest.spyOn(client as any, 'createColumn').mockResolvedValue(true);

            const result = await client.ensureColumns(ListName.QA);

            expect(Array.isArray(result)).toBe(true);
        });
    });

    // ─── ensureListView ───────────────────────────────────────────────────────

    describe('ensureListView', () => {
        it('calls spPost for each column', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: { Id: 'view-123' } });
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            const cols = [{ Title: 'ColA' }, { Title: 'ColB' }] as any[];

            await client.ensureListView(ListName.QA, cols);

            expect(postSpy).toHaveBeenCalledTimes(2);
            expect(postSpy).toHaveBeenCalledWith(expect.stringContaining('ColA'));
        });

        it('does nothing when columns array is empty', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: { Id: 'view-123' } });
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            await client.ensureListView(ListName.QA, []);

            expect(postSpy).not.toHaveBeenCalled();
        });
    });

    // ─── init ─────────────────────────────────────────────────────────────────

    describe('init', () => {
        it('calls initWithCtx when ctx is provided', async () => {
            const spy = jest.spyOn(client, 'initWithCtx').mockResolvedValue();
            await client.init('some-ctx');
            expect(spy).toHaveBeenCalledWith('some-ctx');
        });

        it('calls initWithIds when teamId and channelId are provided', async () => {
            const spy = jest.spyOn(client, 'initWithIds').mockResolvedValue();
            await client.init(undefined, 'team-1', 'ch-1');
            expect(spy).toHaveBeenCalledWith('team-1', 'ch-1');
        });

        it('does nothing when no args are provided', async () => {
            const ctxSpy = jest.spyOn(client, 'initWithCtx').mockResolvedValue();
            const idsSpy = jest.spyOn(client, 'initWithIds').mockResolvedValue();
            await client.init();
            expect(ctxSpy).not.toHaveBeenCalled();
            expect(idsSpy).not.toHaveBeenCalled();
        });
    });

    // ─── initWithCtx ──────────────────────────────────────────────────────────

    describe('initWithCtx', () => {
        it('throws on invalid ctx string when LZString returns null', async () => {
            (LZString.decompressFromEncodedURIComponent as jest.Mock).mockReturnValue(null);
            await expect(client.initWithCtx('bad-ctx')).rejects.toThrow('Invalid ctx format');
        });

        it('throws on invalid ctx string when JSON parse fails', async () => {
            (LZString.decompressFromEncodedURIComponent as jest.Mock).mockReturnValue('not-valid-json{');
            await expect(client.initWithCtx('bad-ctx')).rejects.toThrow('Invalid ctx format');
        });

        it('throws when required keys are missing', async () => {
            (LZString.decompressFromEncodedURIComponent as jest.Mock).mockReturnValue('{"t":"team"}');
            await expect(client.initWithCtx('incomplete-ctx')).rejects.toThrow('Invalid ctx format');
        });

        it('sets instance properties from valid ctx', async () => {
            const payload = { t: 'team-1', c: 'ch-1', n: 'General', s: 'https://example.sharepoint.com/sites/test' };
            (LZString.decompressFromEncodedURIComponent as jest.Mock).mockReturnValue(JSON.stringify(payload));
            jest.spyOn(client as any, 'getUserChannelRole').mockResolvedValue(1);

            await client.initWithCtx('valid-ctx');

            expect(client.getTeamId()).toBe('team-1');
            expect(client.getChannelId()).toBe('ch-1');
            expect(client.getChannelName()).toBe('General');
            expect(client.getAbsoluteUrl()).toBe('https://example.sharepoint.com/sites/test');
            expect(client.getUserRole()).toBe(1);
        });
    });

    // ─── initWithIds ──────────────────────────────────────────────────────────

    describe('initWithIds', () => {
        it('sets instance properties from channel detail', async () => {
            jest.spyOn(client as any, 'getChannelDetail').mockResolvedValue({
                absoluteUrl: 'https://example.sharepoint.com/sites/myteam',
                userRole: 2,
                displayName: 'My Channel',
            });
            jest.spyOn(client as any, 'upgradeLegacyUrl').mockImplementation(() => { /* */ });

            await client.initWithIds('team-xyz', 'channel-xyz');

            expect(client.getTeamId()).toBe('team-xyz');
            expect(client.getChannelId()).toBe('channel-xyz');
            expect(client.getChannelName()).toBe('My Channel');
            expect(client.getAbsoluteUrl()).toBe('https://example.sharepoint.com/sites/myteam');
            expect(client.getUserRole()).toBe(2);
        });

        it('falls back to getTeamSiteUrl when absoluteUrl is empty', async () => {
            jest.spyOn(client as any, 'getChannelDetail').mockResolvedValue({
                absoluteUrl: '',
                userRole: 0,
                displayName: 'Channel',
            });
            jest.spyOn(client as any, 'getTeamSiteUrl').mockResolvedValue('https://example.sharepoint.com/sites/fallback');
            jest.spyOn(client as any, 'upgradeLegacyUrl').mockImplementation(() => { /* */ });

            await client.initWithIds('team-xyz', 'channel-xyz');

            expect(client.getAbsoluteUrl()).toBe('https://example.sharepoint.com/sites/fallback');
        });

        it('calls upgradeLegacyUrl after setting properties', async () => {
            jest.spyOn(client as any, 'getChannelDetail').mockResolvedValue({
                absoluteUrl: 'https://example.sharepoint.com/sites/myteam',
                userRole: 1,
                displayName: 'Channel',
            });
            const upgradeSpy = jest.spyOn(client as any, 'upgradeLegacyUrl').mockImplementation(() => {/* */ });

            await client.initWithIds('team-xyz', 'channel-xyz');

            expect(upgradeSpy).toHaveBeenCalledTimes(1);
        });
    });

    // ─── loadTeamName ─────────────────────────────────────────────────────────

    describe('loadTeamName', () => {
        it('returns cached teamName when already set', async () => {
            (client as any).teamName = 'Cached Team';
            const spy = jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: true, data: {} });

            const result = await (client as any).loadTeamName();

            expect(result).toBe('Cached Team');
            expect(spy).not.toHaveBeenCalled();
        });

        it('fetches from graph when teamName is not set', async () => {
            (client as any).teamName = undefined;
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: true, data: { displayName: 'Graph Team' } });

            const result = await (client as any).loadTeamName();

            expect(result).toBe('Graph Team');
            expect((client as any).teamName).toBe('Graph Team');
        });

        it('returns empty string when graph data is absent', async () => {
            (client as any).teamName = undefined;
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: false, data: null });

            const result = await (client as any).loadTeamName();

            expect(result).toBe('');
        });
    });

    // ─── createListItemsBatch ─────────────────────────────────────────────────

    describe('createListItemsBatch', () => {
        it('sends batch POST request to $batch endpoint', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            await client.createListItemsBatch([{ Title: 'A' }, { Title: 'B' }], ListName.QA);
            expect(spy).toHaveBeenCalledWith(
                expect.stringContaining('$batch'),
                expect.any(String),
                expect.any(Object)
            );
        });

        it('chunks items into groups of 50', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            const items = Array.from({ length: 55 }, (_, i) => ({ Title: `Item ${i}` }));
            await client.createListItemsBatch(items, ListName.QA);
            expect(spy).toHaveBeenCalledTimes(2);
        });
    });

    // ─── moveListItem ─────────────────────────────────────────────────────────

    describe('moveListItem', () => {
        it('calls spPost to move file when item data is found', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({
                ok: true,
                data: { FileRef: '/sites/team/Docs/file.docx', FileDirRef: '/sites/team/Docs', FileLeafRef: 'file.docx' }
            });
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            await client.moveListItem(1, '/sites/team/Archive', ListName.QA);

            expect(postSpy).toHaveBeenCalledWith(expect.stringContaining('moveto'));
        });

        it('does not call spPost when item has no data', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: false, data: null });
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            await client.moveListItem(1, '/target', ListName.QA);

            expect(postSpy).not.toHaveBeenCalled();
        });
    });

    // ─── addAttachmentArrayBuffer ─────────────────────────────────────────────

    describe('addAttachmentArrayBuffer', () => {
        it('uploads buffer and returns API result', async () => {
            jest.spyOn(client as any, 'checkAttachmentExists').mockResolvedValue(false);
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            const result = await client.addAttachmentArrayBuffer(1, 'file.pdf', new ArrayBuffer(8), ListName.QA);

            expect(result.ok).toBe(true);
        });

        it('renames file when attachment with same name already exists', async () => {
            jest.spyOn(client as any, 'checkAttachmentExists')
                .mockResolvedValueOnce(true)
                .mockResolvedValueOnce(false);
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            await client.addAttachmentArrayBuffer(1, 'file.pdf', new ArrayBuffer(8), ListName.QA);

            expect(postSpy).toHaveBeenCalledWith(expect.stringContaining('file (1).pdf'), expect.anything());
        });

        it('returns ok:true immediately when skipIfExists is true and file exists', async () => {
            jest.spyOn(client as any, 'checkAttachmentExists').mockResolvedValue(true);
            const postSpy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });

            const result = await client.addAttachmentArrayBuffer(1, 'file.pdf', new ArrayBuffer(8), ListName.QA, true);

            expect(result).toEqual({ ok: true });
            expect(postSpy).not.toHaveBeenCalled();
        });
    });

    // ─── addAttachmentFile ────────────────────────────────────────────────────

    describe('addAttachmentFile', () => {
        it('delegates to addAttachmentArrayBuffer', async () => {
            const buffer = new ArrayBuffer(4);
            const mockFile = { rawFile: { arrayBuffer: jest.fn().mockResolvedValue(buffer) }, name: 'doc.pdf' } as any;
            const spy = jest.spyOn(client, 'addAttachmentArrayBuffer').mockResolvedValue({ ok: true });

            await client.addAttachmentFile(1, mockFile, ListName.QA);

            expect(spy).toHaveBeenCalledWith(1, 'doc.pdf', buffer, ListName.QA);
        });
    });

    // ─── getTemplatesAttachments ──────────────────────────────────────────────

    describe('getTemplatesAttachments', () => {
        it('returns mapped attachments when item exists', async () => {
            jest.spyOn(client, 'spGet')
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 10 }] })
                .mockResolvedValueOnce({ ok: true, data: [{ FileName: 'a.docx', ServerRelativeUrl: '/a.docx' }] });

            const result = await client.getTemplatesAttachments(ListName.QA);

            expect(result).toEqual([{ name: 'a.docx', serverRelativeUrl: '/a.docx' }]);
        });

        it('returns empty array when no items found', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: [] });

            const result = await client.getTemplatesAttachments(ListName.QA);

            expect(result).toEqual([]);
        });
    });

    // ─── ensureListFolder ─────────────────────────────────────────────────────

    describe('ensureListFolder', () => {
        it('does nothing when folder already exists', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: true, data: {} });
            const createSpy = jest.spyOn(client, 'createListItem').mockResolvedValue({ ok: true, data: { Id: 1 } });

            await client.ensureListFolder('MyFolder', ListName.QA);

            expect(createSpy).not.toHaveBeenCalled();
        });

        it('creates folder when it does not exist', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: false });
            const createSpy = jest.spyOn(client, 'createListItem').mockResolvedValue({ ok: true, data: { Id: 5 } });
            jest.spyOn(client, 'updateListItem').mockResolvedValue(true);

            await client.ensureListFolder('MyFolder', ListName.QA);

            expect(createSpy).toHaveBeenCalled();
        });

        it('calls moveFolder when path is specified and folder does not exist', async () => {
            jest.spyOn(client, 'spGet').mockResolvedValue({ ok: false });
            jest.spyOn(client, 'createListItem').mockResolvedValue({ ok: true, data: { Id: 5 } });
            jest.spyOn(client, 'updateListItem').mockResolvedValue(true);
            const moveSpy = jest.spyOn(client, 'moveFolder').mockResolvedValue(true);

            await client.ensureListFolder('MyFolder', ListName.QA, 'SubPath');

            expect(moveSpy).toHaveBeenCalled();
        });
    });

    // ─── moveFolder ───────────────────────────────────────────────────────────

    describe('moveFolder', () => {
        it('calls spPost with moveTo URL and returns true on success', async () => {
            const spy = jest.spyOn(client, 'spPost').mockResolvedValue({ ok: true });
            const result = await client.moveFolder('/sites/team/old', '/sites/team/new');
            expect(spy).toHaveBeenCalledWith(expect.stringContaining('moveTo'));
            expect(result).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            jest.spyOn(client, 'spPost').mockResolvedValue({ ok: false });
            expect(await client.moveFolder('/a', '/b')).toBe(false);
        });
    });

    // ─── log ──────────────────────────────────────────────────────────────────

    describe('log', () => {
        it('calls ensureListFolder and createListItemUsingPath', async () => {
            const folderSpy = jest.spyOn(client, 'ensureListFolder').mockResolvedValue(undefined);
            const createSpy = jest.spyOn(client, 'createListItemUsingPath').mockResolvedValue({ ok: true });

            await client.log('Test message', { key: 'value' });

            expect(folderSpy).toHaveBeenCalledWith(expect.any(String), ListName.Logs);
            expect(createSpy).toHaveBeenCalled();
        });
    });

    // ─── hideLists ────────────────────────────────────────────────────────────

    describe('hideLists', () => {
        it('calls update for each list name', async () => {
            const updateMock = jest.fn().mockResolvedValue({});
            (client as any).sp = {
                web: { lists: { getByTitle: jest.fn().mockReturnValue({ update: updateMock }) } }
            };

            await client.hideLists([ListName.QA, ListName.Details]);

            expect(updateMock).toHaveBeenCalledTimes(2);
            expect(updateMock).toHaveBeenCalledWith({ Hidden: true });
        });
    });

    // ─── upgradeLegacyUrl ────────────────────────────────────────────────────

    describe('upgradeLegacyUrl', () => {
        let replaceStateSpy: jest.SpyInstance;

        beforeEach(() => {
            replaceStateSpy = jest.spyOn(window.history, 'replaceState').mockImplementation(() => {/* */ });
            (LZString.compressToEncodedURIComponent as jest.Mock).mockReturnValue('encoded-ctx');
        });

        it('calls window.history.replaceState with a ctx query param', () => {
            (client as any).upgradeLegacyUrl();
            expect(replaceStateSpy).toHaveBeenCalledWith(null, '', expect.stringContaining('ctx=encoded-ctx'));
        });

        it('compresses teamId, channelId, channelName and absoluteUrl into ctx', () => {
            (client as any).upgradeLegacyUrl();
            const compressArg = (LZString.compressToEncodedURIComponent as jest.Mock).mock.calls[0][0];
            const parsed = JSON.parse(compressArg);
            expect(parsed.t).toBe('team-abc');
            expect(parsed.c).toBe('ch-123');
            expect(parsed.n).toBe('TestChannel');
            expect(parsed.s).toBe('https://example.sharepoint.com/sites/team');
        });

        it('uses current window.location.pathname as the base path', () => {
            (client as any).upgradeLegacyUrl();
            const [[, , url]] = replaceStateSpy.mock.calls;
            expect(url).toMatch(/^.*\?ctx=encoded-ctx$/);
        });
    });

    // ─── getChannelDetail ─────────────────────────────────────────────────────

    describe('getChannelDetail', () => {
        it('returns absoluteUrl extracted from filesFolder webUrl', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({
                ok: true, data: { webUrl: 'https://example.sharepoint.com/sites/myteam/General' }
            });
            jest.spyOn(client as any, 'getChannelDisplayName').mockResolvedValue('General');
            jest.spyOn(client as any, 'getUserChannelRole').mockResolvedValue(1);

            const result = await (client as any).getChannelDetail('team-1', 'ch-1');

            expect(result.absoluteUrl).toBe('https://example.sharepoint.com/sites/myteam');
            expect(result.displayName).toBe('General');
            expect(result.userRole).toBe(1);
        });

        it('returns empty absoluteUrl when graphGet data is absent', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: false, data: null });
            jest.spyOn(client as any, 'getChannelDisplayName').mockResolvedValue('');
            jest.spyOn(client as any, 'getUserChannelRole').mockResolvedValue(0);

            const result = await (client as any).getChannelDetail('team-1', 'ch-1');

            expect(result.absoluteUrl).toBe('');
        });

        it('calls graphGet with the filesFolder endpoint', async () => {
            const spy = jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: false, data: null });
            jest.spyOn(client as any, 'getChannelDisplayName').mockResolvedValue('');
            jest.spyOn(client as any, 'getUserChannelRole').mockResolvedValue(0);

            await (client as any).getChannelDetail('team-1', 'ch-1');

            expect(spy).toHaveBeenCalledWith(expect.stringContaining('filesFolder'));
        });
    });

    // ─── getChannelDisplayName ────────────────────────────────────────────────

    describe('getChannelDisplayName', () => {
        it('returns displayName from graph response', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: true, data: { displayName: 'General' } });

            const result = await (client as any).getChannelDisplayName('team-1', 'ch-1');

            expect(result).toBe('General');
        });

        it('returns empty string when data is absent', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: false, data: null });

            const result = await (client as any).getChannelDisplayName('team-1', 'ch-1');

            expect(result).toBe('');
        });

        it('calls graphGet with the channels endpoint', async () => {
            const spy = jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: true, data: { displayName: 'Test' } });

            await (client as any).getChannelDisplayName('team-1', 'ch-1');

            expect(spy).toHaveBeenCalledWith(expect.stringContaining('teams/team-1/channels/ch-1'));
        });
    });

    // ─── getTeamSiteUrl ───────────────────────────────────────────────────────

    describe('getTeamSiteUrl', () => {
        it('returns webUrl from graph response', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({
                ok: true, data: { webUrl: 'https://example.sharepoint.com/sites/team' }
            });

            const result = await (client as any).getTeamSiteUrl('team-1');

            expect(result).toBe('https://example.sharepoint.com/sites/team');
        });

        it('returns empty string when data is absent', async () => {
            jest.spyOn(client, 'graphGet').mockResolvedValue({ ok: false, data: null });

            const result = await (client as any).getTeamSiteUrl('team-1');

            expect(result).toBe('');
        });

        it('calls graphGet with the groups sites/root endpoint', async () => {
            const spy = jest.spyOn(client, 'graphGet').mockResolvedValue({
                ok: true, data: { webUrl: 'https://example.com' }
            });

            await (client as any).getTeamSiteUrl('team-1');

            expect(spy).toHaveBeenCalledWith(expect.stringContaining('groups/team-1/sites/root'));
        });
    });
});
