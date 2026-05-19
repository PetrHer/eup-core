/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import UserApiClient from './UserApiClient';
import { SPGroupExternal, SPGroupInternal, SPGroupRO, TeamRole } from '../../enums';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getTeamId: jest.fn().mockReturnValue('team-id-123'),
        getChannelId: jest.fn().mockReturnValue('channel-id-xyz'),
        internalSiteName: 'internal-site',
        templateSiteName: 'template-site',
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPost: jest.fn().mockResolvedValue({ ok: true }),
        spPatch: jest.fn().mockResolvedValue({ ok: true }),
        spDelete: jest.fn().mockResolvedValue({ ok: true }),
        graphGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        graphPost: jest.fn().mockResolvedValue({ ok: true }),
        getUserSPId: jest.fn().mockResolvedValue(42),
        addAttachmentFile: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as SPApiClient;
}

describe('UserApiClient', () => {
    let client: UserApiClient;
    let mockSp: ReturnType<typeof createMockSp>;
    const EUP_NO_REPLY_MAIL = 'noreply@example.com';

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new UserApiClient(mockSp, EUP_NO_REPLY_MAIL);
    });

    // ─── getTeamMembers ───────────────────────────────────────────────────────

    describe('getTeamMembers', () => {
        it('calls graphGet with team ID in URL', async () => {
            await client.getTeamMembers();
            expect(mockSp.graphGet).toHaveBeenCalledWith(
                expect.stringContaining('team-id-123')
            );
        });

        it('returns empty array when graphGet returns empty data', async () => {
            const result = await client.getTeamMembers();
            expect(result).toEqual([]);
        });

        it('returns mapped array from API value', async () => {
            const result = await client.getTeamMembers();
            expect(Array.isArray(result)).toBe(true);
        });

        it('filters team members to only those present in channel members', async () => {
            (mockSp.graphGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ email: 'member@x.com' }] }) // channel members
                .mockResolvedValueOnce({
                    ok: true, data: [  // team members
                        { mail: 'member@x.com', id: 'u1', displayName: 'Member', givenName: 'M', surname: 'B', mobilePhone: '' },
                        { mail: 'other@x.com', id: 'u2', displayName: 'Other', givenName: 'O', surname: 'T', mobilePhone: '' },
                    ]
                });
            const result = await client.getTeamMembers();
            expect(result).toHaveLength(1);
            expect(result[0].email).toBe('member@x.com');
        });
    });

    // ─── getExistingUser ──────────────────────────────────────────────────────

    describe('getExistingUser', () => {
        it('returns undefined when email is empty', async () => {
            const result = await client.getExistingUser('');
            expect(result).toBeUndefined();
        });

        it('calls graphGet with email in filter', async () => {
            await client.getExistingUser('user@example.com');
            expect(mockSp.graphGet).toHaveBeenCalledWith(
                expect.stringContaining('user%40example.com')
            );
        });

        it('returns undefined when API returns no results', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getExistingUser('user@example.com');
            expect(result).toBeUndefined();
        });

        it('returns first user from API results', async () => {
            const user = { id: 'u1', displayName: 'Test User', mail: 'user@example.com' };
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: [user] });
            const result = await client.getExistingUser('user@example.com');
            expect(result).toMatchObject({ id: 'u1' });
        });
    });

    // ─── getUserPhoto ─────────────────────────────────────────────────────────

    describe('getUserPhoto', () => {
        it('returns empty string when data is not a Blob', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getUserPhoto('user@x.com');
            expect(result).toBe('');
        });

        it('returns object URL when data is a Blob', async () => {
            const blob = new Blob(['img'], { type: 'image/jpeg' });
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: blob });
            const originalCreateObjectURL = URL.createObjectURL;
            URL.createObjectURL = jest.fn().mockReturnValue('blob:fake-url');
            const result = await client.getUserPhoto('user@x.com');
            expect(result).toBe('blob:fake-url');
            // eslint-disable-next-line require-atomic-updates
            URL.createObjectURL = originalCreateObjectURL;
        });
    });

    // ─── getCurrentUserGroups ─────────────────────────────────────────────────

    describe('getCurrentUserGroups', () => {
        it('calls spGet with site URL', async () => {
            await client.getCurrentUserGroups();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getCurrentUserGroups();
            expect(result).toEqual([]);
        });

        it('returns array from API data', async () => {
            const result = await client.getCurrentUserGroups();
            expect(Array.isArray(result)).toBe(true);
        });

        it('filters Groups to only those matching known enum values', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true,
                data: {
                    Groups: [
                        { Title: SPGroupInternal.UUO_RM },
                        { Title: SPGroupExternal.EXT_AUD },
                        { Title: 'Random Other Group' },
                    ]
                }
            });
            const result = await client.getCurrentUserGroups();
            expect(result).toContain(SPGroupInternal.UUO_RM);
            expect(result).toContain(SPGroupExternal.EXT_AUD);
            expect(result).not.toContain('Random Other Group');
        });
    });

    // ─── getSiteGroupUsersBatch ───────────────────────────────────────────────

    describe('getSiteGroupUsersBatch', () => {
        it('calls spPost with $batch URL', async () => {
            await client.getSiteGroupUsersBatch([1, 2]);
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('$batch'),
                expect.any(String),
                expect.any(Object)
            );
        });

        it('returns empty Map when response data is not a string', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getSiteGroupUsersBatch([1]);
            expect(result).toBeInstanceOf(Map);
            expect(result.size).toBe(0);
        });

        it('builds batch parts for each group ID', async () => {
            await client.getSiteGroupUsersBatch([5, 10, 15]);
            const [, batchBody] = (mockSp.spPost as jest.Mock).mock.calls[0];
            expect(batchBody).toContain('sitegroups(5)');
            expect(batchBody).toContain('sitegroups(10)');
            expect(batchBody).toContain('sitegroups(15)');
        });
    });

    // ─── addUserToSiteGroup ───────────────────────────────────────────────────

    describe('addUserToSiteGroup', () => {
        it('calls spPost with group ID in URL', async () => {
            await client.addUserToSiteGroup(7, 'user@test.com');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('sitegroups(7)'),
                expect.objectContaining({ LoginName: expect.stringContaining('user@test.com') })
            );
        });
    });

    // ─── removeUserFromsiteGroup ──────────────────────────────────────────────

    describe('removeUserFromsiteGroup', () => {
        it('calls spPost with group and user ID in URL', async () => {
            await client.removeUserFromsiteGroup(3, 99);
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('sitegroups(3)')
            );
        });
    });

    // ─── createInvitation ─────────────────────────────────────────────────────

    describe('createInvitation', () => {
        it('calls spPost with invitation URL', async () => {
            await client.createInvitation('a@x.com', 'Alice', 'Member', 'cs');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('Pozvanky'),
                expect.any(Object)
            );
        });

        it('returns true when spPost succeeds', async () => {
            const result = await client.createInvitation('a@x.com', 'Alice', 'Member', 'cs');
            expect(result).toBe(true);
        });

        it('returns false when spPost fails', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.createInvitation('a@x.com', 'Alice', 'Member', 'cs');
            expect(result).toBe(false);
        });

        it('sets User to email when existingUser is true', async () => {
            await client.createInvitation('a@x.com', 'Alice', 'Member', 'cs', true);
            const [, body] = (mockSp.spPost as jest.Mock).mock.calls[0];
            expect(body.User).toBe('a@x.com');
        });

        it('sets User to empty string when existingUser is false', async () => {
            await client.createInvitation('a@x.com', 'Alice', 'Member', 'cs', false);
            const [, body] = (mockSp.spPost as jest.Mock).mock.calls[0];
            expect(body.User).toBe('');
        });
    });

    // ─── checkTemplatesAccess ─────────────────────────────────────────────────

    describe('checkTemplatesAccess', () => {
        it('returns true when spGet returns ok', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true });
            const result = await client.checkTemplatesAccess();
            expect(result).toBe(true);
        });

        it('returns false when spGet returns not ok', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.checkTemplatesAccess();
            expect(result).toBe(false);
        });

        it('calls spGet with templateSiteName in URL', async () => {
            await client.checkTemplatesAccess();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('template-site')
            );
        });
    });

    // ─── getUserName ──────────────────────────────────────────────────────────

    describe('getUserName', () => {
        it('returns Title when result has data with Title', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: { Title: 'Alice Smith' } });
            const result = await client.getUserName('alice@x.com');
            expect(result).toBe('Alice Smith');
        });

        it('returns empty string when data has no Title', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getUserName('alice@x.com');
            expect(result).toBe('');
        });
    });

    // ─── getCurrentUser ───────────────────────────────────────────────────────

    describe('getCurrentUser', () => {
        it('calls graphGet with /me', async () => {
            await client.getCurrentUser();
            expect(mockSp.graphGet).toHaveBeenCalledWith('/me');
        });

        it('returns mapped person when data is present', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: { id: 'me-1', displayName: 'Me', mail: 'me@x.com' } });
            const result = await client.getCurrentUser();
            expect(result).toBeDefined();
            expect(result?.id).toBe('me-1');
        });

        it('returns undefined when data is null', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getCurrentUser();
            expect(result).toBeUndefined();
        });
    });

    // ─── getInvitations ───────────────────────────────────────────────────────

    describe('getInvitations', () => {
        it('calls spGet with channel ID filter', async () => {
            await client.getInvitations([]);
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('channel-id-xyz')
            );
        });

        it('filters out team members that already have the same email', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true, data: [{ Email: 'existing@x.com', Jmeno: 'Existing' }]
            });
            (mockSp.graphPost as jest.Mock).mockResolvedValue({ ok: true, data: { responses: [] } });
            const teamMembers: any[] = [{ email: 'existing@x.com', name: 'Existing' }];
            const result = await client.getInvitations(teamMembers);
            expect(result).toHaveLength(0);
        });

        it('returns invitations for emails not in team members', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({
                ok: true, data: [{ Email: 'new@x.com', Jmeno: 'New' }]
            });
            (mockSp.graphPost as jest.Mock).mockResolvedValue({
                ok: true, data: { responses: [{ id: '1', status: 200, body: { value: [{ id: 'u1' }] } }] }
            });
            const result = await client.getInvitations([]);
            expect(result).toHaveLength(1);
        });
    });

    // ─── createSiteGroups ─────────────────────────────────────────────────────

    describe('createSiteGroups', () => {
        it('calls spPost for each internal group and invokes increaseProgressCount', async () => {
            const progress = jest.fn();
            await client.createSiteGroups([SPGroupInternal.UUO_RM], progress);
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('sitegroups'),
                expect.objectContaining({ Title: SPGroupInternal.UUO_RM })
            );
            expect(progress).toHaveBeenCalledTimes(1);
        });

        it('calls spPost for RO group and assigns read role', async () => {
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: true, data: { Id: 5 } });
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 99 }] }); // getRoleDefinitionId(Read)
            await client.createSiteGroups([SPGroupRO.RORole], jest.fn());
            expect(mockSp.spPost).toHaveBeenCalledTimes(2); // create group + assignRoleToGroup
        });
    });

    // ─── ensureRole ───────────────────────────────────────────────────────────

    describe('ensureRole', () => {
        it('calls spPost to create role when it does not exist', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [] })    // getRoleDefinitionId(CustomEdit) → none
                .mockResolvedValueOnce({ ok: true, data: [] })    // getMissingRoleAssignments: roleassignments
                .mockResolvedValueOnce({ ok: true, data: [] })    // getRoleDefinitionId(CustomEdit) in getMissing
                .mockResolvedValueOnce({ ok: true, data: [] });   // getRoleDefinitionId(Edit) in getMissing
            await client.ensureRole();
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('roledefinitions'),
                expect.any(Object)
            );
        });

        it('skips role creation when role already exists', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 42 }] }) // getRoleDefinitionId → found
                .mockResolvedValueOnce({ ok: true, data: [] })           // getMissingRoleAssignments: roleassignments
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 42 }] }) // getRoleDefinitionId(CustomEdit)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 43 }] }); // getRoleDefinitionId(Edit)
            await client.ensureRole();
            expect(mockSp.spPost).not.toHaveBeenCalledWith(
                expect.stringContaining('roledefinitions'),
                expect.any(Object)
            );
        });

        it('assigns missing roles for known SP groups', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 42 }] }) // getRoleDefinitionId(CustomEdit) → exists
                .mockResolvedValueOnce({
                    ok: true, data: [             // roleassignments: one group missing role
                        { Member: { LoginName: SPGroupInternal.UUO_RM, Id: 1 }, RoleDefinitionBindings: [] }
                    ]
                })
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 42 }] }) // getRoleDefinitionId(CustomEdit)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 43 }] }); // getRoleDefinitionId(Edit)
            await client.ensureRole();
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('addroleassignment')
            );
        });
    });

    // ─── addEUPAdminsAsOwners ─────────────────────────────────────────────────

    describe('addEUPAdminsAsOwners', () => {
        it('returns false when owners spGet fails', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.addEUPAdminsAsOwners();
            expect(result).toBe(false);
        });

        it('calls spPost for both logins on each owner group ID', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 10 } }) // owners
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 20 }] }); // getGroupIdByName(ADMIN)
            await client.addEUPAdminsAsOwners();
            expect(mockSp.spPost).toHaveBeenCalledTimes(4); // 2 groups × 2 logins
        });

        it('returns false when any spPost fails', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 10 } })
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 20 }] });
            (mockSp.spPost as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.addEUPAdminsAsOwners();
            expect(result).toBe(false);
        });
    });

    // ─── addEUPMembersAsMembers ───────────────────────────────────────────────

    describe('addEUPMembersAsMembers', () => {
        it('returns false when members spGet fails', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.addEUPMembersAsMembers();
            expect(result).toBe(false);
        });

        it('calls spPost for member group and RO role group', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 5 } })     // members group
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 15 }] }); // getGroupIdByName(RORole)
            await client.addEUPMembersAsMembers();
            expect(mockSp.spPost).toHaveBeenCalledTimes(2); // membersId + roRoleId
        });

        it('returns true when all calls succeed', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 5 } })
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 15 }] });
            const result = await client.addEUPMembersAsMembers();
            expect(result).toBe(true);
        });
    });

    // ─── addUserByEmailToTeamAndChannel ──────────────────────────────────────

    describe('addUserByEmailToTeamAndChannel', () => {
        it('returns false when user not found in graph', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.addUserByEmailToTeamAndChannel('missing@x.com');
            expect(result).toBe(false);
        });

        it('adds user to team via graphPost when user found (standard channel)', async () => {
            (mockSp.graphGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ id: 'uid-1', mail: 'u@x.com' }] }) // users filter
                .mockResolvedValueOnce({ ok: true, data: {} }); // channel details (standard)
            (mockSp.graphPost as jest.Mock).mockResolvedValue({ ok: true });
            const result = await client.addUserByEmailToTeamAndChannel('u@x.com');
            expect(mockSp.graphPost).toHaveBeenCalledWith(
                expect.stringContaining('/members'),
                expect.any(Object)
            );
            expect(result).toBe(true);
        });

        it('takes private channel path when channel membershipType is private', async () => {
            (mockSp.graphGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ id: 'uid-1' }] }) // users filter
                .mockResolvedValueOnce({ ok: true, data: { membershipType: 'private' } }); // channel details
            (mockSp.graphPost as jest.Mock).mockResolvedValue({ ok: true });
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 100 } })   // owners in addMemberToPrivateChannel
                .mockResolvedValueOnce({ ok: true, data: { Id: 200 } });  // owners in addUserOrGroupByIdToTeamAndChannel
            const result = await client.addUserByEmailToTeamAndChannel('u@x.com');
            expect(mockSp.graphPost).toHaveBeenCalledTimes(2); // addMemberToTeam + addMemberToPrivateChannel
            expect(result).toBe(true);
        });

        it('calls fixGuestGroup when role is Guest and channel is standard', async () => {
            (mockSp.graphGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ id: 'uid-1' }] })
                .mockResolvedValueOnce({ ok: true, data: {} }); // standard channel
            (mockSp.graphPost as jest.Mock).mockResolvedValue({ ok: true });
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: { Id: 5 } })  // visitors
                .mockResolvedValueOnce({ ok: true, data: { Id: 6 } }); // members
            await client.addUserByEmailToTeamAndChannel('u@x.com', TeamRole.Guest);
            // fixGuestGroup calls spGet twice (visitors + members) then spPost x2
            expect(mockSp.spGet).toHaveBeenCalledTimes(2);
            expect(mockSp.spPost).toHaveBeenCalledTimes(2);
        });
    });

    // ─── checkUserAndCreateInvitation ────────────────────────────────────────

    describe('checkUserAndCreateInvitation', () => {
        it('creates invitation when user does not exist in AD', async () => {
            (mockSp.graphGet as jest.Mock).mockResolvedValue({ ok: false }); // user not in AD
            await client.checkUserAndCreateInvitation('new@x.com', 'New', 'Member', 'cs');
            expect(mockSp.spPost).toHaveBeenCalledWith(
                expect.stringContaining('Pozvanky'),
                expect.any(Object)
            );
        });

        it('calls addUserByEmailToTeamAndChannel when user exists in AD', async () => {
            (mockSp.graphGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true })  // checkADUserExistence → exists
                .mockResolvedValueOnce({ ok: true, data: [] }); // addUserByEmailToTeamAndChannel → no userId
            const result = await client.checkUserAndCreateInvitation('existing@x.com', 'Existing', 'Member', 'cs');
            expect(result).toBe(false); // user found in AD but no graph user found
        });
    });
});
