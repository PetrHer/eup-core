/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import DetailApiClient from './DetailApiClient';
import { ChannelStatus, ContactDesc, UserType } from '../../enums';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelId: jest.fn().mockReturnValue('channel-123'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        getTeamName: jest.fn().mockResolvedValue('TestTeam'),
        internalSiteName: 'internal-site',
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPatch: jest.fn().mockResolvedValue({ ok: true }),
        createListItem: jest.fn().mockResolvedValue({ ok: true }),
        updateListItem: jest.fn().mockResolvedValue(true),
        getUserSPId: jest.fn().mockResolvedValue(42),
    } as unknown as SPApiClient;
}

describe('DetailApiClient', () => {
    let client: DetailApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new DetailApiClient(mockSp);
    });

    describe('createDetail', () => {
        it('calls createListItem', async () => {
            await client.createDetail('active', 1);
            expect(mockSp.createListItem).toHaveBeenCalled();
        });

        it('returns true when createListItem succeeds', async () => {
            const result = await client.createDetail('active', 1);
            expect(result).toBe(true);
        });

        it('returns false when createListItem fails', async () => {
            (mockSp.createListItem as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.createDetail('active', 1);
            expect(result).toBe(false);
        });

        it('includes channel name and status in body', async () => {
            await client.createDetail('pending', 2);
            const [body] = (mockSp.createListItem as jest.Mock).mock.calls[0];
            expect(body.Title).toBe('TestChannel');
            expect(Object.values(body)).toContain('pending');
        });
    });

    describe('getLoanDetail', () => {
        it('calls spGet with channel ID in filter', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.getLoanDetail();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('channel-123')
            );
        });

        it('returns undefined when data array is empty', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: [] });
            const result = await client.getLoanDetail();
            expect(result).toBeUndefined();
        });

        it('returns undefined when data is null', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getLoanDetail();
            expect(result).toBeUndefined();
        });

        it('uses absolute URL in query', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            await client.getLoanDetail();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('fetches IUP link and returns mapped detail when data is present', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 1, Title: 'Test' }] })
                .mockResolvedValueOnce({ ok: true, data: [{ ContactId: 42 }] });

            const result = await client.getLoanDetail();

            expect(mockSp.getTeamName).toHaveBeenCalled();
            expect(result).toBeDefined();
        });

        it('returns mapped detail with empty IUP link when teamName is absent', async () => {
            (mockSp.getTeamName as jest.Mock).mockResolvedValue('');
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({ ok: true, data: [{ Id: 1, Title: 'Test' }] })
                .mockResolvedValueOnce({ ok: true, data: [] });

            const result = await client.getLoanDetail();

            expect(result).toBeDefined();
        });

        it('parses contact JSON before mapping loan detail', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{
                        Id: 1,
                        Title: 'Test',
                        currentStatus: 'Stage1/Phase1',
                        loanText: 'Loan text',
                        ChannelStatus: 'Active',
                        contactData: JSON.stringify({ id: 'abc', phone: '123', desc: 'Sales' }),
                        secondaryData: JSON.stringify({ id: 'def', phone: '456', desc: 'BackOffice' }),
                        version: 3,
                    }]
                })
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{
                        contact: {
                            Name: 'i:0#.f|membership|user@domain.com',
                            EMail: 'user@domain.com',
                            FirstName: 'John',
                            LastName: 'Doe',
                            Title: 'John Doe',
                        },
                        secondaryContact: {
                            Name: 'i:0#.f|membership|secondary@domain.com',
                            EMail: 'secondary@domain.com',
                            FirstName: 'Jane',
                            LastName: 'Doe',
                            Title: 'Jane Doe',
                        }
                    }]
                });

            const result = await client.getLoanDetail();

            expect(result).toEqual(expect.objectContaining({
                id: 1,
                title: 'Test',
                currentStatus: 'Stage1',
                currentFolder: 'Phase1',
                text: 'Loan text',
                version: 3,
                channelStatus: ChannelStatus.Active,
                contactDesc: ContactDesc.Sales,
                secondaryDesc: ContactDesc.BackOffice,
            }));
            expect(result?.contact).toEqual(expect.objectContaining({
                email: 'user@domain.com',
                phone: '123',
                id: 'abc',
                userType: UserType.Internal,
            }));
            expect(result?.secondaryContact).toEqual(expect.objectContaining({
                email: 'secondary@domain.com',
                phone: '456',
                id: 'def',
                userType: UserType.Internal,
            }));
        });

        it('falls back to empty contact data when contact JSON is malformed', async () => {
            (mockSp.spGet as jest.Mock)
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{
                        Id: 1,
                        Title: 'Test',
                        currentStatus: 'Stage1/Phase1',
                        ChannelStatus: 'Active',
                        contactData: '{bad-json',
                        secondaryData: null,
                        version: 3,
                    }]
                })
                .mockResolvedValueOnce({
                    ok: true,
                    data: [{
                        contact: {
                            Name: 'i:0#.f|membership|user@domain.com',
                            EMail: 'user@domain.com',
                            FirstName: 'John',
                            LastName: 'Doe',
                            Title: 'John Doe',
                        }
                    }]
                });

            const result = await client.getLoanDetail();

            expect(result).toBeDefined();
            expect(result?.contactDesc).toBeUndefined();
            expect(result?.contact).toEqual(expect.objectContaining({
                email: 'user@domain.com',
                phone: '',
                id: '',
            }));
        });
    });

    describe('updateDetail', () => {
        it('calls updateListItem with item ID and body', async () => {
            await client.updateDetail(12, { Status: 'active' });
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                12,
                { Status: 'active' },
                expect.any(String)
            );
        });

        it('returns true when updateListItem returns true', async () => {
            const result = await client.updateDetail(12, {});
            expect(result).toBe(true);
        });

        it('returns false when updateListItem returns false', async () => {
            (mockSp.updateListItem as jest.Mock).mockResolvedValue(false);
            const result = await client.updateDetail(12, {});
            expect(result).toBe(false);
        });
    });

    describe('updateContact', () => {
        it('calls getUserSPId when email is provided', async () => {
            await client.updateContact(5, 'data', 'secdata', 'user@example.com');
            expect(mockSp.getUserSPId).toHaveBeenCalledWith('user@example.com');
        });

        it('does not call getUserSPId when email is not provided', async () => {
            await client.updateContact(5, 'data', 'secdata');
            expect(mockSp.getUserSPId).not.toHaveBeenCalled();
        });

        it('calls updateListItem with item ID', async () => {
            await client.updateContact(5, 'data', 'secdata');
            expect(mockSp.updateListItem).toHaveBeenCalledWith(
                5,
                expect.any(Object),
                expect.any(String)
            );
        });

        it('returns true when updateListItem returns true', async () => {
            const result = await client.updateContact(5, 'data', 'secdata');
            expect(result).toBe(true);
        });
    });
});
