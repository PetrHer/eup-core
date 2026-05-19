/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import CalendarApiClient from './CalendarApiClient';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        createListItemUsingPath: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as SPApiClient;
}

describe('CalendarApiClient', () => {
    let client: CalendarApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new CalendarApiClient(mockSp);
    });

    describe('getCalendarEvents', () => {
        it('calls spGet with channel name in URL', async () => {
            await client.getCalendarEvents();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.getCalendarEvents();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getCalendarEvents();
            expect(result).toEqual([]);
        });

        it('returns mapped array from API data', async () => {
            const result = await client.getCalendarEvents();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('createCalendarEvent', () => {
        const eventPayload = {
            subject: 'Meeting',
            startDate: new Date('2024-01-15'),
            endDate: new Date('2024-01-16'),
            type: 'Type1',
        };

        it('calls createListItemUsingPath', async () => {
            await client.createCalendarEvent(eventPayload as any);
            expect(mockSp.createListItemUsingPath).toHaveBeenCalled();
        });

        it('returns true when createListItemUsingPath succeeds', async () => {
            const result = await client.createCalendarEvent(eventPayload as any);
            expect(result).toBe(true);
        });

        it('returns false when createListItemUsingPath fails', async () => {
            (mockSp.createListItemUsingPath as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.createCalendarEvent(eventPayload as any);
            expect(result).toBe(false);
        });

        it('includes event subject in formValues', async () => {
            await client.createCalendarEvent(eventPayload as any);
            const [body] = (mockSp.createListItemUsingPath as jest.Mock).mock.calls[0];
            const titleField = body.formValues.find((f: any) => f.FieldName === 'Title');
            expect(titleField?.FieldValue).toBe('Meeting');
        });
    });
});
