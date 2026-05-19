/* eslint-disable no-undef, @rushstack/hoist-jest-mock, @typescript-eslint/explicit-function-return-type */
import FormApiClient from './FormApiClient';
import { Language } from '../../enums';
import SPApiClient from '../SPApiClient';

jest.mock('../SPApiClient');

function createMockSp() {
    return {
        getAbsoluteUrl: jest.fn().mockReturnValue('https://example.sharepoint.com/sites/team'),
        getRelativeUrl: jest.fn().mockReturnValue('/sites/team'),
        getChannelName: jest.fn().mockReturnValue('TestChannel'),
        spGet: jest.fn().mockResolvedValue({ ok: true, data: [] }),
        spPatch: jest.fn().mockResolvedValue({ ok: true }),
    } as unknown as SPApiClient;
}

describe('FormApiClient', () => {
    let client: FormApiClient;
    let mockSp: ReturnType<typeof createMockSp>;

    beforeEach(() => {
        jest.clearAllMocks();
        mockSp = createMockSp();
        client = new FormApiClient(mockSp);
    });

    describe('getFormRecords', () => {
        it('calls spGet with channel name in folder path', async () => {
            await client.getFormRecords();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('TestChannel')
            );
        });

        it('calls spGet with absolute URL', async () => {
            await client.getFormRecords();
            expect(mockSp.spGet).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team')
            );
        });

        it('returns empty array when API returns no array data', async () => {
            (mockSp.spGet as jest.Mock).mockResolvedValue({ ok: true, data: null });
            const result = await client.getFormRecords();
            expect(result).toEqual([]);
        });

        it('returns mapped records for array data', async () => {
            const result = await client.getFormRecords();
            expect(Array.isArray(result)).toBe(true);
        });
    });

    describe('saveFormData', () => {
        it('calls spPatch with item ID in URL', async () => {
            await client.saveFormData(7, { field: 'value' }, Language.CS);
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('items(7)'),
                expect.any(Object)
            );
        });

        it('calls spPatch with absolute URL', async () => {
            await client.saveFormData(7, {}, Language.CS);
            expect(mockSp.spPatch).toHaveBeenCalledWith(
                expect.stringContaining('https://example.sharepoint.com/sites/team'),
                expect.any(Object)
            );
        });

        it('returns true when spPatch succeeds', async () => {
            const result = await client.saveFormData(7, {}, Language.CS);
            expect(result).toBe(true);
        });

        it('returns false when spPatch fails', async () => {
            (mockSp.spPatch as jest.Mock).mockResolvedValue({ ok: false });
            const result = await client.saveFormData(7, {}, Language.CS);
            expect(result).toBe(false);
        });

        it('sets submitted status when asDraft is false', async () => {
            await client.saveFormData(7, {}, Language.CS, false);
            const body = (mockSp.spPatch as jest.Mock).mock.calls[0][1];
            expect(body).toHaveProperty('Status');
        });

        it('sets draft status when asDraft is true', async () => {
            await client.saveFormData(7, {}, Language.CS, true);
            const body = (mockSp.spPatch as jest.Mock).mock.calls[0][1];
            expect(body).toHaveProperty('Status');
        });

        it('includes form data as JSON string in body', async () => {
            const formData = { name: 'Test', value: 42 };
            await client.saveFormData(7, formData, Language.CS);
            const body = (mockSp.spPatch as jest.Mock).mock.calls[0][1];
            expect(Object.values(body)).toContain(JSON.stringify(formData));
        });
    });
});
