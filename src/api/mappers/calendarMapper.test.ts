/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapCalendarEvents } from './calendarMapper';
import { EventType } from '../../enums';

describe('mapCalendarEvents', () => {
    const rawItem = {
        Id: 1,
        Title: 'Event A',
        UniqueId: 'uid-1',
        Date: '2024-06-01T10:00:00Z',
        EventType: EventType.Covenant,
        Amount: 1000,
        Currency: 'CZK',
        Instruction: 'Do something',
        Category: 'Cat1',
        EventDescription: 'Desc',
        Note: 'Note1',
        Period: 'monthly',
        EndDate: '2024-12-31T00:00:00Z',
    };

    it('returns an empty array for empty input', () => {
        expect(mapCalendarEvents([])).toEqual([]);
    });

    it('maps id and subject from Id and Title', () => {
        const [result] = mapCalendarEvents([rawItem]);
        expect(result.id).toBe(1);
        expect(result.subject).toBe('Event A');
        expect(result.uniqueId).toBe('uid-1');
    });

    it('maps Date to a Date object for time and startDate', () => {
        const [result] = mapCalendarEvents([rawItem]);
        expect(result.time).toBeInstanceOf(Date);
        expect(result.startDate).toBeInstanceOf(Date);
    });

    it('maps EventType correctly', () => {
        const [result] = mapCalendarEvents([rawItem]);
        expect(result.type).toBe(EventType.Covenant);
    });

    it('maps EndDate to a Date object', () => {
        const [result] = mapCalendarEvents([rawItem]);
        expect(result.endDate).toBeInstanceOf(Date);
    });

    it('sets time and startDate to undefined when Date is absent', () => {
        const [result] = mapCalendarEvents([{ ...rawItem, Date: undefined }]);
        expect(result.time).toBeUndefined();
        expect(result.startDate).toBeUndefined();
    });

    it('sets endDate to undefined when EndDate is absent', () => {
        const [result] = mapCalendarEvents([{ ...rawItem, EndDate: undefined }]);
        expect(result.endDate).toBeUndefined();
    });

    it('maps amount, currency, and note fields', () => {
        const [result] = mapCalendarEvents([rawItem]);
        expect(result.amount).toBe(1000);
        expect(result.currency).toBe('CZK');
        expect(result.note).toBe('Note1');
    });

    it('maps multiple items', () => {
        const result = mapCalendarEvents([rawItem, { ...rawItem, Id: 2, Title: 'Event B' }]);
        expect(result).toHaveLength(2);
        expect(result[1].subject).toBe('Event B');
    });
});
