/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapFormRecords } from './formMapper';
import { FormStatus, FormType, Language } from '../../enums';

describe('mapFormRecords', () => {
    const rawItem = {
        Id: 1,
        Title: 'Form A',
        'Title_en': 'Form A EN',
        FormData: '{"name":"Test"}',
        Status: FormStatus.Submitted,
        CategoryId: 5,
        FormType: FormType.KYC,
        Language: Language.CS,
        Completed: true,
        FileId: 7,
    };

    it('returns empty array for empty input', () => {
        expect(mapFormRecords([])).toEqual([]);
    });

    it('maps id and title', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.id).toBe(1);
        expect(result.title).toBe('Form A');
    });

    it('maps localization from Title and Title_en', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.localization[Language.CS]).toBe('Form A');
        expect(result.localization[Language.EN]).toBe('Form A EN');
    });

    it('falls back to Title when Title_en is absent', () => {
        const [result] = mapFormRecords([{ ...rawItem, 'Title_en': undefined }]);
        expect(result.localization[Language.EN]).toBe('Form A');
    });

    it('parses FormData JSON string', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.formData).toBeDefined();
    });

    it('maps status from Status field', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.status).toBe(FormStatus.Submitted);
    });

    it('maps type to FormType.Default when value is not a valid FormType', () => {
        const [result] = mapFormRecords([{ ...rawItem, FormType: 'INVALID' }]);
        expect(result.type).toBe(FormType.Default);
    });

    it('maps valid FormType correctly', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.type).toBe(FormType.KYC);
    });

    it('maps categoryId, completed, and fileId', () => {
        const [result] = mapFormRecords([rawItem]);
        expect(result.categoryId).toBe(5);
        expect(result.completed).toBe(true);
        expect(result.fileId).toBe(7);
    });

    it('maps language to undefined when value is invalid', () => {
        const [result] = mapFormRecords([{ ...rawItem, Language: 'INVALID' }]);
        expect(result.language).toBeUndefined();
    });

    it('maps multiple records', () => {
        const result = mapFormRecords([rawItem, { ...rawItem, Id: 2, Title: 'Form B' }]);
        expect(result).toHaveLength(2);
        expect(result[1].title).toBe('Form B');
    });
});
