/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapRequestOptions, mapRequests } from './requestMapper';
import { Language } from '../../enums';

const rawOption = {
    Id: 1,
    Title: 'Option A',
    'Title_en': 'Option A EN',
    FileTemplate: '/path/a.docx;/path/b.docx',
    FileTemplateEn: '/path/a-en.docx;/path/b-en.docx',
    Description: 'Desc CS',
    'Description_en': 'Desc EN',
    Etapa: 'Stage1',
    Fields: [{ name: 'field1' }],
    FileAttachments: true,
    Notification: ['group1'],
    FileReply: true,
};

const rawRequest = {
    Id: 10,
    Created: '2024-03-01T00:00:00Z',
    RequestType: {
        Id: 1,
        Title: 'Req Type A',
        'Title_en': 'Req Type A EN',
        Notification: ['g1'],
    },
    RequestStatus: 'new',
    FormData: '{"field":"value"}',
    FileTemplate: '/template.docx',
    categoryId: 99,
    FileIds: 'uid1;uid2',
};

describe('mapRequestOptions', () => {
    it('returns empty array for empty input', () => {
        expect(mapRequestOptions([])).toEqual([]);
    });

    it('maps id and type from Id and Title', () => {
        const [result] = mapRequestOptions([rawOption]);
        expect(result.id).toBe(1);
        expect(result.type).toBe('Option A');
    });

    it('maps localization', () => {
        const [result] = mapRequestOptions([rawOption]);
        expect(result.loc[Language.CS]).toBe('Option A');
        expect(result.loc[Language.EN]).toBe('Option A EN');
    });

    it('maps desc from Description fields', () => {
        const [result] = mapRequestOptions([rawOption]);
        expect(result.desc?.[Language.CS]).toBe('Desc CS');
        expect(result.desc?.[Language.EN]).toBe('Desc EN');
    });

    it('splits semicolon-delimited FileTemplate into fileTemplates array', () => {
        const [result] = mapRequestOptions([rawOption]);
        expect(result.fileTemplates).toHaveLength(2);
        expect(result.fileTemplates[0][Language.CS]).toBe('/path/a.docx');
        expect(result.fileTemplates[0][Language.EN]).toBe('/path/a-en.docx');
    });

    it('returns empty fileTemplates array when FileTemplate is absent', () => {
        const [result] = mapRequestOptions([{ ...rawOption, FileTemplate: undefined }]);
        expect(result.fileTemplates).toEqual([]);
    });

    it('maps multiple options', () => {
        const result = mapRequestOptions([rawOption, { ...rawOption, Id: 2, Title: 'Option B' }]);
        expect(result).toHaveLength(2);
    });
});

describe('mapRequests', () => {
    it('returns empty array for empty input', () => {
        expect(mapRequests([])).toEqual([]);
    });

    it('maps id and status', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.id).toBe(10);
        expect(result.status).toBe('new');
    });

    it('maps type from RequestType object', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.type.id).toBe(1);
        expect(result.type.type).toBe('Req Type A');
        expect(result.type.loc[Language.CS]).toBe('Req Type A');
    });

    it('parses FormData JSON string', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.formData).toEqual({ field: 'value' });
    });

    it('maps created as a Date', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.created).toBeInstanceOf(Date);
    });

    it('splits FileIds by semicolon', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.fileIds).toEqual(['uid1', 'uid2']);
    });

    it('returns empty fileIds when FileIds is absent', () => {
        const [result] = mapRequests([{ ...rawRequest, FileIds: undefined }]);
        expect(result.fileIds).toEqual([]);
    });

    it('maps templateRef and categoryId', () => {
        const [result] = mapRequests([rawRequest]);
        expect(result.templateRef).toBe('/template.docx');
        expect(result.categoryId).toBe(99);
    });

    it('maps multiple requests', () => {
        const result = mapRequests([rawRequest, { ...rawRequest, Id: 11 }]);
        expect(result).toHaveLength(2);
    });
});
