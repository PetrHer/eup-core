/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { parseQnAFilesData } from './parser';

describe('parseQnAFilesData', () => {
    it('returns undefined for undefined input', () => {
        expect(parseQnAFilesData(undefined)).toBeUndefined();
    });

    it('returns undefined for empty string', () => {
        expect(parseQnAFilesData('')).toBeUndefined();
    });

    it('returns undefined for invalid JSON', () => {
        expect(parseQnAFilesData('not-json')).toBeUndefined();
    });

    it('returns undefined when parsed value is not an object', () => {
        expect(parseQnAFilesData('"string"')).toBeUndefined();
    });

    it('returns undefined when a key value is not an array', () => {
        const invalid = JSON.stringify({ group1: 'not-an-array' });
        expect(parseQnAFilesData(invalid)).toBeUndefined();
    });

    it('returns undefined when array items are missing required fields', () => {
        const invalid = JSON.stringify({ group1: [{ id: 1 }] });
        expect(parseQnAFilesData(invalid)).toBeUndefined();
    });

    it('returns parsed object for valid QnAFilesData', () => {
        const valid = JSON.stringify({
            Question: [{ name: 'file.pdf', serverRelativeUrl: '/sites/team/file.pdf', inFolder: true }],
        });
        const result = parseQnAFilesData(valid);
        expect(result).toBeDefined();
        expect(result?.Question?.[0].name).toBe('file.pdf');
    });

    it('returns parsed object for empty groups object', () => {
        const valid = JSON.stringify({});
        const result = parseQnAFilesData(valid);
        expect(result).toBeDefined();
    });

    it('accepts items without inFolder field (only name and serverRelativeUrl required)', () => {
        const valid = JSON.stringify({
            Answer: [{ name: 'doc.docx', serverRelativeUrl: '/sites/team/doc.docx' }],
        });
        const result = parseQnAFilesData(valid);
        expect(result).toBeDefined();
    });
});
