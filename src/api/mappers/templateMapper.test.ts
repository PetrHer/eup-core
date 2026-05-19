/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapTemplates, mapTemplateReferences, mapFolderTemplates } from './templateMapper';

const rawTemplate = {
    Id: 1,
    Title: 'Category A',
    Stage: 'Stage1',
    Phase: 'Phase1',
    FolderDescription: 'Description',
    'FolderDescription_en': 'Description EN',
    'Stage_en': 'Stage1 EN',
    'Category_en': 'Category A EN',
    'Phase_en': 'Phase1 EN',
    CategoryOrder: 3,
};

describe('mapTemplates', () => {
    it('returns empty array for empty input', () => {
        expect(mapTemplates([])).toEqual([]);
    });

    it('maps id from item.Id', () => {
        const [result] = mapTemplates([rawTemplate]);
        expect(result.id).toBe(1);
    });

    it('maps stage, phase, category from respective fields', () => {
        const [result] = mapTemplates([rawTemplate]);
        expect(result.stage).toBe('Stage1');
        expect(result.phase).toBe('Phase1');
        expect(result.category).toBe('Category A');
    });

    it('maps EN fields with fallbacks to CS when absent', () => {
        const [result] = mapTemplates([{ ...rawTemplate, 'Stage_en': undefined, 'Category_en': undefined, 'Phase_en': undefined }]);
        expect(result.stageEn).toBe('Stage1');
        expect(result.categoryEn).toBe('Category A');
        expect(result.phaseEn).toBe('Phase1');
    });

    it('maps description and descriptionEn', () => {
        const [result] = mapTemplates([rawTemplate]);
        expect(result.description).toBe('Description');
        expect(result.descriptionEn).toBe('Description EN');
    });

    it('maps multiple items', () => {
        const result = mapTemplates([rawTemplate, { ...rawTemplate, Id: 2, Title: 'Category B' }]);
        expect(result).toHaveLength(2);
    });
});

describe('mapTemplateReferences', () => {
    it('returns empty array for empty input', () => {
        expect(mapTemplateReferences([])).toEqual([]);
    });

    it('maps fileRef and fileLeafRef', () => {
        const [result] = mapTemplateReferences([{ FileRef: '/sites/team/file.docx', FileLeafRef: 'file.docx' }]);
        expect(result.fileRef).toBe('/sites/team/file.docx');
        expect(result.fileLeafRef).toBe('file.docx');
    });

    it('maps multiple references', () => {
        const result = mapTemplateReferences([
            { FileRef: '/a.docx', FileLeafRef: 'a.docx' },
            { FileRef: '/b.docx', FileLeafRef: 'b.docx' },
        ]);
        expect(result).toHaveLength(2);
    });
});

describe('mapFolderTemplates', () => {
    it('returns empty array for empty input', () => {
        expect(mapFolderTemplates([])).toEqual([]);
    });

    it('maps top-level folders (no parent in data)', () => {
        const data = [
            { FileRef: '/root/folderA', FileLeafRef: 'folderA', FileDirRef: '/root', Name_en: 'Folder A EN', WritePermissions: 'w', ReadPermissions: 'r' },
        ];
        const result = mapFolderTemplates(data);
        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('folderA');
        expect(result[0].nameEn).toBe('Folder A EN');
    });

    it('nests child folders under their parent', () => {
        const data = [
            { FileRef: '/root/parent', FileLeafRef: 'parent', FileDirRef: '/root', Name_en: 'Parent EN' },
            { FileRef: '/root/parent/child', FileLeafRef: 'child', FileDirRef: '/root/parent', Name_en: 'Child EN' },
        ];
        const result = mapFolderTemplates(data);
        expect(result).toHaveLength(1);
        expect(result[0].templates).toHaveLength(1);
        expect(result[0].templates?.[0].name).toBe('child');
    });

    it('returns empty templates array for leaf folders', () => {
        const data = [
            { FileRef: '/root/leaf', FileLeafRef: 'leaf', FileDirRef: '/root', Name_en: 'Leaf EN' },
        ];
        const result = mapFolderTemplates(data);
        expect(result[0].templates).toEqual([]);
    });

    it('maps write and read permissions', () => {
        const data = [
            { FileRef: '/root/f', FileLeafRef: 'f', FileDirRef: '/root', Name_en: 'F', WritePermissions: 'UUO_RM', ReadPermissions: 'OAO_CC' },
        ];
        const result = mapFolderTemplates(data);
        expect(result[0].write).toBe('UUO_RM');
        expect(result[0].read).toBe('OAO_CC');
    });
});
