/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapStructure, mapCategoryNodes } from './structureMapper';
import { FolderStatus, Language } from '../../enums';

const rawCategory = {
    Id: 1,
    Title: 'Category A',
    Stage: 'Stage1',
    Phase: 'Phase1',
    'Stage_en': 'Stage1 EN',
    'Phase_en': 'Phase1 EN',
    'Category_en': 'Category A EN',
    FolderDescription: 'Desc CS',
    'FolderDescription_en': 'Desc EN',
    CategoryStatus: FolderStatus.Accepted,
    CategoryOrder: 2,
};

describe('mapCategoryNodes', () => {
    it('returns empty array for empty input', () => {
        expect(mapCategoryNodes([])).toEqual([]);
    });

    it('maps id and name from Id and Title', () => {
        const [result] = mapCategoryNodes([rawCategory]);
        expect(result.id).toBe(1);
        expect(result.name).toBe('Category A');
    });

    it('maps status from CategoryStatus', () => {
        const [result] = mapCategoryNodes([rawCategory]);
        expect(result.status).toBe(FolderStatus.Accepted);
    });

    it('defaults status to FolderStatus.Default when CategoryStatus absent', () => {
        const [result] = mapCategoryNodes([{ ...rawCategory, CategoryStatus: undefined }]);
        expect(result.status).toBe(FolderStatus.Default);
    });

    it('maps description localization', () => {
        const [result] = mapCategoryNodes([rawCategory]);
        expect(result.description?.[Language.CS]).toBe('Desc CS');
        expect(result.description?.[Language.EN]).toBe('Desc EN');
    });

    it('maps order from CategoryOrder', () => {
        const [result] = mapCategoryNodes([rawCategory]);
        expect(result.order).toBe(2);
    });

    it('sets isCategory to true', () => {
        const [result] = mapCategoryNodes([rawCategory]);
        expect(result.isCategory).toBe(true);
    });

    it('maps multiple items', () => {
        const result = mapCategoryNodes([rawCategory, { ...rawCategory, Id: 2, Title: 'Category B' }]);
        expect(result).toHaveLength(2);
    });
});

describe('mapStructure', () => {
    it('returns empty array when both inputs are empty', () => {
        expect(mapStructure([], [])).toEqual([]);
    });

    it('builds a stage > phase > category hierarchy', () => {
        const result = mapStructure([rawCategory], []);
        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('Stage1');
        expect(result[0].nodes).toHaveLength(1);
        expect(result[0].nodes[0].name).toBe('Phase1');
        expect(result[0].nodes[0].nodes).toHaveLength(1);
        expect(result[0].nodes[0].nodes[0].name).toBe('Category A');
    });

    it('groups multiple categories under the same stage and phase', () => {
        const cat2 = { ...rawCategory, Id: 2, Title: 'Category B' };
        const result = mapStructure([rawCategory, cat2], []);
        expect(result[0].nodes[0].nodes).toHaveLength(2);
    });

    it('creates separate stage nodes for different stages', () => {
        const cat2 = { ...rawCategory, Id: 2, Title: 'Category B', Stage: 'Stage2', Phase: 'Phase1' };
        const result = mapStructure([rawCategory, cat2], []);
        expect(result).toHaveLength(2);
    });

    it('adds phasesQnA phases to existing stages', () => {
        const phaseQnA = { Stage: 'Stage1', Phase: 'Phase2', 'Phase_en': 'Phase2 EN' };
        const result = mapStructure([rawCategory], [phaseQnA]);
        const stage = result.find(n => n.name === 'Stage1');
        expect(stage?.nodes).toHaveLength(2);
        expect(stage?.nodes.some(n => n.name === 'Phase2')).toBe(true);
    });

    it('ignores phasesQnA for stages that do not exist in categories', () => {
        const phaseQnA = { Stage: 'NonexistentStage', Phase: 'PhaseX', 'Phase_en': 'PhaseX EN' };
        const result = mapStructure([rawCategory], [phaseQnA]);
        expect(result).toHaveLength(1);
    });

    it('does not duplicate phases already present from categories', () => {
        const phaseQnA = { Stage: 'Stage1', Phase: 'Phase1', 'Phase_en': 'Phase1 EN' };
        const result = mapStructure([rawCategory], [phaseQnA]);
        const stage = result.find(n => n.name === 'Stage1');
        expect(stage?.nodes).toHaveLength(1);
    });

    it('sorts stage and phase nodes alphabetically', () => {
        const catB = { ...rawCategory, Id: 2, Title: 'Category B', Stage: 'StageB', Phase: 'Phase1' };
        const catA = { ...rawCategory, Id: 3, Title: 'Category C', Stage: 'StageA', Phase: 'Phase1' };
        const result = mapStructure([catB, catA], []);
        expect(result[0].name).toBe('StageA');
        expect(result[1].name).toBe('StageB');
    });
});
