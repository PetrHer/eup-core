import { z } from "zod";

import { IStructureApiClient } from './IStructureApiClient';
import { ColumnsCategories, FolderStatus, ListName } from '../../enums';
import { IStructureNode } from '../../interfaces';
import { mapStructure } from '../mappers';
import SPApiClient from "../SPApiClient";

const StructureCategoryItemSchema = z.object({
    Id: z.number(),
    Title: z.string(),
    [ColumnsCategories.Stage]: z.string(),
    [ColumnsCategories.Phase]: z.string(),
    [ColumnsCategories.StageEn]: z.string().nullable().optional(),
    [ColumnsCategories.PhaseEn]: z.string().nullable().optional(),
    [ColumnsCategories.CategoryEn]: z.string().nullable().optional(),
    [ColumnsCategories.Description]: z.string().nullable().optional(),
    [ColumnsCategories.DescriptionEn]: z.string().nullable().optional(),
    [ColumnsCategories.CategoryStatus]: z.nativeEnum(FolderStatus).nullable().optional(),
    [ColumnsCategories.Order]: z.number().nullable().optional(),
});

const StructureQnAPhaseItemSchema = z.object({
    Stage: z.string(),
    Phase: z.string(),
    Stage_en: z.string().optional(),
    Phase_en: z.string().optional(),
});

const parseStructureCategoryItems = (value?: unknown): z.infer<typeof StructureCategoryItemSchema>[] => {
    if (!value || !Array.isArray(value)) {
        return [];
    }
    return value
        .map(item => StructureCategoryItemSchema.safeParse(item))
        .filter((item): item is { success: true; data: z.infer<typeof StructureCategoryItemSchema> } => item.success)
        .map(item => item.data);
};

const parseStructureQnAPhaseItems = (value?: unknown): z.infer<typeof StructureQnAPhaseItemSchema>[] => {
    if (!value || !Array.isArray(value)) {
        return [];
    }

    return value
        .map(item => StructureQnAPhaseItemSchema.safeParse(item))
        .filter((item): item is { success: true; data: z.infer<typeof StructureQnAPhaseItemSchema> } => item.success)
        .map(item => item.data);
};

export default class StructureApiClient implements IStructureApiClient {
    constructor(private readonly sp: SPApiClient) { }

    /**
     * Retrieves all folders within the document structure for the current channel.
     * @returns An array of folders.
     */
    public async getStructure(): Promise<IStructureNode[]> {
        const categories = await this.getCategories();
        const phases = await this.getPhasesFromQnAList();
        return mapStructure(categories, phases);
    }

    /**
     * Retrieves categories from the SharePoint list filtered by the current channel's folder path.
     *
     * @returns A promise that resolves to an array of category items, or an empty array if none found.
     */
    private async getCategories(): Promise<any[]> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.Categories}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${fileDirRef}'`;
        const select = '*,FileDirRef,UniqueId';
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.Categories}')/items?$select=${select}&$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        return parseStructureCategoryItems(response.data);
    }

    /**
 * Retrieves phase and stage information from the QA SharePoint list filtered by the current channel's folder path.
 *
 * @returns A promise that resolves to an array of phase and stage items, or an empty array if none found.
 */
    private async getPhasesFromQnAList(): Promise<any[]> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.QA}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${fileDirRef}'`;
        const select = 'Stage,Phase,Stage_en,Phase_en';
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.QA}')/items?$select=${select}&$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        return parseStructureQnAPhaseItems(response.data);
    }
}