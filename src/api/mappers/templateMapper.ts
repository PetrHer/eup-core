import { ColumnsCategories } from "../../enums";
import { IFolderTemplate, ITemplateContent, ITemplateReference } from "../../interfaces";

/**
 * Maps raw data to an array of template content objects.
 * @param data The raw data array.
 * @returns An array of mapped template content objects.
 */
export const mapTemplates = (data: any[]): ITemplateContent[] => {
    return data.map((item: any) => {
        return {
            stage: item[ColumnsCategories.Stage],
            phase: item[ColumnsCategories.Phase],
            category: item.Title,
            description: item[ColumnsCategories.Description],
            descriptionEn: item[ColumnsCategories.DescriptionEn],
            stageEn: item[ColumnsCategories.StageEn] ?? item[ColumnsCategories.Stage],
            categoryEn: item[ColumnsCategories.CategoryEn] ?? item.Title,
            phaseEn: item[ColumnsCategories.PhaseEn] ?? item[ColumnsCategories.Phase],
            //folder: item[ColumnsCategories.Folder],
            id: item.Id
        };
    });
};

export const mapTemplateReferences = (data: any[]): ITemplateReference[] => {
    return data.map(item => ({
        fileRef: item.FileRef,
        fileLeafRef: item.FileLeafRef
    }));
};

/**
 * Maps raw folder template data to a hierarchical structure of folder templates.
 * Filters out items that are children (folders within folders) to build a nested tree.
 *
 * @param data - Raw data array from the SharePoint API representing folder templates.
 * @returns An array of {@link IFolderTemplate} with nested children in `templates`.
 */
export const mapFolderTemplates = (data: any[]): IFolderTemplate[] => {
    return data
        .filter(item => !data.some(it => it.FileRef === item.FileDirRef))
        .map(item => {
            return {
                name: item.FileLeafRef,
                nameEn: item.Name_en,
                templates: mapFolderTemplates(data.filter(it => (it.FileDirRef as string).startsWith(item.FileRef))),
                write: item.WritePermissions,
                read: item.ReadPermissions,
            };
        });
};