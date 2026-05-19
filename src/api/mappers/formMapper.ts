import { ColumnsForms, FormStatus, FormType, Language } from "../../enums";
import { IFormRecord } from "../../interfaces";
import { parseFormData } from "../../utils/utils";

/**
 * Maps raw data from the SharePoint list to an array of form records.
 * @param data - The raw data array from the SharePoint list
 * @returns IFormRecord[] - An array of form records after mapping the raw data
 */
export const mapFormRecords = (data: any[]): IFormRecord[] => {
    return data.map(item => {
        return {
            id: item.Id,
            title: item.Title,
            formData: parseFormData(item.FormData),
            status: item.Status as FormStatus ?? FormStatus.New,
            localization: {
                [Language.CS]: item.Title,
                [Language.EN]: item.Title_en ?? item.Title
            },
            categoryId: item[`${ColumnsForms.Category}Id`],
            type: Object.values(FormType).includes(item[ColumnsForms.FormType]) ? item[ColumnsForms.FormType] : FormType.Default,
            language: Object.values(Language).includes(item[ColumnsForms.Language]) ? item[ColumnsForms.Language] : undefined,
            completed: item[ColumnsForms.Completed],
            fileId: item[ColumnsForms.FileId]
        };
    });
};