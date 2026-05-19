// import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IFolderTemplate {
    name: string;
    nameEn?: string;
    // description?: ILocalizedStrings;
    // templateId?: number;
    // uniqueId?: string;
    // sequenceNumber?: number;
    templates?: IFolderTemplate[];
    write?: string;
    read?: string;
}