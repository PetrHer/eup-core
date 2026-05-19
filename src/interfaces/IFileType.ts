import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IFileType {
    loc: ILocalizedStrings,
    // children?: IFileType[],
    id: number,
    uploadStatus?: string,
    permissions?: string[],
    folderPath: string,
    folderPathAfter?: string,
    validityStatuses: string[],
    folderPathInternal?: string,
    folderPathInternalAfter?: string,
}