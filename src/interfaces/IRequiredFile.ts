import { IFileType } from ".";
import { FileStatus } from "../enums";
import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IRequiredFile {
    loc: ILocalizedStrings,
    //folderPath: string,
    id: number,
    description?: string;
    type?: IFileType;
    format?: string;
    uploadedFile?: string;
    deadline?: Date;
    status?: FileStatus;
}