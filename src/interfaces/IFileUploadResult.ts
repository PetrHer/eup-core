import { IFile } from "./IFile";

export interface IFileUploadResult {
    fileName: string;
    success: boolean;
    file?: IFile;
    typeId?: number;
}