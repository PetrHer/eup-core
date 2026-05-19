import { IFile } from "./IFile";
import { IRequiredFile } from "./IRequiredFile";

export interface IExtendedFile extends IFile {
    description?: string;
    format?: string;
    stage?: string;
    phase?: string;
    category?: string;
    reupload?: boolean;
    requiredFile?: IRequiredFile;
    deadline?: Date
}