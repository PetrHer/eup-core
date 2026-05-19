import { FolderStatus } from "../enums";
import { IFile } from "./IFile";
import { ILocalizedStrings } from "./ILocalizedStrings";
import { IRequiredFile } from "./IRequiredFile";

export interface IStructureNode {
    uploadedFiles: IFile[];
    requiredFiles?: IRequiredFile[];
    nodes: IStructureNode[];
    //folderPath: string;
    name: string;
    description?: ILocalizedStrings;
    //templateId?: number;
    //uniqueId?: string;
    //sequenceNumber?: number;
    status: FolderStatus;
    localization?: ILocalizedStrings;
    isCategory?: boolean;
    id?: number;
    //deadline?: Date;
    order?: number;
}