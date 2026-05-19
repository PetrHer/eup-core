import { FileStatus } from "../enums";
import { IFileType } from "./IFileType";

/**
 * Represents a remaped file loaded from SharePoint.
 */
export interface IFile {
    /** The name of the file. */
    name: string;
    /** The current status of the file. */
    status: FileStatus;
    /** The URL or path to the file. */
    serverRelativeUrl: string;
    /** A unique identifier for the file. */
    uniqueId: string;
    version: string;
    type?: IFileType;
    created?: Date;
    id?: number;
    nodeId?: number;
    documentation?: boolean;
    provided?: string;
    contractValidity?: string;
    validTo?: Date;
    qaId?: string;
    author?: string;
    typeId?: number;
    modified?: Date;
}