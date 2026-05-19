import { FileInfo } from '@syncfusion/ej2-react-inputs';

import { FinancialDoc } from '../../enums';
import { IFile, IFileInput, IFileType, IFileUploadResult, IFileVersion, IFolder, IRequiredFile, ISiteGroup, IStructureNode, IWorkflowAction } from '../../interfaces';
import IApiResult from '../IApiResult';

export interface IFileApiClient {
    moveFileToBin(fileRef: string, fileName: string): Promise<boolean>;
    loadFiles(id: number): Promise<IFile[]>;
    loadFilesByIds(ids: number[]): Promise<IFile[]>;
    loadFilesByUniqueIds(uniqueIds: string[]): Promise<IFile[]>;
    loadFileByRelativeUrl(serverRelativeUrl: string): Promise<IFile | undefined>;
    loadAllFiles(documentation: boolean): Promise<IFile[]>;
    publishFile(uniqueId: string): Promise<boolean>;
    setFilePermissions(action: IWorkflowAction, siteGroups: ISiteGroup[], uniqueId: string): Promise<void>;
    uploadFiles(files: IFileInput[], secondStage: boolean, uploadedBy?: string, folderId?: number): Promise<IFileUploadResult[]>;
    getFileTypes(webUrl?: string): Promise<IFileType[]>;
    undoFileCheckOut(uniqueId: string): Promise<boolean>;
    getFileByUniqueId(uniqueId: string): Promise<IFile | undefined>;
    getFileVersionHistory(id: number): Promise<IFileVersion[]>;
    getFileVersion(relativeUrl: string, versionId: number): Promise<ArrayBuffer | undefined>;
    getFileToDownload(relativeUrl: string): Promise<ArrayBuffer | undefined>;
    archiveFile(file: IFile, stage?: IStructureNode, phase?: IStructureNode): Promise<boolean>;
    moveFile(fileRef: string, fileName: string, newFileDirRef: string): Promise<boolean>;
    updateFileByUniqueId(uniqueId: string, body?: object): Promise<IApiResult>;
    updateFileByUrl(fileRef: string, body?: object): Promise<IApiResult>;
    uploadFile(folderPath: string, file: FileInfo, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    reuploadFile(folderPath: string, file: FileInfo, fileName: string, uploadedBy?: string, folderId?: number, documentTypeId?: number, status?: string): Promise<IFileUploadResult>;
    getRequiredFiles(id: number): Promise<IRequiredFile[]>;
    getAllRequiredFiles(): Promise<Map<number, IRequiredFile[]>>;
    getFolders(): Promise<IFolder[]>;
    sendFiletoExtraction(file: IFile, type: FinancialDoc): Promise<boolean>;
}
