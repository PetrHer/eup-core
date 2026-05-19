import { FileInfo } from '@syncfusion/ej2-inputs';

import { IFile, IFileInput, IFileType, IRequest, IRequestData, IRequestOption, IStructureNode } from '../../interfaces';
import IApiResult from '../IApiResult';

export interface IRequestApiClient {
    createRequest(data: IRequestData, files: FileInfo[], fileType: IFileType, secondStage: boolean, uploadedBy: string): Promise<{ result: IApiResult; failedUploads: string[] }>;
    getRequests(): Promise<IRequest[]>;
    getRequestsOptions(): Promise<IRequestOption[]>;
    updateRequest(body: object, itemId: number, uploadedBy?: string, secondStage?: boolean, files?: IFileInput[], category?: IStructureNode): Promise<{ files: IFile[]; result: boolean }>;
}
