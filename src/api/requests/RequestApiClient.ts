import { FileInfo } from "@syncfusion/ej2-inputs";
import { z } from "zod";

import { IRequestApiClient } from './IRequestApiClient';
import { ColumnsReqTypes, ColumnsRequests, ListName, RequestStatus } from "../../enums";
import { IFile, IFileInput, IFileType, IRequest, IRequestData, IRequestOption, IStructureNode } from "../../interfaces";
import { getFileExtension } from "../../utils/utils";
import FileApiClient from '../files/FileApiClient';
import IApiResult from "../IApiResult";
import { mapRequestOptions, mapRequests } from '../mappers';
import SPApiClient from '../SPApiClient';

const ParsedFormDataSchema = z.record(z.unknown());

const normalizeRequestFormData = (value?: unknown): string => {
    if (!value || typeof value !== 'string') {
        return '{}';
    }

    try {
        const parsed = JSON.parse(value);
        const parsedFormData = ParsedFormDataSchema.safeParse(parsed);
        return parsedFormData.success ? value : '{}';
    } catch {
        return '{}';
    }
};

const RequestSchema = z.object({
    Id: z.number(),
    Created: z.string(),
    [ColumnsRequests.RequestType]: z.object({
        Id: z.number(),
        Title: z.string(),
        [ColumnsReqTypes.TitleEn]: z.string(),
        [ColumnsReqTypes.Notification]: z.string(),
    }),
    [ColumnsRequests.RequestStatus]: z.nativeEnum(RequestStatus),
    [ColumnsRequests.FormData]: z.string().nullable().optional().transform(normalizeRequestFormData),
    [ColumnsRequests.FileTemplate]: z.string().nullable().optional(),
    [ColumnsRequests.FileIds]: z.string().nullable().optional(),
    [ColumnsRequests.CategoryId]: z.string().nullable().optional(),
});

const RequestsSchema = z.array(RequestSchema);

const parseRequest = (value?: unknown): z.infer<typeof RequestSchema> | undefined => {
    if (!value) {
        return undefined;
    }
    const parsedRequest = RequestSchema.safeParse(value);
    return parsedRequest.success ? parsedRequest.data : undefined;
};

const parseRequests = (value?: unknown): z.infer<typeof RequestsSchema> => {
    if (!value || !Array.isArray(value)) {
        return [];
    }
    const items = value.map(item => parseRequest(item)).filter((item): item is z.infer<typeof RequestSchema> => item !== undefined);
    return items;
};


export default class RequestApiClient implements IRequestApiClient {

    constructor(private readonly files: FileApiClient, private readonly sp: SPApiClient) { }

    /**
         * create request in specific folder using /AddValidateUpdateItemUsingPath endpoint and attach files to it
         * @param data 
         * @param files 
         * @returns boolean
         */
    public async createRequest(data: IRequestData, files: FileInfo[], fileType: IFileType, secondStage: boolean, uploadedBy: string): Promise<{ result: IApiResult, failedUploads: string[] }> {
        const folderPath = secondStage && fileType.folderPathAfter ? fileType.folderPathAfter : fileType.folderPath;
        let fileTemplate = '';
        const fileIds: string[] = [];
        const failedUploads: string[] = [];
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const ext = getFileExtension(file.name);
            const today = new Date();
            const day = String(today.getDate()).padStart(2, "0");
            const month = String(today.getMonth() + 1).padStart(2, "0");
            const year = String(today.getFullYear()).slice(-2);
            const dateSuffix = `${day}${month}${year}`;
            const fileName = data.templateRef && i === 0 ? `${data.templateName}_${dateSuffix}.${ext}` : file.name;
            const uploadResult = await this.files.uploadFile(folderPath, { ...file, name: fileName }, uploadedBy, undefined, fileType.id, fileType.uploadStatus);
            if (uploadResult.file) {
                if (data.templateRef && i === 0) {
                    fileTemplate = uploadResult.file.serverRelativeUrl;
                } else {
                    fileIds.push(uploadResult.file.uniqueId);
                }
            }
            if (!uploadResult.success) {
                failedUploads.push(file.name);
            }
        }

        const body = {
            "formValues": [
                {
                    "FieldName": "Title",
                    "FieldValue": data.templateName
                },
                {
                    "FieldName": ColumnsRequests.FormData,
                    "FieldValue": JSON.stringify(data.formData)
                },
                {
                    "FieldName": ColumnsRequests.RequestType,
                    "FieldValue": data.typeId?.toString()
                },
                {
                    "FieldName": ColumnsRequests.RequestStatus,
                    "FieldValue": RequestStatus.New
                },
                {
                    "FieldName": ColumnsRequests.FileTemplate,
                    "FieldValue": fileTemplate
                },
                {
                    "FieldName": ColumnsRequests.FileIds,
                    "FieldValue": fileIds.join(';')
                },
                {
                    "FieldName": ColumnsRequests.ExportData,
                    "FieldValue": data.exportData ? data.exportData : ''
                },
            ],
        };
        const result = await this.sp.createListItemUsingPath(body, ListName.Requests);
        return { result, failedUploads };
    }

    /**
     * loads request from SP list from folder with name of channel
     * @returns array of {@link IRequest}
     */
    public async getRequests(): Promise<IRequest[]> {
        const folderPath = `${this.sp.getRelativeUrl()}/lists/${ListName.Requests}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${folderPath}'`;
        const select = `*,${ColumnsRequests.RequestType}/Id,${ColumnsRequests.RequestType}/Title,${ColumnsRequests.RequestType}/${ColumnsReqTypes.TitleEn},${ColumnsRequests.RequestType}/${ColumnsReqTypes.Notification}`;
        const expand = `${ColumnsRequests.RequestType}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Requests}')/items?&$select=${select}&$filter=${filter}&$expand=${expand}`;
        const response = await this.sp.spGet(apiUrl);
        const data = parseRequests(response.data);
        return mapRequests(data);
    }

    /**
     * Loads all request options from the SharePoint list specified by {@link ListName.RequestTypes}.
     * @returns A promise that resolves to an array of {@link IRequestOption} objects.
     */
    public async getRequestsOptions(): Promise<IRequestOption[]> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.RequestTypes}')/items`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) ? mapRequestOptions(result.data) : [];
    }

    /**
         * update request and attach new files to it
         * @param data 
         * @param itemId 
         * @param files 
         * @returns boolean
         */
    public async updateRequest(body: object, itemId: number, uploadedBy?: string, secondStage?: boolean, files?: IFileInput[], category?: IStructureNode): Promise<{ files: IFile[], result: boolean }> {
        const uploadedFiles: IFile[] = [];
        if (files) {
            const result = await this.files.uploadFiles(files, !!secondStage, uploadedBy, category?.id);
            const fileIds = result.map(f => f.file?.uniqueId);
            for (const item of result) {
                if (item.file) {
                    uploadedFiles.push(item.file);
                }
            }
            const fileIdsString = fileIds.join(';');
            (body as any)[ColumnsRequests.FileIds] = (body as any)[ColumnsRequests.FileIds] ? (body as any)[ColumnsRequests.FileIds] + ';' + fileIdsString : fileIdsString;
        }
        const result = await this.sp.updateListItem(itemId, body, ListName.Requests);
        return { files: uploadedFiles, result };
    }

}