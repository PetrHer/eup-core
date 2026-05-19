import { ColumnsReqTypes, ColumnsRequests, Language } from "../../enums";
import { IRequest, IRequestOption } from "../../interfaces";
import { parseFormData } from "../../utils/utils";

const mapFileTemplates = (csTemplates?: string, enTemplates?: string): { [Language.CS]?: string;[Language.EN]?: string; }[] => {
    const cs = csTemplates ? csTemplates.split(';') : [];
    const en = enTemplates ? enTemplates.split(';') : [];
    const result: { [Language.CS]?: string;[Language.EN]?: string; }[] = [];
    for (let i = 0; i < cs.length; i++) {
        result.push({ [Language.CS]: cs[i], [Language.EN]: en[i] });

    }
    return result;
};

const mapRequestOption = (item: any): IRequestOption => {
    return {
        id: item.Id,
        type: item.Title,
        fileTemplates: mapFileTemplates(item[ColumnsReqTypes.FileTemplate], item[ColumnsReqTypes.FileTemplateEn]),
        loc: {
            [Language.CS]: item.Title,
            [Language.EN]: item[ColumnsReqTypes.TitleEn]
        },
        desc: {
            [Language.CS]: item[ColumnsReqTypes.Description],
            [Language.EN]: item[ColumnsReqTypes.DescriptionEn]
        },
        stage: item[ColumnsReqTypes.Stage],
        fields: item[ColumnsReqTypes.Fields],
        fileAttachments: item[ColumnsReqTypes.FileAttachments],
        notifGroups: item[ColumnsReqTypes.Notification],
        fileReply: item[ColumnsReqTypes.FileReply]
    };
};

/**
 * Maps raw SharePoint list items to an array of {@link IRequestOption} objects.
 * @param data - Array of raw list items from SharePoint.
 * @returns An array of mapped {@link IRequestOption} objects.
 */
export const mapRequestOptions = (data: any[]): IRequestOption[] => {
    return data.map(item => mapRequestOption(item));
};

const mapRequest = (item: any,/* attachments: IAttachment[]*/): IRequest => {
    const createdDate = new Date(item.Created) ?? new Date();
    createdDate.setHours(0, 0, 0, 0);
    return {
        type: {
            id: item[ColumnsRequests.RequestType].Id,
            type: item[ColumnsRequests.RequestType].Title,
            loc: {
                [Language.CS]: item[ColumnsRequests.RequestType].Title,
                [Language.EN]: item[ColumnsRequests.RequestType][ColumnsReqTypes.TitleEn],
            },
            fileTemplates: [],
            notifGroups: item[ColumnsRequests.RequestType][ColumnsReqTypes.Notification],
        },
        status: item[ColumnsRequests.RequestStatus],
        created: createdDate,
        formData: parseFormData(item[ColumnsRequests.FormData] ?? '{}'),
        id: item.Id,
        //attachments: attachments,
        templateRef: item.FileTemplate,
        categoryId: item[ColumnsRequests.CategoryId],
        fileIds: item[ColumnsRequests.FileIds] ? item[ColumnsRequests.FileIds].split(';').filter((id: string) => id !== '') : []
    };
};

/**
 * map requests from sharepoint rest api to array of {@link IRequest}
 * @param data 
 * @returns array of {@link IRequest}
 */
export const mapRequests = (data: any, /*attachmentMap: Map<number, IAttachment[]>*/): IRequest[] => {
    return data.map((item: any) => {
        //const attachments = attachmentMap.get(item.Id) || [];
        return mapRequest(item,/* attachments*/);
    });
};