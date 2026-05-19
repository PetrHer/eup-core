import { FieldValues } from "react-hook-form";

import { IRequestOption } from ".";
import { RequestStatus } from "../enums";

export interface IRequest {
    status: RequestStatus;
    created: Date;
    id: number;
    //attachments: IAttachment[];
    type: IRequestOption;
    templateRef?: string;
    formData: FieldValues;
    categoryId?: string
    fileIds: string[]
}