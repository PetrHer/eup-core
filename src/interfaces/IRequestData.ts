import { FieldValues } from "react-hook-form";

import { RequestStatus, RequestType } from "../enums";

export interface IRequestData {
    // name: string;
    type?: RequestType;
    typeId?: number;
    // description: string;
    status: RequestStatus;
    templateRef?: boolean;
    templateName?: string;
    formData?: FieldValues;
    exportData?: string;
}