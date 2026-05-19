import { FieldValues } from "react-hook-form";

import { FormStatus, FormType, Language } from "../enums";
import { IFile } from "./IFile";
import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IFormRecord {
    id: number;
    title: string;
    formData: FieldValues;
    status: FormStatus;
    localization: ILocalizedStrings;
    categoryId?: number;
    type: FormType;
    language?: Language;
    completed?: string;
    fileId?: number;
    file?: IFile
}