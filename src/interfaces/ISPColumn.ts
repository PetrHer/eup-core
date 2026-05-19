import { ListName } from "../enums";

export interface ISPColumn {
    FieldTypeKind: string;
    Title: string;
    Choices?: string[];
    lookUpListTitle?: ListName;
    displayName: string;
    allowMultiple?: boolean;
}