import { QAItemType } from "../enums";
import { IQnAAttachment } from "./IQnAAttachment";

export interface IQnAFilesData extends Partial<{
    [K in QAItemType]: IQnAAttachment[];
}> { }