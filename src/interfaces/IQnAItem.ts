import { IQnAFilesData, ISPComment } from ".";
import { QACategory, QAPriority } from "../enums";

export interface IQnAItem {
    question: string,
    answer?: string,
    comments: ISPComment[];
    uniqueId: string;
    id: number;
    status: string;
    category?: QACategory;
    phase: string;
    // fileLink: string;
    // folderLink: string;
    // questionFileLink: string;
    // questionFolderLink: string;
    order?: number;
    stage: string;
    createdText: string;
    answeredText?: string;
    priority?: QAPriority;
    created: Date;
    answeredUser: string;
    answeredDate: string;
    createdUser: string;
    createdDate: string;
    filesData?: IQnAFilesData;
    subCategory?: string;
    modified: string;
}