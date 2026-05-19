import { ListName } from '../../enums';
import { IFileInput, IFileType, IQnAAttachment, IQnAFilesData, IQnAItem, ISPComment } from '../../interfaces';

export interface IQnAApiClient {
    createQnAItem(
        question: string,
        category: string,
        stage: string,
        phase: string,
        priority: string,
        uploadedBy: string,
        fileType?: IFileType,
        order?: number,
        files?: IFileInput[],
        subCategory?: string
    ): Promise<string | undefined>;
    addComment(
        qnAItem: IQnAItem,
        listName: string,
        text: string,
        files: IFileInput[],
        uploadedBy: string
    ): Promise<{ newComment: ISPComment | undefined; filesData: IQnAFilesData | undefined }>;
    uploadQAFiles(files: IFileInput[], uploadedBy: string, qnAItemId: number, fileType?: IFileType, commentId?: number): Promise<IQnAAttachment[]>;
    deleteComment(listName: ListName, uniqueId: string, commentId: number): Promise<boolean>;
    getQnAItems(stage?: string, phase?: string): Promise<IQnAItem[]>;
    getQnACategories(): Promise<string[]>;
}
