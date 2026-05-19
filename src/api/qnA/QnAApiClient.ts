import { IQnAApiClient } from './IQnAApiClient';
import { ColumnsQA, ListName, QAItemType } from '../../enums';
import { IFileInput, IFileType, IQnAAttachment, IQnAFilesData, IQnAItem, ISPComment } from '../../interfaces';
import { getLatestDate, parseNameAndDate, splitFolderPath } from '../../utils/utils';
import FileApiClient from '../files/FileApiClient';
import { parseQnAFilesData } from '../mappers/parser';
import SPApiClient from '../SPApiClient';

export default class QnAApiClient implements IQnAApiClient {
    protected documentFolder: string = 'Sdilene dokumenty';

    constructor(private readonly files: FileApiClient, private readonly sp: SPApiClient) { }

    /**
         * Creates a new Q&A item in the QA list with the specified question, category, stage, phase, and order.
         *
         * @param question - The question text to be added.
         * @param category - The category to which the question belongs.
         * @param stage - The stage related to the question.
         * @param phase - The phase related to the question.
         * @param order - The display order of the question.
         * @returns A promise that resolves to true if the item was created successfully, otherwise false.
         */
    public async createQnAItem(question: string, category: string, stage: string, phase: string, priority: string, uploadedBy: string, fileType?: IFileType, order?: number, files?: IFileInput[], subCategory?: string): Promise<string | undefined> {
        const body = {
            "formValues": [
                {
                    "FieldName": "Title",
                    "FieldValue": question
                },
                {
                    "FieldName": "Category",
                    "FieldValue": category
                },
                {
                    "FieldName": "Status",
                    "FieldValue": 'new'
                },
                {
                    "FieldName": "Stage",
                    "FieldValue": stage
                },
                {
                    "FieldName": "Phase",
                    "FieldValue": phase
                },
                {
                    "FieldName": "QAOrder",
                    "FieldValue": order ? order.toString() : ''
                },
                {
                    "FieldName": ColumnsQA.Priority,
                    "FieldValue": priority
                },
                {
                    "FieldName": ColumnsQA.SubCategory,
                    "FieldValue": subCategory
                }
            ],
        };
        const result = await this.sp.createListItemUsingPath(body, ListName.QA);
        const value = result.data && Array.isArray((result.data as any).value) ? (result.data as any).value : [];
        const idField = value.find((item: any) => item.FieldName === 'Id');
        if (files && files.length > 0) {
            // const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.QA}')/items(${idField.FieldValue})`;
            const addedFiles = await this.uploadQAFiles(files, uploadedBy, Number(idField.FieldValue), fileType);
            const filesData: IQnAFilesData = { [QAItemType.Question]: [...addedFiles] };
            if (result.ok) {
                const body = { [ColumnsQA.FilesData]: JSON.stringify(filesData) };
                await this.sp.updateListItem(idField.FieldValue, body, ListName.QA);
            }
        }
        return idField ? idField.FieldValue : undefined;
    }

    /**
     * Adds a comment to a list item based on its unique ID.
     * @param uniqueId The unique identifier of the list item.
     * @param listName The name of the list where the item exists.
     * @param text The comment text to add.
     * @returns A boolean indicating whether the comment was successfully added.
     */
    public async addComment(qnAItem: IQnAItem, listName: string, text: string, files: IFileInput[], uploadedBy: string)
        : Promise<{ newComment: ISPComment | undefined, filesData: IQnAFilesData | undefined }> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${qnAItem.uniqueId}')/Comments()`;
        const body = {
            "text": text
        };
        const response = await this.sp.spPost(apiUrl, body);
        let filesData: IQnAFilesData | undefined = qnAItem.filesData;
        const fileType = files[0] ? files[0].type : undefined;
        if (files.length > 0 && fileType) {
            const addedFiles = await this.uploadQAFiles(files, uploadedBy, qnAItem.id, fileType, Number((response.data as any).id));
            filesData = {
                ...(qnAItem.filesData ?? {}),
                [QAItemType.Comment]: [
                    ...(qnAItem.filesData?.[QAItemType.Comment] ?? []),
                    ...addedFiles
                ]
            };
            await this.sp.updateItemInList(qnAItem.uniqueId, { [ColumnsQA.FilesData]: JSON.stringify(filesData) }, ListName.QA/*, file ?? [] */);
        }
        return response.ok && response.data ? { newComment: this.mapItemComments([response.data])[0], filesData } : { newComment: undefined, filesData: undefined };
    }

    /**
     * Uploads Q&A files to the specified folder and updates file metadata.
     * @param files Array of files to upload.
     * @param uploadedBy User uploading the files.
     * @param qnAItemId Q&A item ID for association.
     * @param fileType File type information (optional).
     * @param commentId Comment ID for association (optional).
     * @returns Promise resolving to an array of IQnAAttachment objects.
     */
    public async uploadQAFiles(files: IFileInput[], uploadedBy: string, qnAItemId: number, fileType?: IFileType, commentId?: number): Promise<IQnAAttachment[]> {
        if (!fileType || !files?.length) {
            return [];
        }
        const { name, path } = splitFolderPath(fileType.folderPath);
        const parentFolderPath = `${this.sp.getRelativeUrl()}/${this.documentFolder}/${this.sp.getChannelName()}/${path}`;
        await this.sp.ensureFolder(parentFolderPath, name);
        const uploadResult = await this.files.uploadFiles(files, false, uploadedBy);
        const addedFiles = uploadResult
            .filter(res => res.success)
            .map(res => ({
                name: res.file ? res.file.name : res.fileName,
                serverRelativeUrl: res.file?.serverRelativeUrl ?? '',
                commentId: commentId ? commentId : undefined
            }));
        for (const file of addedFiles) {
            if (file.serverRelativeUrl) {
                await this.files.updateFileByUrl(file.serverRelativeUrl, { QAId: qnAItemId.toString().padStart(4, '0') });
            }
        }
        return addedFiles;
    }

    /**
     * Deletes a comment from a list item by its unique ID and comment ID.
     * @param listName The name of the list.
     * @param uniqueId The unique identifier of the list item.
     * @param commentId The ID of the comment to delete.
     * @returns Promise resolving to true if deleted successfully, otherwise false.
     */
    public async deleteComment(listName: ListName, uniqueId: string, commentId: number): Promise<boolean> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${listName}')/GetItemByUniqueId('${uniqueId}')/Comments('${commentId}')`;
        const response = await this.sp.spDelete(apiUrl);
        return !!response.ok;
    }


    /**
 * Retrieves Q&A items from the specified folder in the QA list.
 * @param folderPAth The relative path of the folder within the QA list.
 * @returns An array of Q&A items.
 */
    public async getQnAItems(stage?: string, phase?: string): Promise<IQnAItem[]> {
        const fileDirRef = `${this.sp.getRelativeUrl()}/Lists/${ListName.QA}/${this.sp.getChannelName()}`;
        const expand = 'Author';
        const select = '*,FileDirRef,UniqueId,Author/Title';
        const filter = `startswith(FileDirRef,'${fileDirRef}') and startswith(ContentTypeId,'0x0100') ${stage && phase ? `and Stage eq '${stage}' and Phase eq '${phase}'` : ``}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.QA}')/items?$select=${select}&$filter=${filter}&$expand=${expand}`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) ? this.mapQnAItems(response.data) : [];
    }

    /**
     * Retrieves the choice options for the 'Category' field from the QnA list.
     * 
     * @returns An array of category choice strings available in the QnA list.
     */
    public async getQnACategories(): Promise<any> {
        const columnsApi = `${this.sp.getAbsoluteUrl()}/_api/lists/getbytitle('${ListName.QA}')/fields/getbyinternalnameortitle('Category')`;
        const col = await this.sp.spGet(columnsApi);
        return (col.data as any).Choices ?? [];
    }

    /**
 * Maps raw comment data to a structured array of comment objects.
 * @param data The raw data array of comments.
 * @returns An array of mapped comment objects.
 */
    private mapItemComments(data: any[]): ISPComment[] {
        data.sort((a, b) => (a.createdDate as string).localeCompare(b.createdDate as string));
        return data.map(item => {
            return {
                text: item.text,
                author: {
                    name: item.author.name,
                    email: item.author.email,
                    isExternal: item.author.isExternal
                },
                createdDate: item.createdDate,
                id: Number(item.id)
            };
        });
    }

    /**
         * Maps raw Q&A data to a structured array of Q&A items, including comments for each item.
         * @param data The raw data array of Q&A items.
         * @returns A promise that resolves to an array of mapped Q&A item objects.
         */
    private async mapQnAItems(data: any[]): Promise<IQnAItem[]> {
        return await Promise.all(data.map(async (item) => await this.mapQnAItem(item))
        );
    }

    /**
     * Maps a single raw Q&A item to a structured IQnAItem object, including comments and metadata.
     * @param item The raw Q&A item data.
     * @returns Promise resolving to a mapped IQnAItem object.
     */
    private async mapQnAItem(item: any): Promise<IQnAItem> {
        const comments = await this.getItemComments(item.Id);
        const { name, date } = parseNameAndDate(item[ColumnsQA.Answered]);
        const modified = getLatestDate([...comments.map(c => c.createdDate), item.Modified]);
        return {
            question: item.Title,
            answer: item[ColumnsQA.Answer],
            comments: comments,
            uniqueId: item.UniqueId,
            id: item.Id,
            status: item.Status,
            category: item.Category,
            phase: item.Phase,
            order: item[ColumnsQA.Order],
            stage: item.Stage,
            createdText: `${item.Author.Title} ${new Date(item.Created).toLocaleDateString()}`,
            answeredText: item[ColumnsQA.Answered],
            priority: item[ColumnsQA.Priority],
            created: new Date(item.Created),
            createdUser: item.Author.Title,
            createdDate: new Date(item.Created).toLocaleString(undefined, {
                year: "numeric",
                month: "numeric",
                day: "numeric",
                hour: "2-digit",
                minute: "2-digit"
            }),
            answeredDate: date,
            answeredUser: name,
            filesData: parseQnAFilesData(item[ColumnsQA.FilesData]),
            subCategory: item[ColumnsQA.SubCategory],
            modified: new Date(modified).toLocaleString(undefined, {
                year: "numeric",
                month: "numeric",
                day: "numeric",
                hour: "2-digit",
                minute: "2-digit"
            })
        };
    }

    /**
 * Retrieves comments for a specific item in the QA list based on its ID.
 * @param id The ID of the item to retrieve comments for.
 * @returns An array of comments associated with the item.
 */
    private async getItemComments(id: number): Promise<ISPComment[]> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getByTitle('${ListName.QA}')/items(${id})/comments`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) ? this.mapItemComments(response.data) : [];
    }
}