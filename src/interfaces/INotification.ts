// import { IPerson } from ".";

import { ILocalizedStrings } from "./ILocalizedStrings";

export interface INotification {
    id: number;
    title: ILocalizedStrings;
    task: boolean;
    completed: boolean;
    category: ILocalizedStrings;
    stage: ILocalizedStrings;
    phase: ILocalizedStrings;
    created: string;
    author: string;
    unread: boolean;
    read: number[];
    targetId?: string;
    categoryId: number;
    comment?: string;
    status?: string;
    modified?: string;
    editor?: string;
    attachments?: {
        fileName: string;
        url: string;
    }[];
    link?: string;
    uniqueId: string;
    fileLink?: string;
    // assingedTo: IPerson[];
}