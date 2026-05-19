export interface INewNotification {
    title: string,
    task: boolean,
    categoryId: number,
    userGroups: string,
    targetId?: string,
    titleEn: string,
    link?: string,
    comment?: string,
    fileLink?: string,
}