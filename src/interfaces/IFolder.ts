import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IFolder {
    name: string,
    localization: ILocalizedStrings,
    serverRelativeUrl: string,
    folders: IFolder[]
}