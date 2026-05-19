import { Language } from "../enums";
import { ILocalizedStrings } from "./ILocalizedStrings";

export interface IRequestOption {
    id: number,
    type: string,
    fileTemplates: {
        [Language.CS]?: string;
        [Language.EN]?: string;
    }[],
    loc: ILocalizedStrings,
    stage?: string[],
    desc?: ILocalizedStrings,
    fields?: string[];
    fileAttachments?: boolean;
    notifGroups?: string;
    fileReply?: boolean
}