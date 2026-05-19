import { SPGroupExternal, SPGroupInternal } from "../enums";
import { IPerson } from "./IPerson";

export interface IExtendedPerson extends IPerson {
    // spMembership?: { [key in SPGroupExternal | SPGroupInternal]: boolean }
    spMembership?: (SPGroupExternal | SPGroupInternal)[],
    userExists?: boolean
}