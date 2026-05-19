import { ISiteGroupUser } from ".";
import { SPGroupExternal, SPGroupInternal, SPGroupRO } from "../enums";

export interface ISiteGroup {
    id: number,
    title: SPGroupExternal | SPGroupInternal | SPGroupRO,
    users: ISiteGroupUser[]
}