import { SPGroupExternal, SPGroupInternal, SPGroupRO, TeamRole } from '../../enums';
import { IExtendedPerson, IInvitation, IPerson, ISiteGroupUser } from '../../interfaces';

export interface IUserApiClient {
    readonly eupNoReplyMail: string;
    getTeamMembers(): Promise<IPerson[]>;
    getExistingUser(mail: string): Promise<IPerson | undefined>;
    getUserPhoto(userIdOrEmail: string): Promise<string>;
    getCurrentUserGroups(): Promise<(SPGroupInternal | SPGroupExternal)[]>;
    getSiteGroupUsersBatch(groupIds: number[]): Promise<Map<number, ISiteGroupUser[]>>;
    ensureRole(): Promise<void>;
    createSiteGroups(groupNames: (SPGroupExternal | SPGroupInternal | SPGroupRO)[], increaseProgressCount: () => void): Promise<void>;
    getInvitations(teamMembers: IPerson[]): Promise<IInvitation[]>;
    addUserToSiteGroup(id: number, userEmail: string): Promise<void>;
    removeUserFromsiteGroup(groupId: number, userId: number): Promise<void>;
    createInvitation(email: string, name: string, roles: string, language: string, existingUser?: boolean): Promise<boolean>;
    checkTemplatesAccess(): Promise<boolean>;
    checkUserAndCreateInvitation(email: string, name: string, roles: string, language: string): Promise<boolean>;
    addUserByEmailToTeamAndChannel(email: string, role?: TeamRole): Promise<boolean>;
    addEUPAdminsAsOwners(): Promise<boolean>;
    addEUPMembersAsMembers(): Promise<boolean>;
    getUserName(userEmail: string): Promise<string>;
    getCurrentUser(): Promise<IPerson | undefined>;
    getExistingUsers(mail: string): Promise<IPerson[]>;
    removeUserFromTeamOrChannel(person: IExtendedPerson): Promise<boolean>;
}
