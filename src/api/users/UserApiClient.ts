import { z } from "zod";

import { IUserApiClient } from './IUserApiClient';
import { RoleDefinition, SPGroupExternal, SPGroupInternal, SPGroupRO, SystemAccount, TeamRole, TemplatesListName } from '../../enums';
import { IExtendedPerson, IInvitation, IPerson, ISiteGroupUser } from '../../interfaces';
import IApiResult from '../IApiResult';
import { mapTeamMembers } from '../mappers';
import { mapInvitations, mapPerson, mapPersons, parseBatchSiteGroupUsersResponse } from '../mappers/userMapper';
import SPApiClient from "../SPApiClient";

const RoleAssignmentSchema = z.array(z.object({
    Member: z.object({
        LoginName: z.string(),
        Id: z.number(),
    }),
    RoleDefinitionBindings: z.array(z.object({
        Id: z.number(),
    }))
}));

const parseRoleAssignments = (value?: unknown): z.infer<typeof RoleAssignmentSchema> => {
    if (!value) {
        return [];
    }
    const parsedRoleAssignments = RoleAssignmentSchema.safeParse(value);
    return parsedRoleAssignments.success ? parsedRoleAssignments.data : [];
};

export default class UserApiClient implements IUserApiClient {
    constructor(private readonly sp: SPApiClient, public readonly eupNoReplyMail: string) { }

    /**
         * @param groupId 
         * @returns teamMembers of current team mapped to array of {@link IPerson}
         */
    public async getTeamMembers(): Promise<IPerson[]> {
        const channelMembersApiUrl = `/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/members`;
        const channelMembersResponse = await this.sp.graphGet(channelMembersApiUrl);
        const channelMembers = Array.isArray(channelMembersResponse.data) ? channelMembersResponse.data : [];
        const apiUrl = `groups/${this.sp.getTeamId()}/members`;
        const response = await this.sp.graphGet(apiUrl);
        const teamMembersData = Array.isArray(response.data) ? response.data : [];
        const filteredData: any[] = teamMembersData
            .filter(teamMember =>
                channelMembers.some(channelMember =>
                    channelMember.email === teamMember.mail
                ) &&
                !Object.values(SystemAccount).includes(teamMember.mail)
            )
            .map(teamMember => {
                const match = channelMembers.find(channelMember =>
                    channelMember.email === teamMember.mail
                );

                return {
                    ...teamMember,
                    channelMemberId: match?.id // 👈 THIS is what you were missing
                };
            });
        return mapTeamMembers(filteredData);
    }

    public async getExistingUser(mail: string): Promise<IPerson | undefined> {
        const encodedMail = encodeURIComponent(mail);
        const apiUrl = `/users?$filter=mail eq '${encodedMail}'`;
        const result = await this.sp.graphGet(apiUrl);
        return Array.isArray(result.data) && result.data.length > 0 ? mapPerson(result.data[0]) : undefined;
    }

    public async getExistingUsers(mail: string): Promise<IPerson[]> {
        const encodedMail = encodeURIComponent(mail);
        const apiUrl =
            `/users?$filter=` +
            `startswith(mail,'${encodedMail}') ` +
            `or startswith(displayName,'${encodedMail}') ` +
            `or startswith(userPrincipalName,'${encodedMail}')` +
            `&$select=id,displayName,mail,userPrincipalName` +
            `&$top=10`;
        const result = await this.sp.graphGet(apiUrl);
        const data = Array.isArray(result.data) ? result.data : [];
        return mapPersons(data);
    }

    /**
     * Retrieves the photo of a user based on their ID.
     * @param userIdOrEmail The ID or email of the user whose photo is to be fetched.
     * @returns A promise that resolves to a URL representing the user's photo.
     */
    public async getUserPhoto(userIdOrEmail: string): Promise<string> {
        const photoResponse = await this.sp.graphGet(`/users/${userIdOrEmail}/photo/$value`);
        return photoResponse.data && photoResponse.data instanceof Blob ? URL.createObjectURL(photoResponse.data) : '';
    }

    /**
     * Retrieves the current user's SharePoint groups and filters them to only include
     * groups defined in internal and external enums.
     * 
     * @returns Array of group titles from SPGroupInternal or SPGroupExternal enums the user belongs to.
     */
    public async getCurrentUserGroups(): Promise<(SPGroupInternal | SPGroupExternal)[]> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/currentuser?$expand=Groups`;
        const result = await this.sp.spGet(apiUrl);

        const groups: any[] = result.data && Array.isArray((result.data as any).Groups) ? (result.data as any).Groups : [];
        const allEnumGroupTitles = [
            ...Object.values(SPGroupInternal),
            ...Object.values(SPGroupExternal),
        ];

        const relevantGroupTitles: (SPGroupInternal | SPGroupExternal)[] = groups
            .map((g: any) => {
                const parsed = z.object({
                    Title: z.string(),
                }).safeParse(g);
                return parsed.success ? parsed.data.Title : undefined;
            })
            .filter(title => title && allEnumGroupTitles.includes(title as SPGroupInternal | SPGroupExternal)) as (SPGroupInternal | SPGroupExternal)[];

        return relevantGroupTitles;
    }

    public async getSiteGroupUsersBatch(groupIds: number[]): Promise<Map<number, ISiteGroupUser[]>> {

        const boundary = "batch_" + Date.now();
        const batchParts = groupIds.map(groupId => {
            const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users`;
            return [
                `--${boundary}`,
                'Content-Type: application/http',
                'Content-Transfer-Encoding: binary',
                `Content-ID: ${groupId}`,
                '',
                `GET ${apiUrl} HTTP/1.1`,
                'Accept: application/json',
                '',
                ''
            ].join('\r\n');
        });
        const batchBody = batchParts.join('\r\n')
            + `\r\n--${boundary}--\r\n`;

        const response = await this.sp.spPost(
            `${this.sp.getAbsoluteUrl()}/_api/$batch`,
            batchBody,
            {
                "Content-Type": `multipart/mixed; boundary=${boundary}`,
                "Accept": "application/json",
                "OData-Version": "4.0",
                "OData-MaxVersion": "4.0"
            }
        );
        const groupMap = response.data && typeof response.data === 'string' ? parseBatchSiteGroupUsersResponse(response.data as string, groupIds) : new Map();
        return groupMap;
    }

    public async ensureRole(): Promise<void> {
        const roleId = await this.getRoleDefinitionId(RoleDefinition.CustomEdit);
        if (!roleId) {
            const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roledefinitions`;
            const body = {
                BasePermissions: { High: 432, Low: 1011028583 },
                Description: "Read, Add, Edit without Delete",
                Name: RoleDefinition.CustomEdit,
                Order: 2147483647
            };
            await this.sp.spPost(apiUrl, body);
        }
        const missingRoles = await this.getMissingRoleAssignments();
        await this.assingMissingRoles(missingRoles);
    }

    private async getRoleDefinitionId(roleName: string): Promise<number | undefined> {
        const filter = `Name eq '${roleName}'`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roledefinitions`;
        const result = await this.sp.spGet(`${apiUrl}?$filter=${filter}`);
        return Array.isArray(result.data) && result.data[0] ? result.data[0].Id : undefined;
    }

    /**
     * Creates multiple site groups based on the provided group names.
     * @param groupNames An array of group names to be created (either internal or external).
     * @param increaseProgressCount A callback function to track progress during the group creation process.
     * @returns A promise that resolves when all site groups are created.
     */
    public async createSiteGroups(groupNames: (SPGroupExternal | SPGroupInternal | SPGroupRO)[], increaseProgressCount: () => void): Promise<void> {
        for (let i = 0; i < groupNames.length; i++) {
            if (Object.values(SPGroupRO).includes(groupNames[i] as SPGroupRO)) {
                await this.createSiteROGroup(groupNames[i] as SPGroupRO);
            } else {
                await this.createSiteGroup(groupNames[i] as (SPGroupExternal | SPGroupInternal), increaseProgressCount);
            }
        }
    }

    /**
     * Creates a single site group based on the provided group name.
     * @param groupName The name of the group to be created (either internal, external, or read-only).
     * @param increaseProgressCount A callback function to track progress after each group creation.
     * @returns A promise that resolves when the site group is created.
     */
    private async createSiteGroup(groupName: SPGroupExternal | SPGroupInternal, increaseProgressCount: () => void): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups`;
        const body = { Title: groupName };
        await this.sp.spPost(apiUrl, body);
        increaseProgressCount();
    }

    private async createSiteROGroup(groupName: SPGroupRO): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups`;
        const body = { Title: groupName };
        const result = await this.sp.spPost(apiUrl, body);
        const groupId = result.data ? (result?.data as any).Id : undefined;
        const roleId = await this.getRoleDefinitionId(RoleDefinition.Read);
        if (groupId && roleId) {
            await this.assignRoleToGroup(groupId, roleId);
        }
    }

    private async assingMissingRoles(missingRoles: { groupId: number, roleId?: number }[]): Promise<void> {
        for (let i = 0; i < missingRoles.length; i++) {
            const { groupId, roleId } = missingRoles[i];
            if (roleId) {
                await this.assignRoleToGroup(groupId, roleId);
            }
        }
    }

    private async assignRoleToGroup(groupId: number, roleId: number): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roleassignments/addroleassignment(principalid=${groupId}, roledefid=${roleId})`;
        await this.sp.spPost(apiUrl);
    }

    private async getMissingRoleAssignments(): Promise<{ groupId: number, roleId?: number }[]> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/roleassignments?$expand=Member,RoleDefinitionBindings`;
        const result = await this.sp.spGet(apiUrl);
        const customEditId = await this.getRoleDefinitionId(RoleDefinition.CustomEdit);
        const editId = await this.getRoleDefinitionId(RoleDefinition.Edit);
        const missingAssignments: { groupId: number, roleId?: number }[] = [];
        const data = Array.isArray(result.data) ? parseRoleAssignments(result.data) : [];
        for (let i = 0; i < data.length; i++) {
            const item = data[i];
            const name = item.Member.LoginName;
            if (!Object.values(SPGroupInternal).includes(name as SPGroupInternal) && !Object.values(SPGroupExternal).includes(name as SPGroupExternal)) {
                continue;
            }
            const roleId = Object.values(SPGroupInternal).includes(name as SPGroupInternal) ? editId : customEditId;
            const role = item.RoleDefinitionBindings.find(r => r.Id === roleId);
            if (!role) {
                missingAssignments.push({ groupId: item.Member.Id, roleId: roleId });
            }
        }
        return missingAssignments;
    }

    /**
     * Retrieves all invitations from the Invitations list.
     * @returns A promise that resolves to an array of invitation objects.
     */
    public async getInvitations(teamMembers: IPerson[]): Promise<IInvitation[]> {
        const filter = `ChannelId eq '${this.sp.getChannelId()}'`;
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        const apiUrl = `${webUrl}/_api/web/GetListByTitle('${TemplatesListName.Invitations}')/items?$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        const data = Array.isArray(response.data) ? response.data : [];
        const filteredData = data.filter(invitation => !teamMembers.some(teamMember => teamMember.email === invitation.Email));
        const userExistenceMap = await this.checkADUsersExistence(filteredData.map(item => item.Email));
        return Array.isArray(response.data) ? mapInvitations(filteredData, userExistenceMap) : [];
    }

    /**
     * Adds a user to a specific SharePoint site group.
     * @param id The ID of the site group.
     * @param userEmail The email address of the user to be added.
     * @returns A promise that resolves when the user is added to the group.
     */
    public async addUserToSiteGroup(id: number, userEmail: string): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
        const body = { LoginName: `i:0#.f|membership|${userEmail}` };
        await this.sp.spPost(apiUrl, body);
    }

    /**
     * Removes a user from a specific SharePoint site group.
     * @param groupId The ID of the site group.
     * @param userId The ID of the user to be removed.
     * @returns A promise that resolves when the user is removed from the group.
     */
    public async removeUserFromsiteGroup(groupId: number, userId: number): Promise<void> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users/removeById(${userId})`;
        await this.sp.spPost(apiUrl);
    }

    private async removeUserFromsiteGroupByEmail(groupId: number, email: string): Promise<void> {
        const loginName = `i:0#.f|membership|${email}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${groupId})/users/removeByLoginName('${encodeURIComponent(loginName)}')`;
        await this.sp.spPost(apiUrl);
    }

    /**
     * Creates a new invitation in the Invitations list.
     * @param email The email address of the person being invited.
     * @param name The name of the person being invited.
     * @param roles The roles assigned to the person being invited.
     */
    public async createInvitation(email: string, name: string, roles: string, language: string, existingUser: boolean = false): Promise<boolean> {
        const webUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}`;
        const apiUrl = `${webUrl}/_api/web/GetListByTitle('${TemplatesListName.Invitations}')/items`;
        const body: any = {
            "Title": name,
            'Jmeno': name,
            'Email': email,
            'ChannelId': this.sp.getChannelId(),
            'Roles': roles,
            'Language': language,
            'User': existingUser ? email : ''
        };
        const result = await this.sp.spPost(apiUrl, body);
        return !!result.ok;
    }

    /**
     * Checks if an Active Directory user exists by their email address.
     * @param email The email address of the user.
     * @returns A Promise that resolves to true if the user exists, otherwise false.
     */
    private async checkADUserExistence(email: string): Promise<boolean> {
        const result = await this.sp.graphGet(`/users/${email}`);
        return result.ok ?? false;
    }

    private async checkADUsersExistence(emails: string[]): Promise<Map<string, boolean>> {
        const resultMap = new Map<string, boolean>();
        const chunkSize = 20;

        for (let i = 0; i < emails.length; i += chunkSize) {
            const chunk = emails.slice(i, i + chunkSize);

            const batchBody = {
                requests: chunk.map((email, index) => ({
                    id: `${index + 1}`,
                    method: "GET",
                    url: `/users/?$filter=mail eq '${encodeURIComponent(email)}'`
                }))
            };

            const result = await this.sp.graphPost(`/$batch`, batchBody);
            const responses = (result.data as any)?.responses || [];

            for (const response of responses) {
                const email = chunk[parseInt(response.id, 10) - 1];
                resultMap.set(email, response.status === 200 && Array.isArray(response.body.value) && response.body.value.length > 0);
            }
        }

        return resultMap;
    }
    /**
        * checks if user has access to EUP-Sablony site
        */
    public async checkTemplatesAccess(): Promise<boolean> {
        const apiUrl = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.templateSiteName}/_api/web?$select=Id`;
        const result = await this.sp.spGet(apiUrl);
        return !!result.ok;
    }

    /**
     * Checks if a user exists in Active Directory by email, 
     * creates an invitation if the user does not exist, 
     * or adds the existing user to the team.
     * 
     * @param email - The email address of the user to check or invite.
     * @param name - The display name of the user for the invitation.
     */
    public async checkUserAndCreateInvitation(email: string, name: string, roles: string, language: string): Promise<boolean> {
        const userExists = await this.checkADUserExistence(email);
        if (!userExists) {
            return await this.createInvitation(email, name, roles, language);
        } else {
            return await this.addUserByEmailToTeamAndChannel(email);
        }
    }

    // public async checkUserAndAddToTeam(email:string):Promise<

    /**
     * Adds a member to a team by their email address.
     * @param email The email address of the user to add to the team.
     * @returns A promise indicating the completion of the operation.
     */
    public async addUserByEmailToTeamAndChannel(email: string, role: TeamRole = TeamRole.Member): Promise<boolean> {
        const userData = await this.sp.graphGet(`/users/?$filter=mail eq '${encodeURIComponent(email)}'`);
        const userId = Array.isArray(userData.data) ? userData.data[0]?.id : undefined;
        if (!userId) {
            return false;
        }
        const result = await this.addUserOrGroupByIdToTeamAndChannel(userId, role, email);
        return result;
    }

    public async addEUPAdminsAsOwners(): Promise<boolean> {
        const ownersApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id,Title,LoginName`;
        const ownersResult = await this.sp.spGet(ownersApiUrl);
        if (!ownersResult.ok) {
            return false;
        }
        const ownersId = (ownersResult.data as any).Id;
        const adminId = await this.getGroupIdByName(SPGroupInternal.ADMIN);
        const ids = [ownersId, adminId];

        const aadGroupClaim = `c:0t.c|tenant|${SystemAccount.EUPAdminsId}`;
        const emailClaim = `i:0#.f|membership|${this.eupNoReplyMail}`;
        let result = true;

        for (const id of ids) {

            const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
            for (const login of [aadGroupClaim, emailClaim]) {
                const response = await this.sp.spPost(apiUrl, { LoginName: login });
                if (!response.ok) {
                    result = false;
                }
            }
        }
        return result;
    }

    public async addEUPMembersAsMembers(): Promise<boolean> {
        const memberssApiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/associatedmembergroup?$select=Id,Title,LoginName`;
        const membersResult = await this.sp.spGet(memberssApiUrl);
        if (!membersResult.ok) {
            return false;
        }
        const membersId = (membersResult.data as any).Id;
        const roRoleId = await this.getGroupIdByName(SPGroupRO.RORole);

        const ids = [membersId, ...(roRoleId !== undefined ? [roRoleId] : [])];
        const aadGroupClaim = `c:0t.c|tenant|${SystemAccount.EUPMembersId}`;

        let result = true;
        for (const id of ids) {
            const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups(${id})/users`;
            const response = await this.sp.spPost(apiUrl, { LoginName: aadGroupClaim });
            if (!response.ok) { result = false; }
        }
        return result;
    }

    private async getGroupIdByName(groupName: string): Promise<number | undefined> {
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/sitegroups?$filter=Title eq '${groupName}'`;

        const result = await this.sp.spGet(apiUrl);

        if (!result.ok) {
            return undefined;
        }

        const groups: any[] = Array.isArray(result.data) ? result.data : [];
        return groups.length > 0 ? groups[0].Id : undefined;
    }

    private async addUserOrGroupByIdToTeamAndChannel(userId: string, role: TeamRole = TeamRole.Member, email: string): Promise<boolean> {
        const addedToTeam = await this.addMemberToTeam(userId, this.sp.getTeamId(), role);
        const isPrivateChannel = await this.sp.isPrivateChannel();
        if (isPrivateChannel) {
            const addedToChannel = await this.addMemberToPrivateChannel(this.sp.getTeamId(), this.sp.getChannelId(), userId, role, email);
            const owners = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id`);
            await this.addUserToSiteGroup((owners.data as any)?.Id, email);
            return addedToChannel;
        }
        if (role === TeamRole.Guest) {
            await this.fixGuestGroup(email);
        }
        return addedToTeam;
    }

    public async removeUserFromTeamOrChannel(person: IExtendedPerson): Promise<boolean> {
        const isPrivateChannel = await this.sp.isPrivateChannel();
        if (isPrivateChannel) {
            const result = await this.sp.graphDelete(`/teams/${this.sp.getTeamId()}/channels/${this.sp.getChannelId()}/members/${person.channelMemberId}`);
            return !!result.ok;
        } else {
            const result = await this.sp.graphDelete(`/groups/${this.sp.getTeamId()}/members/${person.id}/$ref`);
            return !!result.ok;
        }
    }

    private async addMemberToTeam(userId: string, teamId: string, role: TeamRole): Promise<boolean> {
        const body = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": [role],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
        };
        const apiUrl = `/teams/${teamId}/members`;
        const result = await this.sp.graphPost(apiUrl, body);
        return !!result.ok;
    }

    private async addMemberToPrivateChannel(teamId: string, channelId: string, userId: string, role: TeamRole, email: string): Promise<boolean> {
        const body = {
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": [role],
            "user@odata.bind": `https://graph.microsoft.com/v1.0/users('${userId}')`
        };
        const apiUrl = `/teams/${teamId}/channels/${channelId}/members`;
        const result = await this.sp.graphPost(apiUrl, body);
        const owners = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedownergroup?$select=Id`);
        await this.addUserToSiteGroup((owners.data as any)?.Id, email);
        if (role === TeamRole.Guest) {
            await this.fixGuestGroup(email);
        }
        return !!result.ok;
    }

    private async fixGuestGroup(email: string): Promise<void> {
        const visitors = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedvisitorgroup?$select=Id`);
        const members = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/associatedmembergroup?$select=Id`);
        await this.addUserToSiteGroup((visitors.data as any).Id, email);
        await this.removeUserFromsiteGroupByEmail((members.data as any).Id, email);
    }

    public async getUserName(userEmail: string): Promise<string> {
        const loginName = encodeURIComponent(`i:0#.f|membership|${userEmail}`);
        const result: IApiResult = await this.sp.spGet(`${this.sp.getAbsoluteUrl()}/_api/web/siteusers/getByLoginName('${loginName}')`);
        return result.data && (result.data as any).Title
            ? (result.data as any).Title
            : '';
    }

    public async getCurrentUser(): Promise<IPerson | undefined> {
        const result = await this.sp.graphGet('/me');
        return result.data ? mapPerson(result.data as any) : undefined;
    }
}