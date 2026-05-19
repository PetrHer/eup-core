// eslint-disable-next-line import/no-unresolved
import { User } from "@microsoft/microsoft-graph-types";
import { z } from "zod";

import { UserType } from "../../enums";
import { IInvitation, IPerson, ISiteGroupUser } from "../../interfaces";

/**
 * Maps raw invitation data to a structured format.
 * @param data The raw invitation data to be mapped.
 * @returns An array of mapped invitations.
 */
export const mapInvitations = (data: any[], userExistenceMap: Map<string, boolean>): IInvitation[] => {
    return data.map(item => {
        return {
            name: item.Jmeno,
            email: item.Email,
            userExists: userExistenceMap.get(item.Email) ?? false
        };
    });
};

/**
 * Maps raw site group user data to a structured array of user objects.
 * @param data The raw data array of site group users.
 * @returns An array of mapped site group user objects with relevant user details.
 */
export const mapSiteGroupUsers = (data: any[]): ISiteGroupUser[] => {
    return data.map(item => {
        return {
            id: item.Id,
            email: item.Email,
            name: item.Title
        };
    });
};

export const mapPerson = (user: any): IPerson => {
    return {
        email: user.mail ?? '',
        name: user.displayName ?? '',
        phone: user.mobilePhone ?? '',
        photo: '',
        firstName: user.givenName ?? '',
        lastName: user.surname ?? '',
        id: user.id,
        userType: user.userPrincipalName && !user.userPrincipalName.includes('#EXT#') ? UserType.Internal : UserType.External,
        loadedPhoto: false,
        channelMemberId: user.channelMemberId
    };
};

/**
 * Maps raw team member data to a structured array of person objects.
 * @param users The raw data array of team members.
 * @returns An array of mapped person objects with relevant user details.
 */
export const mapPersons = (users: any[]): IPerson[] => {
    return users.map((user: User) => mapPerson(user));
};

const UserSchema = z.array(z.object({
    Id: z.number(),
    Title: z.string(),
    Email: z.string(),
}));

export const parseBatchSiteGroupUsersResponse = (raw: string, groupIds: number[]): Map<number, ISiteGroupUser[]> => {
    const result = new Map<number, ISiteGroupUser[]>();

    const parts = raw.split(/--batchresponse_[\s\S]*?Content-Type: application\/http/gi).filter(p => p !== '');
    parts.forEach((part, index) => {
        const jsonMatch = part.match(/\{[\s\S]*\}/m);
        if (!jsonMatch) return;
        let json: any;
        try {
            json = JSON.parse(jsonMatch[0]);
        } catch {
            return;
        }
        const array = UserSchema.safeParse(json.value);
        if (!array.success) {
            return;
        }
        const users: ISiteGroupUser[] = array.data.map((u: any) => ({
            id: u.Id,
            name: u.Title,
            email: u.Email
        }));

        const groupId = groupIds[index];
        if (!groupId) return;
        result.set(groupId, users);
    });

    return result;
};