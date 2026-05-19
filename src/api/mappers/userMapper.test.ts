/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapInvitations, mapSiteGroupUsers, mapPerson, mapPersons, parseBatchSiteGroupUsersResponse } from './userMapper';
import { UserType } from '../../enums';

describe('mapInvitations', () => {
    it('returns empty array for empty input', () => {
        expect(mapInvitations([], new Map())).toEqual([]);
    });

    it('maps name and email from raw item', () => {
        const [result] = mapInvitations([{ Jmeno: 'John Doe', Email: 'john@test.com' }], new Map());
        expect(result.name).toBe('John Doe');
        expect(result.email).toBe('john@test.com');
    });

    it('maps userExists from existenceMap by email', () => {
        const map = new Map([['john@test.com', true]]);
        const [result] = mapInvitations([{ Jmeno: 'John', Email: 'john@test.com' }], map);
        expect(result.userExists).toBe(true);
    });

    it('defaults userExists to false when email not in map', () => {
        const [result] = mapInvitations([{ Jmeno: 'Jane', Email: 'jane@test.com' }], new Map());
        expect(result.userExists).toBe(false);
    });
});

describe('mapSiteGroupUsers', () => {
    it('returns empty array for empty input', () => {
        expect(mapSiteGroupUsers([])).toEqual([]);
    });

    it('maps id, email, name', () => {
        const [result] = mapSiteGroupUsers([{ Id: 1, Email: 'u@t.com', Title: 'User A' }]);
        expect(result.id).toBe(1);
        expect(result.email).toBe('u@t.com');
        expect(result.name).toBe('User A');
    });

    it('maps multiple users', () => {
        const result = mapSiteGroupUsers([
            { Id: 1, Email: 'a@t.com', Title: 'A' },
            { Id: 2, Email: 'b@t.com', Title: 'B' },
        ]);
        expect(result).toHaveLength(2);
    });
});

describe('mapPerson', () => {
    const internalUser: any = {
        mail: 'john@domain.com',
        displayName: 'John Doe',
        mobilePhone: '+420123456',
        givenName: 'John',
        surname: 'Doe',
        id: 'user-id-1',
        userPrincipalName: 'john@domain.com',
    };

    it('maps email, name, phone, firstName, lastName, id', () => {
        const result = mapPerson(internalUser);
        expect(result.email).toBe('john@domain.com');
        expect(result.name).toBe('John Doe');
        expect(result.phone).toBe('+420123456');
        expect(result.firstName).toBe('John');
        expect(result.lastName).toBe('Doe');
        expect(result.id).toBe('user-id-1');
    });

    it('marks as Internal when userPrincipalName has no #EXT#', () => {
        const result = mapPerson(internalUser);
        expect(result.userType).toBe(UserType.Internal);
    });

    it('marks as External when userPrincipalName contains #EXT#', () => {
        const externalUser: any = { ...internalUser, userPrincipalName: 'john_domain.com#EXT#@tenant.com' };
        const result = mapPerson(externalUser);
        expect(result.userType).toBe(UserType.External);
    });

    it('defaults optional fields to empty string when absent', () => {
        const result = mapPerson({} as any);
        expect(result.email).toBe('');
        expect(result.name).toBe('');
        expect(result.phone).toBe('');
        expect(result.firstName).toBe('');
        expect(result.lastName).toBe('');
    });

    it('sets loadedPhoto to false', () => {
        const result = mapPerson(internalUser);
        expect(result.loadedPhoto).toBe(false);
    });
});

describe('mapPersons', () => {
    it('returns empty array for empty input', () => {
        expect(mapPersons([])).toEqual([]);
    });

    it('maps each user via mapPerson', () => {
        const users: any[] = [
            { mail: 'a@t.com', displayName: 'A', id: '1', userPrincipalName: 'a@t.com' },
            { mail: 'b@t.com', displayName: 'B', id: '2', userPrincipalName: 'b@t.com' },
        ];
        const result = mapPersons(users);
        expect(result).toHaveLength(2);
        expect(result[0].email).toBe('a@t.com');
    });
});

describe('parseBatchSiteGroupUsersResponse', () => {
    it('returns an empty Map for empty response', () => {
        const result = parseBatchSiteGroupUsersResponse('', [1, 2]);
        expect(result).toBeInstanceOf(Map);
    });

    it('parses user data from a valid batch response', () => {
        const json = JSON.stringify({ value: [{ Id: 1, Title: 'User A', Email: 'a@t.com' }] });
        const raw = `--batchresponse_abc\r\nContent-Type: application/http\r\n\r\nHTTP/1.1 200 OK\r\nContent-Type: application/json\r\n\r\n${json}`;
        const result = parseBatchSiteGroupUsersResponse(raw, [10]);
        expect(result.has(10)).toBe(true);
        expect(result.get(10)?.[0].name).toBe('User A');
    });
});
