/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapDetail } from './detailMapper';
import { ChannelStatus, UserType } from '../../enums';

const baseItem = {
    Id: 1,
    Title: 'Loan 1',
    currentStatus: 'Stage1/Phase1',
    loanText: 'Some loan text',
    ChannelStatus: 'Active',
    contactData: null,
    secondaryData: null,
    version: 3,
};

describe('mapDetail', () => {
    it('maps id and title from item', () => {
        const result = mapDetail(baseItem, 'https://iup.link');
        expect(result.id).toBe(1);
        expect(result.title).toBe('Loan 1');
    });

    it('splits currentStatus into currentStatus and currentFolder', () => {
        const result = mapDetail(baseItem, 'https://iup.link');
        expect(result.currentStatus).toBe('Stage1');
        expect(result.currentFolder).toBe('Phase1');
    });

    it('sets currentStatus and currentFolder to empty strings when field is absent', () => {
        const result = mapDetail({ ...baseItem, currentStatus: undefined }, '');
        expect(result.currentStatus).toBe('');
        expect(result.currentFolder).toBe('');
    });

    it('maps text from loanText', () => {
        const result = mapDetail(baseItem, '');
        expect(result.text).toBe('Some loan text');
    });

    it('maps iupLink from second argument', () => {
        const result = mapDetail(baseItem, 'https://my-iup-link.com');
        expect(result.iupLink).toBe('https://my-iup-link.com');
    });

    it('maps version field', () => {
        const result = mapDetail(baseItem, '');
        expect(result.version).toBe(3);
    });

    it('defaults version to 0 when absent', () => {
        const result = mapDetail({ ...baseItem, version: undefined }, '');
        expect(result.version).toBe(0);
    });

    it('maps channelStatus from ChannelStatus field', () => {
        const result = mapDetail({ ...baseItem, ChannelStatus: 'Inactive' }, '');
        expect(result.channelStatus).toBe(ChannelStatus.Inactive);
    });

    it('defaults channelStatus to Active when field does not match enum', () => {
        const result = mapDetail({ ...baseItem, ChannelStatus: 'Unknown' }, '');
        expect(result.channelStatus).toBe(ChannelStatus.Active);
    });

    it('sets contact to undefined when contactsData has no contact', () => {
        const result = mapDetail(baseItem, '', undefined);
        expect(result.contact).toBeUndefined();
    });

    it('maps internal contact when contactsData.contact has a Name without #EXT#', () => {
        const contactsData = {
            contact: {
                Name: 'i:0#.f|membership|user@domain.com',
                EMail: 'user@domain.com',
                FirstName: 'John',
                LastName: 'Doe',
                Title: 'John Doe',
            },
            secondaryContact: null,
        };
        const result = mapDetail(
            { ...baseItem, contactData: { phone: '123', id: 'abc', desc: 'Sales' } },
            '',
            contactsData
        );
        expect(result.contact).toBeDefined();
        expect(result.contact?.email).toBe('user@domain.com');
        expect(result.contact?.userType).toBe(UserType.Internal);
    });

    it('maps external contact when Name contains #EXT#', () => {
        const contactsData = {
            contact: {
                Name: 'user_domain.com#EXT#@tenant.onmicrosoft.com',
                EMail: 'user@domain.com',
                FirstName: 'Jane',
                LastName: 'Smith',
                Title: 'Jane Smith',
            },
            secondaryContact: null,
        };
        const result = mapDetail(baseItem, '', contactsData);
        expect(result.contact?.userType).toBe(UserType.External);
    });
});
