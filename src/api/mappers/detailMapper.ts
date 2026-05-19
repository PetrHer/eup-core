import { ChannelStatus, ColumnsDetails, ContactDesc, UserType } from "../../enums";
import { ILoanDetail } from "../../interfaces";

type ContactData = {
    id?: string;
    phone?: string;
    desc?: string;
};

/**
 * map detail from shrepoint rest api to {@link ILoanDetail}
 * @param item 
 * @returns detail {@link ILoanDetail}
 */
export const mapDetail = (item: any, iupLink: string, contactsData?: any): ILoanDetail => {
    const status = item[ColumnsDetails.CurrentStatus] ? item[ColumnsDetails.CurrentStatus].split('/') : '';
    const channelStatus = Object.values(ChannelStatus).find(v => v === item[ColumnsDetails.ChannelStatus]);
    const contactData = (item[ColumnsDetails.ContactData] ?? {}) as ContactData;
    const secondaryData = (item[ColumnsDetails.SecondaryData] ?? {}) as ContactData;
    return {
        id: item.Id,
        currentStatus: status[0] ?? '',
        currentFolder: status[1] ?? '',
        text: item[ColumnsDetails.LoanText],
        title: item.Title,
        contactDesc: Object.values(ContactDesc).includes(contactData.desc as ContactDesc) ? contactData.desc as ContactDesc : undefined,
        secondaryDesc: Object.values(ContactDesc).includes(secondaryData.desc as ContactDesc) ? secondaryData.desc as ContactDesc : undefined,
        contact: contactsData && contactsData[ColumnsDetails.Contact] && contactsData[ColumnsDetails.Contact].Name
            ? {
                email: contactsData[ColumnsDetails.Contact].EMail,
                firstName: contactsData[ColumnsDetails.Contact].FirstName,
                lastName: contactsData[ColumnsDetails.Contact].LastName,
                name: contactsData[ColumnsDetails.Contact].Title,
                phone: contactData.phone ? contactData.phone : '',
                photo: '',
                userType: contactsData[ColumnsDetails.Contact].Name.includes('#EXT#') ? UserType.External : UserType.Internal,
                id: contactData.id ? contactData.id : '',
                loadedPhoto: false
            }
            : undefined,
        secondaryContact: contactsData && contactsData[ColumnsDetails.SecondaryContact] && contactsData[ColumnsDetails.SecondaryContact].Name
            ? {
                email: contactsData[ColumnsDetails.SecondaryContact].EMail,
                firstName: contactsData[ColumnsDetails.SecondaryContact].FirstName,
                lastName: contactsData[ColumnsDetails.SecondaryContact].LastName,
                name: contactsData[ColumnsDetails.SecondaryContact].Title,
                phone: secondaryData.phone ? secondaryData.phone : '',
                photo: '',
                userType: contactsData[ColumnsDetails.SecondaryContact].Name.includes('#EXT#') ? UserType.External : UserType.Internal,
                id: secondaryData.id ? secondaryData.id : '',
                loadedPhoto: false
            }
            : undefined,
        channelStatus: channelStatus ? channelStatus : ChannelStatus.Active,
        iupLink: iupLink,
        version: item[ColumnsDetails.Version] ? item[ColumnsDetails.Version] : 0
    };
};