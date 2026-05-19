import { z } from "zod";

import { IDetailApiClient } from "./IDetailApiClient";
import { ChannelStatus, ColumnsDetails, ListName } from '../../enums';
import { ILoanDetail } from '../../interfaces';
import IApiResult from '../IApiResult';
import { mapDetail } from '../mappers';
import SPApiClient from "../SPApiClient";

const ContactDataSchema = z.object({
    id: z.string().nullable().optional(),
    phone: z.string().nullable().optional(),
    desc: z.string().nullable().optional(),
});

const parseContactData = (value?: unknown): z.infer<typeof ContactDataSchema> => {
    if (!value) {
        return {};
    }

    try {
        if (typeof value !== 'string') {
            return {};
        }

        const parsedJson = JSON.parse(value);
        const parsedContactData = ContactDataSchema.safeParse(parsedJson);

        return parsedContactData.success ? parsedContactData.data : {};
    } catch {
        return {};
    }
};

const SPDetailItemSchema = z.object({
    Id: z.number(),
    Title: z.string().nullable().optional(),

    [ColumnsDetails.CurrentStatus]: z.string().nullable().optional(),
    [ColumnsDetails.ChannelStatus]: z.string().nullable().optional(),
    [ColumnsDetails.LoanText]: z.string().nullable().optional(),
    [ColumnsDetails.ContactData]: z.string().nullable().optional().transform(parseContactData),
    [ColumnsDetails.SecondaryData]: z.string().nullable().optional().transform(parseContactData),
    [ColumnsDetails.Version]: z.number().nullable().optional(),
});

export default class DetailApiClient implements IDetailApiClient {
    private documentFolder: string = 'Sdilene dokumenty';

    /**
     * Constructs a DetailApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    /**
         * create detail of loan when installing
         * @param channelId 
         * @param title 
         * @param currentStatus 
         */
    public async createDetail(currentStatus: string, version: number): Promise<boolean> {
        const body: any = {
            "Title": this.sp.getChannelName(),
            [ColumnsDetails.ChannelId]: this.sp.getChannelId(),
            [ColumnsDetails.CurrentStatus]: currentStatus,
            [ColumnsDetails.ChannelStatus]: ChannelStatus.Active,
            [ColumnsDetails.Version]: version
        };
        const result = await this.sp.createListItem(body, ListName.Details);
        return result.ok ?? false;
    }

    /**
     * @param channelId 
     * @returns detail {@link ILoanDetail}
     */
    public async getLoanDetail(): Promise<ILoanDetail | undefined> {
        const filter = `${ColumnsDetails.ChannelId} eq '${this.sp.getChannelId()}'`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Details}')/items?$filter=${filter}&$select=*`;
        const response = await this.sp.spGet(apiUrl);
        const contactResponse = await this.getLoanDetailContacts(this.sp.getChannelId());
        const contactsData = Array.isArray(contactResponse.data) ? contactResponse.data[0] : undefined;
        const rawItem = Array.isArray(response.data) ? response.data[0] : undefined;

        if (!rawItem) {
            return undefined;
        }
        const parsed = SPDetailItemSchema.safeParse(rawItem);
        if (!parsed.success) {
            console.error(parsed.error);
            return undefined;
        }

        const item = parsed.data;
        const iupLink = await this.getIUPLink();
        return mapDetail(item, iupLink, contactsData);
    }

    /**
   * updates detail using {@link updateListItem}
   * @param itemId 
   * @param body 
   * @returns boolean
   */
    public async updateDetail(itemId: number, body: object): Promise<boolean> {
        const result = await this.sp.updateListItem(itemId, body, ListName.Details);
        return result;
    }

    /**
     * ensure user and update detail using {@link updateListItem}
     * @param itemId 
     * @param email 
     * @returns boolean
     */
    public async updateContact(itemId: number, contactData: string, secondaryData: string, email?: string, secondaryEmail?: string): Promise<boolean> {
        const userId = email ? await this.sp.getUserSPId(email) : null;
        const secondaryId = secondaryEmail ? await this.sp.getUserSPId(secondaryEmail) : null;
        const body = {
            [`${ColumnsDetails.Contact}Id`]: userId,
            [`${ColumnsDetails.SecondaryContact}Id`]: secondaryId,
            [ColumnsDetails.ContactData]: contactData ?? '',
            [ColumnsDetails.SecondaryData]: secondaryData ?? ''
        };
        const result = await this.sp.updateListItem(itemId, body, ListName.Details);
        return result;
    }

    private async getLoanDetailContacts(channelId?: string): Promise<IApiResult> {
        const filter = `${ColumnsDetails.ChannelId} eq '${channelId}'`;
        const select = [
            '*',
            `${ColumnsDetails.Contact}/FirstName`,
            `${ColumnsDetails.Contact}/LastName`,
            `${ColumnsDetails.Contact}/Title`,
            `${ColumnsDetails.Contact}/Name`,
            `${ColumnsDetails.Contact}/UserName`,
            `${ColumnsDetails.Contact}/EMail`,
            `${ColumnsDetails.Contact}/Id`,
            `${ColumnsDetails.Contact}/MobilePhone`,
            `${ColumnsDetails.Contact}/WorkPhone`,
            `${ColumnsDetails.SecondaryContact}/FirstName`,
            `${ColumnsDetails.SecondaryContact}/LastName`,
            `${ColumnsDetails.SecondaryContact}/Title`,
            `${ColumnsDetails.SecondaryContact}/Name`,
            `${ColumnsDetails.SecondaryContact}/UserName`,
            `${ColumnsDetails.SecondaryContact}/EMail`,
            `${ColumnsDetails.SecondaryContact}/Id`,
            `${ColumnsDetails.SecondaryContact}/MobilePhone`,
            `${ColumnsDetails.SecondaryContact}/WorkPhone`,
        ].join(',');
        const expand = `${ColumnsDetails.Contact},${ColumnsDetails.SecondaryContact}`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Details}')/items?$filter=${filter}&$select=${select}&$expand=${expand}`;
        const response = await this.sp.spGet(apiUrl);
        return response;
    }

    private async getIUPLink(): Promise<string> {
        const teamName = await this.sp.getTeamName();
        if (!teamName) {
            return '';
        }

        const updatedTeamName = teamName.replace('_ext', '');
        const id = `/sites/${this.sp.internalSiteName}/${this.documentFolder}/${updatedTeamName}/${this.sp.getChannelName()}`;
        const url = `${this.sp.getAbsoluteUrl().replace(this.sp.getRelativeUrl(), '')}/sites/${this.sp.internalSiteName}/${this.documentFolder}/Forms/AllItems.aspx?id=${id}`;
        return url;
    }
}