import { IPerson } from ".";
import { ChannelStatus, ContactDesc } from "../enums";

export interface ILoanDetail {
    contactDesc?: ContactDesc;
    secondaryDesc?: ContactDesc;
    id: number;
    currentStatus?: string;
    currentFolder?: string;
    text?: string;
    title: string;
    contact?: IPerson,
    secondaryContact?: IPerson,
    channelStatus: ChannelStatus,
    iupLink: string,
    version: number,
}