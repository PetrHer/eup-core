import { UserType } from "../enums";

export interface IPerson {
    name: string;
    email: string;
    phone: string;
    photo: string;
    firstName: string;
    lastName: string;
    userType: UserType;
    id?: string;
    loadedPhoto: boolean;
    channelMemberId?: string;
}