import { ILoanDetail } from '../../interfaces';

export interface IDetailApiClient {
    createDetail(currentStatus: string, version: number): Promise<boolean>;
    getLoanDetail(): Promise<ILoanDetail | undefined>;
    updateDetail(itemId: number, body: object): Promise<boolean>;
    updateContact(itemId: number, contactData: string, secondaryData: string, email?: string, secondaryEmail?: string): Promise<boolean>;
}
