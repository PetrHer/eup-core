import { FieldValues } from 'react-hook-form';

import { IFormApiClient } from './IFormApiClient';
import { ColumnsForms, FormStatus, Language, ListName } from '../../enums';
import { IFormRecord } from '../../interfaces';
import { mapFormRecords } from '../mappers';
import SPApiClient from "../SPApiClient";

export default class FormApiClient implements IFormApiClient {
    /**
     * Constructs a FormApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    /**
     * Retrieves form records from the SharePoint list.
     * @returns Promise<IFormRecord[]> - A promise that resolves to an array of form records
     */
    public async getFormRecords(): Promise<IFormRecord[]> {
        const folderPath = `${this.sp.getRelativeUrl()}/lists/${ListName.Forms}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Forms}')/items?$filter=${filter}`;
        const result = await this.sp.spGet(apiUrl);
        return Array.isArray(result.data) ? mapFormRecords(result.data) : [];
    }

    /**
     * Saves the form data to a SharePoint list and updates its status.
     * @param id - The ID of the form record to update
     * @param formData - The form data to be saved
     * @param asDraft - A flag indicating whether the form data should be saved as a draft (optional, defaults to false)
     * @returns Promise<boolean> - A promise that resolves to true if the data is successfully saved, otherwise false
     */
    public async saveFormData(id: number, formData: FieldValues, language: Language, asDraft: boolean = false): Promise<boolean> {
        const body = {
            [ColumnsForms.FormData]: JSON.stringify(formData),
            [ColumnsForms.Status]: asDraft ? FormStatus.New : FormStatus.Submitted,
            [ColumnsForms.Language]: language
        };
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Forms}')/items(${id})`;
        const result = await this.sp.spPatch(apiUrl, body);
        return result.ok ?? false;
    }
}