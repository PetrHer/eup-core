import { ICalendarApiClient } from "./ICalendarApiClient";
import { ColumnsDuties, ListName } from '../../enums';
import { ICalendarEvent } from '../../interfaces';
import { mapCalendarEvents } from '../mappers';
import SPApiClient from "../SPApiClient";

export default class CalendarApiClient implements ICalendarApiClient {
    /**
     * Constructs a CalendarApiClient instance.
     * @param sp The SPApiClient instance for SharePoint operations.
     */
    constructor(private readonly sp: SPApiClient) { }

    /**
     * Retrieves calendar events from the Duties SharePoint list.
     * Filters items by folder and content type, then maps them to structured calendar event objects.
     * @returns Promise resolving to an array of ICalendarEvent objects, or an empty array if no events are found.
     */
    public async getCalendarEvents(): Promise<ICalendarEvent[]> {
        const folderPath = `${this.sp.getRelativeUrl()}/lists/${ListName.Duties}/${this.sp.getChannelName()}`;
        const filter = `FileDirRef eq '${folderPath}' and startswith(ContentTypeId,'0x0100')`;
        const select = '*,UniqueId';
        const apiUrl = `${this.sp.getAbsoluteUrl()}/_api/web/lists/getbytitle('${ListName.Duties}')/items?$select=${select}&$filter=${filter}`;
        const response = await this.sp.spGet(apiUrl);
        return Array.isArray(response.data) ? mapCalendarEvents(response.data) : [];
    }

    /**
     * Creates a new calendar event item in the Duties SharePoint list.
     * Normalizes event dates to noon and formats as ISO strings before saving.
     * @param event The calendar event data to create.
     * @returns Promise resolving to true if creation succeeded, otherwise false.
     */
    public async createCalendarEvent(event: ICalendarEvent): Promise<boolean> {
        const date = event.startDate;
        date?.setHours(12);
        const formattedDate = date ? date.toISOString() : '';

        const endDate = event.endDate;
        endDate?.setHours(12);
        const formattedEndDate = endDate ? endDate.toISOString() : '';

        const body = {
            "formValues": [
                {
                    "FieldName": "Title",
                    "FieldValue": event.subject
                },
                {
                    "FieldName": ColumnsDuties.Date,
                    "FieldValue": formattedDate
                },
                {
                    "FieldName": ColumnsDuties.EventType,
                    "FieldValue": event.type
                },
                {
                    "FieldName": ColumnsDuties.Amount,
                    "FieldValue": event.amount?.toString()
                },
                {
                    "FieldName": ColumnsDuties.Currency,
                    "FieldValue": event.currency
                },
                {
                    "FieldName": ColumnsDuties.Instruction,
                    "FieldValue": event.instruction
                },
                {
                    "FieldName": ColumnsDuties.Category,
                    "FieldValue": event.category
                },
                {
                    "FieldName": ColumnsDuties.EventDescription,
                    "FieldValue": event.description
                },
                {
                    "FieldName": ColumnsDuties.Note,
                    "FieldValue": event.note
                },
                {
                    "FieldName": ColumnsDuties.Period,
                    "FieldValue": event.period
                },
                {
                    "FieldName": ColumnsDuties.EndDate,
                    "FieldValue": formattedEndDate
                }
            ],
        };
        const result = await this.sp.createListItemUsingPath(body, ListName.Duties);
        return result.ok ?? false;
    }
}