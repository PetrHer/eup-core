import { ColumnsDuties, EventPeriodicity, EventType } from "../../enums";
import { ICalendarEvent } from "../../interfaces";
import { parseSharePointDate } from "../../utils/utils";

/**
 * Maps raw data from the SharePoint "Duties" list to an array of ICalendarEvent objects.
 * @param data The raw data retrieved from the "Duties" list.
 * @returns An array of ICalendarEvent objects with the mapped properties.
 */
export const mapCalendarEvents = (data: any[]): ICalendarEvent[] => {
    return data.map(item => {
        return {
            id: item.Id,
            subject: item.Title,
            uniqueId: item.UniqueId,
            time: parseSharePointDate(item[ColumnsDuties.Date]),
            startDate: parseSharePointDate(item[ColumnsDuties.Date]),
            type: item[ColumnsDuties.EventType] as EventType,
            amount: item[ColumnsDuties.Amount],
            currency: item[ColumnsDuties.Currency],
            instruction: item[ColumnsDuties.Instruction],
            category: item[ColumnsDuties.Category],
            description: item[ColumnsDuties.EventDescription],
            note: item[ColumnsDuties.Note],
            period: item[ColumnsDuties.Period] ? item[ColumnsDuties.Period] as EventPeriodicity : undefined,
            endDate: parseSharePointDate(item[ColumnsDuties.EndDate]),
        };
    });
};