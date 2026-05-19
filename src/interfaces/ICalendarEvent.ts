import { IStructureNode } from ".";
import { EventPeriodicity } from "../enums";

export interface ICalendarEvent {
    uniqueId: string,
    id: number,
    subject: string,
    time: Date | undefined,
    startDate: Date | undefined,
    type: string,
    amount?: number,
    currency?: string,
    instruction?: string,
    category?: string,
    description?: string,
    note?: string,
    node?: IStructureNode,
    period?: EventPeriodicity,
    endDate?: Date
}