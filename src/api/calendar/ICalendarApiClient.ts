import { ICalendarEvent } from '../../interfaces';

export interface ICalendarApiClient {
    getCalendarEvents(): Promise<ICalendarEvent[]>;
    createCalendarEvent(event: ICalendarEvent): Promise<boolean>;
}
