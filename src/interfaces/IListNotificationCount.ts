import { ListName } from "../enums";

export interface IListNotificationCount {
    [ListName.Requests]: number,
    [ListName.Forms]: number,
    'Total': number,
}