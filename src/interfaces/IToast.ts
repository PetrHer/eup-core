import { ToastType } from "../enums";

export interface IToast {
    message: string;
    type: ToastType;
}