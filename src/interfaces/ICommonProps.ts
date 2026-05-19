// import * as React from 'react';

// import { IToast } from "./IToast";
// import { IApiClient } from "../api/IApiClient";
//import { Permissions } from "../enums";

/** 
 * props used in multiple components
 */
export interface ICommonProps {
    // openModal: (body: React.ReactNode, title: string) => void;
    // closeModal: () => void;
    // apiClient: IApiClient;
    // isDarkTheme: boolean;
    refreshCallback: () => Promise<void>;
    //permissions: Permissions;
    // showToast: (toast: IToast | undefined) => void;
}