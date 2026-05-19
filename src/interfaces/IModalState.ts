import * as React from 'react';

export interface IModalState {
    body: React.ReactNode;
    title: string;
    width: number | string;
}