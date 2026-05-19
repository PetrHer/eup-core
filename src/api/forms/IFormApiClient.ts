import { FieldValues } from 'react-hook-form';

import { Language } from '../../enums';
import { IFormRecord } from '../../interfaces';

export interface IFormApiClient {
    getFormRecords(): Promise<IFormRecord[]>;
    saveFormData(id: number, formData: FieldValues, language: Language, asDraft?: boolean): Promise<boolean>;
}
