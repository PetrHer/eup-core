import { FileInfo } from "@syncfusion/ej2-react-inputs";

import { IFileType } from ".";

export interface IFileInput {
    file: FileInfo,
    type?: IFileType
}