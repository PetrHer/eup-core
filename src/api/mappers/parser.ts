import { IQnAFilesData } from "../../interfaces";

function isQnAFilesData(obj: unknown): obj is IQnAFilesData {
    if (!obj || typeof obj !== "object") return false;

    return Object.values(obj as Record<string, unknown>).every(isFileArray);
}

function isFileArray(val: unknown): val is IQnAFilesData[keyof IQnAFilesData] {
    if (!Array.isArray(val)) return false;

    return val.every(isValidFileItem);
}

function isValidFileItem(item: unknown): item is {
    name: string;
    serverRelativeUrl: string;
    inFolder: boolean;
} {
    if (!item || typeof item !== "object") return false;

    const f = item as Record<string, unknown>;

    return (
        typeof f.name === "string" &&
        typeof f.serverRelativeUrl === "string"
    );
}

export const parseQnAFilesData = (
    value?: string
): IQnAFilesData | undefined => {
    if (!value) return undefined;
    try {
        const parsed = JSON.parse(value) as IQnAFilesData;

        if (!isQnAFilesData(parsed)) {
            return undefined;
        }

        return parsed;
    } catch (error) {
        console.warn("Failed to parse FilesData:", error);
        return undefined;
    }
};