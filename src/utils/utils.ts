import { MenuItemModel } from "@syncfusion/ej2-react-navigations";
import JSZip from "jszip";
import { FieldValues } from "react-hook-form";

import { IApiClient } from "../api/IApiClient";
import { Language, RoleAlias, SPGroupExternal, SPGroupInternal } from "../enums";
import { IExtendedFile, IFormRecord, IStructureNode } from "../interfaces";
import { IFile } from "../interfaces/IFile";

/**
 * Adds a class to all child elements within a dropdown item to set its width correctly
 * @param element - The parent HTML element containing the dropdown items
 * @param className - The class name to be added to each child element
 * @returns void
 */
export const addClassToDropdownItem = (element: HTMLElement, className: string): void => {
    const arr = element.children[0] && element.children[0].children ? element.children[0].children : [];
    for (let i = 0; i < arr.length; i++) {
        const element = arr[i];
        element.classList.add(className);
    }
};

export const addClassToComboBoxItem = (element: HTMLElement | Element, className: string, width: number): void => {
    element.classList.add(className);
    const content = element.querySelector('.e-content') as HTMLElement;
    if (!content) return;
    content.classList.add(className);

    const applyWidth = (): void => {
        if (content.children[0]) {
            const contentChildren = content.children[0].children;
            const updatedWidth = contentChildren.length <= 8 ? width + 14 : width;
            for (let i = 0; i < contentChildren.length; i++) {
                (contentChildren[i] as HTMLElement).style.setProperty('width', `${updatedWidth}px`, 'important');
            }
        }
    };

    requestAnimationFrame(applyWidth);

    const observer = new MutationObserver(() => {
        requestAnimationFrame(applyWidth);
    });
    observer.observe(content, { childList: true, subtree: true });
};

/**
 * Creates a unique file name by appending a counter to the base name if the file already exists in the uploaded files list.
 * @param fileName - The original file name
 * @param uploadedFiles - The list of already uploaded files
 * @returns The unique file name
 */
export const createUniqueFileName = (fileName: string, uploadedFiles: IFile[]): string => {
    const extensionIndex = fileName.lastIndexOf('.');
    const baseName = fileName.substring(0, extensionIndex);
    const extension = fileName.substring(extensionIndex);
    let counter = 1;
    let newFileName = fileName;
    const fileExists = (name: string): boolean => uploadedFiles.some(file => file.name === name);
    while (fileExists(newFileName)) {
        newFileName = `${baseName} (${counter})${extension}`;
        counter++;
    }
    return newFileName;
};

/**
 * Extracts the initials from a full name by taking the first letter of each part of the name.
 * @param name - The full name (e.g., "John Doe")
 * @returns A string containing the initials (e.g., "JD")
 */
export const getInitials = (name: string): string => {
    return name.split(' ').map(part => part[0]).join('');
};

/**
* Converts a UI version number into a human-readable version label.
* @param uiVersion - The UI version number (e.g., 513)
* @returns A string representing the version label (e.g., "1.1")
*/
export const convertUIVersionToLabel = (uiVersion: number): string => {
    const SP_UI_VERSION_MAJOR_DIVISOR = 512;
    const majorVersion = Math.floor(uiVersion / SP_UI_VERSION_MAJOR_DIVISOR);
    const minorVersion = uiVersion % SP_UI_VERSION_MAJOR_DIVISOR;
    return `${majorVersion}.${minorVersion}`;
};

/**
 * Extracts the file extension from a given file name.
 * @param fileName - The file name (e.g., "document.pdf")
 * @returns The file extension (e.g., "pdf") or an empty string if no extension exists.
 */
export const getFileExtension = (fileName?: string): string => {
    if (!fileName) {
        return '';
    }
    const parts = fileName.split('.');
    return parts && parts.length > 1 ? parts.pop() ?? '' : '';
};

/**
 * Extracts the file name without its extension.
 * @param fileName - The file name (e.g., "document.pdf")
 * @returns The file name without the extension (e.g., "document")
 */
export const getFileNameWOExt = (fileName: string): string => {
    const ext = getFileExtension(fileName);
    return ext ? fileName.replace(`.${ext}`, '') : fileName;
};

export const getFileName = (relativeUrl?: string): string => {
    if (!relativeUrl) {
        return '';
    }
    return relativeUrl.split('/').pop() ?? '';
};

/**
 * Removes the file extension and trailing numeric suffix (e.g., " (1)") from a file name
 * @param fileName - The full name of the file including extension and optional suffix
 * @returns The cleaned base name of the file without extension or numeric suffix
 */
export const getCleanBaseName = (fileName: string): string => {
    const ext = getFileExtension(fileName);
    let nameWithoutExt = fileName;

    if (ext && fileName.toLowerCase().endsWith(`.${ext.toLowerCase()}`)) {
        nameWithoutExt = fileName.slice(0, -1 * (ext.length + 1));
    }
    const cleanName = nameWithoutExt.replace(/\s\(\d+\)$/, '');
    return cleanName;
};

/**
 * Cleans the folder name by removing a leading two-digit number and a space, if present.
 * @param folderName - The folder name (e.g., "01 FolderName")
 * @returns The cleaned folder name without the leading number (e.g., "FolderName")
 */
export const cleanFolderName = (folderName: string): string => {
    if (!folderName) {
        return '';
    }
    const pattern = /^\d{2} /;
    // Remove leading newline characters
    const cleanedFolderName = folderName.replace(/^\n+/, '');

    if (pattern.test(cleanedFolderName)) {
        return cleanedFolderName.replace(pattern, '');
    }
    return cleanedFolderName;
};

/**
     * Calculate the height of a context menu based on the number of items
     * @param items - The list of menu items in the context menu
     * @returns number - The calculated height of the context menu
     */
export const getContextMenuHeight = (items: MenuItemModel[]): number => {
    return items.filter(item => !item.separator).length * 36;
};


/**
 * Recursively searches for a node with a specific ID within a tree structure
 * @param nodes - The array of structure nodes to search through
 * @param id - The ID of the node to find
 * @returns The matching node if found, otherwise undefined
 */
export const findNodeById = (nodes: IStructureNode[], id: number): IStructureNode | undefined => {
    for (const node of nodes) {
        if (node.id === id) {
            return node;
        }
        const found = findNodeById(node.nodes, id);
        if (found) {
            return found;
        }
    }
    return undefined;
};

const getNodesByCategoryId = (
    nodes: IStructureNode[],
    id: number,
): { stageNode: IStructureNode; phaseNode: IStructureNode; categoryNode: IStructureNode } | undefined => {
    for (const level1 of nodes) {
        for (const level2 of level1.nodes) {
            for (const level3 of level2.nodes) {
                if (level3.id === id) {
                    return {
                        stageNode: level1,
                        phaseNode: level2,
                        categoryNode: level3,
                    };
                }
            }
        }
    }
    return undefined;
};

/**
 * Retrieves the localized stage, phase, and category names for a node by its ID
 * @param nodes - The hierarchical array of structure nodes (stages → phases → categories)
 * @param id - The ID of the target category node
 * @param locale - The language code used to retrieve localized names
 * @returns An object containing the localized stage, phase, and category names, or undefined if not found
 */
export const getLocalizationPathById = (
    nodes: IStructureNode[],
    id: number,
    locale: Language
): { stage: string; phase: string; category: string } | undefined => {
    const nodesById = getNodesByCategoryId(nodes, id);
    if (nodesById) {
        return {
            stage: nodesById.stageNode.localization?.[locale] ?? '',
            phase: nodesById.phaseNode.localization?.[locale] ?? '',
            category: nodesById.categoryNode.localization?.[locale] ?? '',
        };
    }
    return undefined;
};

// REFACTOR: `getStageByCategoryId`, `getStageNodeByCategoryId`, `getPhaseNodeByCategoryId`, and
// `getLocalizationPathById` (above) all contain the same triple-nested loop over a 3-level tree.
// Extract a single helper, e.g.:
//   function findInTree<T>(nodes: IStructureNode[], id: number, pick: (l1, l2, l3) => T): T | undefined
// and have each function call it with the appropriate projection.
export const getStageByCategoryId = (
    nodes: IStructureNode[],
    id?: number
): string | undefined => {
    if (id) {
        const nodesById = getNodesByCategoryId(nodes, id);
        if (nodesById) {
            return nodesById.stageNode.name ?? '';
        }
    }
    return undefined;
};

export const getStageNodeByCategoryId = (
    nodes: IStructureNode[],
    id?: number
): IStructureNode | undefined => {
    if (id) {
        const nodesById = getNodesByCategoryId(nodes, id);
        if (nodesById) {
            return nodesById.stageNode;
        }
    }
    return undefined;
};

export const getPhaseNodeByCategoryId = (
    nodes: IStructureNode[],
    id?: number
): IStructureNode | undefined => {
    if (id) {
        const nodesById = getNodesByCategoryId(nodes, id);
        if (nodesById) {
            return nodesById.phaseNode;
        }
    }
    return undefined;
};

/**
 * Parses the form data from a string to an object.
 * @param formData - The form data in string format
 * @returns FieldValues - The parsed form data object, or an empty object if parsing fails
 */
export const parseFormData = (formData: string): FieldValues => {
    try {
        return JSON.parse(formData);
    } catch (error) {
        console.error("Failed to parse formData:", formData, error);
        return {};
    }
};

export const parseSharePointDate = (value?: string): Date | undefined => {
    if (!value) {
        return undefined;
    }

    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? undefined : date;
};

/**
 * Type guard to check if an item is a FormRecord based on presence of 'formData' property
 * @param item - The item to check, which can be either IExtendedFile or IFormRecord
 * @returns True if the item is an IFormRecord, otherwise false
 */
export const isFormRecord = (item: IExtendedFile | IFormRecord): item is IFormRecord => {
    return 'formData' in item;
};

/**
 * Opens a file URL in a new browser tab, appending a query parameter to indicate web view
 * @param url - The URL of the file to open
 * @returns void
 */
export const openFileInTab = (url: string): void => {
    window.open(`${url}?web=1`, '_blank');
};

const getAppToOpen = (url: string): ('word' | 'excel' | 'powerpoint') => {
    const ext = getFileExtension(url);
    switch (ext) {
        case 'xlsx':
            return 'excel';
        case 'pptx':
            return 'powerpoint';
        default:
            return 'word';
    }
};

export const openFileInDesktopApp = (url: string): void => {
    const protocolMap = {
        word: 'ms-word:ofe|u|',
        excel: 'ms-excel:ofe|u|',
        powerpoint: 'ms-powerpoint:ofe|u|'
    };
    const app = getAppToOpen(url);
    const protocol = protocolMap[app];
    const uri = `${protocol}${encodeURI(url)}`;
    // Preferred: navigate current tab to protocol (usually launches app without leaving a blank tab)
    try {
        window.location.href = uri;
    } catch (e) {
        // Fallback: create an invisible link and click it (sometimes works around restrictions)
        const a = document.createElement('a');
        a.href = uri;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }
};

/**
 * Extracts the leaf (most specific) part of a hierarchical file type title.
 *
 * If the title includes a '>' delimiter (e.g., "Parent > Child"), it returns the child part.
 * Otherwise, returns the original title (trimmed).
 *
 * @param title - The full title string, optionally containing a '>' hierarchy.
 * @returns The simplified or leaf title, or an empty string if input is undefined or empty.
 */
export const getFileTypeTitle = (title?: string): string => {
    if (!title) return '';
    const parts = title.split('>').map(p => p.trim());
    return parts.length > 1 ? parts[1] : parts[0];
};


export const formatDateForSharePoint = (date: Date): string => {
    const month = date.getMonth() + 1; // JS months are 0-based
    const day = date.getDate();
    const year = date.getFullYear();
    let hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, '0');
    const ampm = hours >= 12 ? 'PM' : 'AM';

    hours = hours % 12;
    if (hours === 0) hours = 12;

    return `${month}/${day}/${year} ${hours}:${minutes} ${ampm}`;
};

const GROUP_ALIAS_MAP: Record<RoleAlias, string[]> = {
    [RoleAlias.UUO]: [SPGroupInternal.UUO_BO, SPGroupInternal.UUO_DEP, SPGroupInternal.UUO_DIR, SPGroupInternal.UUO_RM, SPGroupInternal.UUO_RSMA],
    [RoleAlias.OAO]: [SPGroupInternal.OAO_CC, SPGroupInternal.OAO_DEP, SPGroupInternal.OAO_DIR, SPGroupInternal.OAO_LAW],
    [RoleAlias.Admin]: [SPGroupInternal.ADMIN, SPGroupInternal.BUS_ADMIN],
    [RoleAlias.Internal]: Object.values(SPGroupInternal),
    [RoleAlias.Client]: [SPGroupExternal.EXT_CLIFULL, SPGroupExternal.EXT_CLILOW],
    [RoleAlias.External]: Object.values(SPGroupExternal),
    [RoleAlias.AML]: [SPGroupInternal.AML_DIR, SPGroupInternal.AML_S],
    [RoleAlias.WO]: [SPGroupInternal.WO_DEP, SPGroupInternal.WO_DIR, SPGroupInternal.WO_S],
    [RoleAlias.ORKR]: [SPGroupInternal.ORKR_GCCA, SPGroupInternal.ORKR_RA, SPGroupInternal.ORKR_RAPP, SPGroupInternal.ORKR_RD],
};

const resolveGroupAlias = (alias: string): string[] => {
    return GROUP_ALIAS_MAP[alias as RoleAlias] ?? [alias];
};

export const expandRoles = (roles: string): string[] => {
    const roleList = roles.split(";").map(r => r.trim());
    const groupSet = new Set<string>();
    roleList.forEach(role => {
        resolveGroupAlias(role).forEach(g => groupSet.add(g));
    });
    return Array.from(groupSet);
};

export const expandRoleAliases = (aliases: RoleAlias[]): string[] => {
    const groupSet = new Set<string>();
    aliases.forEach(alias => {
        resolveGroupAlias(alias).forEach(g => groupSet.add(g));
    });
    return Array.from(groupSet);
};

export const delay = (ms: number): Promise<void> => new Promise(resolve => setTimeout(resolve, ms));

const triggerBlobDownload = (blob: Blob, name: string): void => {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = name;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
};

export const downloadFile = async (name: string, serverRelativeUrl: string, apiClient: IApiClient): Promise<void> => {
    const arrayBuffer = await apiClient.files.getFileToDownload(serverRelativeUrl);
    if (!arrayBuffer) {
        return;
    }
    const blob = new Blob([arrayBuffer]);
    triggerBlobDownload(blob, name);
};

export const classifyUploader = (currentUserGroups: (SPGroupInternal | SPGroupExternal)[]): string => {
    if (currentUserGroups.some(g => g === SPGroupExternal.EXT_CLIFULL || g === SPGroupExternal.EXT_CLILOW)) {
        return 'client';
    } else if (currentUserGroups.some(g => g === SPGroupExternal.EXT_ADVISER || g === SPGroupExternal.EXT_AUD)) {
        return 'external';
    } else if (currentUserGroups.some(g => Object.values(SPGroupInternal).includes(g as SPGroupInternal))) {
        return 'bank';
    } else {
        return '';
    }
};

export const splitFolderPath = (fullPath: string): { path: string, name: string } => {
    const normalized = fullPath.replace(/\\/g, "/").replace(/\/+/g, "/");
    const lastSlash = normalized.lastIndexOf("/");
    if (lastSlash === -1) {
        return { path: "", name: normalized };
    }
    return {
        path: normalized.substring(0, lastSlash),
        name: normalized.substring(lastSlash + 1),
    };
};

export const createCategoryOptions = (nodes: IStructureNode[], locale: Language): { [key: string]: any }[] => {
    return nodes.map(node => ({
        value: node.id ? node.id.toString() : node.name,
        name: node.localization ? node.localization[locale] : node.name,
        expanded: true,
        selectable: !!node.id,
        subChild: createCategoryOptions(node.nodes, locale)
    }));
};

const downloadFilesAsZip = async (files: IFile[], apiClient: IApiClient, zipName: string = 'files.zip'): Promise<void> => {
    const zip = new JSZip();
    const nameMap: Record<string, number> = {};

    for (const file of files) {
        try {
            const arrayBuffer = await apiClient.files.getFileToDownload(file.serverRelativeUrl);

            let finalName = file.name;
            if (nameMap[file.name]) {
                const baseName = getFileNameWOExt(file.name);
                const ext = getFileExtension(file.name);
                finalName = ext
                    ? `${baseName} (${nameMap[file.name]}).${ext}`
                    : `${baseName} (${nameMap[file.name]})`;
                nameMap[file.name]++;
            } else {
                nameMap[file.name] = 1;
            }
            if (!arrayBuffer) {
                continue;
            }
            zip.file(finalName, arrayBuffer);
        } catch (err) {
            console.error(`Failed to download file ${file.name}:`, err);
        }
    }

    // Generate the ZIP blob once all files are added
    const blob = await zip.generateAsync({ type: 'blob' });
    triggerBlobDownload(blob, zipName);
};

export const downloadFiles = async (files: IFile[], apiClient: IApiClient, zipName: string = 'files.zip'): Promise<void> => {
    if (files.length === 1) {
        await downloadFile(files[0].name, files[0].serverRelativeUrl, apiClient);
    } else {
        await downloadFilesAsZip(files, apiClient, zipName);
    }
};

export const parseNameAndDate = (
    input?: string
): { name: string; date: string } => {
    if (!input || typeof input !== "string") {
        return { name: "", date: "" };
    }

    const match = input.match(
        /^(.*?)\s(\d{1,2}\.\s*\d{1,2}\.\s*\d{4}\s+\d{1,2}:\d{2})$/
    );

    if (!match) {
        return { name: "", date: "" };
    }

    return {
        name: match[1].trim(),
        date: match[2].trim()
    };
};

export const getLatestDate = (dates: string[]): string => {
    return dates.reduce((latest, dateStr) => {
        const date = new Date(dateStr);
        return date > new Date(latest) ? dateStr : latest;
    }, dates[0]);
};

export const copyLink = async (fileLink: string): Promise<void> => {
    await navigator.clipboard.writeText(fileLink);
};


export const normalizeDocType = (s?: string): string =>
    (s ?? "")
        .normalize("NFKD")                 // split accents (safe even if none)
        .replace(/[\u0300-\u036f]/g, "")   // remove diacritics
        .toLowerCase()
        .trim()
        .replace(/\u00A0/g, " ")           // replace non-breaking spaces
        .replace(/\s*-\s*/g, "-")          // normalize hyphen spacing
        .replace(/\s+/g, "")               // remove remaining spaces
        .replace(/[^a-z0-9-]/g, "");

export const transformPath = (path: string): string => {
    return path
        .split('/')
        .map(segment => {
            return encodeURIComponent(segment)
                .replace(/%20/g, '+');
        })
        .join('%252f');
};