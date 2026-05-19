/* eslint-disable no-undef */
jest.mock('jszip', () => {
    const MockJSZip = jest.fn().mockImplementation(() => ({
        file: jest.fn(),
        generateAsync: jest.fn().mockResolvedValue(new Blob(['zip'])),
    }));
    (MockJSZip as any).default = MockJSZip;
    return MockJSZip;
});

import { MenuItemModel } from '@syncfusion/ej2-react-navigations';

import {
    addClassToDropdownItem,
    classifyUploader,
    cleanFolderName,
    convertUIVersionToLabel,
    copyLink,
    createCategoryOptions,
    createUniqueFileName,
    downloadFile,
    downloadFiles,
    expandRoles,
    findNodeById,
    formatDateForSharePoint,
    getCleanBaseName,
    getContextMenuHeight,
    getFileExtension,
    getFileName,
    getFileNameWOExt,
    getFileTypeTitle,
    getInitials,
    getLatestDate,
    getLocalizationPathById,
    getPhaseNodeByCategoryId,
    getStageByCategoryId,
    getStageNodeByCategoryId,
    isFormRecord,
    normalizeDocType,
    openFileInDesktopApp,
    parseFormData,
    parseSharePointDate,
    parseNameAndDate,
    splitFolderPath,
    transformPath,
} from './utils';
import { FileStatus, FolderStatus, Language, SPGroupExternal, SPGroupInternal } from '../enums';
import { IFile } from '../interfaces';
import { IStructureNode } from '../interfaces/IStructureNode';

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const makeNode = (
    id: number | undefined,
    name: string,
    nodes: IStructureNode[] = [],
    localization?: { [Language.CS]: string;[Language.EN]: string }
): IStructureNode => ({
    id,
    name,
    nodes,
    uploadedFiles: [],
    status: FolderStatus.Default,
    localization,
});

// Three-level tree used across structure-traversal tests
//   Stage1
//     Phase1
//       Cat1  (id: 10)
//       Cat2  (id: 11)
//     Phase2
//       Cat3  (id: 12)
//   Stage2
//     Phase3
//       Cat4  (id: 20)
const buildTree = (): IStructureNode[] => [
    makeNode(1, 'Stage1', [
        makeNode(2, 'Phase1', [
            makeNode(10, 'Cat1', [], { [Language.CS]: 'Kategorie 1', [Language.EN]: 'Category 1' }),
            makeNode(11, 'Cat2', [], { [Language.CS]: 'Kategorie 2', [Language.EN]: 'Category 2' }),
        ], { [Language.CS]: 'Fáze 1', [Language.EN]: 'Phase 1' }),
        makeNode(3, 'Phase2', [
            makeNode(12, 'Cat3', [], { [Language.CS]: 'Kategorie 3', [Language.EN]: 'Category 3' }),
        ], { [Language.CS]: 'Fáze 2', [Language.EN]: 'Phase 2' }),
    ], { [Language.CS]: 'Etapa 1', [Language.EN]: 'Stage 1' }),
    makeNode(4, 'Stage2', [
        makeNode(5, 'Phase3', [
            makeNode(20, 'Cat4', [], { [Language.CS]: 'Kategorie 4', [Language.EN]: 'Category 4' }),
        ], { [Language.CS]: 'Fáze 3', [Language.EN]: 'Phase 3' }),
    ], { [Language.CS]: 'Etapa 2', [Language.EN]: 'Stage 2' }),
];

// ---------------------------------------------------------------------------
// addClassToDropdownItem
// ---------------------------------------------------------------------------

describe('addClassToDropdownItem', () => {
    let mockElement: HTMLElement;

    beforeEach(() => {
        mockElement = document.createElement('div');
        const wrapper = document.createElement('div');
        wrapper.appendChild(document.createElement('div'));
        wrapper.appendChild(document.createElement('div'));
        mockElement.appendChild(wrapper);
    });

    it('adds the specified class to all children of the first child', () => {
        addClassToDropdownItem(mockElement, 'test-class');
        const firstChild = mockElement.children[0];
        expect(firstChild.children[0].classList.contains('test-class')).toBe(true);
        expect(firstChild.children[1].classList.contains('test-class')).toBe(true);
    });

    it('does not throw if the first child has no children', () => {
        const el = document.createElement('div');
        el.appendChild(document.createElement('div'));
        expect(() => addClassToDropdownItem(el, 'test-class')).not.toThrow();
    });

    it('does nothing if the element has no children', () => {
        const el = document.createElement('div');
        addClassToDropdownItem(el, 'test-class');
        expect(el.children.length).toBe(0);
    });

    it('adds the class alongside existing classes', () => {
        const firstGrandchild = mockElement.children[0].children[0];
        firstGrandchild.classList.add('existing-class');
        addClassToDropdownItem(mockElement, 'new-class');
        expect(firstGrandchild.classList.contains('existing-class')).toBe(true);
        expect(firstGrandchild.classList.contains('new-class')).toBe(true);
    });
});

// ---------------------------------------------------------------------------
// createUniqueFileName
// ---------------------------------------------------------------------------

describe('createUniqueFileName', () => {
    const uploadedFiles: IFile[] = [
        { name: 'file1.txt', status: FileStatus.Processing, serverRelativeUrl: '/files/file1.txt', uniqueId: '1', version: '1.0' },
        { name: 'file2.txt', status: FileStatus.Processing, serverRelativeUrl: '/files/file2.txt', uniqueId: '2', version: '1.0' },
        { name: 'document.pdf', status: FileStatus.Processing, serverRelativeUrl: '/files/document.pdf', uniqueId: '3', version: '1.0' },
    ];

    it('returns the original name when it does not exist yet', () => {
        expect(createUniqueFileName('newFile.txt', uploadedFiles)).toBe('newFile.txt');
    });

    it('appends (1) when the name already exists', () => {
        expect(createUniqueFileName('file1.txt', uploadedFiles)).toBe('file1 (1).txt');
    });

    it('increments counter until a unique name is found', () => {
        const files = [
            ...uploadedFiles,
            { name: 'file1 (1).txt', status: FileStatus.Processing, serverRelativeUrl: '', uniqueId: '4', version: '1.0' },
        ];
        expect(createUniqueFileName('file1.txt', files)).toBe('file1 (2).txt');
    });

    it('returns the original name for files without an extension', () => {
        expect(createUniqueFileName('fileWithoutExtension', uploadedFiles)).toBe('fileWithoutExtension');
    });

    it('returns the original name when a counter-suffixed name does not already exist', () => {
        expect(createUniqueFileName('file1 (2).txt', uploadedFiles)).toBe('file1 (2).txt');
    });

    it('handles empty string input gracefully', () => {
        expect(createUniqueFileName('', uploadedFiles)).toBe('');
    });
});

// ---------------------------------------------------------------------------
// getInitials
// ---------------------------------------------------------------------------

describe('getInitials', () => {
    it('returns initials for a two-word name', () => {
        expect(getInitials('John Doe')).toBe('JD');
    });

    it('returns an empty string for an empty name', () => {
        expect(getInitials('')).toBe('');
    });

    it('returns initials for a hyphenated first name', () => {
        expect(getInitials('Anna-Marie Smith')).toBe('AS');
    });

    it('returns a single initial for a one-word name', () => {
        expect(getInitials('Alice')).toBe('A');
    });
});

// ---------------------------------------------------------------------------
// convertUIVersionToLabel
// ---------------------------------------------------------------------------

describe('convertUIVersionToLabel', () => {
    it('converts a version number to major.minor label', () => {
        expect(convertUIVersionToLabel(1025)).toBe('2.1');
    });

    it('handles exact multiples of 512 (zero minor version)', () => {
        expect(convertUIVersionToLabel(512)).toBe('1.0');
    });

    it('handles a version number smaller than 512', () => {
        expect(convertUIVersionToLabel(300)).toBe('0.300');
    });
});

// ---------------------------------------------------------------------------
// getFileExtension
// ---------------------------------------------------------------------------

describe('getFileExtension', () => {
    it('extracts the file extension', () => {
        expect(getFileExtension('document.pdf')).toBe('pdf');
    });

    it('returns an empty string when there is no extension', () => {
        expect(getFileExtension('fileWithoutExtension')).toBe('');
    });

    it('returns the last extension for files with multiple dots', () => {
        expect(getFileExtension('archive.tar.gz')).toBe('gz');
    });

    it('returns an empty string for undefined input', () => {
        expect(getFileExtension(undefined)).toBe('');
    });
});

// ---------------------------------------------------------------------------
// getFileNameWOExt
// ---------------------------------------------------------------------------

describe('getFileNameWOExt', () => {
    it('removes the file extension', () => {
        expect(getFileNameWOExt('document.pdf')).toBe('document');
    });

    it('returns the original name when there is no extension', () => {
        expect(getFileNameWOExt('fileWithoutExtension')).toBe('fileWithoutExtension');
    });

    it('removes only the last extension for multi-dot file names', () => {
        expect(getFileNameWOExt('archive.tar.gz')).toBe('archive.tar');
    });
});

// ---------------------------------------------------------------------------
// getFileName
// ---------------------------------------------------------------------------

describe('getFileName', () => {
    it('extracts the file name from a relative URL', () => {
        expect(getFileName('/sites/repo/files/document.pdf')).toBe('document.pdf');
    });

    it('returns the input unchanged when there is no slash', () => {
        expect(getFileName('document.pdf')).toBe('document.pdf');
    });

    it('returns an empty string for undefined input', () => {
        expect(getFileName(undefined)).toBe('');
    });

    it('returns an empty string for an empty string input', () => {
        expect(getFileName('')).toBe('');
    });
});

// ---------------------------------------------------------------------------
// getCleanBaseName
// ---------------------------------------------------------------------------

describe('getCleanBaseName', () => {
    it('removes the extension and a trailing counter suffix', () => {
        expect(getCleanBaseName('file1 (1).txt')).toBe('file1');
    });

    it('removes the extension but leaves the name unchanged when there is no counter', () => {
        expect(getCleanBaseName('document.pdf')).toBe('document');
    });

    it('returns the name unchanged when there is no extension', () => {
        expect(getCleanBaseName('fileWithoutExt')).toBe('fileWithoutExt');
    });

    it('strips any numeric counter regardless of its value', () => {
        expect(getCleanBaseName('file (42).docx')).toBe('file');
    });
});

// ---------------------------------------------------------------------------
// cleanFolderName
// ---------------------------------------------------------------------------

describe('cleanFolderName', () => {
    it('removes a leading two-digit number and space', () => {
        expect(cleanFolderName('12 MyFolder')).toBe('MyFolder');
    });

    it('leaves the name unchanged when it does not start with two digits', () => {
        expect(cleanFolderName('MyFolder')).toBe('MyFolder');
    });

    it('handles larger two-digit prefixes correctly', () => {
        expect(cleanFolderName('99 Folder Name Here')).toBe('Folder Name Here');
    });

    it('returns an empty string for empty input', () => {
        expect(cleanFolderName('')).toBe('');
    });

    it('strips a leading newline before checking for the digit prefix', () => {
        expect(cleanFolderName('\n01 Phase')).toBe('Phase');
    });
});

// ---------------------------------------------------------------------------
// getContextMenuHeight
// ---------------------------------------------------------------------------

describe('getContextMenuHeight', () => {
    it('returns items-count × 36 for non-separator items', () => {
        const items = [
            { text: 'Item 1' },
            { text: 'Item 2' },
            { text: 'Item 3' },
        ] as MenuItemModel[];
        expect(getContextMenuHeight(items)).toBe(108);
    });

    it('excludes separator items from the count', () => {
        const items = [
            { text: 'Item 1' },
            { separator: true },
            { text: 'Item 2' },
        ] as MenuItemModel[];
        expect(getContextMenuHeight(items)).toBe(72);
    });

    it('returns 0 for an empty list', () => {
        expect(getContextMenuHeight([])).toBe(0);
    });
});

// ---------------------------------------------------------------------------
// findNodeById
// ---------------------------------------------------------------------------

describe('findNodeById', () => {
    const tree = buildTree();

    it('finds a top-level node', () => {
        expect(findNodeById(tree, 1)?.name).toBe('Stage1');
    });

    it('finds a deeply nested node', () => {
        expect(findNodeById(tree, 10)?.name).toBe('Cat1');
    });

    it('finds a node in a second branch', () => {
        expect(findNodeById(tree, 20)?.name).toBe('Cat4');
    });

    it('returns undefined when the id does not exist', () => {
        expect(findNodeById(tree, 999)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// getLocalizationPathById
// ---------------------------------------------------------------------------

describe('getLocalizationPathById', () => {
    const tree = buildTree();

    it('returns the correct localized stage/phase/category path', () => {
        const result = getLocalizationPathById(tree, 10, Language.EN);
        expect(result).toEqual({ stage: 'Stage 1', phase: 'Phase 1', category: 'Category 1' });
    });

    it('returns Czech localization when requested', () => {
        const result = getLocalizationPathById(tree, 12, Language.CS);
        expect(result).toEqual({ stage: 'Etapa 1', phase: 'Fáze 2', category: 'Kategorie 3' });
    });

    it('returns undefined when the category id is not found', () => {
        expect(getLocalizationPathById(tree, 999, Language.EN)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// getStageByCategoryId
// ---------------------------------------------------------------------------

describe('getStageByCategoryId', () => {
    const tree = buildTree();

    it('returns the stage name for a known category id', () => {
        expect(getStageByCategoryId(tree, 10)).toBe('Stage1');
    });

    it('returns the correct stage when the category is in the second stage', () => {
        expect(getStageByCategoryId(tree, 20)).toBe('Stage2');
    });

    it('returns undefined for an unknown category id', () => {
        expect(getStageByCategoryId(tree, 999)).toBeUndefined();
    });

    it('returns undefined when id is not provided', () => {
        expect(getStageByCategoryId(tree, undefined)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// getStageNodeByCategoryId
// ---------------------------------------------------------------------------

describe('getStageNodeByCategoryId', () => {
    const tree = buildTree();

    it('returns the stage node containing the given category', () => {
        expect(getStageNodeByCategoryId(tree, 11)?.id).toBe(1);
    });

    it('returns undefined for an unknown category id', () => {
        expect(getStageNodeByCategoryId(tree, 999)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// getPhaseNodeByCategoryId
// ---------------------------------------------------------------------------

describe('getPhaseNodeByCategoryId', () => {
    const tree = buildTree();

    it('returns the phase node containing the given category', () => {
        expect(getPhaseNodeByCategoryId(tree, 12)?.id).toBe(3);
    });

    it('returns undefined for an unknown category id', () => {
        expect(getPhaseNodeByCategoryId(tree, 999)).toBeUndefined();
    });
});

// ---------------------------------------------------------------------------
// parseFormData
// ---------------------------------------------------------------------------

describe('parseFormData', () => {
    it('parses a valid JSON string into an object', () => {
        const result = parseFormData('{"name":"Alice","age":30}');
        expect(result).toEqual({ name: 'Alice', age: 30 });
    });

    it('returns an empty object for invalid JSON', () => {
        const result = parseFormData('not json');
        expect(result).toEqual({});
    });

    it('returns an empty object for an empty string', () => {
        expect(parseFormData('')).toEqual({});
    });
});

// ---------------------------------------------------------------------------
// parseSharePointDate
// ---------------------------------------------------------------------------

describe('parseSharePointDate', () => {
    it('returns undefined for undefined input', () => {
        expect(parseSharePointDate(undefined)).toBeUndefined();
    });

    it('returns undefined for an empty string', () => {
        expect(parseSharePointDate('')).toBeUndefined();
    });

    it('returns undefined for an invalid date string', () => {
        expect(parseSharePointDate('not-a-date')).toBeUndefined();
    });

    it('returns a Date for a valid ISO date string', () => {
        const result = parseSharePointDate('2024-01-15T10:20:30.000Z');

        expect(result).toBeInstanceOf(Date);
        expect(result?.toISOString()).toBe('2024-01-15T10:20:30.000Z');
    });
});

// ---------------------------------------------------------------------------
// isFormRecord
// ---------------------------------------------------------------------------

describe('isFormRecord', () => {
    it('returns true for an object with a formData property', () => {
        const record = { id: 1, title: 'Test', formData: { field: 'value' } } as any;
        expect(isFormRecord(record)).toBe(true);
    });

    it('returns false for an IExtendedFile without a formData property', () => {
        const file = { name: 'doc.pdf', status: FileStatus.Processing, serverRelativeUrl: '/doc.pdf', uniqueId: '1', version: '1.0' } as any;
        expect(isFormRecord(file)).toBe(false);
    });
});

// ---------------------------------------------------------------------------
// getFileTypeTitle
// ---------------------------------------------------------------------------

describe('getFileTypeTitle', () => {
    it('returns the part after > for a hierarchical title', () => {
        expect(getFileTypeTitle('Parent > Child')).toBe('Child');
    });

    it('returns the original title when there is no >', () => {
        expect(getFileTypeTitle('SimpleTitle')).toBe('SimpleTitle');
    });

    it('returns an empty string for undefined input', () => {
        expect(getFileTypeTitle(undefined)).toBe('');
    });

    it('returns an empty string for an empty string', () => {
        expect(getFileTypeTitle('')).toBe('');
    });
});

// ---------------------------------------------------------------------------
// formatDateForSharePoint
// ---------------------------------------------------------------------------

describe('formatDateForSharePoint', () => {
    it('formats an afternoon date as M/D/YYYY h:mm PM', () => {
        const date = new Date(2024, 0, 15, 14, 30); // 15 Jan 2024, 14:30 local
        expect(formatDateForSharePoint(date)).toBe('1/15/2024 2:30 PM');
    });

    it('formats midnight as 12:XX AM', () => {
        const date = new Date(2024, 5, 20, 0, 5); // 20 Jun 2024, 00:05 local
        expect(formatDateForSharePoint(date)).toBe('6/20/2024 12:05 AM');
    });

    it('formats noon as 12:XX PM', () => {
        const date = new Date(2024, 11, 1, 12, 0); // 1 Dec 2024, 12:00 local
        expect(formatDateForSharePoint(date)).toBe('12/1/2024 12:00 PM');
    });

    it('pads minutes with a leading zero', () => {
        const date = new Date(2024, 2, 3, 9, 7); // 3 Mar 2024, 09:07 local
        expect(formatDateForSharePoint(date)).toBe('3/3/2024 9:07 AM');
    });
});

// ---------------------------------------------------------------------------
// expandRoles
// ---------------------------------------------------------------------------

describe('expandRoles', () => {
    it('expands "UUO" alias to all UUO internal groups', () => {
        const result = expandRoles('UUO');
        expect(result).toEqual(expect.arrayContaining([
            SPGroupInternal.UUO_BO,
            SPGroupInternal.UUO_DEP,
            SPGroupInternal.UUO_DIR,
            SPGroupInternal.UUO_RM,
            SPGroupInternal.UUO_RSMA,
        ]));
        expect(result.length).toBe(5);
    });

    it('passes through an unknown role as-is', () => {
        expect(expandRoles('CustomRole')).toEqual(['CustomRole']);
    });

    it('handles multiple roles separated by semicolons', () => {
        const result = expandRoles('AML; WO');
        expect(result).toEqual(expect.arrayContaining([
            SPGroupInternal.AML_DIR,
            SPGroupInternal.AML_S,
            SPGroupInternal.WO_DEP,
            SPGroupInternal.WO_DIR,
            SPGroupInternal.WO_S,
        ]));
        expect(result.length).toBe(5);
    });

    it('deduplicates groups that appear in multiple aliases', () => {
        const result = expandRoles('admin; admin');
        expect(result.length).toBe(new Set(result).size);
    });
});

// ---------------------------------------------------------------------------
// classifyUploader
// ---------------------------------------------------------------------------

describe('classifyUploader', () => {
    it('returns "client" for EXT_CLIFULL', () => {
        expect(classifyUploader([SPGroupExternal.EXT_CLIFULL])).toBe('client');
    });

    it('returns "client" for EXT_CLILOW', () => {
        expect(classifyUploader([SPGroupExternal.EXT_CLILOW])).toBe('client');
    });

    it('returns "external" for EXT_ADVISER', () => {
        expect(classifyUploader([SPGroupExternal.EXT_ADVISER])).toBe('external');
    });

    it('returns "external" for EXT_AUD', () => {
        expect(classifyUploader([SPGroupExternal.EXT_AUD])).toBe('external');
    });

    it('returns "bank" for an internal group', () => {
        expect(classifyUploader([SPGroupInternal.UUO_RM])).toBe('bank');
    });

    it('returns "client" when mixed with internal groups (client takes priority)', () => {
        expect(classifyUploader([SPGroupInternal.UUO_RM, SPGroupExternal.EXT_CLIFULL])).toBe('client');
    });

    it('returns an empty string for an empty array', () => {
        expect(classifyUploader([])).toBe('');
    });
});

// ---------------------------------------------------------------------------
// splitFolderPath
// ---------------------------------------------------------------------------

describe('splitFolderPath', () => {
    it('splits a forward-slash path into path and name', () => {
        expect(splitFolderPath('/sites/repo/folder')).toEqual({ path: '/sites/repo', name: 'folder' });
    });

    it('normalises backslashes to forward slashes', () => {
        expect(splitFolderPath('sites\\repo\\folder')).toEqual({ path: 'sites/repo', name: 'folder' });
    });

    it('returns an empty path when there is no slash', () => {
        expect(splitFolderPath('folder')).toEqual({ path: '', name: 'folder' });
    });

    it('collapses duplicate slashes', () => {
        expect(splitFolderPath('sites//repo//folder')).toEqual({ path: 'sites/repo', name: 'folder' });
    });
});

// ---------------------------------------------------------------------------
// parseNameAndDate
// ---------------------------------------------------------------------------

describe('parseNameAndDate', () => {
    it('parses a name and date from a correctly formatted string', () => {
        expect(parseNameAndDate('John Doe 15. 1. 2024 14:30')).toEqual({ name: 'John Doe', date: '15. 1. 2024 14:30' });
    });

    it('returns empty strings for a string that does not match the expected format', () => {
        expect(parseNameAndDate('no date here')).toEqual({ name: '', date: '' });
    });

    it('returns empty strings for undefined input', () => {
        expect(parseNameAndDate(undefined)).toEqual({ name: '', date: '' });
    });

    it('handles a date without extra spaces', () => {
        expect(parseNameAndDate('Jane Smith 3.3.2023 9:05')).toEqual({ name: 'Jane Smith', date: '3.3.2023 9:05' });
    });
});

// ---------------------------------------------------------------------------
// getLatestDate
// ---------------------------------------------------------------------------

describe('getLatestDate', () => {
    it('returns the latest date string from an array', () => {
        const dates = ['2024-01-01', '2024-06-15', '2023-12-31'];
        expect(getLatestDate(dates)).toBe('2024-06-15');
    });

    it('returns the only element when the array has one item', () => {
        expect(getLatestDate(['2024-03-01'])).toBe('2024-03-01');
    });

    it('returns the correct result when dates are already sorted descending', () => {
        expect(getLatestDate(['2025-01-01', '2024-06-01', '2023-01-01'])).toBe('2025-01-01');
    });
});

// ---------------------------------------------------------------------------
// normalizeDocType
// ---------------------------------------------------------------------------

describe('normalizeDocType', () => {
    it('removes diacritics and lowercases the string', () => {
        expect(normalizeDocType('Smlouva')).toBe('smlouva');
    });

    it('normalises hyphen spacing', () => {
        expect(normalizeDocType('test - value')).toBe('test-value');
    });

    it('removes all whitespace', () => {
        expect(normalizeDocType('hello world')).toBe('helloworld');
    });

    it('strips non-alphanumeric characters except hyphens', () => {
        expect(normalizeDocType('file (1)')).toBe('file1');
    });

    it('returns an empty string for undefined input', () => {
        expect(normalizeDocType(undefined)).toBe('');
    });

    it('handles Czech diacritics correctly', () => {
        expect(normalizeDocType('Úvěr')).toBe('uver');
    });
});

// ---------------------------------------------------------------------------
// transformPath
// ---------------------------------------------------------------------------

describe('transformPath', () => {
    it('encodes path segments and joins them with double-encoded slashes', () => {
        expect(transformPath('sites/test/folder')).toBe('sites%252ftest%252ffolder');
    });

    it('replaces spaces with + signs inside segments', () => {
        expect(transformPath('my folder/file name')).toBe('my+folder%252ffile+name');
    });

    it('handles an empty string', () => {
        expect(transformPath('')).toBe('');
    });

    it('encodes non-ASCII characters in segments', () => {
        const result = transformPath('složka/soubor');
        expect(result).toContain('%252f');
        expect(result).not.toContain('/');
    });
});

// ---------------------------------------------------------------------------
// openFileInDesktopApp
// ---------------------------------------------------------------------------

describe('openFileInDesktopApp', () => {
    it('sets window.location.href to ms-word URI for .docx files', () => {
        const original = window.location.href;
        Object.defineProperty(window, 'location', {
            value: { href: '' },
            writable: true,
        });

        openFileInDesktopApp('https://example.com/file.docx');

        expect(window.location.href).toContain('ms-word:ofe|u|');
        expect(window.location.href).toContain('file.docx');

        Object.defineProperty(window, 'location', { value: { href: original }, writable: true });
    });

    it('uses ms-excel URI for .xlsx files', () => {
        Object.defineProperty(window, 'location', { value: { href: '' }, writable: true });
        openFileInDesktopApp('https://example.com/report.xlsx');
        expect(window.location.href).toContain('ms-excel:ofe|u|');
    });

    it('uses ms-powerpoint URI for .pptx files', () => {
        Object.defineProperty(window, 'location', { value: { href: '' }, writable: true });
        openFileInDesktopApp('https://example.com/slides.pptx');
        expect(window.location.href).toContain('ms-powerpoint:ofe|u|');
    });

    it('falls back to anchor click when window.location.href throws', () => {
        Object.defineProperty(window, 'location', {
            get: () => { throw new Error('blocked'); },
            configurable: true,
        });
        const appendSpy = jest.spyOn(document.body, 'appendChild').mockImplementation(node => node);
        const removeSpy = jest.spyOn(document.body, 'removeChild').mockImplementation(node => node);

        expect(() => openFileInDesktopApp('https://example.com/file.docx')).not.toThrow();

        appendSpy.mockRestore();
        removeSpy.mockRestore();
        Object.defineProperty(window, 'location', { value: { href: '' }, writable: true });
    });
});

// ---------------------------------------------------------------------------
// downloadFile
// ---------------------------------------------------------------------------

describe('downloadFile', () => {
    it('fetches buffer, creates blob, and triggers download', async () => {
        const buffer = new ArrayBuffer(4);
        const mockApiClient = {
            files: { getFileToDownload: jest.fn().mockResolvedValue(buffer) }
        } as any;

        const createObjectURL = jest.fn().mockReturnValue('blob:fake-url');
        const revokeObjectURL = jest.fn();
        window.URL.createObjectURL = createObjectURL;
        window.URL.revokeObjectURL = revokeObjectURL;

        const clickSpy = jest.fn();
        jest.spyOn(document, 'createElement').mockReturnValueOnce({
            href: '', download: '', click: clickSpy,
            style: {},
        } as any);
        jest.spyOn(document.body, 'appendChild').mockImplementation(n => n);
        jest.spyOn(document.body, 'removeChild').mockImplementation(n => n);

        await downloadFile('report.pdf', '/sites/team/report.pdf', mockApiClient);

        expect(mockApiClient.files.getFileToDownload).toHaveBeenCalledWith('/sites/team/report.pdf');
        expect(createObjectURL).toHaveBeenCalled();
        expect(clickSpy).toHaveBeenCalled();

        jest.restoreAllMocks();
    });
});

// ---------------------------------------------------------------------------
// downloadFiles
// ---------------------------------------------------------------------------

describe('downloadFiles', () => {
    const makeApiClient = (): any => ({
        files: { getFileToDownload: jest.fn().mockResolvedValue(new ArrayBuffer(4)) }
    } as any);

    beforeEach(() => {
        window.URL.createObjectURL = jest.fn().mockReturnValue('blob:url');
        window.URL.revokeObjectURL = jest.fn();
        jest.spyOn(document.body, 'appendChild').mockImplementation(n => n);
        jest.spyOn(document.body, 'removeChild').mockImplementation(n => n);
        jest.spyOn(document, 'createElement').mockReturnValue({
            href: '', download: '', click: jest.fn(), style: {},
        } as any);
    });

    afterEach(() => jest.restoreAllMocks());

    it('calls downloadFile directly for a single file', async () => {
        const apiClient = makeApiClient();
        const files: IFile[] = [{ name: 'a.pdf', serverRelativeUrl: '/a.pdf' } as IFile];

        await downloadFiles(files, apiClient);

        expect(apiClient.files.getFileToDownload).toHaveBeenCalledWith('/a.pdf');
    });

    it('creates a zip for multiple files', async () => {
        const apiClient = makeApiClient();
        const files: IFile[] = [
            { name: 'a.pdf', serverRelativeUrl: '/a.pdf' } as IFile,
            { name: 'b.pdf', serverRelativeUrl: '/b.pdf' } as IFile,
        ];

        await downloadFiles(files, apiClient, 'archive.zip');

        expect(apiClient.files.getFileToDownload).toHaveBeenCalledTimes(2);
    });

    it('handles duplicate file names in zip by appending counter', async () => {
        const apiClient = makeApiClient();
        const files: IFile[] = [
            { name: 'doc.pdf', serverRelativeUrl: '/1/doc.pdf' } as IFile,
            { name: 'doc.pdf', serverRelativeUrl: '/2/doc.pdf' } as IFile,
        ];

        await downloadFiles(files, apiClient);

        expect(apiClient.files.getFileToDownload).toHaveBeenCalledTimes(2);
    });
});

// ---------------------------------------------------------------------------
// createCategoryOptions
// ---------------------------------------------------------------------------

describe('createCategoryOptions', () => {
    it('maps nodes to option objects with correct fields', () => {
        const nodes = [makeNode(1, 'Stage1', [], { [Language.CS]: 'Etapa 1', [Language.EN]: 'Stage 1' })];
        const result = createCategoryOptions(nodes, Language.EN);

        expect(result).toHaveLength(1);
        expect(result[0].value).toBe('1');
        expect(result[0].name).toBe('Stage 1');
        expect(result[0].expanded).toBe(true);
        expect(result[0].selectable).toBe(true);
    });

    it('uses node name as value when id is undefined', () => {
        const nodes = [makeNode(undefined, 'NoId')];
        const result = createCategoryOptions(nodes, Language.EN);
        expect(result[0].value).toBe('NoId');
        expect(result[0].selectable).toBe(false);
    });

    it('recursively maps child nodes into subChild', () => {
        const tree = [makeNode(1, 'Parent', [makeNode(2, 'Child')])];
        const result = createCategoryOptions(tree, Language.EN);
        expect(result[0].subChild).toHaveLength(1);
        expect(result[0].subChild[0].value).toBe('2');
    });

    it('returns empty array for empty input', () => {
        expect(createCategoryOptions([], Language.EN)).toEqual([]);
    });
});

// ---------------------------------------------------------------------------
// copyLink
// ---------------------------------------------------------------------------

describe('copyLink', () => {
    it('calls navigator.clipboard.writeText with the provided link', async () => {
        const writeText = jest.fn().mockResolvedValue(undefined);
        Object.defineProperty(navigator, 'clipboard', {
            value: { writeText },
            writable: true,
            configurable: true,
        });

        await copyLink('https://example.com/file.pdf');

        expect(writeText).toHaveBeenCalledWith('https://example.com/file.pdf');
    });
});

// ---------------------------------------------------------------------------
// getPhaseDesc
// ---------------------------------------------------------------------------

describe('getPhaseDesc', () => {
    const makeStrings = (): any => ({
        PhaseDesc: {
            BeforeApproval: {
                Request: 'Request desc',
                Preparation: 'Preparation desc',
                Acceptance: 'Acceptance desc',
                DocumentDelivery: 'Documents desc',
                Approval: 'Approval desc',
                ContractPreparation: 'Contract prep desc',
            },
            AfterApproval: {
                ContractForSignature: 'Contract sign desc',
                SignedContract: 'Signed desc',
                FirstDisbursement: 'First disb desc',
                SubsequentConditions: 'Subsequent cond desc',
                SubsequentDisbursement: 'Subsequent disb desc',
                Monitoring: 'Monitoring desc',
            },
        },
    } as any);

});