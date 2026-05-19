/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapItemAsFile, mapFile, mapItemsAsFiles, mapFolders, mapFileVersions, mapFileTypes } from './fileFolderMapper';
import { FileStatus, Language } from '../../enums';

const rawFileItem = {
    Id: 1,
    FileLeafRef: 'document.pdf',
    FileRef: '/sites/team/docs/document.pdf',
    Status: FileStatus.Accepted,
    UniqueId: 'uid-abc',
    OData__UIVersionString: '1.0',
    DocumentType: {
        Id: 5,
        Title: 'Contract CS',
        'Title_en': 'Contract EN',
        FolderPath: '/path',
        FolderPathAfter: '/after',
        ValidityStatuses: 'accepted;waiting',
        FolderPathInternal: '/internal',
        FolderPathInternalAfter: '/internal-after',
    },
    Created: '2024-01-01T00:00:00Z',
    StrukturaDokumentaceId: 10,
    Documentation: true,
    UploadedBy: 'John',
    ContractValidity: 'valid',
    ValidTo: '2025-01-01T00:00:00Z',
    QAId: 'qa-1',
    Author: { Title: 'Jane Smith' },
    Modified: '2024-02-01T00:00:00Z',
};

describe('mapItemAsFile', () => {
    it('maps name, serverRelativeUrl, uniqueId', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.name).toBe('document.pdf');
        expect(result.serverRelativeUrl).toBe('/sites/team/docs/document.pdf');
        expect(result.uniqueId).toBe('uid-abc');
    });

    it('maps status from Status field', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.status).toBe(FileStatus.Accepted);
    });

    it('defaults status to FileStatus.Processing when Status is absent', () => {
        const result = mapItemAsFile({ ...rawFileItem, Status: undefined });
        expect(result.status).toBe(FileStatus.Processing);
    });

    it('maps type with localized titles when DocumentType is present', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.type).toBeDefined();
        expect(result.type?.id).toBe(5);
    });

    it('sets type to undefined when DocumentType is absent', () => {
        const result = mapItemAsFile({ ...rawFileItem, DocumentType: undefined });
        expect(result.type).toBeUndefined();
    });

    it('maps created and modified as Date instances', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.created).toBeInstanceOf(Date);
        expect(result.modified).toBeInstanceOf(Date);
    });

    it('maps author from Author.Title', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.author).toBe('Jane Smith');
    });

    it('sets author to empty string when Author is absent', () => {
        const result = mapItemAsFile({ ...rawFileItem, Author: undefined });
        expect(result.author).toBe('');
    });

    it('splits ValidityStatuses by semicolon', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.type?.validityStatuses).toEqual(['accepted', 'waiting']);
    });

    it('maps qaId, nodeId, documentation, provided', () => {
        const result = mapItemAsFile(rawFileItem);
        expect(result.qaId).toBe('qa-1');
        expect(result.nodeId).toBe(10);
        expect(result.documentation).toBe(true);
        expect(result.provided).toBe('John');
    });
});

describe('mapFile', () => {
    it('maps core fields from raw file object', () => {
        const raw = {
            FileLeafRef: 'file.pdf',
            FileRef: '/sites/team/file.pdf',
            Status: FileStatus.Waiting,
            UniqueId: 'uid-xyz',
            OData__UIVersionString: '2.0',
            StrukturaDokumentaceId: 5,
            DocumentTypeId: 3,
        };
        const result = mapFile(raw);
        expect(result.name).toBe('file.pdf');
        expect(result.serverRelativeUrl).toBe('/sites/team/file.pdf');
        expect(result.status).toBe(FileStatus.Waiting);
        expect(result.uniqueId).toBe('uid-xyz');
        expect(result.typeId).toBe(3);
    });
});

describe('mapItemsAsFiles', () => {
    it('returns empty array for empty input', () => {
        expect(mapItemsAsFiles([])).toEqual([]);
    });

    it('maps multiple items', () => {
        const result = mapItemsAsFiles([rawFileItem, { ...rawFileItem, Id: 2 }]);
        expect(result).toHaveLength(2);
    });
});

describe('mapFolders', () => {
    it('returns empty array for empty input', () => {
        expect(mapFolders([])).toEqual([]);
    });

    it('returns top-level folders (those not inside another folder in data)', () => {
        const data = [
            { FileRef: '/root/folderA', FileDirRef: '/root', FileLeafRef: 'folderA', Name_en: 'Folder A EN' },
        ];
        const result = mapFolders(data);
        expect(result).toHaveLength(1);
        expect(result[0].name).toBe('folderA');
    });

    it('nests child folders inside parent', () => {
        const data = [
            { FileRef: '/root/parent', FileDirRef: '/root', FileLeafRef: 'parent', Name_en: 'Parent EN' },
            { FileRef: '/root/parent/child', FileDirRef: '/root/parent', FileLeafRef: 'child', Name_en: 'Child EN' },
        ];
        const result = mapFolders(data);
        expect(result).toHaveLength(1);
        expect(result[0].folders).toHaveLength(1);
        expect(result[0].folders[0].name).toBe('child');
    });

    it('maps localization from FileLeafRef and Name_en', () => {
        const data = [
            { FileRef: '/root/folderA', FileDirRef: '/root', FileLeafRef: 'folderA', Name_en: 'Folder A EN' },
        ];
        const result = mapFolders(data);
        expect(result[0].localization[Language.CS]).toBe('folderA');
        expect(result[0].localization[Language.EN]).toBe('Folder A EN');
    });
});

describe('mapFileVersions', () => {
    it('returns empty array for empty input', () => {
        expect(mapFileVersions([], 'teamSite')).toEqual([]);
    });

    it('maps editor, status, versionLabel', () => {
        const raw = {
            FileRef: '/sites/teamSite/file.pdf',
            Editor: { LookupValue: 'Jane' },
            Modified: '2024-01-01',
            Status: FileStatus.Accepted,
            VersionLabel: '1.0',
            IsCurrentVersion: true,
            VersionId: 512,
        };
        const [result] = mapFileVersions([raw], 'teamSite');
        expect(result.editor).toBe('Jane');
        expect(result.status).toBe(FileStatus.Accepted);
        expect(result.versionLabel).toBe('1.0');
        expect(result.isCurrentVersion).toBe(true);
    });

    it('uses versioned URL for non-current versions', () => {
        const raw = {
            FileRef: '/sites/teamSite/file.pdf',
            Editor: { LookupValue: 'Jane' },
            Modified: '2024-01-01',
            Status: FileStatus.Accepted,
            VersionLabel: '0.1',
            IsCurrentVersion: false,
            VersionId: 256,
        };
        const [result] = mapFileVersions([raw], 'teamSite');
        expect(result.relativeUrl).toContain('_vti_history');
    });
});

describe('mapFileTypes', () => {
    it('returns empty array for empty input', () => {
        expect(mapFileTypes([])).toEqual([]);
    });

    it('filters out items whose Title contains ">"', () => {
        const data = [
            { Id: 1, Title: 'Parent Type', 'Title_en': 'Parent EN', Permissions: undefined, ValidityStatuses: undefined },
            { Id: 2, Title: 'Parent Type > Child', 'Title_en': 'Child EN', Permissions: undefined, ValidityStatuses: undefined },
        ];
        const result = mapFileTypes(data);
        expect(result).toHaveLength(1);
        expect(result[0].id).toBe(1);
    });

    it('maps localized title for top-level types', () => {
        const data = [{ Id: 1, Title: 'Contract', 'Title_en': 'Contract EN', Permissions: undefined, ValidityStatuses: undefined }];
        const [result] = mapFileTypes(data);
        expect(result.loc[Language.CS]).toBe('Contract');
        expect(result.loc[Language.EN]).toBe('Contract EN');
    });

    it('splits ValidityStatuses by semicolon', () => {
        const data = [{ Id: 1, Title: 'T', 'Title_en': 'T', Permissions: undefined, ValidityStatuses: 'a;b;c' }];
        const [result] = mapFileTypes(data);
        expect(result.validityStatuses).toEqual(['a', 'b', 'c']);
    });

    it('returns empty validityStatuses array when field is absent', () => {
        const data = [{ Id: 1, Title: 'T', 'Title_en': 'T', Permissions: undefined, ValidityStatuses: undefined }];
        const [result] = mapFileTypes(data);
        expect(result.validityStatuses).toEqual([]);
    });
});
