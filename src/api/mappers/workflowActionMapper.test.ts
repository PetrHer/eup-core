/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapWorkflowActions, mapScheduledAction, mapScheduledActions } from './workflowActionMapper';
import { Language } from '../../enums';

const rawAction = {
    Id: 1,
    Title: 'Approve',
    'Title_en': 'Approve EN',
    DocumentType: 'pdf\ndocx',
    DocumentStatus: 'accepted\nwaiting',
    Notification: 'You have a task',
    'Notification_en': 'You have a task EN',
    ResultStatus: 'approved',
    Publish: true,
    AllowComment: false,
    ActionGroups: 'UUO_RM',
    NotificationGroups: 'UUO_RM',
    Task: 'Review this',
    'Task_en': 'Review this EN',
    TaskGroups: 'UUO_RM',
    Phase: 'P1',
    ReadPermissions: 'read',
    WritePermissions: 'write',
    validity: 'valid',
    Reupload: false,
    FileLink: '/link',
    onMajorVersion: true,
    FileDirRef: '/sites/team/Lists/WorkflowActions/templates',
};

describe('mapWorkflowActions', () => {
    it('returns empty array for empty input', () => {
        expect(mapWorkflowActions([])).toEqual([]);
    });

    it('maps id as string from item Id', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.id).toBe('1');
    });

    it('maps localized title', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.title[Language.CS]).toBe('Approve');
        expect(result.title[Language.EN]).toBe('Approve EN');
    });

    it('splits DocumentType by newline into fileTypes array', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.fileTypes).toEqual(['pdf', 'docx']);
    });

    it('splits DocumentStatus by newline into statuses array', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.statuses).toEqual(['accepted', 'waiting']);
    });

    it('sets fileTypes to empty array when DocumentType is absent', () => {
        const [result] = mapWorkflowActions([{ ...rawAction, DocumentType: undefined }]);
        expect(result.fileTypes).toEqual([]);
    });

    it('maps notification as localized object when present', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.notification?.[Language.CS]).toBe('You have a task');
    });

    it('sets notification to undefined when absent', () => {
        const [result] = mapWorkflowActions([{ ...rawAction, Notification: undefined }]);
        expect(result.notification).toBeUndefined();
    });

    it('derives type from last segment of FileDirRef', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.type).toBe('templates');
    });

    it('maps resultStatus, publish, allowComment', () => {
        const [result] = mapWorkflowActions([rawAction]);
        expect(result.resultStatus).toBe('approved');
        expect(result.publish).toBe(true);
        expect(result.allowComment).toBe(false);
    });
});

describe('mapScheduledAction', () => {
    it('maps actionId and fileId as numbers', () => {
        const result = mapScheduledAction({ Id: 1, ActionId: '5', FileId: '10', Processed: undefined });
        expect(result.actionId).toBe(5);
        expect(result.fileId).toBe(10);
    });

    it('defaults actionId and fileId to 0 when absent', () => {
        const result = mapScheduledAction({ Id: 1 });
        expect(result.actionId).toBe(0);
        expect(result.fileId).toBe(0);
    });

    it('maps processed to Date when present', () => {
        const result = mapScheduledAction({ Id: 1, ActionId: '1', FileId: '2', Processed: '2024-01-01T00:00:00Z' });
        expect(result.processed).toBeInstanceOf(Date);
    });

    it('maps processed to undefined when absent', () => {
        const result = mapScheduledAction({ Id: 1 });
        expect(result.processed).toBeUndefined();
    });
});

describe('mapScheduledActions', () => {
    it('returns empty Map for empty array', () => {
        expect(mapScheduledActions([])).toBeInstanceOf(Map);
        expect(mapScheduledActions([]).size).toBe(0);
    });

    it('indexes actions by fileId', () => {
        const map = mapScheduledActions([
            { Id: 1, ActionId: '2', FileId: '42', Processed: undefined },
        ]);
        expect(map.has(42)).toBe(true);
        expect(map.get(42)?.actionId).toBe(2);
    });

    it('skips items with fileId of 0', () => {
        const map = mapScheduledActions([{ Id: 1, ActionId: '1', FileId: '0' }]);
        expect(map.size).toBe(0);
    });

    it('indexes multiple items by their fileIds', () => {
        const map = mapScheduledActions([
            { Id: 1, ActionId: '1', FileId: '10' },
            { Id: 2, ActionId: '2', FileId: '20' },
        ]);
        expect(map.size).toBe(2);
        expect(map.has(10)).toBe(true);
        expect(map.has(20)).toBe(true);
    });
});
