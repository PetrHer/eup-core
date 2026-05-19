/* eslint-disable no-undef, @typescript-eslint/explicit-function-return-type */
import { mapNotifications } from './notificationMapper';
import { Language } from '../../enums';

const baseItem = {
    Id: 1,
    Title: 'Notification 1',
    Title_en: 'Notification 1 EN',
    Task: null,
    Completed: false,
    Category: {
        Title: 'Cat CS',
        'Category_en': 'Cat EN',
        Stage: 'S1',
        'Stage_en': 'S1en',
        Phase: 'P1',
        'Phase_en': 'P1en',
        Id: 10,
    },
    Created: '2024-01-15T12:00:00Z',
    Author: { Title: 'John Doe' },
    ReadId: null,
    TargetId: 99,
    Status: 'active',
    Comment: 'A comment',
    Modified: '2024-01-16T12:00:00Z',
    Editor: { Title: 'Jane Smith' },
    Link: 'https://example.com',
    UniqueId: 'uid-1',
    FileLink: 'https://file.link',
};

describe('mapNotifications', () => {
    it('returns empty array for empty input', () => {
        expect(mapNotifications([])).toEqual([]);
    });

    it('maps id, title and targetId', () => {
        const [result] = mapNotifications([baseItem]);
        expect(result.id).toBe(1);
        expect(result.title[Language.CS]).toBe('Notification 1');
        expect(result.title[Language.EN]).toBe('Notification 1 EN');
        expect(result.targetId).toBe(99);
    });

    it('maps author from Author.Title', () => {
        const [result] = mapNotifications([baseItem]);
        expect(result.author).toBe('John Doe');
    });

    it('maps category localization from Category object', () => {
        const [result] = mapNotifications([baseItem]);
        expect(result.category[Language.CS]).toBe('Cat CS');
        expect(result.stage[Language.CS]).toBe('S1');
        expect(result.phase[Language.CS]).toBe('P1');
    });

    it('sets empty strings for category fields when Category is null', () => {
        const [result] = mapNotifications([{ ...baseItem, Category: null }]);
        expect(result.category[Language.CS]).toBe('');
        expect(result.stage[Language.CS]).toBe('');
    });

    it('marks notification as unread when ReadId is null and Task is false', () => {
        const [result] = mapNotifications([{ ...baseItem, ReadId: null, Task: false }], 5);
        expect(result.unread).toBe(true);
    });

    it('marks notification as read when user id is in ReadId', () => {
        const [result] = mapNotifications([{ ...baseItem, ReadId: [5], Task: false }], 5);
        expect(result.unread).toBe(false);
    });

    it('maps editor from Editor.Title', () => {
        const [result] = mapNotifications([baseItem]);
        expect(result.editor).toBe('Jane Smith');
    });

    it('sets editor to empty string when Editor is null', () => {
        const [result] = mapNotifications([{ ...baseItem, Editor: null }]);
        expect(result.editor).toBe('');
    });

    it('maps attachments from attachmentsMap', () => {
        const attachmentsMap = new Map([[1, [{ fileName: 'file.pdf', url: '/file.pdf' }]]]);
        const [result] = mapNotifications([baseItem], undefined, attachmentsMap);
        expect(result.attachments).toHaveLength(1);
        expect(result.attachments?.[0].fileName).toBe('file.pdf');
    });

    it('returns empty attachments array when no map entry', () => {
        const [result] = mapNotifications([baseItem]);
        expect(result.attachments).toEqual([]);
    });

    it('maps multiple items', () => {
        const result = mapNotifications([baseItem, { ...baseItem, Id: 2, Title: 'Notif 2' }]);
        expect(result).toHaveLength(2);
    });
});
