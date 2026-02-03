/**
 * Unit tests for AppleScript template generation.
 */

import { describe, it, expect } from 'vitest';
import { respondToEvent, deleteEvent, sendEmail } from '../../../src/applescript/scripts.js';

describe('respondToEvent', () => {
  it('should generate accept script with comment', () => {
    const script = respondToEvent({
      eventId: 123,
      response: 'accept',
      sendResponse: true,
      comment: 'I will be there',
    });

    expect(script).toContain('calendar event id 123');
    expect(script).toContain('accept');
    expect(script).toContain('I will be there');
  });

  it('should generate decline script without sending response', () => {
    const script = respondToEvent({
      eventId: 456,
      response: 'decline',
      sendResponse: false,
    });

    expect(script).toContain('calendar event id 456');
    expect(script).toContain('decline');
  });

  it('should generate tentative accept script', () => {
    const script = respondToEvent({
      eventId: 789,
      response: 'tentative',
      sendResponse: true,
    });

    expect(script).toContain('calendar event id 789');
    expect(script).toContain('tentative');
  });
});

describe('deleteEvent', () => {
  it('should generate script for single instance', () => {
    const script = deleteEvent({ eventId: 123, applyTo: 'this_instance' });
    expect(script).toContain('calendar event id 123');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting single instance');
  });

  it('should generate script for all in series', () => {
    const script = deleteEvent({ eventId: 456, applyTo: 'all_in_series' });
    expect(script).toContain('calendar event id 456');
    expect(script).toContain('delete');
    expect(script).toContain('Deleting entire series');
  });

  it('should include success output format', () => {
    const script = deleteEvent({ eventId: 789, applyTo: 'this_instance' });
    expect(script).toContain('success{{=}}true');
    expect(script).toContain('eventId{{=}}');
  });
});

describe('sendEmail', () => {
  it('should generate plain text email with single recipient', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test Subject',
      body: 'Test body',
      bodyType: 'plain',
    });

    expect(script).toContain('Test Subject');
    expect(script).toContain('Test body');
    expect(script).toContain('plain text content');
    expect(script).toContain('test@example.com');
  });

  it('should generate HTML email', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'HTML Test',
      body: '<p>HTML body</p>',
      bodyType: 'html',
    });

    expect(script).toContain('HTML Test');
    expect(script).toContain('html content');
    expect(script).toContain('HTML body');
  });

  it('should include CC and BCC recipients', () => {
    const script = sendEmail({
      to: ['to@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      cc: ['cc1@example.com', 'cc2@example.com'],
      bcc: ['bcc@example.com'],
    });

    expect(script).toContain('cc1@example.com');
    expect(script).toContain('cc2@example.com');
    expect(script).toContain('bcc@example.com');
    expect(script).toContain('recipient cc');
    expect(script).toContain('recipient bcc');
  });

  it('should include reply-to address', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      replyTo: 'reply@example.com',
    });

    expect(script).toContain('reply to of newMessage to "reply@example.com"');
  });

  it('should include attachments', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      attachments: [
        { path: '/path/to/file.pdf' },
        { path: '/path/to/image.png', name: 'screenshot.png' },
      ],
    });

    expect(script).toContain('POSIX file "/path/to/file.pdf"');
    expect(script).toContain('POSIX file "/path/to/image.png"');
    expect(script).toContain('make new attachment');
  });

  it('should include account ID', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
      accountId: 123,
    });

    expect(script).toContain('account id 123');
  });

  it('should handle special characters in subject and body', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test "quotes" and \\backslash',
      body: 'Body with "quotes"',
      bodyType: 'plain',
    });

    expect(script).toContain('Test');
    expect(script).toContain('Body with');
  });

  it('should include success output format', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
    });

    expect(script).toContain('success{{=}}true');
    expect(script).toContain('messageId{{=}}');
    expect(script).toContain('sentAt{{=}}');
  });

  it('should include error handling', () => {
    const script = sendEmail({
      to: ['test@example.com'],
      subject: 'Test',
      body: 'Body',
      bodyType: 'plain',
    });

    expect(script).toContain('on error errMsg');
    expect(script).toContain('success{{=}}false');
    expect(script).toContain('error{{=}}');
  });
});
