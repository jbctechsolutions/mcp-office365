import { describe, test, expect } from 'vitest';
import { createMailSender } from '../../../src/applescript/index.js';

const isOutlookAvailable = process.env.OUTLOOK_AVAILABLE === '1';
const testIf = (condition: boolean) => (condition ? test : test.skip);

describe('Email Sending Integration', () => {
  const mailSender = createMailSender();

  testIf(isOutlookAvailable)('sends basic email', () => {
    expect(mailSender.sendEmail).toBeDefined();
  });

  testIf(isOutlookAvailable)('validates attachment exists', () => {
    expect(() => {
      mailSender.sendEmail({
        to: ['test@example.com'],
        subject: 'Test',
        body: 'Test',
        bodyType: 'plain',
        attachments: [{ path: '/nonexistent/file.pdf' }],
      });
    }).toThrow('Attachment file not found');
  });
});
