/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Mail sender interface for sending emails via AppleScript.
 */

import { existsSync } from 'fs';
import { executeAppleScriptOrThrow } from './executor.js';
import * as scripts from './scripts.js';
import { parseSendEmailResult } from './parser.js';
import { AppleScriptError, AttachmentNotFoundError, MailSendError } from '../utils/errors.js';

export interface Attachment {
  readonly path: string;
  readonly name?: string;
}

export interface SendEmailParams {
  readonly to: readonly string[];
  readonly subject: string;
  readonly body: string;
  readonly bodyType: 'plain' | 'html';
  readonly cc?: readonly string[];
  readonly bcc?: readonly string[];
  readonly replyTo?: string;
  readonly attachments?: readonly Attachment[];
  readonly accountId?: number;
}

export interface SentEmail {
  readonly messageId: string;
  readonly sentAt: string;
}

export interface IMailSender {
  sendEmail(params: SendEmailParams): SentEmail;
}

export class AppleScriptMailSender implements IMailSender {
  sendEmail(params: SendEmailParams): SentEmail {
    // Validate attachments exist
    if (params.attachments != null) {
      for (const attachment of params.attachments) {
        if (!existsSync(attachment.path)) {
          throw new AttachmentNotFoundError(attachment.path);
        }
      }
    }

    let scriptParams: scripts.SendEmailParams = {
      to: params.to,
      subject: params.subject,
      body: params.body,
      bodyType: params.bodyType,
    };

    if (params.cc != null) scriptParams = { ...scriptParams, cc: params.cc };
    if (params.bcc != null) scriptParams = { ...scriptParams, bcc: params.bcc };
    if (params.replyTo != null) scriptParams = { ...scriptParams, replyTo: params.replyTo };
    if (params.attachments != null) scriptParams = { ...scriptParams, attachments: params.attachments };
    if (params.accountId != null) scriptParams = { ...scriptParams, accountId: params.accountId };

    const script = scripts.sendEmail(scriptParams);

    const output = executeAppleScriptOrThrow(script);
    const result = parseSendEmailResult(output);

    if (result == null) {
      throw new AppleScriptError('Failed to parse send email response');
    }

    if (!result.success) {
      throw new MailSendError(result.error ?? 'Unknown error');
    }

    return {
      messageId: result.messageId!,
      sentAt: result.sentAt!,
    };
  }
}

export function createMailSender(): IMailSender {
  return new AppleScriptMailSender();
}
