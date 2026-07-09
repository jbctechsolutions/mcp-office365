/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Adapter that bridges GraphRepository's xxxAsync() methods
 * to the IMailboxRepository interface.
 */

import type { IMailboxRepository, FolderRow, EmailRow } from '../database/repository.js';
import type { GraphRepository } from './repository.js';

export class GraphMailboxAdapter implements IMailboxRepository {
  constructor(private readonly graph: GraphRepository) {}

  // Read
  getEmail(id: string | number): Promise<EmailRow | undefined> {
    return this.graph.getEmailAsync(id);
  }

  getFolder(id: number): Promise<FolderRow | undefined> {
    return this.graph.getFolderAsync(id);
  }

  // Email organization
  moveEmail(emailId: string | number, destinationFolderId: number): Promise<void> {
    return this.graph.moveEmailAsync(emailId, destinationFolderId);
  }

  deleteEmail(emailId: string | number): Promise<void> {
    return this.graph.deleteEmailAsync(emailId);
  }

  archiveEmail(emailId: string | number): Promise<void> {
    return this.graph.archiveEmailAsync(emailId);
  }

  junkEmail(emailId: string | number): Promise<void> {
    return this.graph.junkEmailAsync(emailId);
  }

  markEmailRead(emailId: string | number, isRead: boolean): Promise<void> {
    return this.graph.markEmailReadAsync(emailId, isRead);
  }

  setEmailFlag(emailId: string | number, flagStatus: number): Promise<void> {
    return this.graph.setEmailFlagAsync(emailId, flagStatus);
  }

  setEmailCategories(emailId: string | number, categories: string[]): Promise<void> {
    return this.graph.setEmailCategoriesAsync(emailId, categories);
  }

  setEmailImportance(emailId: string | number, importance: string): Promise<void> {
    return this.graph.setEmailImportanceAsync(emailId, importance);
  }

  // Folder management
  createFolder(name: string, parentFolderId?: number): Promise<FolderRow> {
    return this.graph.createFolderAsync(name, parentFolderId);
  }

  deleteFolder(folderId: number): Promise<void> {
    return this.graph.deleteFolderAsync(folderId);
  }

  renameFolder(folderId: number, newName: string): Promise<void> {
    return this.graph.renameFolderAsync(folderId, newName);
  }

  moveFolder(folderId: number, destinationParentId: number): Promise<void> {
    return this.graph.moveFolderAsync(folderId, destinationParentId);
  }

  emptyFolder(folderId: number): Promise<void> {
    return this.graph.emptyFolderAsync(folderId);
  }
}
