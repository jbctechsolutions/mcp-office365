/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Builds Graph beta plannerTaskChatMessage request bodies with @mention markup.
 *
 * Graph expects sanitized HTML with Skype mention spans plus a parallel
 * `mentions[]` array. Agents pass plain text + user ids; this helper wires both.
 */

export interface PlannerTaskChatMention {
  mentioned: string;
  position: number;
  mentionType: 'user';
}

/**
 * Builds `content` + `mentions` for create/update planner task chat messages.
 * When `mention_user_ids` is empty, `content` is sent as-is (plain text or HTML).
 */
export function buildPlannerTaskMessagePayload(
  content: string,
  mentionUserIds: readonly string[],
): { content: string; mentions: PlannerTaskChatMention[] } {
  if (mentionUserIds.length === 0) {
    return { content, mentions: [] };
  }

  const mentions: PlannerTaskChatMention[] = mentionUserIds.map((mentioned, position) => ({
    mentioned,
    position,
    mentionType: 'user',
  }));

  // Caller supplied hand-crafted mention HTML — only attach the mentions array.
  if (content.includes('<span itemid=')) {
    return { content, mentions };
  }

  const mentionSpans = mentions
    .map((m) => `<span itemid="${m.position}" itemtype="https://schema.skype.com/Mention/Person"></span>`)
    .join('');

  const escaped = content
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  return {
    content: `<div>${mentionSpans} ${escaped}</div>`,
    mentions,
  };
}
