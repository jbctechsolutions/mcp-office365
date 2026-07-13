/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

import { describe, it, expect } from 'vitest';
import { buildPlannerTaskMessagePayload } from '../../../src/graph/planner-task-message-payload.js';

describe('buildPlannerTaskMessagePayload', () => {
  it('returns plain content when there are no mentions', () => {
    expect(buildPlannerTaskMessagePayload('Hello team', [])).toEqual({
      content: 'Hello team',
      mentions: [],
    });
  });

  it('wraps plain text with mention spans and a mentions array', () => {
    const result = buildPlannerTaskMessagePayload('Please review', ['user-1', 'user-2']);

    expect(result.mentions).toEqual([
      { mentioned: 'user-1', position: 0, mentionType: 'user' },
      { mentioned: 'user-2', position: 1, mentionType: 'user' },
    ]);
    expect(result.content).toContain('<span itemid="0" itemtype="https://schema.skype.com/Mention/Person"></span>');
    expect(result.content).toContain('<span itemid="1" itemtype="https://schema.skype.com/Mention/Person"></span>');
    expect(result.content).toContain('Please review');
  });

  it('passes through hand-crafted mention HTML unchanged', () => {
    const html = '<div><span itemid="0" itemtype="https://schema.skype.com/Mention/Person"></span> Hi</div>';
    const result = buildPlannerTaskMessagePayload(html, ['user-1']);

    expect(result.content).toBe(html);
    expect(result.mentions).toHaveLength(1);
  });
});
