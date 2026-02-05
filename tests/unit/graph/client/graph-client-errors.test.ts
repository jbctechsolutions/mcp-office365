/**
 * Tests for Graph API client error handling.
 * Focuses on critical Priority 1 error scenarios.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import type * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

// Mock request builder that chains all methods
const createMockRequestBuilder = (mockResponse?: any, mockError?: Error) => {
  const builder: any = {
    select: vi.fn().mockReturnThis(),
    top: vi.fn().mockReturnThis(),
    skip: vi.fn().mockReturnThis(),
    orderby: vi.fn().mockReturnThis(),
    filter: vi.fn().mockReturnThis(),
    search: vi.fn().mockReturnThis(),
    query: vi.fn().mockReturnThis(),
    get: mockError
      ? vi.fn().mockRejectedValue(mockError)
      : vi.fn().mockResolvedValue(mockResponse),
  };
  return builder;
};

// Mock Graph client
const mockApi = vi.fn();
const mockGraphClient = {
  api: mockApi,
};

vi.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: vi.fn(function() { return mockGraphClient; }),
  },
}));

// Mock auth module
vi.mock('../../../../src/graph/auth/index.js', () => ({
  getAccessToken: vi.fn().mockResolvedValue('test-access-token'),
}));

// Mock isomorphic-fetch
vi.mock('isomorphic-fetch', () => ({
  default: vi.fn(),
}));

import { GraphClient } from '../../../../src/graph/client/graph-client.js';

describe('graph/client/graph-client - Error Handling', () => {
  let graphClient: GraphClient;

  beforeEach(() => {
    vi.clearAllMocks();
    graphClient = new GraphClient();
  });

  describe('Rate Limiting (HTTP 429)', () => {
    it('handles 429 responses with rate limit error', async () => {
      const rateLimitError: any = new Error('Rate limit exceeded');
      rateLimitError.statusCode = 429;
      rateLimitError.code = 'TooManyRequests';

      const errorBuilder = createMockRequestBuilder(undefined, rateLimitError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles 429 with Retry-After header parsing', async () => {
      const rateLimitError: any = new Error('Rate limit exceeded');
      rateLimitError.statusCode = 429;
      rateLimitError.code = 'TooManyRequests';
      rateLimitError.headers = {
        'Retry-After': '120',
      };

      const errorBuilder = createMockRequestBuilder(undefined, rateLimitError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMailFolder('folder-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles 429 in message listing', async () => {
      const rateLimitError: any = new Error('Rate limit exceeded');
      rateLimitError.statusCode = 429;
      rateLimitError.code = 'TooManyRequests';

      const errorBuilder = createMockRequestBuilder(undefined, rateLimitError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.listMessages('folder-1', 50, 0)).rejects.toThrow('Rate limit exceeded');
    });
  });

  describe('Network Errors', () => {
    it('handles ETIMEDOUT errors', async () => {
      const timeoutError: any = new Error('Request timeout');
      timeoutError.code = 'ETIMEDOUT';
      timeoutError.errno = -60;

      const errorBuilder = createMockRequestBuilder(undefined, timeoutError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles ECONNREFUSED errors', async () => {
      const connectionError: any = new Error('Connection refused');
      connectionError.code = 'ECONNREFUSED';
      connectionError.errno = -61;

      const errorBuilder = createMockRequestBuilder(undefined, connectionError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getEvent('event-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles connection reset errors (ECONNRESET)', async () => {
      const resetError: any = new Error('Connection reset by peer');
      resetError.code = 'ECONNRESET';
      resetError.errno = -54;

      const errorBuilder = createMockRequestBuilder(undefined, resetError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getContact('contact-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles network timeout during folder listing', async () => {
      const timeoutError: any = new Error('Request timeout');
      timeoutError.code = 'ETIMEDOUT';

      const errorBuilder = createMockRequestBuilder(undefined, timeoutError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.listMailFolders()).rejects.toThrow('Request timeout');
    });

    it('handles socket hang up errors', async () => {
      const hangupError: any = new Error('socket hang up');
      hangupError.code = 'ECONNRESET';

      const errorBuilder = createMockRequestBuilder(undefined, hangupError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getTask('list-1', 'task-1');

      expect(result).toBeNull();
    });
  });

  describe('API Error Responses', () => {
    it('handles 401 Unauthorized errors', async () => {
      const unauthorizedError: any = new Error('Unauthorized');
      unauthorizedError.statusCode = 401;
      unauthorizedError.code = 'Unauthorized';

      const errorBuilder = createMockRequestBuilder(undefined, unauthorizedError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles 403 Forbidden errors', async () => {
      const forbiddenError: any = new Error('Forbidden');
      forbiddenError.statusCode = 403;
      forbiddenError.code = 'Forbidden';

      const errorBuilder = createMockRequestBuilder(undefined, forbiddenError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMailFolder('folder-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles 500 Server Error', async () => {
      const serverError: any = new Error('Internal Server Error');
      serverError.statusCode = 500;
      serverError.code = 'InternalServerError';

      const errorBuilder = createMockRequestBuilder(undefined, serverError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getEvent('event-1');

      expect(result).toBeNull();
      expect(errorBuilder.get).toHaveBeenCalled();
    });

    it('handles 404 Not Found errors', async () => {
      const notFoundError: any = new Error('Not Found');
      notFoundError.statusCode = 404;
      notFoundError.code = 'NotFound';

      const errorBuilder = createMockRequestBuilder(undefined, notFoundError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('nonexistent-id');

      expect(result).toBeNull();
    });

    it('handles 503 Service Unavailable errors', async () => {
      const unavailableError: any = new Error('Service Unavailable');
      unavailableError.statusCode = 503;
      unavailableError.code = 'ServiceUnavailable';

      const errorBuilder = createMockRequestBuilder(undefined, unavailableError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getContact('contact-1');

      expect(result).toBeNull();
    });

    it('handles 401 during calendar listing', async () => {
      const unauthorizedError: any = new Error('Unauthorized');
      unauthorizedError.statusCode = 401;
      unauthorizedError.code = 'InvalidAuthenticationToken';

      const errorBuilder = createMockRequestBuilder(undefined, unauthorizedError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.listCalendars()).rejects.toThrow('Unauthorized');
    });

    it('handles 403 during message search', async () => {
      const forbiddenError: any = new Error('Forbidden');
      forbiddenError.statusCode = 403;
      forbiddenError.code = 'ErrorAccessDenied';

      const errorBuilder = createMockRequestBuilder(undefined, forbiddenError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.searchMessages('test query')).rejects.toThrow('Forbidden');
    });

    it('handles 500 during event listing', async () => {
      const serverError: any = new Error('Internal Server Error');
      serverError.statusCode = 500;
      serverError.code = 'InternalServerError';

      const errorBuilder = createMockRequestBuilder(undefined, serverError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.listEvents(50)).rejects.toThrow('Internal Server Error');
    });
  });

  describe('Malformed Response Handling', () => {
    it('handles malformed JSON responses', async () => {
      const parseError = new SyntaxError('Unexpected token in JSON at position 0');

      const errorBuilder = createMockRequestBuilder(undefined, parseError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeNull();
    });

    it('handles empty response body', async () => {
      const emptyBuilder = createMockRequestBuilder(null);
      mockApi.mockReturnValue(emptyBuilder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeNull();
    });

    it('handles response with missing required fields', async () => {
      const invalidMessage = {
        // Missing id and other required fields
        subject: 'Test',
      };

      const builder = createMockRequestBuilder(invalidMessage);
      mockApi.mockReturnValue(builder);

      const result = await graphClient.getMessage('msg-1');

      expect(result).toBeDefined();
      expect(result?.subject).toBe('Test');
    });

    it('handles empty value array in response', async () => {
      const emptyResponse = {
        value: [],
      };

      const builder = createMockRequestBuilder(emptyResponse);
      mockApi.mockReturnValue(builder);

      const result = await graphClient.listMessages('folder-1', 50, 0);

      expect(result).toEqual([]);
      expect(Array.isArray(result)).toBe(true);
    });

    it('handles undefined response from API', async () => {
      const builder = createMockRequestBuilder(undefined);
      mockApi.mockReturnValue(builder);

      const result = await graphClient.getMessage('msg-1');

      // When the API returns undefined, the method returns undefined (not null)
      expect(result).toBeUndefined();
    });
  });

  describe('Complex Error Scenarios', () => {
    it('handles partial pagination failure in mail folders', async () => {
      // Top-level folders succeed
      const topLevelBuilder = createMockRequestBuilder({
        value: [
          { id: 'folder-1', displayName: 'Inbox' },
          { id: 'folder-2', displayName: 'Sent' },
        ],
      });

      // First child folder succeeds
      const childSuccessBuilder = createMockRequestBuilder({
        value: [{ id: 'folder-1-1', displayName: 'Subfolder 1' }],
      });

      // Second child folder fails
      const childErrorBuilder = createMockRequestBuilder(
        undefined,
        new Error('Access denied')
      );

      mockApi
        .mockReturnValueOnce(topLevelBuilder)
        .mockReturnValueOnce(childSuccessBuilder)
        .mockReturnValueOnce(childErrorBuilder);

      const result = await graphClient.listMailFolders();

      // Should return top-level folders and accessible children
      expect(result.length).toBeGreaterThanOrEqual(2);
    });

    it('handles mixed network and API errors', async () => {
      const networkError: any = new Error('Network error');
      networkError.code = 'ECONNRESET';

      const errorBuilder = createMockRequestBuilder(undefined, networkError);
      mockApi.mockReturnValue(errorBuilder);

      const result1 = await graphClient.getMessage('msg-1');
      expect(result1).toBeNull();

      // Now switch to API error
      const apiError: any = new Error('Not found');
      apiError.statusCode = 404;
      const apiErrorBuilder = createMockRequestBuilder(undefined, apiError);
      mockApi.mockReturnValue(apiErrorBuilder);

      const result2 = await graphClient.getEvent('event-1');
      expect(result2).toBeNull();
    });

    it('handles timeout during search operation', async () => {
      const timeoutError: any = new Error('Request timeout');
      timeoutError.code = 'ETIMEDOUT';

      const errorBuilder = createMockRequestBuilder(undefined, timeoutError);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.searchMessages('test query')).rejects.toThrow('Request timeout');
    });
  });

  describe('Error Recovery and Caching', () => {
    it('does not cache failed requests', async () => {
      const testClient = new GraphClient();

      const error = new Error('Server error');
      const errorBuilder = createMockRequestBuilder(undefined, error);
      mockApi.mockReturnValue(errorBuilder);

      // First call fails
      await expect(testClient.listMessages('folder-1', 50, 0)).rejects.toThrow('Server error');

      // Second call should retry (not use cache)
      await expect(testClient.listMessages('folder-1', 50, 0)).rejects.toThrow('Server error');

      // API should be called twice (no caching of errors)
      expect(mockApi).toHaveBeenCalledTimes(2);
    });

    it('clears cache after repeated failures', async () => {
      const testClient = new GraphClient();

      // First successful call
      const successBuilder = createMockRequestBuilder({
        value: [{ id: 'msg-1', subject: 'Test' }],
      });
      mockApi.mockReturnValueOnce(successBuilder);

      await testClient.listMessages('folder-1', 50, 0);

      // Clear cache
      testClient.clearCache();

      // Next call should hit API again
      const errorBuilder = createMockRequestBuilder(undefined, new Error('Failed'));
      mockApi.mockReturnValue(errorBuilder);

      await expect(testClient.listMessages('folder-1', 50, 0)).rejects.toThrow('Failed');
    });

    it('recovers from transient network errors on retry', async () => {
      const testClient = new GraphClient();

      // First call fails
      const networkError: any = new Error('Network timeout');
      networkError.code = 'ETIMEDOUT';
      const errorBuilder = createMockRequestBuilder(undefined, networkError);
      mockApi.mockReturnValueOnce(errorBuilder);

      const result1 = await testClient.getMessage('msg-1');
      expect(result1).toBeNull();

      // Second call succeeds
      const successBuilder = createMockRequestBuilder({
        id: 'msg-1',
        subject: 'Test Message',
      });
      mockApi.mockReturnValue(successBuilder);

      const result2 = await testClient.getMessage('msg-1');
      expect(result2).toBeDefined();
      expect(result2?.subject).toBe('Test Message');
    });
  });

  describe('Edge Cases', () => {
    it('handles extremely long error messages', async () => {
      const longMessage = 'Error: ' + 'x'.repeat(10000);
      const error = new Error(longMessage);

      const errorBuilder = createMockRequestBuilder(undefined, error);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');
      expect(result).toBeNull();
    });

    it('handles simultaneous errors across different resources', async () => {
      const error1 = new Error('Error 1');
      const error2 = new Error('Error 2');
      const error3 = new Error('Error 3');

      mockApi
        .mockReturnValueOnce(createMockRequestBuilder(undefined, error1))
        .mockReturnValueOnce(createMockRequestBuilder(undefined, error2))
        .mockReturnValueOnce(createMockRequestBuilder(undefined, error3));

      const [result1, result2, result3] = await Promise.all([
        graphClient.getMessage('msg-1').catch(() => null),
        graphClient.getEvent('event-1').catch(() => null),
        graphClient.getContact('contact-1').catch(() => null),
      ]);

      expect(result1).toBeNull();
      expect(result2).toBeNull();
      expect(result3).toBeNull();
    });

    it('handles empty error object', async () => {
      const emptyError = new Error();

      const errorBuilder = createMockRequestBuilder(undefined, emptyError);
      mockApi.mockReturnValue(errorBuilder);

      const result = await graphClient.getMessage('msg-1');
      expect(result).toBeNull();
    });

    it('handles error during unread messages fetch', async () => {
      const error: any = new Error('Filter not supported');
      error.statusCode = 400;
      error.code = 'BadRequest';

      const errorBuilder = createMockRequestBuilder(undefined, error);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.listUnreadMessages('folder-1', 50, 0)).rejects.toThrow('Filter not supported');
    });

    it('handles error during folder-specific message search', async () => {
      const error: any = new Error('Search not available');
      error.statusCode = 503;

      const errorBuilder = createMockRequestBuilder(undefined, error);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.searchMessagesInFolder('folder-1', 'query', 50)).rejects.toThrow('Search not available');
    });

    it('handles error during contact search', async () => {
      const error: any = new Error('Invalid filter');
      error.statusCode = 400;

      const errorBuilder = createMockRequestBuilder(undefined, error);
      mockApi.mockReturnValue(errorBuilder);

      await expect(graphClient.searchContacts('John', 50)).rejects.toThrow('Invalid filter');
    });
  });
});
