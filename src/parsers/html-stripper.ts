/**
 * Copyright (c) 2026 JBC Tech Solutions, LLC
 * Licensed under the MIT License. See LICENSE file in the project root.
 */

/**
 * Simple HTML to plain text converter.
 *
 * Uses regex-based approach to avoid heavy dependencies.
 * Handles common HTML patterns found in emails.
 */

/**
 * Named HTML entities and their text equivalents.
 */
const HTML_ENTITIES: Record<string, string> = {
  '&nbsp;': ' ',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&apos;': "'",
  '&#39;': "'",
  '&copy;': '\u00A9',
  '&reg;': '\u00AE',
  '&trade;': '\u2122',
  '&mdash;': '\u2014',
  '&ndash;': '\u2013',
  '&hellip;': '\u2026',
  '&lsquo;': '\u2018',
  '&rsquo;': '\u2019',
  '&ldquo;': '\u201C',
  '&rdquo;': '\u201D',
  '&bull;': '\u2022',
  '&middot;': '\u00B7',
  '&euro;': '\u20AC',
  '&pound;': '\u00A3',
  '&yen;': '\u00A5',
  '&cent;': '\u00A2',
};

/**
 * Tags that should add a newline when stripped.
 */
const BLOCK_TAGS = new Set([
  'p',
  'div',
  'br',
  'hr',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'li',
  'tr',
  'blockquote',
  'pre',
  'article',
  'section',
  'header',
  'footer',
  'nav',
  'aside',
  'table',
  'thead',
  'tbody',
  'tfoot',
]);

/**
 * Tags whose content should be completely removed.
 */
const INVISIBLE_TAGS = new Set(['script', 'style', 'head', 'meta', 'link', 'noscript']);

/**
 * Interface for HTML stripper options.
 */
export interface StripHtmlOptions {
  /**
   * Whether to preserve whitespace (default: false).
   */
  readonly preserveWhitespace?: boolean;

  /**
   * Maximum length of output (0 = no limit, default: 0).
   */
  readonly maxLength?: number;
}

/**
 * Strips HTML tags and converts to plain text.
 *
 * @param html - The HTML string to convert
 * @param options - Optional configuration
 * @returns Plain text string
 */
export function stripHtml(html: string | null | undefined, options: StripHtmlOptions = {}): string {
  if (html == null || html === '') {
    return '';
  }

  const { preserveWhitespace = false, maxLength = 0 } = options;

  let text = html;

  // Remove invisible tag content (script, style, etc.)
  for (const tag of INVISIBLE_TAGS) {
    const regex = new RegExp(`<${tag}[^>]*>[\\s\\S]*?</${tag}>`, 'gi');
    text = text.replace(regex, '');
  }

  // Remove HTML comments
  text = text.replace(/<!--[\s\S]*?-->/g, '');

  // Remove CDATA sections
  text = text.replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, '');

  // Handle list items with bullets (before block element processing)
  text = text.replace(/<li[^>]*>/gi, '\n\u2022 ');
  text = text.replace(/<\/li>/gi, '');

  // Add newlines for block elements (excluding li which is handled above)
  for (const tag of BLOCK_TAGS) {
    if (tag === 'li') continue;
    // Opening tags
    text = text.replace(new RegExp(`<${tag}[^>]*>`, 'gi'), '\n');
    // Self-closing tags (like <br />)
    text = text.replace(new RegExp(`<${tag}[^>]*/>`, 'gi'), '\n');
    // Closing tags
    text = text.replace(new RegExp(`</${tag}>`, 'gi'), '\n');
  }

  // Remove all remaining HTML tags
  text = text.replace(/<[^>]+>/g, '');

  // Decode numeric HTML entities
  text = text.replace(/&#(\d+);/g, (_, code: string) => {
    const num = parseInt(code, 10);
    return String.fromCharCode(num);
  });

  // Decode hex HTML entities
  text = text.replace(/&#x([a-fA-F0-9]+);/g, (_, code: string) => {
    const num = parseInt(code, 16);
    return String.fromCharCode(num);
  });

  // Decode named HTML entities
  for (const [entity, replacement] of Object.entries(HTML_ENTITIES)) {
    text = text.split(entity).join(replacement);
  }

  if (!preserveWhitespace) {
    // Normalize whitespace
    text = text
      // Replace tabs with spaces
      .replace(/\t/g, ' ')
      // Replace multiple spaces with single space
      .replace(/ +/g, ' ')
      // Replace multiple newlines with double newline (preserve paragraphs)
      .replace(/\n\s*\n/g, '\n\n')
      // Remove leading/trailing whitespace from lines
      .replace(/^[ \t]+/gm, '')
      .replace(/[ \t]+$/gm, '')
      // Trim overall
      .trim();
  }

  // Apply max length if specified
  if (maxLength > 0 && text.length > maxLength) {
    text = text.substring(0, maxLength - 3) + '...';
  }

  return text;
}

/**
 * Checks if a string appears to contain HTML.
 *
 * @param text - The string to check
 * @returns True if the string appears to contain HTML tags
 */
export function containsHtml(text: string | null | undefined): boolean {
  if (text == null || text === '') {
    return false;
  }
  // Check for common HTML patterns
  return /<[a-zA-Z][^>]*>/.test(text);
}

/**
 * Extracts plain text from an email body, detecting HTML automatically.
 *
 * @param body - The email body (HTML or plain text)
 * @param options - Optional configuration
 * @returns Plain text string
 */
export function extractPlainText(
  body: string | null | undefined,
  options: StripHtmlOptions = {}
): string {
  if (body == null || body === '') {
    return '';
  }

  if (containsHtml(body)) {
    return stripHtml(body, options);
  }

  // Already plain text, just apply max length if needed
  const { maxLength = 0 } = options;
  if (maxLength > 0 && body.length > maxLength) {
    return body.substring(0, maxLength - 3) + '...';
  }

  return body;
}
