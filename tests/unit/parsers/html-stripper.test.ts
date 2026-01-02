import { describe, it, expect } from 'vitest';
import {
  stripHtml,
  containsHtml,
  extractPlainText,
} from '../../../src/parsers/html-stripper.js';

describe('stripHtml', () => {
  describe('basic tag removal', () => {
    it('removes simple HTML tags', () => {
      expect(stripHtml('<p>Hello World</p>')).toBe('Hello World');
    });

    it('removes nested tags', () => {
      expect(stripHtml('<div><p><strong>Hello</strong></p></div>')).toBe('Hello');
    });

    it('removes self-closing tags', () => {
      expect(stripHtml('Hello<br/>World')).toBe('Hello\nWorld');
    });

    it('removes tags with attributes', () => {
      expect(stripHtml('<a href="http://example.com" target="_blank">Link</a>')).toBe('Link');
    });
  });

  describe('block elements', () => {
    it('adds newlines for paragraphs', () => {
      expect(stripHtml('<p>First</p><p>Second</p>')).toBe('First\n\nSecond');
    });

    it('adds newlines for divs', () => {
      expect(stripHtml('<div>First</div><div>Second</div>')).toBe('First\n\nSecond');
    });

    it('handles br tags', () => {
      expect(stripHtml('Line 1<br>Line 2<br/>Line 3')).toBe('Line 1\nLine 2\nLine 3');
    });

    it('handles headings', () => {
      expect(stripHtml('<h1>Title</h1><p>Content</p>')).toBe('Title\n\nContent');
    });

    it('handles lists with bullets', () => {
      const html = '<ul><li>Item 1</li><li>Item 2</li></ul>';
      const result = stripHtml(html);
      expect(result).toContain('\u2022 Item 1');
      expect(result).toContain('\u2022 Item 2');
    });
  });

  describe('invisible content removal', () => {
    it('removes script tags and content', () => {
      expect(stripHtml('<p>Hello</p><script>alert("hi")</script><p>World</p>')).toBe(
        'Hello\n\nWorld'
      );
    });

    it('removes style tags and content', () => {
      expect(stripHtml('<style>body { color: red; }</style><p>Text</p>')).toBe('Text');
    });

    it('removes head content', () => {
      const html = '<html><head><title>Title</title></head><body>Content</body></html>';
      expect(stripHtml(html)).toBe('Content');
    });
  });

  describe('HTML entity decoding', () => {
    it('decodes named entities', () => {
      expect(stripHtml('&amp; &lt; &gt; &quot;')).toBe('& < > "');
    });

    it('decodes nbsp', () => {
      expect(stripHtml('Hello&nbsp;World')).toBe('Hello World');
    });

    it('decodes numeric entities', () => {
      expect(stripHtml('&#65;&#66;&#67;')).toBe('ABC');
    });

    it('decodes hex entities', () => {
      expect(stripHtml('&#x41;&#x42;&#x43;')).toBe('ABC');
    });

    it('decodes typographic entities', () => {
      expect(stripHtml('&mdash; &ndash; &hellip;')).toBe('\u2014 \u2013 \u2026');
    });

    it('decodes quote entities', () => {
      expect(stripHtml('&lsquo;Hello&rsquo; &ldquo;World&rdquo;')).toBe('\u2018Hello\u2019 \u201CWorld\u201D');
    });
  });

  describe('whitespace handling', () => {
    it('normalizes multiple spaces', () => {
      expect(stripHtml('Hello    World')).toBe('Hello World');
    });

    it('normalizes multiple newlines', () => {
      expect(stripHtml('<p>First</p>\n\n\n<p>Second</p>')).toBe('First\n\nSecond');
    });

    it('trims leading and trailing whitespace', () => {
      expect(stripHtml('  <p>  Hello  </p>  ')).toBe('Hello');
    });

    it('preserves whitespace when option is set', () => {
      const result = stripHtml('Hello    World', { preserveWhitespace: true });
      expect(result).toBe('Hello    World');
    });
  });

  describe('maxLength option', () => {
    it('truncates long text', () => {
      const result = stripHtml('<p>This is a very long text that should be truncated</p>', {
        maxLength: 20,
      });
      expect(result.length).toBe(20);
      expect(result.endsWith('...')).toBe(true);
    });

    it('does not truncate short text', () => {
      const result = stripHtml('<p>Short</p>', { maxLength: 20 });
      expect(result).toBe('Short');
    });

    it('ignores zero maxLength', () => {
      const result = stripHtml('<p>This is a long text</p>', { maxLength: 0 });
      expect(result).toBe('This is a long text');
    });
  });

  describe('edge cases', () => {
    it('returns empty string for null', () => {
      expect(stripHtml(null)).toBe('');
    });

    it('returns empty string for undefined', () => {
      expect(stripHtml(undefined)).toBe('');
    });

    it('returns empty string for empty string', () => {
      expect(stripHtml('')).toBe('');
    });

    it('handles plain text without tags', () => {
      expect(stripHtml('Hello World')).toBe('Hello World');
    });

    it('removes HTML comments', () => {
      expect(stripHtml('Hello <!-- comment --> World')).toBe('Hello World');
    });

    it('handles CDATA sections', () => {
      expect(stripHtml('Hello <![CDATA[some data]]> World')).toBe('Hello World');
    });

    it('handles malformed HTML gracefully', () => {
      expect(stripHtml('<p>Unclosed paragraph')).toBe('Unclosed paragraph');
      // Incomplete tags without closing > are preserved (not valid HTML)
      expect(stripHtml('No tags <missing')).toBe('No tags <missing');
      // But complete tags are removed even if unclosed
      expect(stripHtml('<div>Missing close div')).toBe('Missing close div');
    });
  });

  describe('real-world email HTML', () => {
    it('handles typical email HTML', () => {
      const emailHtml = `
        <html>
          <head><title>Email</title></head>
          <body>
            <div style="font-family: Arial;">
              <p>Dear User,</p>
              <p>Thank you for your <strong>purchase</strong>!</p>
              <p>Best regards,<br>The Team</p>
            </div>
          </body>
        </html>
      `;
      const result = stripHtml(emailHtml);
      expect(result).toContain('Dear User,');
      expect(result).toContain('Thank you for your purchase!');
      expect(result).toContain('Best regards,');
      expect(result).toContain('The Team');
    });
  });
});

describe('containsHtml', () => {
  it('returns true for HTML content', () => {
    expect(containsHtml('<p>Hello</p>')).toBe(true);
    expect(containsHtml('<div class="test">Content</div>')).toBe(true);
    expect(containsHtml('Text <br> more text')).toBe(true);
  });

  it('returns false for plain text', () => {
    expect(containsHtml('Hello World')).toBe(false);
    expect(containsHtml('Email: test@example.com')).toBe(false);
    expect(containsHtml('Price: $100 < $200')).toBe(false);
  });

  it('returns false for null/undefined/empty', () => {
    expect(containsHtml(null)).toBe(false);
    expect(containsHtml(undefined)).toBe(false);
    expect(containsHtml('')).toBe(false);
  });

  it('returns false for angle brackets without valid tags', () => {
    expect(containsHtml('1 < 2 > 0')).toBe(false);
    expect(containsHtml('<<<>>>')).toBe(false);
  });
});

describe('extractPlainText', () => {
  it('strips HTML when HTML is detected', () => {
    expect(extractPlainText('<p>Hello</p>')).toBe('Hello');
  });

  it('returns plain text unchanged', () => {
    expect(extractPlainText('Hello World')).toBe('Hello World');
  });

  it('handles null/undefined', () => {
    expect(extractPlainText(null)).toBe('');
    expect(extractPlainText(undefined)).toBe('');
  });

  it('applies maxLength to plain text', () => {
    const result = extractPlainText('This is a long plain text', { maxLength: 15 });
    expect(result.length).toBe(15);
    expect(result.endsWith('...')).toBe(true);
  });

  it('applies maxLength to HTML', () => {
    const result = extractPlainText('<p>This is a long HTML text</p>', { maxLength: 15 });
    expect(result.length).toBe(15);
    expect(result.endsWith('...')).toBe(true);
  });
});
