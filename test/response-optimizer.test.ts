import { describe, it, expect } from 'vitest';
import {
  stripHtmlTags,
  removeEmbeddedImages,
  removeInlineAttachments,
  optimizeMessageBody,
  optimizeMailMessage,
  optimizeMailCollection,
  DEFAULT_LLM_OPTIMIZATION,
} from '../src/response-optimizer.js';

describe('Response Optimizer', () => {
  describe('stripHtmlTags', () => {
    it('should remove HTML tags and preserve text', () => {
      const html = '<p>Hello <b>world</b>!</p><br/><div>Test</div>';
      const result = stripHtmlTags(html);
      expect(result).toBe('Hello world!\n\nTest');
    });

    it('should handle empty string', () => {
      expect(stripHtmlTags('')).toBe('');
    });

    it('should decode HTML entities', () => {
      const html = 'Hello &amp; welcome &quot;friend&quot; &lt;test&gt;';
      const result = stripHtmlTags(html);
      expect(result).toBe('Hello & welcome "friend" <test>');
    });

    it('should decode numeric HTML entities and remove invisible characters', () => {
      const html = 'Test &#65279;͏ &#65279;͏ content with invisible chars';
      const result = stripHtmlTags(html);
      expect(result).toBe('Test content with invisible chars');
    });

    it('should handle complex email content', () => {
      const html =
        'New posts from Future Processing S.A. &#65279;͏ &#65279;͏ &#65279;͏<br>Future Processing S.A.';
      const result = stripHtmlTags(html);
      expect(result).toBe('New posts from Future Processing S.A. \nFuture Processing S.A.');
    });
  });

  describe('removeEmbeddedImages', () => {
    it('should remove base64 images', () => {
      const html = '<p>Text</p><img src="data:image/png;base64,abc123"/><p>More text</p>';
      const result = removeEmbeddedImages(html);
      expect(result).toBe('<p>Text</p>[IMAGE_REMOVED]<p>More text</p>');
    });

    it('should handle empty string', () => {
      expect(removeEmbeddedImages('')).toBe('');
    });
  });

  describe('removeInlineAttachments', () => {
    it('should remove cid references', () => {
      const html = '<p>Text</p><img src="cid:image001.jpg@01D26CD8.6C05F070"/><p>More text</p>';
      const result = removeInlineAttachments(html);
      expect(result).toBe('<p>Text</p>[ATTACHMENT_REMOVED]<p>More text</p>');
    });
  });

  describe('optimizeMessageBody', () => {
    it('should optimize HTML body content', () => {
      const body = {
        contentType: 'html',
        content: '<p>Hello <b>world</b>!</p><img src="data:image/png;base64,abc123"/>',
      };

      const result = optimizeMessageBody(body, DEFAULT_LLM_OPTIMIZATION);

      expect(result.contentType).toBe('text');
      expect(result.content).toBe('Hello world!\n[IMAGE_REMOVED]');
    });

    it('should truncate long content', () => {
      const longContent = 'a'.repeat(3000);
      const body = {
        contentType: 'html',
        content: `<p>${longContent}</p>`,
      };

      const result = optimizeMessageBody(body, DEFAULT_LLM_OPTIMIZATION);

      expect(result.content.length).toBeLessThanOrEqual(
        DEFAULT_LLM_OPTIMIZATION.maxHtmlContentSize + 20
      ); // +20 for truncation marker
      expect(result.content).toContain('...[TRUNCATED]');
    });
  });

  describe('optimizeMailMessage', () => {
    it('should optimize a complete mail message', () => {
      const message = {
        id: 'test-id',
        subject: 'Test Subject',
        body: {
          contentType: 'html',
          content: '<p>Hello <b>world</b>!</p><img src="data:image/png;base64,abc123"/>',
        },
        toRecipients: [
          {
            emailAddress: {
              name: 'John Doe',
              address: 'john@example.com',
            },
          },
        ],
        conversationId: 'conv-123',
        webLink: 'https://outlook.com/...',
        changeKey: 'change-123',
      };

      const result = optimizeMailMessage(message, DEFAULT_LLM_OPTIMIZATION);

      expect(result.id).toBe('test-id');
      expect(result.subject).toBe('Test Subject');
      expect(result.body.contentType).toBe('text');
      expect(result.body.content).toBe('Hello world!\n[IMAGE_REMOVED]');
      expect(result.toRecipients[0]).toEqual({
        name: 'John Doe',
        address: 'john@example.com',
      });

      // Check that unnecessary fields are removed
      expect(result.conversationId).toBeUndefined();
      expect(result.webLink).toBeUndefined();
      expect(result.changeKey).toBeUndefined();
    });
  });

  describe('optimizeMailCollection', () => {
    it('should optimize a collection of mail messages', () => {
      const collection = {
        value: [
          {
            id: 'msg1',
            subject: 'Message 1',
            body: {
              contentType: 'html',
              content: '<p>Content 1</p>',
            },
          },
          {
            id: 'msg2',
            subject: 'Message 2',
            body: {
              contentType: 'html',
              content: '<p>Content 2</p>',
            },
          },
        ],
        '@odata.count': 2,
      };

      const result = optimizeMailCollection(collection, DEFAULT_LLM_OPTIMIZATION);

      expect(result.value).toHaveLength(2);
      expect(result.value[0].body.contentType).toBe('text');
      expect(result.value[0].body.content).toBe('Content 1');
      expect(result['@odata.count']).toBe(2);
    });

    it('should limit collection size', () => {
      const largeCollection = {
        value: Array.from({ length: 100 }, (_, i) => ({
          id: `msg${i}`,
          subject: `Message ${i}`,
        })),
      };

      const result = optimizeMailCollection(largeCollection, DEFAULT_LLM_OPTIMIZATION);

      expect(result.value).toHaveLength(DEFAULT_LLM_OPTIMIZATION.maxItemsInCollection);
    });
  });
});
