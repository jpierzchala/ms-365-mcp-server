import logger from './logger.js';

/**
 * Configuration for optimizing responses for LLM agents
 */
export interface OptimizationConfig {
  // Maximum size of HTML content to return (characters)
  maxHtmlContentSize: number;
  // Whether to strip HTML tags and return only text
  stripHtmlToText: boolean;
  // Whether to remove base64 embedded images
  removeEmbeddedImages: boolean;
  // Whether to remove inline attachments
  removeInlineAttachments: boolean;
  // Fields to select for mail messages to minimize response size
  mailSelectFields: string[];
  // Maximum number of items to return in collections
  maxItemsInCollection: number;
}

/**
 * Default optimization configuration for LLM agents
 */
export const DEFAULT_LLM_OPTIMIZATION: OptimizationConfig = {
  maxHtmlContentSize: 2000, // Limit HTML content to 2000 characters
  stripHtmlToText: true, // Convert HTML to plain text
  removeEmbeddedImages: true, // Remove base64 images
  removeInlineAttachments: true, // Remove inline attachments
  mailSelectFields: [
    'id',
    'subject', 
    'sender',
    'from',
    'toRecipients',
    'ccRecipients',
    'receivedDateTime',
    'sentDateTime',
    'hasAttachments',
    'importance',
    'isRead',
    'bodyPreview', // Short preview instead of full body
    'body' // We'll optimize this separately
  ],
  maxItemsInCollection: 50 // Limit collections to 50 items
};

/**
 * Strips HTML tags and returns plain text
 */
export function stripHtmlTags(html: string): string {
  if (!html) return '';
  
  // Remove HTML tags but preserve line breaks
  let text = html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]*>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  // Clean up extra whitespace and line breaks
  text = text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n\s*\n/g, '\n\n')
    .trim();

  return text;
}

/**
 * Removes base64 embedded images from HTML content
 */
export function removeEmbeddedImages(html: string): string {
  if (!html) return html;
  
  // Remove img tags with data: URLs (base64 images)
  return html.replace(/<img[^>]*src\s*=\s*["']data:[^"']*["'][^>]*>/gi, '[IMAGE_REMOVED]');
}

/**
 * Removes inline attachments references
 */
export function removeInlineAttachments(html: string): string {
  if (!html) return html;
  
  // Remove img tags with cid: URLs (inline attachments)
  return html.replace(/<img[^>]*src\s*=\s*["']cid:[^"']*["'][^>]*>/gi, '[ATTACHMENT_REMOVED]');
}

/**
 * Optimizes message body content for LLM consumption
 */
export function optimizeMessageBody(body: any, config: OptimizationConfig): any {
  if (!body || typeof body !== 'object') return body;
  
  let optimizedBody = { ...body };
  
  if (body.content && typeof body.content === 'string') {
    let content = body.content;
    
    // Remove embedded images if configured
    if (config.removeEmbeddedImages) {
      content = removeEmbeddedImages(content);
    }
    
    // Remove inline attachments if configured
    if (config.removeInlineAttachments) {
      content = removeInlineAttachments(content);
    }
    
    // Strip HTML if configured or if content is too long
    if (config.stripHtmlToText || content.length > config.maxHtmlContentSize) {
      if (body.contentType === 'html' || content.includes('<')) {
        content = stripHtmlTags(content);
        optimizedBody.contentType = 'text';
      }
    }
    
    // Truncate if still too long
    if (content.length > config.maxHtmlContentSize) {
      content = content.substring(0, config.maxHtmlContentSize) + '...[TRUNCATED]';
    }
    
    optimizedBody.content = content;
  }
  
  return optimizedBody;
}

/**
 * Optimizes a single mail message for LLM consumption
 */
export function optimizeMailMessage(message: any, config: OptimizationConfig): any {
  if (!message || typeof message !== 'object') return message;
  
  const optimized = { ...message };
  
  // Optimize body content
  if (message.body) {
    optimized.body = optimizeMessageBody(message.body, config);
  }
  
  // Remove unnecessary fields that take up space
  delete optimized.conversationId;
  delete optimized.conversationIndex;
  delete optimized.internetMessageId;
  delete optimized.webLink;
  delete optimized.changeKey;
  delete optimized.parentFolderId;
  delete optimized.inferenceClassification;
  delete optimized.flag;
  delete optimized.categories;
  
  // Simplify recipient objects to just email and name
  ['toRecipients', 'ccRecipients', 'bccRecipients'].forEach(field => {
    if (optimized[field] && Array.isArray(optimized[field])) {
      optimized[field] = optimized[field].map((recipient: any) => ({
        name: recipient.emailAddress?.name,
        address: recipient.emailAddress?.address
      }));
    }
  });
  
  // Simplify sender and from objects
  ['sender', 'from'].forEach(field => {
    if (optimized[field]?.emailAddress) {
      optimized[field] = {
        name: optimized[field].emailAddress.name,
        address: optimized[field].emailAddress.address
      };
    }
  });
  
  return optimized;
}

/**
 * Optimizes a collection of mail messages
 */
export function optimizeMailCollection(response: any, config: OptimizationConfig): any {
  if (!response || typeof response !== 'object') return response;
  
  const optimized = { ...response };
  
  if (response.value && Array.isArray(response.value)) {
    // Limit number of items
    let items = response.value;
    if (items.length > config.maxItemsInCollection) {
      items = items.slice(0, config.maxItemsInCollection);
      logger.info(`Collection truncated from ${response.value.length} to ${config.maxItemsInCollection} items`);
    }
    
    // Optimize each message
    optimized.value = items.map((message: any) => optimizeMailMessage(message, config));
    
    // Update count if present
    if (optimized['@odata.count']) {
      optimized['@odata.count'] = optimized.value.length;
    }
  }
  
  return optimized;
}

/**
 * Generates optimized $select parameter for mail messages
 */
export function getOptimizedMailSelect(config: OptimizationConfig): string {
  return config.mailSelectFields.join(',');
}

/**
 * Applies response optimization based on endpoint and configuration
 */
export function optimizeResponse(
  response: any, 
  endpoint: string, 
  config: OptimizationConfig = DEFAULT_LLM_OPTIMIZATION
): any {
  if (!response || typeof response !== 'object') return response;
  
  // Check if this is a mail-related endpoint
  if (endpoint.includes('/messages') || endpoint.includes('mail')) {
    if (response.value && Array.isArray(response.value)) {
      // Collection of messages
      return optimizeMailCollection(response, config);
    } else if (response.subject || response.body) {
      // Single message
      return optimizeMailMessage(response, config);
    }
  }
  
  return response;
}
