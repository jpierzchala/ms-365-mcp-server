# Response Optimization for LLM Agents

This document describes the new response optimization features added to ms-365-mcp-server to reduce token consumption when used with LLM agents.

## Problem Description

When fetching email messages from Microsoft Graph API, the server returns large amounts of data that consumes excessive tokens:

1. **Full HTML content** - Emails contain complete HTML with formatting, CSS styles, and embedded content
2. **Base64 encoded images** - Inline images are embedded as base64 strings
3. **Unnecessary metadata** - Many fields like conversationId, webLink, changeKey are not needed for LLM analysis
4. **All fields returned** - Without $select filtering, all available fields are returned

Example of problematic response:
```json
{
  "body": {
    "contentType": "html",
    "content": "<html><head>...</head><body>...very long HTML with embedded images...</body></html>"
  },
  "conversationId": "...",
  "webLink": "...",
  "changeKey": "...",
  // many other fields
}
```

## Solution

The optimization system includes several strategies:

### 1. Automatic Field Selection for Mail Endpoints

For mail-related endpoints (`/messages`), the system automatically applies optimized `$select` parameters:
- `id`, `subject`, `sender`, `from`, `toRecipients`, `ccRecipients`
- `receivedDateTime`, `sentDateTime`, `hasAttachments`, `importance`, `isRead`
- `bodyPreview` (short text preview) and `body` (which gets further optimized)

### 2. HTML Content Optimization

**HTML to Text Conversion:**
- Strips HTML tags while preserving structure (line breaks)
- Converts HTML entities (`&amp;` → `&`, `&quot;` → `"`)
- Preserves readability while dramatically reducing size

**Embedded Content Removal:**
- Removes base64 embedded images (`data:image/*`)
- Removes inline attachments (`cid:*`)
- Replaces with `[IMAGE_REMOVED]` and `[ATTACHMENT_REMOVED]` markers

**Content Size Limiting:**
- Truncates content exceeding maximum size (default: 2000 characters)
- Adds `...[TRUNCATED]` marker when content is cut

### 3. Object Structure Optimization

**Unnecessary Field Removal:**
- Removes metadata fields: `conversationId`, `conversationIndex`, `internetMessageId`, `webLink`, `changeKey`, `parentFolderId`, `inferenceClassification`
- Simplifies recipient objects to just `name` and `address`

**Collection Limiting:**
- Limits number of items in collections (default: 50 items)
- Prevents overwhelming responses for large mailboxes

## Configuration Options

### CLI Options

```bash
# Disable optimization completely
--no-llm-optimization

# Keep HTML formatting (don't convert to text)
--keep-html

# Adjust content size limit (default: 2000)
--max-content-size 1000

# Adjust collection size limit (default: 50)
--max-items 25
```

### Environment Variables

You can also use environment variables:
```bash
MS365_MCP_LLM_OPTIMIZATION=false
MS365_MCP_MAX_CONTENT_SIZE=1000
MS365_MCP_MAX_ITEMS=25
MS365_MCP_KEEP_HTML=true
```

### Programmatic Configuration

```typescript
import { OptimizationConfig } from './response-optimizer.js';

const customConfig: Partial<OptimizationConfig> = {
  stripHtmlToText: false,      // Keep HTML
  maxHtmlContentSize: 1000,    // Smaller limit
  maxItemsInCollection: 25,    // Fewer items
  removeEmbeddedImages: true,  // Still remove images
  removeInlineAttachments: true
};
```

## Example Results

### Before Optimization
```json
{
  "value": [
    {
      "id": "AAMkAGViZTkxNmM3...",
      "subject": "Meeting Notes",
      "body": {
        "contentType": "html", 
        "content": "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><style><!--@font-face{font-family:Helvetica}--></style></head><body>Hello,<br><br>Please find attached...<img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAA...very long base64 string...></body></html>"
      },
      "toRecipients": [
        {
          "emailAddress": {
            "name": "John Doe",
            "address": "john@example.com"
          }
        }
      ],
      "conversationId": "AAQkAGViZTkxNmM3...",
      "webLink": "https://outlook.office365.com/owa/...",
      "changeKey": "CQAAABYAAAADIdxtY...",
      // ... many more fields
    }
  ]
}
```

### After Optimization
```json
{
  "value": [
    {
      "id": "AAMkAGViZTkxNmM3...",
      "subject": "Meeting Notes",
      "body": {
        "contentType": "text",
        "content": "Hello,\n\nPlease find attached...[IMAGE_REMOVED]"
      },
      "toRecipients": [
        {
          "name": "John Doe", 
          "address": "john@example.com"
        }
      ],
      "receivedDateTime": "2025-08-04T08:41:57Z",
      "hasAttachments": false,
      "isRead": false
    }
  ]
}
```

## Token Savings

Typical token savings observed:
- **60-90% reduction** for HTML-heavy emails
- **70-95% reduction** for emails with embedded images
- **40-60% reduction** for plain text emails (due to field filtering)

Example: An email response that was 15,000 characters became 2,800 characters (81% reduction).

## Monitoring

The system logs optimization results:
```
info: Auto-applied optimized $select for mail endpoint: id,subject,sender,from...
info: Response optimized: 15234 → 2847 chars (81% reduction)
```

## Compatibility

- All existing functionality is preserved
- Optimization is applied transparently 
- Can be disabled if needed for specific use cases
- Works with all MCP clients (Claude, LibreChat, etc.)

## Future Enhancements

Potential additional optimizations:
- Smart content summarization for very long emails
- Attachment metadata optimization
- Calendar event content optimization
- SharePoint document content filtering
