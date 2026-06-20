import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { getAuth } from './google.js';
import { searchDrive, trashFile, renameFile, moveFile, copyFile, listFolder, createFolder, getFileMetadata, addComment, listComments, resolveComment, listPermissions, createPermission, updatePermission, deletePermission, updateComment, deleteComment, listReplies, createReply, deleteReply, listRevisions, getRevision, aboutGet } from './tools/drive.js';
import { createDoc, copyDoc, formatDoc, exportDoc, listDocTabs, updateHeaderFooter, insertImage, renameDoc, renameDocTab, readDoc, editDoc, insertTable, modifyTable, updateList, insertPageBreak, createNamedRange, deleteNamedRange, listNamedRanges, updateDocumentStyle, updateNamedStyle, insertSectionBreak, insertDate, insertPerson, insertRichLink, mergeTableCells, unmergeTableCells, formatTable, pinTableHeaderRows, replaceImage, deletePositionedObject, replaceNamedRangeContent, createFootnote, deleteHeader, deleteFooter, addDocumentTab, deleteDocTab, updateSectionStyle } from './tools/docs.js';
import { TOOLS as SHEET_TOOLS } from './tools/sheets.js';
import { TOOLS as FORMAT_TOOLS } from './tools/sheets-format.js';
import { TOOLS as STRUCTURE_TOOLS } from './tools/sheets-structure.js';

const ALL_TOOLS = [...SHEET_TOOLS, ...FORMAT_TOOLS, ...STRUCTURE_TOOLS];

// Build a lookup map: tool name → handler function
const HANDLERS = Object.fromEntries(ALL_TOOLS.map(t => [t.name, t.fn]));

const DRIVE_DOC_TOOLS = [
  {
    name: 'search_drive',
    description: 'Search for files in Google Drive by name or content',
    inputSchema: { type: 'object', properties: { query: { type: 'string' }, mimeType: { type: 'string', description: 'Optional MIME type filter' }, maxResults: { type: 'number', description: 'Max results (default 10, max 50)' } }, required: ['query'] },
  },
  {
    name: 'create_doc',
    description: 'Create a new Google Doc with optional body text. Returns the doc ID and URL.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Document title' },
        folderId: { type: 'string', description: 'Optional Google Drive folder ID to place the doc in' },
        body: { type: 'string', description: 'Optional initial text content' },
      },
      required: ['title'],
    },
  },
  {
    name: 'read_doc',
    description: 'Read a Google Doc as plain text / markdown',
    inputSchema: { type: 'object', properties: { fileId: { type: 'string', description: 'Google Doc file ID or URL' } }, required: ['fileId'] },
  },
  {
    name: 'trash_file',
    description: 'Move a Google Drive file or folder to trash (or permanently delete it). Use search_drive first to find the file ID.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'The ID of the file to trash or delete' },
        permanent: { type: 'boolean', description: 'If true, permanently delete instead of trashing (default: false)' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'edit_doc',
    description: 'Edit a Google Doc in-place. Supports insert, delete, replace (find/replace all), and append.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        operations: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['insert', 'delete', 'replace', 'append'] },
              text: { type: 'string' },
              index: { type: 'number' },
              startIndex: { type: 'number' },
              endIndex: { type: 'number' },
              find: { type: 'string' },
              replaceWith: { type: 'string' },
            },
            required: ['type'],
          },
        },
      },
      required: ['fileId', 'operations'],
    },
  },
  {
    name: 'copy_doc',
    description: 'Copy an existing Google Doc. Optionally rename and/or move to a different folder.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL to copy' },
        title: { type: 'string', description: 'Title for the copy (defaults to "Copy of ...")' },
        folderId: { type: 'string', description: 'Optional folder ID to place the copy in' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'format_doc',
    description: 'Format text or paragraphs in a Google Doc. Supports bold, italic, underline, strikethrough, font size/family, colors, heading styles, alignment, and spacing.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        operations: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['textStyle', 'paragraphStyle'] },
              startIndex: { type: 'number' },
              endIndex: { type: 'number' },
              bold: { type: 'boolean' },
              italic: { type: 'boolean' },
              underline: { type: 'boolean' },
              strikethrough: { type: 'boolean' },
              fontSize: { type: 'number', description: 'Font size in points' },
              fontFamily: { type: 'string', description: 'e.g. "Arial", "Times New Roman"' },
              foregroundColor: { type: 'object', description: '{ red, green, blue } values 0.0-1.0', properties: { red: { type: 'number' }, green: { type: 'number' }, blue: { type: 'number' } } },
              backgroundColor: { type: 'object', description: '{ red, green, blue } values 0.0-1.0', properties: { red: { type: 'number' }, green: { type: 'number' }, blue: { type: 'number' } } },
              link: { type: 'string', description: 'URL to hyperlink the text to' },
              headingLink: { type: 'string', description: 'Heading ID to link to within the document' },
              bookmarkLink: { type: 'string', description: 'Bookmark ID to link to within the document' },
              removeLink: { type: 'boolean', description: 'Set true to remove an existing hyperlink' },
              namedStyleType: { type: 'string', enum: ['NORMAL_TEXT', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6', 'TITLE', 'SUBTITLE'], description: 'Paragraph style (paragraphStyle type only)' },
              alignment: { type: 'string', enum: ['START', 'CENTER', 'END', 'JUSTIFIED'], description: 'Paragraph alignment (paragraphStyle type only)' },
              lineSpacing: { type: 'number', description: 'Line spacing as percentage, e.g. 100 for single, 200 for double (paragraphStyle type only)' },
              spaceAbove: { type: 'number', description: 'Space above paragraph in points (paragraphStyle type only)' },
              spaceBelow: { type: 'number', description: 'Space below paragraph in points (paragraphStyle type only)' },
            },
            required: ['type', 'startIndex', 'endIndex'],
          },
        },
      },
      required: ['fileId', 'operations'],
    },
  },
  {
    name: 'export_doc',
    description: 'Export a Google Doc to pdf, docx, txt, html, rtf, or epub.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        format: { type: 'string', enum: ['pdf', 'docx', 'txt', 'html', 'rtf', 'epub'] },
      },
      required: ['fileId', 'format'],
    },
  },
  {
    name: 'list_doc_tabs',
    description: 'List all tabs in a Google Doc (title, ID, index).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'update_header_footer',
    description: 'Create or replace the default header or footer in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        type: { type: 'string', enum: ['header', 'footer'] },
        text: { type: 'string', description: 'Text content for the header or footer' },
      },
      required: ['fileId', 'type', 'text'],
    },
  },
  {
    name: 'insert_image',
    description: 'Insert an inline image into a Google Doc from a public URL.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        imageUrl: { type: 'string', description: 'Public URL of the image' },
        index: { type: 'number', description: 'Document index to insert at (default: 1, start of doc)' },
        width: { type: 'number', description: 'Image width in points (optional)' },
        height: { type: 'number', description: 'Image height in points (optional)' },
      },
      required: ['fileId', 'imageUrl'],
    },
  },
  {
    name: 'rename_doc',
    description: 'Rename a Google Doc file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        title: { type: 'string', description: 'New title for the document' },
      },
      required: ['fileId', 'title'],
    },
  },
  {
    name: 'rename_doc_tab',
    description: 'Rename a tab in a Google Doc. Use list_doc_tabs to find tab IDs.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        tabId: { type: 'string', description: 'Tab ID to rename (from list_doc_tabs)' },
        title: { type: 'string', description: 'New title for the tab' },
      },
      required: ['fileId', 'tabId', 'title'],
    },
  },
  // --- Drive generic tools ---
  {
    name: 'rename_file',
    description: 'Rename any Google Drive file or folder.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        name: { type: 'string', description: 'New name for the file' },
      },
      required: ['fileId', 'name'],
    },
  },
  {
    name: 'move_file',
    description: 'Move a Google Drive file or folder to a different parent folder.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        folderId: { type: 'string', description: 'Destination folder ID' },
      },
      required: ['fileId', 'folderId'],
    },
  },
  {
    name: 'copy_file',
    description: 'Copy any Google Drive file. Optionally rename and/or place in a specific folder.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL to copy' },
        name: { type: 'string', description: 'Name for the copy (defaults to "Copy of ...")' },
        folderId: { type: 'string', description: 'Optional folder ID to place the copy in' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'list_folder',
    description: 'List contents of a Google Drive folder.',
    inputSchema: {
      type: 'object',
      properties: {
        folderId: { type: 'string', description: 'Folder ID' },
        maxResults: { type: 'number', description: 'Max results (default 20, max 100)' },
        orderBy: { type: 'string', description: 'Sort order: name, modifiedTime, createdTime (default: name)' },
      },
      required: ['folderId'],
    },
  },
  {
    name: 'create_folder',
    description: 'Create a new folder in Google Drive.',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Folder name' },
        parentFolderId: { type: 'string', description: 'Optional parent folder ID' },
      },
      required: ['name'],
    },
  },
  {
    name: 'get_file_metadata',
    description: 'Get metadata for a Google Drive file (name, type, size, owner, dates, parents, etc.).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
      },
      required: ['fileId'],
    },
  },
  // --- Comments (work on any Drive file) ---
  {
    name: 'add_comment',
    description: 'Add a comment to a Google Drive file. Optionally anchor it to quoted text.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        content: { type: 'string', description: 'Comment text' },
        quotedText: { type: 'string', description: 'Optional text to anchor the comment to' },
      },
      required: ['fileId', 'content'],
    },
  },
  {
    name: 'list_comments',
    description: 'List comments on a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        includeResolved: { type: 'boolean', description: 'Include resolved comments (default: false)' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'resolve_comment',
    description: 'Resolve a comment on a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string', description: 'Comment ID to resolve' },
      },
      required: ['fileId', 'commentId'],
    },
  },
  // --- Permissions ---
  {
    name: 'list_permissions',
    description: 'List who has access to a Google Drive file.',
    inputSchema: { type: 'object', properties: { fileId: { type: 'string', description: 'File ID or URL' } }, required: ['fileId'] },
  },
  {
    name: 'create_permission',
    description: 'Share a Google Drive file with a user, group, domain, or anyone.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        role: { type: 'string', enum: ['reader', 'writer', 'commenter', 'owner'], description: 'Permission role' },
        type: { type: 'string', enum: ['user', 'group', 'domain', 'anyone'], description: 'Grantee type' },
        emailAddress: { type: 'string', description: 'Email (for user/group type)' },
        domain: { type: 'string', description: 'Domain (for domain type)' },
        sendNotificationEmail: { type: 'boolean', description: 'Send notification (default: false)' },
      },
      required: ['fileId', 'role', 'type'],
    },
  },
  {
    name: 'update_permission',
    description: 'Update a permission role on a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        permissionId: { type: 'string', description: 'Permission ID' },
        role: { type: 'string', enum: ['reader', 'writer', 'commenter', 'owner'] },
      },
      required: ['fileId', 'permissionId', 'role'],
    },
  },
  {
    name: 'delete_permission',
    description: 'Remove a permission from a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        permissionId: { type: 'string', description: 'Permission ID' },
      },
      required: ['fileId', 'permissionId'],
    },
  },
  // --- Comment/reply management ---
  {
    name: 'update_comment',
    description: 'Edit a comment on a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string' },
        content: { type: 'string', description: 'New comment text' },
      },
      required: ['fileId', 'commentId', 'content'],
    },
  },
  {
    name: 'delete_comment',
    description: 'Delete a comment from a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string' },
      },
      required: ['fileId', 'commentId'],
    },
  },
  {
    name: 'list_replies',
    description: 'List replies on a comment.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string' },
      },
      required: ['fileId', 'commentId'],
    },
  },
  {
    name: 'create_reply',
    description: 'Reply to a comment on a Google Drive file.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string' },
        content: { type: 'string', description: 'Reply text' },
        action: { type: 'string', enum: ['resolve', 'reopen'], description: 'Optional action' },
      },
      required: ['fileId', 'commentId', 'content'],
    },
  },
  {
    name: 'delete_reply',
    description: 'Delete a reply from a comment.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        commentId: { type: 'string' },
        replyId: { type: 'string' },
      },
      required: ['fileId', 'commentId', 'replyId'],
    },
  },
  // --- Revisions ---
  {
    name: 'list_revisions',
    description: 'List revision history of a Google Drive file.',
    inputSchema: { type: 'object', properties: { fileId: { type: 'string', description: 'File ID or URL' } }, required: ['fileId'] },
  },
  {
    name: 'get_revision',
    description: 'Get details of a specific file revision.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'File ID or URL' },
        revisionId: { type: 'string' },
      },
      required: ['fileId', 'revisionId'],
    },
  },
  // --- About ---
  {
    name: 'about',
    description: 'Get Google Drive account info and storage quota.',
    inputSchema: { type: 'object', properties: {} },
  },
  // --- Docs: tables, lists, page breaks, named ranges ---
  {
    name: 'insert_table',
    description: 'Insert a table into a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        rows: { type: 'number', description: 'Number of rows' },
        columns: { type: 'number', description: 'Number of columns' },
        index: { type: 'number', description: 'Document index to insert at (default: 1)' },
      },
      required: ['fileId', 'rows', 'columns'],
    },
  },
  {
    name: 'modify_table',
    description: 'Insert or delete rows/columns in a Google Doc table. Requires the table start index (from read_doc structure).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        operations: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['insertRow', 'insertColumn', 'deleteRow', 'deleteColumn'] },
              tableStartIndex: { type: 'number', description: 'Start index of the table element' },
              rowIndex: { type: 'number', description: '0-based row index' },
              columnIndex: { type: 'number', description: '0-based column index' },
              insertBelow: { type: 'boolean', description: 'For insertRow: insert below (default true)' },
              insertRight: { type: 'boolean', description: 'For insertColumn: insert right (default true)' },
            },
            required: ['type', 'tableStartIndex', 'rowIndex', 'columnIndex'],
          },
        },
      },
      required: ['fileId', 'operations'],
    },
  },
  {
    name: 'update_list',
    description: 'Apply or remove bulleted/numbered list formatting on a range of paragraphs in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        startIndex: { type: 'number', description: 'Start index of the range' },
        endIndex: { type: 'number', description: 'End index of the range' },
        type: { type: 'string', enum: ['bullet', 'numbered', 'remove'], description: 'List type to apply, or "remove" to clear' },
      },
      required: ['fileId', 'startIndex', 'endIndex', 'type'],
    },
  },
  {
    name: 'insert_page_break',
    description: 'Insert a page break at a position in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        index: { type: 'number', description: 'Document index to insert the page break at' },
      },
      required: ['fileId', 'index'],
    },
  },
  {
    name: 'create_named_range',
    description: 'Create a named range (bookmark) in a Google Doc that can be linked to.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        name: { type: 'string', description: 'Name for the range' },
        startIndex: { type: 'number', description: 'Start index' },
        endIndex: { type: 'number', description: 'End index' },
      },
      required: ['fileId', 'name', 'startIndex', 'endIndex'],
    },
  },
  {
    name: 'delete_named_range',
    description: 'Delete a named range from a Google Doc by name or ID.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        name: { type: 'string', description: 'Named range name (deletes all with this name)' },
        namedRangeId: { type: 'string', description: 'Named range ID (deletes specific range)' },
      },
      required: ['fileId'],
    },
  },
  {
    name: 'list_named_ranges',
    description: 'List all named ranges (bookmarks) in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
      },
      required: ['fileId'],
    },
  },
  // --- Docs: document & section styles ---
  {
    name: 'update_document_style',
    description: 'Update page-level style of a Google Doc (margins, page size, background color).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        style: {
          type: 'object',
          description: 'Style properties',
          properties: {
            marginTop: { type: 'number', description: 'Top margin in PT' },
            marginBottom: { type: 'number', description: 'Bottom margin in PT' },
            marginLeft: { type: 'number', description: 'Left margin in PT' },
            marginRight: { type: 'number', description: 'Right margin in PT' },
            pageSize: { type: 'object', properties: { width: { type: 'number' }, height: { type: 'number' } }, description: 'Page dimensions in PT' },
            background: { type: 'object', description: 'Background color as { color: { rgbColor: { red, green, blue } } }' },
          },
        },
      },
      required: ['fileId', 'style'],
    },
  },
  {
    name: 'update_named_style',
    description: 'Update a named style (e.g. HEADING_1, NORMAL_TEXT) in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        namedStyleType: { type: 'string', enum: ['NORMAL_TEXT', 'HEADING_1', 'HEADING_2', 'HEADING_3', 'HEADING_4', 'HEADING_5', 'HEADING_6', 'TITLE', 'SUBTITLE'] },
        textStyle: { type: 'object', description: 'Text style properties (bold, italic, fontSize, fontFamily, etc.)' },
        paragraphStyle: { type: 'object', description: 'Paragraph style properties (alignment, lineSpacing, etc.)' },
      },
      required: ['fileId', 'namedStyleType'],
    },
  },
  {
    name: 'insert_section_break',
    description: 'Insert a section break in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        index: { type: 'number', description: 'Document index' },
        type: { type: 'string', enum: ['NEXT_PAGE', 'CONTINUOUS'], description: 'Section break type (default: NEXT_PAGE)' },
      },
      required: ['fileId', 'index'],
    },
  },
  {
    name: 'update_section_style',
    description: 'Update section-level style in a Google Doc (columns, margins, headers/footers).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        startIndex: { type: 'number' },
        endIndex: { type: 'number' },
        style: { type: 'object', description: 'Section style properties (columnProperties, sectionType, etc.)' },
      },
      required: ['fileId', 'startIndex', 'endIndex', 'style'],
    },
  },
  // --- Docs: insert special elements ---
  {
    name: 'insert_date',
    description: 'Insert a date chip into a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        index: { type: 'number', description: 'Document index' },
      },
      required: ['fileId', 'index'],
    },
  },
  {
    name: 'insert_person',
    description: 'Insert a person mention (@ chip) into a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        email: { type: 'string', description: 'Email address of the person' },
        index: { type: 'number', description: 'Document index' },
      },
      required: ['fileId', 'email', 'index'],
    },
  },
  {
    name: 'insert_rich_link',
    description: 'Insert a rich link (smart chip) to a Google file into a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        url: { type: 'string', description: 'URL of the Google resource to link' },
        index: { type: 'number', description: 'Document index' },
      },
      required: ['fileId', 'url', 'index'],
    },
  },
  {
    name: 'create_footnote',
    description: 'Create a footnote at a position in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        index: { type: 'number', description: 'Document index' },
      },
      required: ['fileId', 'index'],
    },
  },
  // --- Docs: table formatting ---
  {
    name: 'merge_table_cells',
    description: 'Merge cells in a Google Doc table.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        tableStartIndex: { type: 'number', description: 'Table start index' },
        rowStart: { type: 'number' }, rowEnd: { type: 'number' },
        columnStart: { type: 'number' }, columnEnd: { type: 'number' },
      },
      required: ['fileId', 'tableStartIndex', 'rowStart', 'rowEnd', 'columnStart', 'columnEnd'],
    },
  },
  {
    name: 'unmerge_table_cells',
    description: 'Unmerge cells in a Google Doc table.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        tableStartIndex: { type: 'number', description: 'Table start index' },
        rowStart: { type: 'number' }, rowEnd: { type: 'number' },
        columnStart: { type: 'number' }, columnEnd: { type: 'number' },
      },
      required: ['fileId', 'tableStartIndex', 'rowStart', 'rowEnd', 'columnStart', 'columnEnd'],
    },
  },
  {
    name: 'format_table',
    description: 'Format table cells, columns, or rows in a Google Doc (borders, backgrounds, padding, widths, heights).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        operations: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              type: { type: 'string', enum: ['cellStyle', 'columnProperties', 'rowStyle'] },
              tableStartIndex: { type: 'number' },
              rowIndex: { type: 'number' }, columnIndex: { type: 'number' },
              style: { type: 'object', description: 'For cellStyle: backgroundColor, borders, contentAlignment, padding' },
              width: { type: 'number', description: 'For columnProperties: width in PT' },
              widthType: { type: 'string', enum: ['EVENLY_DISTRIBUTED', 'FIXED_WIDTH'] },
              minHeight: { type: 'number', description: 'For rowStyle: min height in PT' },
              minHeightType: { type: 'string', enum: ['AT_LEAST', 'EXACTLY'] },
            },
            required: ['type', 'tableStartIndex'],
          },
        },
      },
      required: ['fileId', 'operations'],
    },
  },
  {
    name: 'pin_table_header_rows',
    description: 'Pin header rows in a Google Doc table.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        tableStartIndex: { type: 'number' },
        pinnedCount: { type: 'number', description: 'Number of header rows to pin' },
      },
      required: ['fileId', 'tableStartIndex', 'pinnedCount'],
    },
  },
  // --- Docs: images & objects ---
  {
    name: 'replace_image',
    description: 'Replace an existing image in a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        imageObjectId: { type: 'string', description: 'Object ID of the image to replace' },
        uri: { type: 'string', description: 'New image URL' },
      },
      required: ['fileId', 'imageObjectId', 'uri'],
    },
  },
  {
    name: 'delete_positioned_object',
    description: 'Delete a positioned object (e.g. floating image) from a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        objectId: { type: 'string', description: 'Object ID to delete' },
      },
      required: ['fileId', 'objectId'],
    },
  },
  {
    name: 'replace_named_range_content',
    description: 'Replace the content of a named range in a Google Doc (useful for templates).',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        namedRangeId: { type: 'string', description: 'Named range ID' },
        name: { type: 'string', description: 'Named range name (alternative to ID)' },
        text: { type: 'string', description: 'Replacement text' },
      },
      required: ['fileId', 'text'],
    },
  },
  // --- Docs: headers, footers, tabs ---
  {
    name: 'delete_header',
    description: 'Delete a header from a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        headerId: { type: 'string', description: 'Header ID to delete' },
      },
      required: ['fileId', 'headerId'],
    },
  },
  {
    name: 'delete_footer',
    description: 'Delete a footer from a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        footerId: { type: 'string', description: 'Footer ID to delete' },
      },
      required: ['fileId', 'footerId'],
    },
  },
  {
    name: 'add_doc_tab',
    description: 'Add a new tab to a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        title: { type: 'string', description: 'Tab title' },
      },
      required: ['fileId', 'title'],
    },
  },
  {
    name: 'delete_doc_tab',
    description: 'Delete a tab from a Google Doc.',
    inputSchema: {
      type: 'object',
      properties: {
        fileId: { type: 'string', description: 'Google Doc file ID or URL' },
        tabId: { type: 'string', description: 'Tab ID to delete' },
      },
      required: ['fileId', 'tabId'],
    },
  },
];

const server = new Server(
  { name: 'google-mcp', version: '2.0.0' },
  { capabilities: { tools: {} } }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    ...DRIVE_DOC_TOOLS,
    ...ALL_TOOLS.map(({ name, description, inputSchema }) => ({ name, description, inputSchema })),
  ],
}));

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;
  try {
    const auth = await getAuth();

    let result;
    if (name === 'search_drive') result = await searchDrive(auth, args);
    else if (name === 'trash_file') result = await trashFile(auth, args);
    else if (name === 'create_doc') result = await createDoc(auth, args);
    else if (name === 'copy_doc') result = await copyDoc(auth, args);
    else if (name === 'format_doc') result = await formatDoc(auth, args);
    else if (name === 'export_doc') result = await exportDoc(auth, args);
    else if (name === 'list_doc_tabs') result = await listDocTabs(auth, args);
    else if (name === 'update_header_footer') result = await updateHeaderFooter(auth, args);
    else if (name === 'insert_image') result = await insertImage(auth, args);
    else if (name === 'rename_doc') result = await renameDoc(auth, args);
    else if (name === 'rename_doc_tab') result = await renameDocTab(auth, args);
    else if (name === 'read_doc') result = await readDoc(auth, args);
    else if (name === 'edit_doc') result = await editDoc(auth, args);
    else if (name === 'rename_file') result = await renameFile(auth, args);
    else if (name === 'move_file') result = await moveFile(auth, args);
    else if (name === 'copy_file') result = await copyFile(auth, args);
    else if (name === 'list_folder') result = await listFolder(auth, args);
    else if (name === 'create_folder') result = await createFolder(auth, args);
    else if (name === 'get_file_metadata') result = await getFileMetadata(auth, args);
    else if (name === 'add_comment') result = await addComment(auth, args);
    else if (name === 'list_comments') result = await listComments(auth, args);
    else if (name === 'resolve_comment') result = await resolveComment(auth, args);
    else if (name === 'insert_table') result = await insertTable(auth, args);
    else if (name === 'modify_table') result = await modifyTable(auth, args);
    else if (name === 'update_list') result = await updateList(auth, args);
    else if (name === 'insert_page_break') result = await insertPageBreak(auth, args);
    else if (name === 'create_named_range') result = await createNamedRange(auth, args);
    else if (name === 'delete_named_range') result = await deleteNamedRange(auth, args);
    else if (name === 'list_named_ranges') result = await listNamedRanges(auth, args);
    else if (name === 'list_permissions') result = await listPermissions(auth, args);
    else if (name === 'create_permission') result = await createPermission(auth, args);
    else if (name === 'update_permission') result = await updatePermission(auth, args);
    else if (name === 'delete_permission') result = await deletePermission(auth, args);
    else if (name === 'update_comment') result = await updateComment(auth, args);
    else if (name === 'delete_comment') result = await deleteComment(auth, args);
    else if (name === 'list_replies') result = await listReplies(auth, args);
    else if (name === 'create_reply') result = await createReply(auth, args);
    else if (name === 'delete_reply') result = await deleteReply(auth, args);
    else if (name === 'list_revisions') result = await listRevisions(auth, args);
    else if (name === 'get_revision') result = await getRevision(auth, args);
    else if (name === 'about') result = await aboutGet(auth);
    else if (name === 'update_document_style') result = await updateDocumentStyle(auth, args);
    else if (name === 'update_named_style') result = await updateNamedStyle(auth, args);
    else if (name === 'insert_section_break') result = await insertSectionBreak(auth, args);
    else if (name === 'update_section_style') result = await updateSectionStyle(auth, args);
    else if (name === 'insert_date') result = await insertDate(auth, args);
    else if (name === 'insert_person') result = await insertPerson(auth, args);
    else if (name === 'insert_rich_link') result = await insertRichLink(auth, args);
    else if (name === 'create_footnote') result = await createFootnote(auth, args);
    else if (name === 'merge_table_cells') result = await mergeTableCells(auth, args);
    else if (name === 'unmerge_table_cells') result = await unmergeTableCells(auth, args);
    else if (name === 'format_table') result = await formatTable(auth, args);
    else if (name === 'pin_table_header_rows') result = await pinTableHeaderRows(auth, args);
    else if (name === 'replace_image') result = await replaceImage(auth, args);
    else if (name === 'delete_positioned_object') result = await deletePositionedObject(auth, args);
    else if (name === 'replace_named_range_content') result = await replaceNamedRangeContent(auth, args);
    else if (name === 'delete_header') result = await deleteHeader(auth, args);
    else if (name === 'delete_footer') result = await deleteFooter(auth, args);
    else if (name === 'add_doc_tab') result = await addDocumentTab(auth, args);
    else if (name === 'delete_doc_tab') result = await deleteDocTab(auth, args);
    else if (HANDLERS[name]) result = await HANDLERS[name](auth, args);
    else throw new Error(`Unknown tool: ${name}`);

    return { content: [{ type: 'text', text: result }] };
  } catch (err) {
    return { content: [{ type: 'text', text: `Error: ${err.message}` }], isError: true };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
