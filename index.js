import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { getAuth } from './google.js';
import { searchDrive } from './tools/drive.js';
import { readDoc, editDoc } from './tools/docs.js';
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
    name: 'read_doc',
    description: 'Read a Google Doc as plain text / markdown',
    inputSchema: { type: 'object', properties: { fileId: { type: 'string', description: 'Google Doc file ID or URL' } }, required: ['fileId'] },
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
    else if (name === 'read_doc') result = await readDoc(auth, args);
    else if (name === 'edit_doc') result = await editDoc(auth, args);
    else if (HANDLERS[name]) result = await HANDLERS[name](auth, args);
    else throw new Error(`Unknown tool: ${name}`);

    return { content: [{ type: 'text', text: result }] };
  } catch (err) {
    return { content: [{ type: 'text', text: `Error: ${err.message}` }], isError: true };
  }
});

const transport = new StdioServerTransport();
await server.connect(transport);
