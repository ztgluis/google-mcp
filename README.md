# google-mcp

MCP server for Google Drive, Docs, and Sheets — built for Claude Code.

Gives Claude Code direct read/write access to Google Drive (file management, sharing, comments), Google Docs (full editing, formatting, tables, lists, hyperlinks), and Google Sheets (cell-level edits, formatting, structure).

## Setup

### Prerequisites

- Node.js 18+ (`brew install node` on macOS)
- A Google Cloud project with the Drive, Docs, and Sheets APIs enabled
- An OAuth 2.0 Client ID (Desktop app type) — download as `credentials.json`

### Install

```bash
git clone git@github.com:ztgluis/google-mcp.git ~/dev/google-mcp
cd ~/dev/google-mcp
npm install
```

### Authenticate

Place your `credentials.json` in the project root, then:

```bash
npm run auth
```

This opens your browser for Google sign-in and saves `token.json` locally (never committed).

### Register with Claude Code

```bash
claude mcp add google -s user -- node "$HOME/dev/google-mcp/index.js"
```

Use `-s user` (not `-s project`) so the registration is stored locally in `~/.claude.json`. If your project directory is on a synced drive (e.g., Google Drive), `-s project` would create a `.mcp.json` with a machine-specific absolute path that breaks on other machines.

Restart Claude Code to load the server.

## Tools (52)

### Drive
- `search_drive` — search files by name or content
- `get_file_metadata` — file info (name, type, size, owner, dates, parents)
- `rename_file` — rename any file or folder
- `move_file` — move a file to a different folder
- `copy_file` — copy any file, optionally rename and relocate
- `create_folder` — create a new folder
- `trash_file` — move to trash or permanently delete
- `list_folder` — list contents of a folder

### Comments
- `add_comment` — add a comment to any Drive file, optionally anchored to quoted text
- `list_comments` — list comments on a file
- `resolve_comment` — resolve a comment

### Docs
- `create_doc` — create a new Google Doc
- `read_doc` — read a Google Doc as plain text / markdown
- `edit_doc` — insert, delete, replace, or append text
- `format_doc` — bold, italic, underline, font, colors, headings, alignment, hyperlinks
- `copy_doc` — copy a Google Doc
- `rename_doc` — rename a Google Doc
- `export_doc` — export to PDF, DOCX, TXT, HTML, RTF, or EPUB
- `insert_image` — insert an inline image from a URL
- `insert_table` — insert a table
- `modify_table` — insert/delete rows and columns in a table
- `update_list` — apply or remove bulleted/numbered list formatting
- `insert_page_break` — insert a page break
- `create_named_range` — create a named range (bookmark)
- `delete_named_range` — delete a named range
- `list_named_ranges` — list all named ranges
- `list_doc_tabs` — list tabs in a doc
- `rename_doc_tab` — rename a doc tab
- `update_header_footer` — create or replace headers/footers

### Sheets — Data
- `read_sheet` — read a range
- `batch_read_sheet` — read multiple ranges in one call
- `append_sheet_row` — append a row after the last data row (supports `dataRange` for complex layouts)
- `update_sheet_cell` — update a single cell or range
- `batch_update_sheet` — write to multiple ranges in one call
- `clear_sheet` — clear values in a range
- `insert_rows` — insert rows at a position, optionally with data
- `delete_rows` — delete rows
- `delete_columns` — delete columns
- `find_replace` — find and replace text
- `sort_range` — sort a range by one or more columns
- `export_sheet` — export to CSV, TSV, XLSX, PDF, or HTML

### Sheets — Formatting
- `format_cells` — colors, fonts, alignment, number formats, wrap strategy
- `merge_cells` — merge a range
- `unmerge_cells` — unmerge a range
- `add_conditional_formatting` — boolean or gradient rules
- `freeze` — freeze rows/columns (with merged-cell conflict detection)
- `auto_resize` — auto-fit columns or rows to content

### Sheets — Structure
- `get_sheet_metadata` — title, URL, sheet names, IDs, dimensions, frozen rows/cols
- `create_spreadsheet` — create a new spreadsheet
- `insert_sheet` — add a tab
- `delete_sheet` — remove a tab
- `rename_sheet` — rename a tab
- `duplicate_sheet` — duplicate a tab within the same spreadsheet
- `copy_sheet_to` — copy a tab to a different spreadsheet
- `update_sheet_properties` — change tab title, color, grid size, frozen rows/cols

## Security

`credentials.json` and `token.json` are gitignored and must never be committed. Each machine generates its own `token.json` via `npm run auth`.

