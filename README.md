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

## Tools (130)

### Drive — Files
- `search_drive` — search files by name or content
- `get_file_metadata` — file info (name, type, size, owner, dates, parents)
- `rename_file` — rename any file or folder
- `move_file` — move a file to a different folder
- `copy_file` — copy any file, optionally rename and relocate
- `create_folder` — create a new folder
- `trash_file` — move to trash or permanently delete
- `list_folder` — list contents of a folder
- `upload_file` — upload a local file to Google Drive
- `download_file` — download a Drive file to local filesystem
- `about` — account info and storage quota

### Drive — Permissions
- `list_permissions` — list who has access to a file
- `create_permission` — share a file with a user, group, domain, or anyone
- `update_permission` — change a permission's role
- `delete_permission` — remove access

### Drive — Comments & Replies
- `add_comment` — add a comment, optionally anchored to quoted text
- `list_comments` — list comments on a file
- `resolve_comment` — resolve a comment
- `update_comment` — edit a comment
- `delete_comment` — delete a comment
- `list_replies` — list replies on a comment
- `create_reply` — reply to a comment (with optional resolve/reopen action)
- `delete_reply` — delete a reply

### Drive — Revisions
- `list_revisions` — list revision history
- `get_revision` — get details of a specific revision

### Docs — Content
- `create_doc` — create a new Google Doc
- `read_doc` — read a Google Doc as plain text / markdown
- `read_doc_structure` — read full structure with indices, formatting, object IDs, styles
- `edit_doc` — insert, delete, replace, or append text
- `copy_doc` — copy a Google Doc
- `rename_doc` — rename a Google Doc
- `export_doc` — export to PDF, DOCX, TXT, HTML, RTF, or EPUB

### Docs — Formatting
- `format_doc` — bold, italic, underline, font, colors, headings, alignment, hyperlinks
- `update_document_style` — page margins, size, background color
- `update_named_style` — modify heading/paragraph styles globally
- `update_section_style` — section-level columns, margins, headers/footers
- `update_list` — apply or remove bulleted/numbered list formatting

### Docs — Tables
- `insert_table` — insert a table
- `modify_table` — insert/delete rows and columns
- `merge_table_cells` — merge cells in a table
- `unmerge_table_cells` — unmerge cells in a table
- `format_table` — cell borders, backgrounds, padding, column widths, row heights
- `pin_table_header_rows` — pin header rows in a table

### Docs — Special Elements
- `insert_image` — insert an inline image from a URL
- `replace_image` — replace an existing image
- `delete_positioned_object` — delete a floating image or object
- `insert_page_break` — insert a page break
- `insert_section_break` — insert a section break (next page or continuous)
- `insert_date` — insert a date chip
- `insert_person` — insert a person mention (@ chip)
- `insert_rich_link` — insert a smart chip linking to a Google file
- `create_footnote` — create a footnote

### Docs — Named Ranges
- `create_named_range` — create a named range (bookmark)
- `delete_named_range` — delete a named range
- `list_named_ranges` — list all named ranges
- `replace_named_range_content` — replace content in a named range (templates)

### Docs — Headers, Footers & Tabs
- `update_header_footer` — create or replace headers/footers
- `delete_header` — delete a header
- `delete_footer` — delete a footer
- `list_doc_tabs` — list tabs in a doc
- `rename_doc_tab` — rename a doc tab
- `add_doc_tab` — add a new tab
- `delete_doc_tab` — delete a tab

### Sheets — Data
- `read_sheet` — read a range
- `batch_read_sheet` — read multiple ranges in one call
- `append_sheet_row` — append a row after the last data row
- `update_sheet_cell` — update a single cell or range
- `batch_update_sheet` — write to multiple ranges in one call
- `clear_sheet` — clear values in a range
- `batch_clear_sheet` — clear multiple ranges in one call
- `insert_rows` — insert rows at a position, optionally with data
- `delete_rows` — delete rows
- `delete_columns` — delete columns
- `find_replace` — find and replace text
- `sort_range` — sort a range by one or more columns
- `export_sheet` — export to CSV, TSV, XLSX, PDF, or HTML
- `read_cell_format` — read cell formatting, borders, validation, notes, hyperlinks, rich text runs
- `set_data_validation` — set dropdowns, number rules, or custom formulas
- `clear_data_validation` — clear validation rules from a range
- `text_to_columns` — split text into columns by delimiter
- `set_cell_note` — set, update, or clear a note on a cell
- `set_cell_rich_text` — per-character formatting within a cell (mixed bold/color/links)
- `copy_paste` — copy a range with paste options (values, format, all, transpose)
- `cut_paste` — cut and paste a range
- `auto_fill` — extend a pattern or formula (fill handle)
- `trim_whitespace` — trim leading/trailing whitespace from cells
- `update_spreadsheet_properties` — update title, locale, timezone, recalculation
- `append_dimension` — add empty rows or columns at the end of a sheet
- `set_dimension_size` — set specific row height or column width in pixels
- `add_pivot_table` — create a pivot table from a data range

### Sheets — Formatting
- `format_cells` — colors, fonts, alignment, number formats, wrap strategy
- `update_borders` — cell borders (solid, dashed, dotted, thick, double)
- `merge_cells` — merge a range
- `unmerge_cells` — unmerge a range
- `add_conditional_formatting` — boolean or gradient rules
- `read_conditional_formats` — read full rule details (conditions, values, formats, gradients)
- `update_conditional_formatting` — update an existing rule by index
- `delete_conditional_formatting` — delete a rule by index
- `add_banding` — alternating row colors
- `update_banding` — update banding colors
- `delete_banding` — remove banding
- `freeze` — freeze rows/columns
- `auto_resize` — auto-fit columns or rows to content

### Sheets — Objects & Charts
- `list_sheet_objects` — list all charts (with full specs), filters, banding, protected ranges, named ranges
- `add_chart` — create a chart (bar, line, pie, column, area, scatter, combo)
- `update_chart` — update an existing chart (type, data, colors, title, axes)
- `delete_embedded_object` — delete a chart or embedded object
- `set_basic_filter` — set auto-filter on a range
- `clear_basic_filter` — remove auto-filter
- `add_filter_view` — create a named filter view
- `update_filter_view` — update filter view (title, range, criteria, sort)
- `delete_filter_view` — delete a filter view
- `add_slicer` — add an interactive slicer (filter control)

### Sheets — Named Ranges
- `add_sheet_named_range` — create a named range (e.g. "SalesData" -> Sheet1!A1:D100)
- `update_sheet_named_range` — rename or change a named range reference
- `delete_sheet_named_range` — delete a named range

### Sheets — Structure
- `get_sheet_metadata` — title, URL, sheet names, IDs, dimensions, frozen rows/cols
- `create_spreadsheet` — create a new spreadsheet
- `insert_sheet` — add a tab
- `delete_sheet` — remove a tab
- `rename_sheet` — rename a tab
- `duplicate_sheet` — duplicate a tab within the same spreadsheet
- `copy_sheet_to` — copy a tab to a different spreadsheet
- `update_sheet_properties` — change tab title, color, grid size, frozen rows/cols
- `add_protected_range` — protect a range (with optional warning-only mode)
- `update_protected_range` — modify a protected range
- `delete_protected_range` — remove range protection
- `move_dimension` — move rows or columns to a new position
- `add_dimension_group` — create collapsible row/column groups
- `delete_dimension_group` — remove row/column groups

## Security

`credentials.json` and `token.json` are gitignored and must never be committed. Each machine generates its own `token.json` via `npm run auth`.

