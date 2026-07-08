# docx_tools

A paragraph-level Word document editing API, exposed as a [Model Context Protocol (MCP)](https://modelcontextprotocol.io/) server so AI assistants (Claude, etc.) can read and edit `.docx` files directly.

Two interfaces in one package:

- **`DocxEditor`** — a Python class you can import and use in your own scripts
- **`server.py`** — an MCP server that exposes `DocxEditor` as tools callable by Claude Code or any MCP-compatible client

---

## Requirements

- Python 3.10+
- [Claude Code](https://claude.ai/code) (if using the MCP server)

---

## Installation

### 1. Clone the repo

```bash
git clone git@github.com:NORCE-Teknologi/docx_tools.git
cd docx_tools
```

### 2. Create a virtual environment and install dependencies

**Linux / macOS**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**Windows (PowerShell)**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## MCP server setup (Claude Code)

Register the server once, globally, so it's available in all your Claude Code workspaces:

**Linux / macOS**
```bash
claude mcp add --scope user docx-editor \
  /path/to/docx_tools/.venv/bin/python \
  /path/to/docx_tools/server.py
```

**Windows (Git Bash)**
```bash
claude mcp add --scope user docx-editor \
  "C:/path/to/docx_tools/.venv/Scripts/python.exe" \
  "C:/path/to/docx_tools/server.py"
```

Verify it registered:
```bash
claude mcp list
```

You should see `docx-editor` with status `Connected`.

### Optional: auto-approve all docx-editor tools

Add this to `~/.claude/settings.json` so Claude never prompts for permission on these tools:

```json
{
  "permissions": {
    "allow": [
      "mcp__docx-editor__*"
    ]
  }
}
```

---

## Quick start example

After installing and registering the MCP server, you can drive it from the Claude CLI in one shot:

**Linux / macOS (bash)**
```bash
# 1. Create a minimal test document in the current directory
python -c "
from docx import Document
d = Document()
d.add_heading('Hello World', level=1)
d.add_paragraph('This is a test document created by docx_tools MCP server.')
d.save('hello.docx')
"

# 2. Ask Claude to open it — $(pwd) expands to the current directory automatically
claude -p "Using the docx-editor MCP tools: open $(pwd)/hello.docx, list all paragraphs, and tell me what the document contains." --allowedTools "mcp__docx-editor__*"
```

**Windows (PowerShell)**
```powershell
# 1. Create a minimal test document in the current directory
python -c "
from docx import Document
d = Document()
d.add_heading('Hello World', level=1)
d.add_paragraph('This is a test document created by docx_tools MCP server.')
d.save('hello.docx')
"

# 2. Ask Claude to open it — $PWD expands to the current directory automatically
claude -p "Using the docx-editor MCP tools: open $PWD\hello.docx, list all paragraphs, and tell me what the document contains." --allowedTools "mcp__docx-editor__*"
```

Expected output (approximate):

```
The document contains 2 paragraphs:
  P0  [Heading 1]  Hello World
  P1  [Normal]     This is a test document created by docx_tools.
```

You can go further in the same command — ask Claude to edit, insert, or reformat paragraphs and save the result:

**bash:** `claude -p "Using the docx-editor MCP tools: open $(pwd)/hello.docx, replace the heading text with 'Goodbye World', then save the document." --allowedTools "mcp__docx-editor__*"`

**PowerShell:** `claude -p "Using the docx-editor MCP tools: open $PWD\hello.docx, replace the heading text with 'Goodbye World', then save the document." --allowedTools "mcp__docx-editor__*"`

---

## MCP tool reference

All tools operate on a single document kept open in server memory. Call `open_document` first, then use read/write tools freely, and call `save_document` when done.

### Document lifecycle

| Tool | Parameters | Description |
|------|-----------|-------------|
| `open_document` | `path` | Open a `.docx` file. Must be called first. |
| `save_document` | `path?` | Save in-place (no path) or to a new path. |
| `document_info` | — | Returns path and paragraph count. |

### Read tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `list_paragraphs` | `start`, `end`, `include_empty` | List paragraphs in index range with style and text. |
| `get_paragraph` | `idx` | Get a single paragraph by index. |
| `get_context` | `idx`, `before`, `after` | Window of paragraphs around a target index. |
| `find_paragraphs` | `query`, `style?`, `regex?` | Search paragraphs by text or regex, optionally filtered by style. |
| `get_headings` | — | All headings with level and index. |
| `list_styles` | — | All style names available in the document. |
| `get_comments` | — | All Word comments with author, date, and text. |
| `get_tracked_changes` | — | All tracked insertions and deletions. |
| `get_tables` | — | All tables as lists of cell text. |
| `get_fields` | `idx?` | List Word field codes (SEQ, REF, ADDIN…) for one paragraph or the whole doc — distinguishes an auto-numbered caption (SEQ field) from a hand-typed number. |
| `get_paragraph_detail` | `idx` | Style, plain text, `accepted_text`, `raw_text` (deletions shown), field codes, and `has_tracked_changes` — for paragraphs garbled by tracked changes. |

### Write tools

| Tool | Parameters | Description |
|------|-----------|-------------|
| `replace_paragraph` | `idx`, `new_text` | Replace paragraph text (preserves character formatting of first run). |
| `insert_paragraph` | `after_idx`, `text`, `style` | Insert a new paragraph after the given index. |
| `insert_paragraph_before` | `before_idx`, `text`, `style` | Insert a new paragraph before the given index. |
| `delete_paragraph` | `idx` | Delete a paragraph. |
| `set_paragraph_style` | `idx`, `style` | Change a paragraph's style (e.g. `Heading 2`, `Body Text`). |
| `set_caption` | `idx`, `text`, `label?`, `sep?`, `style?` | Rewrite a paragraph as a clean auto-numbered caption `<label> {SEQ <label> \* ARABIC><sep><text>`, clearing stray manual numbers / half-deleted field cruft. |
| `accept_changes` | `author?` | Accept all tracked changes (or just one author's). |
| `reject_changes` | `author?` | Reject all tracked changes (or just one author's). |

> **Index stability**: paragraph indices shift after insert/delete operations. Re-query with `list_paragraphs` or `find_paragraphs` if you need to locate paragraphs after editing.

---

## Python API

You can use `DocxEditor` directly without the MCP server:

```python
from docx_tools import DocxEditor

ed = DocxEditor("my_document.docx")

# List paragraphs 10-20
for p in ed.paragraphs()[10:20]:
    print(p["idx"], p["style"], p["text"])

# Find a section heading
results = ed.find("Introduction", style="Heading 1")

# Replace paragraph text
ed.replace(42, "Updated text here.")

# Insert a new paragraph after index 15
ed.insert_after(15, "New paragraph content.", style="Body Text")

# Save
ed.save()                      # in-place
ed.save("output_copy.docx")   # to a new file
```

### `DocxEditor` methods

| Method | Returns | Description |
|--------|---------|-------------|
| `paragraphs(include_empty=False)` | `list[dict]` | All paragraphs as `{idx, style, text}`. |
| `get(idx)` | `dict` | Single paragraph by index. |
| `find(query, style=None, regex=False)` | `list[dict]` | Search paragraphs. |
| `headings()` | `list[dict]` | All headings with `level`. |
| `styles()` | `list[str]` | All style names. |
| `paragraph_count()` | `int` | Total paragraph count. |
| `context(idx, before=2, after=2)` | `list[dict]` | Window around an index. |
| `comments()` | `list[dict]` | Word comments. |
| `tracked_changes()` | `list[dict]` | Tracked insertions/deletions. |
| `tables()` | `list[dict]` | Tables as cell-text grids. |
| `fields(idx=None)` | `list[dict]` | Field codes per paragraph (SEQ/REF/ADDIN…). |
| `detail(idx)` | `dict` | Style, plain/accepted/raw text, fields, `has_tracked_changes`. |
| `replace(idx, new_text)` | `str` | Replace text, returns old text. |
| `insert_after(idx, text, style)` | `int` | Insert after idx, returns new index. |
| `insert_before(idx, text, style)` | `int` | Insert before idx, returns new index. |
| `delete(idx)` | `str` | Delete paragraph, returns old text. |
| `set_style(idx, style)` | `None` | Change paragraph style. |
| `set_caption(idx, text, label='Figure', sep=': ', style=None)` | `str` | Rewrite a paragraph as an auto-numbered SEQ-field caption. |
| `accept_changes(author=None)` | `int` | Accept tracked changes, returns count resolved. |
| `reject_changes(author=None)` | `int` | Reject tracked changes, returns count resolved. |
| `save(path=None)` | `Path` | Save document. |

---

## Notes

- Only one document can be open at a time in the MCP server (module-level state). Re-calling `open_document` replaces the current document.
- `replace_paragraph` preserves the character formatting (`w:rPr`) of the first run in the paragraph — bold, italic, font, etc. — on the replacement text.
- Style names must match exactly. Use `list_styles` / `styles()` to discover valid names for a given document.
- Word documents must not be open in Microsoft Word when you call `open_document` — Word locks the file.

---

## License

[GNU Affero General Public License v3.0 or later (AGPL-3.0-or-later)](LICENSE.md) — free to use, modify, and distribute under the terms of the AGPL. Any modified version deployed as a network service must also be released under the AGPL.
