"""
MCP server exposing DocxEditor as tools for Claude.

Usage:
  python server.py          # stdio mode (for Claude Code / claude_desktop_config.json)

Register with Claude Code:
  claude mcp add docx-editor python C:\\Users\\jeco\\Dev\\ODIN\\docx_tools\\server.py
"""

import json
import sys
from pathlib import Path
from typing import Optional

# Allow running from this directory or as a package
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx_tools.editor import DocxEditor
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("docx-editor")

# Module-level state: one document open at a time
_editor: Optional[DocxEditor] = None


def _ed() -> DocxEditor:
    if _editor is None:
        raise RuntimeError("No document open. Call open_document first.")
    return _editor


# ── Document lifecycle ────────────────────────────────────────────────────────

@mcp.tool()
def create_document(path: str) -> str:
    """Create a new blank Word .docx file at path and open it for editing."""
    global _editor
    _editor = DocxEditor.create(path)
    return f"Created and opened: {path}\n0 paragraphs total."


@mcp.tool()
def open_document(path: str) -> str:
    """Open a Word .docx file for editing. Must be called before any other tool."""
    global _editor
    _editor = DocxEditor(path)
    n = _editor.paragraph_count()
    return f"Opened: {path}\n{n} paragraphs total."


@mcp.tool()
def save_document(path: Optional[str] = None) -> str:
    """
    Save the open document.
    If path is omitted, saves in-place (overwrites the source).
    If path is provided, saves a copy there.
    """
    out = _ed().save(path)
    return f"Saved to: {out}"


@mcp.tool()
def document_info() -> str:
    """Return the open document's path and paragraph count."""
    ed = _ed()
    return json.dumps({"path": str(ed.path), "paragraph_count": ed.paragraph_count()})


# ── Read ─────────────────────────────────────────────────────────────────────

@mcp.tool()
def list_paragraphs(start: int = 0, end: int = 50,
                    include_empty: bool = False) -> str:
    """
    List paragraphs in index range [start, end).
    Each entry: {idx, style, text}.
    Use include_empty=true to include blank paragraphs.
    """
    paras = _ed().paragraphs(include_empty=include_empty)
    subset = [p for p in paras if start <= p["idx"] < end]
    return json.dumps(subset, ensure_ascii=False, indent=2)


@mcp.tool()
def get_paragraph(idx: int) -> str:
    """Get a single paragraph by its index. Returns {idx, style, text}."""
    return json.dumps(_ed().get(idx), ensure_ascii=False)


@mcp.tool()
def get_context(idx: int, before: int = 2, after: int = 2) -> str:
    """Return a window of paragraphs around idx for orientation."""
    return json.dumps(_ed().context(idx, before, after), ensure_ascii=False, indent=2)


@mcp.tool()
def find_paragraphs(query: str, style: Optional[str] = None,
                    regex: bool = False) -> str:
    """
    Search for paragraphs whose text contains query (case-insensitive).
    Optionally filter by style name (e.g. 'Heading 2').
    Set regex=true to use query as a regular expression.
    Returns list of {idx, style, text}.
    """
    results = _ed().find(query, style=style, regex=regex)
    return json.dumps(results, ensure_ascii=False, indent=2)


@mcp.tool()
def get_headings() -> str:
    """Return all headings with {idx, style, level, text}."""
    return json.dumps(_ed().headings(), ensure_ascii=False, indent=2)


@mcp.tool()
def list_styles() -> str:
    """List all style names available in the document (useful before insert_paragraph)."""
    return json.dumps(_ed().styles(), ensure_ascii=False, indent=2)


@mcp.tool()
def get_comments() -> str:
    """Extract all Word comments. Returns list of {id, author, date, text}."""
    return json.dumps(_ed().comments(), ensure_ascii=False, indent=2)


@mcp.tool()
def get_tracked_changes() -> str:
    """Extract all tracked changes (insertions and deletions) from the document."""
    return json.dumps(_ed().tracked_changes(), ensure_ascii=False, indent=2)


@mcp.tool()
def get_tables() -> str:
    """Extract all tables as {table_idx, rows} where rows is a list of lists of cell text."""
    return json.dumps(_ed().tables(), ensure_ascii=False, indent=2)


@mcp.tool()
def get_fields(idx: Optional[int] = None) -> str:
    """
    List Word field codes (SEQ, REF, ADDIN, etc.), which plain-text tools can't see.
    Pass idx for a single paragraph, or omit for the whole document.
    Returns [{idx, fields:[{kind, type, instr, result}]}]. Use this to tell an
    auto-numbered caption (a SEQ field) from a hand-typed number (no field).
    """
    return json.dumps(_ed().fields(idx), ensure_ascii=False, indent=2)


@mcp.tool()
def get_paragraph_detail(idx: int) -> str:
    """
    Full detail for one paragraph: style, plain text, accepted_text (tracked
    changes accepted), raw_text (deletions shown), field codes, and
    has_tracked_changes. Use when a paragraph looks garbled due to tracked changes.
    """
    return json.dumps(_ed().detail(idx), ensure_ascii=False, indent=2)


# ── Write ─────────────────────────────────────────────────────────────────────

@mcp.tool()
def replace_paragraph(idx: int, new_text: str) -> str:
    """
    Replace the full text of the paragraph at idx.
    Preserves the character formatting (bold, italic, font) of the first run.
    Returns old and new text for confirmation.
    """
    old = _ed().replace(idx, new_text)
    return f"P{idx} replaced.\nOLD: {old[:200]}\nNEW: {new_text[:200]}"


@mcp.tool()
def append_paragraph(text: str, style: str = "Body Text") -> str:
    """
    Append a new paragraph at the end of the document.
    Use list_styles() to see valid style names.
    Returns the index of the newly appended paragraph.
    """
    new_idx = _ed().append(text, style)
    return f"Appended at P{new_idx} (style: '{style}')."


@mcp.tool()
def insert_paragraph(after_idx: int, text: str, style: str = "Body Text") -> str:
    """
    Insert a new paragraph immediately after after_idx.
    Use list_styles() to see valid style names.
    Returns the index of the newly inserted paragraph.
    """
    new_idx = _ed().insert_after(after_idx, text, style)
    return f"Inserted at P{new_idx} (style: '{style}')."


@mcp.tool()
def insert_paragraph_before(before_idx: int, text: str, style: str = "Body Text") -> str:
    """
    Insert a new paragraph immediately before before_idx.
    Returns the index of the newly inserted paragraph.
    """
    new_idx = _ed().insert_before(before_idx, text, style)
    return f"Inserted at P{new_idx} (style: '{style}')."


@mcp.tool()
def delete_paragraph(idx: int) -> str:
    """
    Delete the paragraph at idx.
    Returns the deleted text for confirmation.
    """
    old = _ed().delete(idx)
    return f"Deleted P{idx}: '{old[:200]}'"


@mcp.tool()
def set_paragraph_style(idx: int, style: str) -> str:
    """Change the style of the paragraph at idx (e.g. 'Heading 2', 'Body Text')."""
    _ed().set_style(idx, style)
    return f"P{idx} style set to '{style}'."


@mcp.tool()
def set_caption(idx: int, text: str, label: str = "Figure",
                sep: str = ": ", style: Optional[str] = None) -> str:
    """
    Rewrite paragraph idx as a clean AUTO-NUMBERED caption:
    '<label> {SEQ <label> * ARABIC}<sep><text>'.
    Clears any existing runs, stray manual numbers, and half-deleted field cruft.
    label is the sequence name ('Figure' or 'Table'); sep defaults to ': ';
    pass style='Caption' to also set the paragraph style.
    The number recomputes in Word on open / Ctrl+A+F9. Returns the new plain text.
    """
    new_text = _ed().set_caption(idx, text, label=label, sep=sep, style=style)
    return f"P{idx} set to caption: {new_text[:200]}"


@mcp.tool()
def accept_changes(author: Optional[str] = None) -> str:
    """
    Accept all tracked changes in the document (insertions kept, deletions removed,
    format-change markers cleared). Pass author to accept only one person's changes.
    Returns the number of change elements resolved.
    """
    n = _ed().accept_changes(author)
    return f"Accepted {n} tracked-change element(s)" + (f" by {author}" if author else "") + "."


@mcp.tool()
def reject_changes(author: Optional[str] = None) -> str:
    """
    Reject all tracked changes (insertions removed, deletions restored, format
    changes reverted). Pass author to reject only one person's changes.
    Returns the number of change elements resolved.
    """
    n = _ed().reject_changes(author)
    return f"Rejected {n} tracked-change element(s)" + (f" by {author}" if author else "") + "."


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    mcp.run()
