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
def get_core_properties() -> str:
    """
    Read the document's core properties (title, author, subject, keywords, dates).
    Useful for spotting metadata inherited from a template or another document.
    """
    return json.dumps(_ed().core_properties(), ensure_ascii=False, indent=2)


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
def replace_in_paragraph(idx: int, old: str, new: str, occurrence: int = 1) -> str:
    """
    Replace one occurrence of `old` with `new` inside paragraph idx, editing only
    the runs that hold it. Keeps field codes (SEQ caption numbers, REF
    cross-references, EndNote citations) and the formatting of untouched runs;
    use this instead of replace_paragraph on any paragraph get_fields() lists.
    `old` must lie outside field codes and may span several runs.
    Returns the paragraph's new text.
    """
    text = _ed().replace_in_paragraph(idx, old, new, occurrence)
    return f"P{idx} edited.\nNEW: {text[:300]}"


@mcp.tool()
def set_field_result(idx: int, instr_contains: str, text: str) -> str:
    """
    Set the displayed result of a complex field in paragraph idx (the first
    field whose instruction contains `instr_contains`, e.g. "SEQ Figure"),
    such as a stale caption number. The field code is kept, so Word still
    recomputes it when fields are updated. Returns the paragraph's text.
    """
    return f"P{idx} field result set.\nNEW: {_ed().set_field_result(idx, instr_contains, text)[:300]}"


@mcp.tool()
def unlink_field(idx: int, instr_contains: str, text: Optional[str] = None) -> str:
    """
    Replace a complex field in paragraph idx (the first whose instruction
    contains `instr_contains`, e.g. "REF _Ref243577315") by plain text: `text`
    if given, else the field's current result. For cross-references whose
    target was deleted ("Error! Reference source not found."). Returns the
    paragraph's text.
    """
    return f"P{idx} field unlinked.\nNEW: {_ed().unlink_field(idx, instr_contains, text)[:300]}"


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
def get_inline_objects(idx: int) -> str:
    """
    List a paragraph's non-run inline objects (content controls, hyperlinks,
    bookmarks) with their positions. Cross-references and reference-manager
    citations live here and are invisible to get_paragraph's text, so check this
    before rewriting a paragraph that may contain them. Feed the positions to
    set_paragraph_segments.
    """
    return json.dumps(_ed().inline_objects(idx), ensure_ascii=False, indent=2)


@mcp.tool()
def set_paragraph_segments(idx: int, segments: list) -> str:
    """
    Rewrite a paragraph as an ordered mix of text and preserved inline objects.
    Each segment is either a string (literal text) or an integer position from
    get_inline_objects(idx). Use this to reposition or de-duplicate Word
    cross-references and citations without destroying them — replace_paragraph
    would silently drop them.
    """
    new_text = _ed().set_segments(idx, segments)
    return f"P{idx} rebuilt: {new_text[:200]}"


@mcp.tool()
def set_paragraph_runs(idx: int, parts: list) -> str:
    """
    Rewrite a paragraph as runs with explicit formatting, for paragraphs that mix
    bold/italic (e.g. a bold term followed by an italic definition).
    parts is a list of objects: [{"text": "AMR ", "bold": true},
    {"text": "Antimicrobial Resistance", "italic": true}].
    Runs inherit the paragraph's existing font and language.
    """
    new_text = _ed().set_runs(idx, parts)
    return f"P{idx} rebuilt with {len(parts)} run(s): {new_text[:200]}"


@mcp.tool()
def get_list_numbering(idx: int) -> str:
    """
    Report whether a paragraph carries automatic list numbering, as {num_id, level}
    or null. Applying a heading style strips numbering, which silently breaks
    "Section N" cross-references elsewhere in the text — check this after any
    style change on a numbered heading.
    """
    return json.dumps(_ed().list_numbering(idx), ensure_ascii=False)


@mcp.tool()
def set_list_numbering(idx: int, num_id: Optional[str] = None, level: str = "0") -> str:
    """
    Attach automatic list numbering to a paragraph, or remove it with num_id=null.
    Take num_id from a sibling that still numbers correctly (get_list_numbering).
    Use this to restore numbering lost when a numbered paragraph was restyled as a
    heading.
    """
    result = _ed().set_list_numbering(idx, num_id, level)
    return f"P{idx} numbering: {json.dumps(result, ensure_ascii=False)}"


@mcp.tool()
def set_table_cell(table_idx: int, row: int, col: int, text: str) -> str:
    """
    Replace the text of a single table cell, keeping its character formatting.
    Use get_tables() first to find the table index and row/column positions.
    """
    old = _ed().set_cell_text(table_idx, row, col, text)
    return f"Table {table_idx} cell ({row},{col}) updated.\nOLD: {old[:150]}\nNEW: {text[:150]}"


@mcp.tool()
def insert_paragraph_like(model_idx: int, after_idx: int, text: str) -> str:
    """
    Insert a paragraph after after_idx, cloning the formatting of the paragraph at
    model_idx (style, list numbering, indentation, character formatting).
    Use this instead of insert_paragraph when adding an item to a numbered or
    bulleted list, which a plain style name cannot reproduce.
    Returns the index of the newly inserted paragraph.
    """
    new_idx = _ed().insert_like(model_idx, after_idx, text)
    return f"Inserted at P{new_idx}, formatted like P{model_idx}."


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
def set_core_properties(title: Optional[str] = None, subject: Optional[str] = None,
                        author: Optional[str] = None, keywords: Optional[str] = None,
                        comments: Optional[str] = None, category: Optional[str] = None,
                        content_status: Optional[str] = None,
                        last_modified_by: Optional[str] = None,
                        version: Optional[str] = None) -> str:
    """
    Set document core properties. Only the arguments you pass are changed.
    Use this to clear metadata inherited from a template (e.g. a wrong title
    or another document's author) before sharing a file.
    """
    props = _ed().set_core_properties(
        title=title, subject=subject, author=author, keywords=keywords,
        comments=comments, category=category, content_status=content_status,
        last_modified_by=last_modified_by, version=version)
    return json.dumps(props, ensure_ascii=False, indent=2)


@mcp.tool()
def insert_table(after_idx: int, rows: list, style: str = "Table Grid",
                 header: bool = True) -> str:
    """
    Insert a table immediately after the paragraph at after_idx.
    rows is a list of row lists, e.g. [["Field","Value"],["WP","WP5"]];
    short rows are padded. The first row is bolded unless header=false.
    Returns the new table's index.
    """
    tbl_idx = _ed().insert_table(after_idx, rows, style=style, header=header)
    return f"Inserted table {tbl_idx} after P{after_idx} ({len(rows)} rows, style '{style}')."


@mcp.tool()
def delete_table(table_idx: int) -> str:
    """
    Delete the table at table_idx (from get_tables()).
    Returns the removed table's dimensions and first row for confirmation.
    """
    rows = _ed().delete_table(table_idx)
    head = rows[0] if rows else []
    return (f"Deleted table {table_idx} ({len(rows)} rows). "
            f"First row was: {json.dumps(head, ensure_ascii=False)[:200]}")


@mcp.tool()
def delete_table_row(table_idx: int, row: int) -> str:
    """
    Delete a single row from the table at table_idx.
    Use get_tables() first to find the table and row indices.
    Returns the removed row's cell texts for confirmation.
    """
    cells = _ed().delete_table_row(table_idx, row)
    return f"Deleted row {row} of table {table_idx}: {json.dumps(cells, ensure_ascii=False)[:200]}"


@mcp.tool()
def insert_image(after_idx: int, image_path: str,
                 width_cm: Optional[float] = None, center: bool = True) -> str:
    """
    Insert an image (PNG/JPG/...) in its own paragraph immediately after the
    paragraph at after_idx. width_cm scales the image, keeping aspect ratio;
    omit it to use the natural size capped at the text-column width.
    The image is centered unless center=false. Pair with set_caption on a
    following paragraph for an auto-numbered figure caption.
    Returns the new paragraph's index.
    """
    new_idx = _ed().insert_image(after_idx, image_path, width_cm=width_cm, center=center)
    return f"Inserted image '{image_path}' at P{new_idx}."


@mcp.tool()
def add_bookmark(idx: int, name: str) -> str:
    """
    Bookmark the paragraph at idx so internal links can target it (typically a
    heading). Names must be unique, start with a letter or underscore, and
    contain only letters, digits, or underscores (max 40 chars).
    """
    _ed().add_bookmark(idx, name)
    return f"Bookmark '{name}' added on P{idx}."


@mcp.tool()
def bookmark_caption(idx: int, name: Optional[str] = None) -> str:
    """
    Bookmark the label and number of the caption at idx ("Figure 7": the text up
    to and including its SEQ field), as Word does for cross-reference targets.
    Returns the bookmark name (an existing one if the caption is already
    bookmarked; a new Word-style _Ref… name otherwise).
    """
    return _ed().bookmark_caption(idx, name)


@mcp.tool()
def add_cross_reference(idx: int, text: str, bookmark: str, occurrence: int = 1) -> str:
    """
    Turn the nth occurrence of `text` in paragraph idx (a hand-typed "Figure 7")
    into a real cross-reference: a REF field to `bookmark` (from
    bookmark_caption) displaying `text`, which renumbers on field update.
    The text must lie within one plain run. Returns the paragraph's text.
    """
    return f"P{idx} cross-reference added.\nNEW: {_ed().add_cross_reference(idx, text, bookmark, occurrence)[:300]}"


@mcp.tool()
def add_internal_link(idx: int, text: str, bookmark: str, occurrence: int = 1) -> str:
    """
    Turn the nth occurrence of `text` in paragraph idx into an internal
    hyperlink to `bookmark` (Ctrl+Click navigation in Word; the visible text is
    unchanged). Create the target first with add_bookmark. The occurrence must
    lie within a single run — check get_paragraph_detail if the match fails.
    """
    _ed().add_internal_link(idx, text, bookmark, occurrence=occurrence)
    return f"Linked '{text}' (occurrence {occurrence}) in P{idx} -> bookmark '{bookmark}'."


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
