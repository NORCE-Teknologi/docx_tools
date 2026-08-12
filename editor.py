"""
Core Word document editing library.
Wraps python-docx with a convenience API for paragraph-level read/write.
"""

import copy
import re
from pathlib import Path
from typing import Optional

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


class DocxEditor:
    """Open a .docx, query and edit paragraphs, save back."""

    def __init__(self, path: str):
        self.path = Path(path)
        self.doc = docx.Document(str(self.path))

    @classmethod
    def create(cls, path: str) -> "DocxEditor":
        """Create a new blank .docx at path and return an editor for it."""
        new_doc = docx.Document()
        # Remove the default empty paragraph Word always adds
        for p in list(new_doc.paragraphs):
            p._element.getparent().remove(p._element)
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        new_doc.save(str(p))
        return cls(str(p))

    # ── Read ──────────────────────────────────────────────────────────────────

    def paragraphs(self, include_empty: bool = False) -> list:
        """Return all paragraphs as a list of dicts {idx, style, text}."""
        out = []
        for i, p in enumerate(self.doc.paragraphs):
            if not include_empty and not p.text.strip():
                continue
            out.append({"idx": i, "style": p.style.name, "text": p.text})
        return out

    def get(self, idx: int) -> dict:
        """Get a single paragraph by index."""
        p = self.doc.paragraphs[idx]
        return {"idx": idx, "style": p.style.name, "text": p.text}

    def find(self, query: str, style: Optional[str] = None,
             regex: bool = False, case_sensitive: bool = False) -> list:
        """
        Search paragraphs containing query.
        Returns list of {idx, style, text}.
        """
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(query, flags) if regex else None
        results = []
        for i, p in enumerate(self.doc.paragraphs):
            t = p.text
            if regex:
                matched = bool(pattern.search(t))
            else:
                matched = query.lower() in t.lower() if not case_sensitive else query in t
            if matched:
                if style is None or p.style.name == style:
                    results.append({"idx": i, "style": p.style.name, "text": t})
        return results

    def headings(self) -> list:
        """Return all heading paragraphs with their level."""
        out = []
        for i, p in enumerate(self.doc.paragraphs):
            name = p.style.name
            if name.startswith("Heading"):
                try:
                    level = int(name.split()[-1])
                except ValueError:
                    level = 0
                out.append({"idx": i, "style": name, "level": level, "text": p.text})
        return out

    def styles(self) -> list:
        """Return all style names available in the document."""
        return sorted(s.name for s in self.doc.styles if s.name)

    def paragraph_count(self) -> int:
        return len(self.doc.paragraphs)

    def context(self, idx: int, before: int = 2, after: int = 2) -> list:
        """Return a window of paragraphs around idx."""
        paras = self.doc.paragraphs
        start = max(0, idx - before)
        end = min(len(paras), idx + after + 1)
        return [{"idx": i, "style": paras[i].style.name, "text": paras[i].text}
                for i in range(start, end)]

    def comments(self) -> list:
        """Extract all Word comments. Returns list of {id, author, date, text}."""
        out = []
        try:
            for c in self.doc.comments:
                ts = c.timestamp
                date = ts.isoformat() if ts is not None else ""
                out.append({"id": str(c.comment_id), "author": c.author or "",
                            "date": date, "text": c.text})
        except Exception as e:
            out.append({"error": str(e)})
        return out

    def tracked_changes(self) -> list:
        """Extract tracked insertions and deletions. Returns list of {type, text, author}."""
        ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        out = []
        for elem in self.doc.element.body.iter():
            tag = elem.tag
            if tag == f"{{{ns}}}ins":
                text = "".join(t.text or "" for t in elem.iter(f"{{{ns}}}t"))
                out.append({"type": "insert", "author": elem.get(f"{{{ns}}}author", ""), "text": text})
            elif tag == f"{{{ns}}}del":
                text = "".join(t.text or "" for t in elem.iter(f"{{{ns}}}delText"))
                out.append({"type": "delete", "author": elem.get(f"{{{ns}}}author", ""), "text": text})
        return out

    def tables(self) -> list:
        """Extract all tables as list of {table_idx, rows}."""
        out = []
        for i, table in enumerate(self.doc.tables):
            rows = [[cell.text for cell in row.cells] for row in table.rows]
            out.append({"table_idx": i, "rows": rows})
        return out

    def core_properties(self) -> dict:
        """Return the document's core properties (docProps/core.xml)."""
        cp = self.doc.core_properties
        return {
            "title": cp.title or "",
            "subject": cp.subject or "",
            "author": cp.author or "",
            "last_modified_by": cp.last_modified_by or "",
            "keywords": cp.keywords or "",
            "comments": cp.comments or "",
            "category": cp.category or "",
            "content_status": cp.content_status or "",
            "version": cp.version or "",
            "revision": cp.revision,
            "created": cp.created.isoformat() if cp.created else "",
            "modified": cp.modified.isoformat() if cp.modified else "",
        }

    # ── Fields & tracked-change inspection ──────────────────────────────────────

    def _para_fields(self, pel) -> list:
        """List Word field codes in a paragraph element (simple + complex)."""
        fields = []
        # simple fields: <w:fldSimple w:instr="...">result</w:fldSimple>
        for fs in pel.iter(qn("w:fldSimple")):
            instr = (fs.get(qn("w:instr")) or "").strip()
            result = "".join(t.text or "" for t in fs.iter(qn("w:t")))
            fields.append({"kind": "simple", "instr": instr,
                           "type": instr.split()[0] if instr else "", "result": result})
        # complex fields: fldChar begin / instrText / separate / result / end
        state = None
        instr = result = ""
        for r in pel.iter(qn("w:r")):
            fc = r.find(qn("w:fldChar"))
            it = r.find(qn("w:instrText"))
            t = r.find(qn("w:t"))
            if fc is not None:
                typ = fc.get(qn("w:fldCharType"))
                if typ == "begin":
                    state, instr, result = "instr", "", ""
                elif typ == "separate":
                    state = "result"
                elif typ == "end":
                    s = instr.strip()
                    fields.append({"kind": "complex", "instr": s,
                                   "type": s.split()[0] if s else "", "result": result})
                    state = None
            elif it is not None and state == "instr":
                instr += it.text or ""
            elif t is not None and state == "result":
                result += t.text or ""
        return fields

    def _para_text(self, pel):
        """Return (accepted_text, raw_text) for a paragraph, honouring ins/del."""
        acc, raw = [], []

        def walk(el, indel=False):
            for ch in el:
                tag = ch.tag
                if tag in (qn("w:del"), qn("w:moveFrom")):
                    walk(ch, True)
                elif tag in (qn("w:ins"), qn("w:moveTo")):
                    walk(ch, indel)
                elif tag == qn("w:r"):
                    for s in ch:
                        if s.tag == qn("w:t"):
                            raw.append(s.text or "")
                            if not indel:
                                acc.append(s.text or "")
                        elif s.tag == qn("w:delText"):
                            raw.append(s.text or "")
                else:
                    walk(ch, indel)
        walk(pel)
        return "".join(acc), "".join(raw)

    def fields(self, idx: Optional[int] = None) -> list:
        """
        List field codes. If idx is given, just that paragraph; else the whole doc.
        Each entry: {idx, fields:[{kind, type, instr, result}]}.
        """
        if idx is not None:
            paras = [(idx, self.doc.paragraphs[idx])]
        else:
            paras = list(enumerate(self.doc.paragraphs))
        out = []
        for i, p in paras:
            fs = self._para_fields(p._element)
            if fs:
                out.append({"idx": i, "fields": fs})
        return out

    def detail(self, idx: int) -> dict:
        """Full paragraph detail: style, text, accepted/raw text, and field codes."""
        p = self.doc.paragraphs[idx]
        acc, raw = self._para_text(p._element)
        return {"idx": idx, "style": p.style.name, "text": p.text,
                "accepted_text": acc, "raw_text": raw,
                "fields": self._para_fields(p._element),
                "has_tracked_changes": acc != raw}

    # ── Write ─────────────────────────────────────────────────────────────────

    def replace(self, idx: int, new_text: str) -> str:
        """
        Replace text of paragraph at idx.
        Preserves the character formatting (w:rPr) of the first run.
        Returns old text.
        """
        para = self.doc.paragraphs[idx]
        old_text = para.text
        first_rpr = None
        for r in para.runs:
            rpr = r._r.find(qn("w:rPr"))
            if rpr is not None:
                first_rpr = copy.deepcopy(rpr)
                break
        for r in para.runs:
            r._r.getparent().remove(r._r)
        new_run = para.add_run(new_text)
        if first_rpr is not None:
            new_run._r.insert(0, first_rpr)
        return old_text

    def append(self, text: str, style: str = "Body Text") -> int:
        """Append a paragraph at the end of the document. Returns its index."""
        body = self.doc.element.body
        el = self._make_para(text, style)
        sectPr = body.find(qn("w:sectPr"))
        if sectPr is not None:
            sectPr.addprevious(el)   # sectPr must remain the body's last child
        else:
            body.append(el)
        return len(self.doc.paragraphs) - 1

    def insert_after(self, idx: int, text: str, style: str = "Body Text") -> int:
        """Insert paragraph after idx. Returns new paragraph's index."""
        ref = self.doc.paragraphs[idx]
        ref._element.addnext(self._make_para(text, style))
        return idx + 1

    def insert_before(self, idx: int, text: str, style: str = "Body Text") -> int:
        """Insert paragraph before idx. Returns new paragraph's index."""
        ref = self.doc.paragraphs[idx]
        ref._element.addprevious(self._make_para(text, style))
        return idx

    def inline_objects(self, idx: int) -> list:
        """
        List the paragraph's non-run inline objects (content controls w:sdt,
        hyperlinks, bookmarks). These can hold visible text that python-docx's
        paragraph .text does not expose — cross-references and reference-manager
        citations usually live here. Positions are the ids used by set_segments.
        """
        p = self.doc.paragraphs[idx]._element
        out = []
        pos = 0
        for child in p:
            if child.tag in (qn("w:pPr"), qn("w:r")):
                continue
            text = "".join(t.text or "" for t in child.iter(qn("w:t")))
            out.append({"pos": pos, "tag": child.tag.split("}")[-1], "text": text})
            pos += 1
        return out

    def set_segments(self, idx: int, segments: list) -> str:
        """
        Rewrite a paragraph as an ordered sequence of text and preserved inline
        objects. Each segment is either a string (literal text) or an integer
        referring to a `pos` from inline_objects(idx). Use this to reposition
        cross-references or citations without destroying them, which replace()
        cannot do. Text runs inherit the first plain run's character formatting.
        Returns the paragraph's new visible text.
        """
        p = self.doc.paragraphs[idx]._element
        pPr = p.find(qn("w:pPr"))
        kept = [copy.deepcopy(c) for c in p
                if c.tag not in (qn("w:pPr"), qn("w:r"))]
        base_rPr = None
        for r in p.findall(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            if rPr is not None:
                base_rPr = copy.deepcopy(rPr)
                break
        used = set()
        for seg in segments:
            if isinstance(seg, bool) or not isinstance(seg, (int, str)):
                raise ValueError(f"segment must be a string or an int, got {seg!r}")
            if isinstance(seg, int):
                if not 0 <= seg < len(kept):
                    raise ValueError(f"no inline object at position {seg}; "
                                     f"paragraph has {len(kept)}")
                if seg in used:
                    raise ValueError(f"inline object {seg} used more than once")
                used.add(seg)
        for child in list(p):
            if child is not pPr:
                p.remove(child)
        for seg in segments:
            if isinstance(seg, int):
                p.append(kept[seg])
            else:
                run = self._run(seg)
                if base_rPr is not None:
                    run.insert(0, copy.deepcopy(base_rPr))
                p.append(run)
        return self.doc.paragraphs[idx].text

    def set_runs(self, idx: int, parts: list) -> str:
        """
        Rewrite a paragraph as runs with explicit character formatting.
        parts is a list of dicts: {"text": str, "bold": bool, "italic": bool}.
        Each run inherits the paragraph's existing first-run formatting (font,
        size, language) with bold/italic overridden as requested — use this when
        one paragraph mixes formatting, which replace() cannot express.
        Returns the paragraph's new text.
        """
        p = self.doc.paragraphs[idx]._element
        pPr = p.find(qn("w:pPr"))
        base_rPr = None
        for r in p.findall(qn("w:r")):
            rPr = r.find(qn("w:rPr"))
            if rPr is not None:
                base_rPr = copy.deepcopy(rPr)
                break
        for child in list(p):
            if child is not pPr:
                p.remove(child)
        for part in parts:
            run = self._run(part.get("text", ""))
            rPr = copy.deepcopy(base_rPr) if base_rPr is not None else OxmlElement("w:rPr")
            for flag, tags in (("bold", ("w:b", "w:bCs")), ("italic", ("w:i", "w:iCs"))):
                want = part.get(flag)
                if want is None:
                    continue
                for tag in tags:
                    for existing in rPr.findall(qn(tag)):
                        rPr.remove(existing)
                    if want:
                        rPr.append(OxmlElement(tag))
            if len(rPr):
                run.insert(0, rPr)
            p.append(run)
        return self.doc.paragraphs[idx].text

    def list_numbering(self, idx: int) -> Optional[dict]:
        """Return {num_id, level} if the paragraph carries list numbering, else None."""
        pPr = self.doc.paragraphs[idx]._element.find(qn("w:pPr"))
        numPr = pPr.find(qn("w:numPr")) if pPr is not None else None
        if numPr is None:
            return None
        num_id = numPr.find(qn("w:numId"))
        ilvl = numPr.find(qn("w:ilvl"))
        return {"num_id": num_id.get(qn("w:val")) if num_id is not None else None,
                "level": ilvl.get(qn("w:val")) if ilvl is not None else "0"}

    def set_list_numbering(self, idx: int, num_id: Optional[str],
                           level: str = "0") -> Optional[dict]:
        """
        Attach or remove automatic list numbering on a paragraph. Pass num_id=None
        to remove it. Applying a heading style in Word strips numbering from a
        previously numbered paragraph, which silently breaks any "Section N"
        cross-references; use this to restore it. Copy num_id from a paragraph
        that still numbers correctly (see list_numbering).
        """
        p = self.doc.paragraphs[idx]._element
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            p.insert(0, pPr)
        for existing in pPr.findall(qn("w:numPr")):
            pPr.remove(existing)
        if num_id is None:
            return None
        numPr = OxmlElement("w:numPr")
        ilvl = OxmlElement("w:ilvl")
        ilvl.set(qn("w:val"), str(level))
        nid = OxmlElement("w:numId")
        nid.set(qn("w:val"), str(num_id))
        numPr.append(ilvl)
        numPr.append(nid)
        # numPr must precede most other pPr children; pStyle comes first if present.
        pStyle = pPr.find(qn("w:pStyle"))
        if pStyle is not None:
            pStyle.addnext(numPr)
        else:
            pPr.insert(0, numPr)
        return self.list_numbering(idx)

    def set_cell_text(self, table_idx: int, row: int, col: int, text: str) -> str:
        """
        Replace the text of one table cell, preserving the character formatting of
        its first formatted run. Returns the previous text.
        """
        cell = self.doc.tables[table_idx].rows[row].cells[col]
        old = cell.text
        first_rPr = None
        for para in cell.paragraphs:
            for r in para.runs:
                rPr = r._r.find(qn("w:rPr"))
                if rPr is not None:
                    first_rPr = copy.deepcopy(rPr)
                    break
            if first_rPr is not None:
                break
        cell.text = text
        if first_rPr is not None:
            for para in cell.paragraphs:
                for r in para.runs:
                    r._r.insert(0, copy.deepcopy(first_rPr))
        return old

    def insert_like(self, model_idx: int, after_idx: int, text: str) -> int:
        """
        Insert a paragraph after after_idx, cloning the paragraph properties of the
        paragraph at model_idx (style, list numbering, indentation) and the character
        formatting of its first formatted run. Use this to extend a numbered list or
        match surrounding body text, which insert_after cannot do.
        Returns the new paragraph's index.
        """
        model = self.doc.paragraphs[model_idx]._element
        new_para = OxmlElement("w:p")
        model_pPr = model.find(qn("w:pPr"))
        if model_pPr is not None:
            new_para.append(copy.deepcopy(model_pPr))
        model_rPr = None
        for r in model.iter(qn("w:r")):
            found = r.find(qn("w:rPr"))
            if found is not None:
                model_rPr = copy.deepcopy(found)
                break
        run = self._run(text)
        if model_rPr is not None:
            run.insert(0, model_rPr)
        new_para.append(run)
        self.doc.paragraphs[after_idx]._element.addnext(new_para)
        return after_idx + 1

    def delete(self, idx: int) -> str:
        """Delete paragraph at idx. Returns its old text."""
        para = self.doc.paragraphs[idx]
        old_text = para.text
        para._element.getparent().remove(para._element)
        return old_text

    def set_style(self, idx: int, style: str) -> None:
        """Change the paragraph style at idx."""
        self.doc.paragraphs[idx].style = self.doc.styles[style]

    def set_caption(self, idx: int, text: str, label: str = "Figure",
                    sep: str = ": ", style: Optional[str] = None) -> str:
        """
        Rewrite paragraph idx as a clean auto-numbered caption:
        '<label> {SEQ <label> \\* ARABIC}<sep><text>'.
        Removes any existing runs/fields/tracked-change cruft in the paragraph.
        Optionally set the paragraph style (e.g. 'Caption'). Returns new plain text.
        """
        if style:
            self.set_style(idx, style)
        p = self.doc.paragraphs[idx]._element
        pPr = p.find(qn("w:pPr"))
        for child in list(p):
            if child is not pPr:
                p.remove(child)
        p.append(self._run(f"{label} "))
        for el in self._seq_field(label):
            p.append(el)
        p.append(self._run(f"{sep}{text}"))
        return self.doc.paragraphs[idx].text

    def set_core_properties(self, **kwargs) -> dict:
        """
        Set document core properties. Accepts title, subject, author, keywords,
        comments, category, content_status, last_modified_by, version.
        Ignores None values. Returns the properties after the update.
        """
        allowed = {"title", "subject", "author", "keywords", "comments",
                   "category", "content_status", "last_modified_by", "version"}
        unknown = set(kwargs) - allowed
        if unknown:
            raise ValueError(f"Unknown core property: {', '.join(sorted(unknown))}. "
                             f"Allowed: {', '.join(sorted(allowed))}")
        cp = self.doc.core_properties
        for key, value in kwargs.items():
            if value is not None:
                setattr(cp, key, value)
        return self.core_properties()

    def insert_table(self, after_idx: int, rows: list, style: str = "Table Grid",
                     header: bool = True) -> int:
        """
        Insert a table immediately after the paragraph at after_idx.
        rows is a list of row lists; short rows are padded with empty cells.
        Bolds the first row when header is true. Returns the new table's index.
        """
        if not rows:
            raise ValueError("rows must contain at least one row")
        ncols = max(len(r) for r in rows)
        if ncols == 0:
            raise ValueError("rows must contain at least one column")

        table = self.doc.add_table(rows=len(rows), cols=ncols)
        try:
            table.style = self.doc.styles[style]
        except KeyError:
            table.style = self.doc.styles["Table Grid"]

        for i, row in enumerate(rows):
            for j in range(ncols):
                cell = table.cell(i, j)
                cell.text = "" if j >= len(row) else str(row[j])
                if header and i == 0:
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.bold = True

        # add_table appends to the body end — move it into position.
        self.doc.paragraphs[after_idx]._element.addnext(table._element)
        # doc.tables rebuilds its wrappers on each access, so match on the element.
        for i, t in enumerate(self.doc.tables):
            if t._element is table._element:
                return i
        return len(self.doc.tables) - 1

    def delete_table(self, table_idx: int) -> list:
        """Delete the table at table_idx. Returns its rows (list of lists) for confirmation."""
        table = self.doc.tables[table_idx]
        rows = [[cell.text for cell in row.cells] for row in table.rows]
        table._element.getparent().remove(table._element)
        return rows

    def delete_table_row(self, table_idx: int, row: int) -> list:
        """Delete one row from a table. Returns the removed row's cell texts."""
        table = self.doc.tables[table_idx]
        tr = table.rows[row]
        cells = [cell.text for cell in tr.cells]
        tr._tr.getparent().remove(tr._tr)
        return cells

    def insert_image(self, after_idx: int, image_path: str,
                     width_cm: Optional[float] = None, center: bool = True) -> int:
        """
        Insert an image in its own new paragraph immediately after the paragraph
        at after_idx. width_cm scales the image (aspect ratio kept); omit it to
        use the image's natural size, capped at the text-column width.
        Returns the new paragraph's index.
        """
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm, Emu

        img = Path(image_path)
        if not img.is_file():
            raise FileNotFoundError(f"image not found: {image_path}")

        # Resolve the anchor BEFORE adding the picture paragraph: add_paragraph
        # changes the paragraph list, which would shift index-based lookups.
        ref = self.doc.paragraphs[after_idx]._element
        para = self.doc.add_paragraph()
        run = para.add_run()
        if width_cm is not None:
            run.add_picture(str(img), width=Cm(width_cm))
        else:
            pic = run.add_picture(str(img))
            sect = self.doc.sections[-1]
            max_w = sect.page_width - sect.left_margin - sect.right_margin
            if pic.width > max_w:
                pic.height = Emu(int(pic.height * (max_w / pic.width)))
                pic.width = max_w
        if center:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ref.addnext(para._element)
        for i, p in enumerate(self.doc.paragraphs):
            if p._element is para._element:
                return i
        return after_idx + 1

    def _bookmark_names(self) -> set:
        return {bm.get(qn("w:name"))
                for bm in self.doc.element.body.iter(qn("w:bookmarkStart"))}

    def add_bookmark(self, idx: int, name: str) -> str:
        """
        Bookmark the paragraph at idx (wraps its whole content) so internal
        hyperlinks can target it. Bookmark names must be unique, start with a
        letter, and contain no spaces. Returns the bookmark name.
        """
        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]{0,39}", name):
            raise ValueError("bookmark name must start with a letter/underscore, "
                             "use only letters, digits or underscores, max 40 chars")
        if name in self._bookmark_names():
            raise ValueError(f"bookmark '{name}' already exists")
        ids = [int(bm.get(qn("w:id"), "0"))
               for bm in self.doc.element.body.iter(qn("w:bookmarkStart"))]
        bm_id = str(max(ids, default=0) + 1)
        p = self.doc.paragraphs[idx]._element
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), bm_id)
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), bm_id)
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            pPr.addnext(start)
        else:
            p.insert(0, start)
        p.append(end)
        return name

    def add_internal_link(self, idx: int, text: str, bookmark: str,
                          occurrence: int = 1) -> str:
        """
        Turn the nth occurrence of `text` inside paragraph idx into an internal
        hyperlink to `bookmark` (Ctrl+Click navigation in Word). The occurrence
        must lie within a single run. Keeps the run's character formatting.
        Returns the linked text.
        """
        if bookmark not in self._bookmark_names():
            raise ValueError(f"bookmark '{bookmark}' does not exist "
                             f"(create it with add_bookmark)")
        p = self.doc.paragraphs[idx]._element
        seen = 0
        for r in list(p.iter(qn("w:r"))):
            if r.getparent().tag != qn("w:p"):
                continue  # already inside a hyperlink or other wrapper
            ts = r.findall(qn("w:t"))
            if len(ts) != 1 or not (ts[0].text or ""):
                continue
            full = ts[0].text
            pos = 0
            while True:
                found = full.find(text, pos)
                if found == -1:
                    break
                seen += 1
                if seen == occurrence:
                    rPr = r.find(qn("w:rPr"))
                    parent = r.getparent()
                    at = list(parent).index(r)
                    new_elems = []
                    if found > 0:
                        pre = self._run(full[:found])
                        if rPr is not None:
                            pre.insert(0, copy.deepcopy(rPr))
                        new_elems.append(pre)
                    link = OxmlElement("w:hyperlink")
                    link.set(qn("w:anchor"), bookmark)
                    link.set(qn("w:history"), "1")
                    mid = self._run(text)
                    if rPr is not None:
                        mid.insert(0, copy.deepcopy(rPr))
                    link.append(mid)
                    new_elems.append(link)
                    tail = full[found + len(text):]
                    if tail:
                        post = self._run(tail)
                        if rPr is not None:
                            post.insert(0, copy.deepcopy(rPr))
                        new_elems.append(post)
                    for el in reversed(new_elems):
                        parent.insert(at, el)
                    parent.remove(r)
                    return text
                pos = found + 1
        raise ValueError(f"occurrence {occurrence} of {text!r} not found within a "
                         f"single run of paragraph {idx}")

    def accept_changes(self, author: Optional[str] = None) -> int:
        """Accept all tracked changes (optionally only from one author). Returns count."""
        return self._resolve(accept=True, author=author)

    def reject_changes(self, author: Optional[str] = None) -> int:
        """Reject all tracked changes (optionally only from one author). Returns count."""
        return self._resolve(accept=False, author=author)

    # ── Save ──────────────────────────────────────────────────────────────────

    def save(self, path: Optional[str] = None) -> Path:
        out = Path(path) if path else self.path
        self.doc.save(str(out))
        return out

    # ── Internal ──────────────────────────────────────────────────────────────

    def _style_id(self, style_name: str) -> str:
        """Resolve display name → Word internal styleId."""
        for s in self.doc.styles:
            if s.name == style_name:
                sid = s.element.get(qn("w:styleId"))
                return sid if sid else style_name
        return style_name

    def _make_para(self, text: str, style_name: str):
        """Build a raw w:p element with the given style and text."""
        new_para = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        pStyle = OxmlElement("w:pStyle")
        pStyle.set(qn("w:val"), self._style_id(style_name))
        pPr.append(pStyle)
        new_para.append(pPr)
        r = OxmlElement("w:r")
        t_el = OxmlElement("w:t")
        t_el.text = text
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(t_el)
        new_para.append(r)
        return new_para

    _XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

    def _run(self, text: str):
        """Build a w:r run wrapping preserved-space text."""
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = text
        t.set(self._XML_SPACE, "preserve")
        r.append(t)
        return r

    def _seq_field(self, label: str) -> list:
        """Build the run sequence for a { SEQ <label> \\* ARABIC } complex field."""
        def fldchar(kind):
            r = OxmlElement("w:r")
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            r.append(fc)
            return r
        instr_run = OxmlElement("w:r")
        it = OxmlElement("w:instrText")
        it.set(self._XML_SPACE, "preserve")
        it.text = f" SEQ {label} \\* ARABIC "
        instr_run.append(it)
        result_run = OxmlElement("w:r")
        rt = OxmlElement("w:t")
        rt.text = "1"
        result_run.append(rt)
        return [fldchar("begin"), instr_run, fldchar("separate"), result_run, fldchar("end")]

    @staticmethod
    def _unwrap(el):
        """Replace el with its children in place (keep content, drop the wrapper)."""
        parent = el.getparent()
        if parent is None:
            return
        pos = list(parent).index(el)
        for child in list(el):
            el.remove(child)
            parent.insert(pos, child)
            pos += 1
        parent.remove(el)

    def _restore_prchange(self, change_el):
        """Reject a w:pPrChange/w:rPrChange: restore the recorded previous properties."""
        parent = change_el.getparent()  # the w:pPr or w:rPr
        if parent is None:
            return
        old = None
        for c in change_el:
            if c.tag in (qn("w:pPr"), qn("w:rPr")):
                old = c
                break
        for c in list(parent):
            if c is not change_el:
                parent.remove(c)
        if old is not None:
            pos = list(parent).index(change_el)
            for c in list(old):
                old.remove(c)
                parent.insert(pos, c)
                pos += 1
        parent.remove(change_el)

    def _resolve(self, accept: bool, author: Optional[str] = None) -> int:
        """Accept or reject all tracked changes under the document body."""
        root = self.doc.element.body
        count = 0

        def match(el):
            return author is None or el.get(qn("w:author")) == author

        if accept:
            for tag in ("w:del", "w:moveFrom"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None and match(el):
                        el.getparent().remove(el)
                        count += 1
            for tag in ("w:ins", "w:moveTo"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None and match(el):
                        self._unwrap(el)
                        count += 1
            for tag in ("w:rPrChange", "w:pPrChange", "w:tblPrChange",
                        "w:trPrChange", "w:tcPrChange", "w:sectPrChange"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None:
                        el.getparent().remove(el)
                        count += 1
        else:
            for tag in ("w:ins", "w:moveTo"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None and match(el):
                        el.getparent().remove(el)
                        count += 1
            for tag in ("w:del", "w:moveFrom"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None and match(el):
                        for dt in el.iter(qn("w:delText")):
                            dt.tag = qn("w:t")
                        self._unwrap(el)
                        count += 1
            for tag in ("w:rPrChange", "w:pPrChange"):
                for el in list(root.iter(qn(tag))):
                    if el.getparent() is not None:
                        self._restore_prchange(el)
                        count += 1
        return count
