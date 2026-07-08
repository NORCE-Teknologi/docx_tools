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
        self.doc.element.body.append(self._make_para(text, style))
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

    def delete(self, idx: int) -> str:
        """Delete paragraph at idx. Returns its old text."""
        para = self.doc.paragraphs[idx]
        old_text = para.text
        para._element.getparent().remove(para._element)
        return old_text

    def set_style(self, idx: int, style: str) -> None:
        """Change the paragraph style at idx."""
        self.doc.paragraphs[idx].style = self.doc.styles[style]

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
