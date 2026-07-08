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
