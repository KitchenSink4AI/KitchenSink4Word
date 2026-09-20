"""Regression tests for the 2026-09-20 field-test fix batch.

Findings credit: the author's live dissertation-consolidation production run
(report under Developer Feedback/General Feedback/Beta Test Report -
2026-09-20 - Dissertation Consolidation.md; punchlist #860-#868).
Each test reproduces the reported defect first, then asserts the fix.
"""

from __future__ import annotations

from lxml import etree

from docx import Document

import word_mcp.server as srv
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import read as rd


# --------------------------------------------------------------- helpers


def _build(path, texts=("Alpha", "Bravo", "Charlie")):
    d = Document()
    for t in texts:
        d.add_paragraph(t)
    d.save(str(path))
    return path


def _style_el(pkg, style_id):
    for s in pkg.root("word/styles.xml").findall(qn("w:style")):
        if s.get(qn("w:styleId")) == style_id:
            return s
    return None


def _ordered_sub(parent, local, order):
    """SubElement that lands in schema position (styles need pPr before rPr)."""
    el = etree.Element(qn(f"w:{local}"))
    rank = order.index(local)
    for child in parent:
        name = etree.QName(child).localname
        if name in order and order.index(name) > rank:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


_STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
    "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
]


def _stamp_normal_style(path, *, size_pt=None, line=None, after=None):
    """Put explicit formatting on the Normal STYLE (not docDefaults) --
    exactly what python-docx-generated chapter files carry."""
    pkg = DocxPackage(path)
    s = _style_el(pkg, "Normal")
    if line is not None or after is not None:
        ppr = _ordered_sub(s, "pPr", _STYLE_ORDER)
        sp = etree.SubElement(ppr, qn("w:spacing"))
        if line is not None:
            sp.set(qn("w:line"), str(line))
            sp.set(qn("w:lineRule"), "auto")
        if after is not None:
            sp.set(qn("w:after"), str(after))
    if size_pt is not None:
        rpr = _ordered_sub(s, "rPr", _STYLE_ORDER)
        for tag in ("w:sz", "w:szCs"):
            etree.SubElement(rpr, qn(tag)).set(
                qn("w:val"), str(int(size_pt * 2))
            )
    pkg.mark_dirty("word/styles.xml")
    pkg.save(do_backup=False)


def _body_paras(path):
    pkg = DocxPackage(path)
    return [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"], pkg


def _spacing_of(p):
    sp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    if sp is None:
        return None
    return {
        "line": sp.get(qn("w:line")),
        "after": sp.get(qn("w:after")),
        "lineRule": sp.get(qn("w:lineRule")),
    }


def _run_sizes(p):
    out = []
    for r in p.iter(qn("w:r")):
        sz = r.find(f"{qn('w:rPr')}/{qn('w:sz')}")
        out.append(sz.get(qn("w:val")) if sz is not None else None)
    return out


# ==================================================================
# #860 CRITICAL: insert_document formatting='source' loses formatting
# defined on a SHARED STYLE NAME (one layer up from docDefaults)
# ==================================================================


def test_insert_document_bakes_shared_style_name_differences(tmp_path):
    """Source defines its body look on the Normal STYLE (python-docx's
    habit), target's Normal says something else. By-name reconciliation
    hands the target's Normal to the carried paragraphs, so the source
    values must be baked explicit or the insertion silently reformats."""
    source = _build(tmp_path / "ch5.docx", ("Ch5 body one.", "Ch5 body two."))
    _stamp_normal_style(source, size_pt=12, line=480, after=0)
    target = _build(tmp_path / "diss.docx", ("T1",))
    _stamp_normal_style(target, size_pt=10, line=240, after=240)

    out = srv.insert_document(str(target), str(source))
    dd = out["document_defaults"]
    assert dd["differ"] is True
    assert "pPr.spacing.line" in dd["differing_properties"]
    assert "rPr.sz.val" in dd["differing_properties"]
    assert dd["paragraphs_baked"] >= 2
    assert dd["runs_baked"] >= 2, "font size was never carried"
    names = [s["style_name"] for s in dd.get("styles_reconciled", [])]
    assert "Normal" in names

    paras, _pkg = _body_paras(target)
    # Target's own paragraph untouched.
    assert _spacing_of(paras[0]) is None
    assert _run_sizes(paras[0]) == [None]
    for p in paras[1:]:
        sp = _spacing_of(p)
        assert sp is not None, "inserted paragraph lost the source look"
        assert sp["line"] == "480"
        assert sp["after"] == "0"
        assert _run_sizes(p) == ["24"], "12pt source text resolved to the target"


def test_insert_document_shared_style_no_difference_bakes_nothing(tmp_path):
    """Same style definitions on both sides: nothing baked, no report."""
    source = _build(tmp_path / "src.docx", ("Body.",))
    _stamp_normal_style(source, size_pt=12, line=480, after=0)
    target = _build(tmp_path / "tgt.docx", ("T1",))
    _stamp_normal_style(target, size_pt=12, line=480, after=0)
    out = srv.insert_document(str(target), str(source))
    assert "document_defaults" not in out
    paras, _pkg = _body_paras(target)
    assert _spacing_of(paras[-1]) is None


def test_insert_document_inherited_chain_difference_is_baked(tmp_path):
    """The difference is two links up the basedOn chain: Body -> Normal.
    Only the shared name (Normal) differs, so the chain must be resolved."""
    source = _build(tmp_path / "src.docx", ("Styled body.",))
    _stamp_normal_style(source, line=480)
    target = _build(tmp_path / "tgt.docx", ("T1",))
    _stamp_normal_style(target, line=240)
    # A "Body Text" style in BOTH files, identical, based on Normal.
    for path in (source, target):
        pkg = DocxPackage(path)
        root = pkg.root("word/styles.xml")
        st = etree.SubElement(root, qn("w:style"))
        st.set(qn("w:type"), "paragraph")
        st.set(qn("w:styleId"), "BodyText")
        etree.SubElement(st, qn("w:name")).set(qn("w:val"), "Body Text")
        etree.SubElement(st, qn("w:basedOn")).set(qn("w:val"), "Normal")
        pkg.mark_dirty("word/styles.xml")
        pkg.save(do_backup=False)
    pkg = DocxPackage(source)
    p = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"][0]
    ppr = etree.Element(qn("w:pPr"))
    p.insert(0, ppr)
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "BodyText")
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    srv.insert_document(str(target), str(source))
    paras, _pkg = _body_paras(target)
    assert _spacing_of(paras[-1])["line"] == "480"


def test_insert_document_cloned_style_is_not_baked(tmp_path):
    """A source style with no name match is CLONED with its definition, so
    its values still resolve the same way: nothing to bake."""
    source = _build(tmp_path / "src.docx", ("Quote.",))
    pkg = DocxPackage(source)
    root = pkg.root("word/styles.xml")
    st = etree.SubElement(root, qn("w:style"))
    st.set(qn("w:type"), "paragraph")
    st.set(qn("w:styleId"), "PullQuote")
    etree.SubElement(st, qn("w:name")).set(qn("w:val"), "Pull Quote")
    sppr = etree.SubElement(st, qn("w:pPr"))
    etree.SubElement(sppr, qn("w:spacing")).set(qn("w:line"), "300")
    p = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"][0]
    ppr = etree.Element(qn("w:pPr"))
    p.insert(0, ppr)
    etree.SubElement(ppr, qn("w:pStyle")).set(qn("w:val"), "PullQuote")
    pkg.mark_dirty()
    pkg.mark_dirty("word/styles.xml")
    pkg.save(do_backup=False)

    target = _build(tmp_path / "tgt.docx", ("T1",))
    srv.insert_document(str(target), str(source))
    paras, _pkg = _body_paras(target)
    sp = _spacing_of(paras[-1])
    assert sp is None or sp["line"] != "300", "cloned style needs no baking"
