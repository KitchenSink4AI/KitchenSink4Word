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


# ==================================================================
# #861: define_style on an existing style must not strip attributes
# and children the call does not address (w:default, rsid, eastAsia)
# ==================================================================


def test_define_style_preserves_default_rsid_and_eastasia(tmp_path):
    """Redefining Normal is the documented read-one-define-one round trip.
    It must not drop w:default='1' (Word's document-default marker), the
    rsid, or the eastAsia font slot the call never mentioned."""
    from word_mcp.ops import structure as sx

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    s = _style_el(pkg, "Normal")
    rpr = _ordered_sub(s, "rPr", _STYLE_ORDER)
    rf = etree.SubElement(rpr, qn("w:rFonts"))
    rf.set(qn("w:ascii"), "Batang")
    rf.set(qn("w:eastAsia"), "Batang")
    pkg.mark_dirty("word/styles.xml")
    pkg.save(do_backup=False)

    pkg = DocxPackage(path)
    sx.define_style(
        pkg, style_id="Normal", name="Normal", style_type="paragraph",
        based_on=None,
        character_formatting={"font": "Times New Roman", "size_pt": 12},
    )
    pkg.save(do_backup=False)

    s = _style_el(DocxPackage(path), "Normal")
    assert s.get(qn("w:default")) == "1", "lost the document-default marker"
    assert s.find(qn("w:rsid")) is not None, "lost the rsid"
    rf = s.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
    assert rf.get(qn("w:ascii")) == "Times New Roman"
    assert rf.get(qn("w:eastAsia")) == "Batang", "lost the eastAsia slot"
    sz = s.find(f"{qn('w:rPr')}/{qn('w:sz')}")
    assert sz.get(qn("w:val")) == "24"


def test_define_style_never_writes_self_referential_based_on(tmp_path):
    """based_on defaults to 'Normal'; redefining Normal itself must not
    produce <w:basedOn w:val='Normal'/> on Normal."""
    from word_mcp.ops import structure as sx

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    out = sx.define_style(
        pkg, style_id="Normal", name="Normal",
        character_formatting={"size_pt": 12},
    )
    pkg.save(do_backup=False)
    s = _style_el(DocxPackage(path), "Normal")
    assert s.find(qn("w:basedOn")) is None
    assert out.get("based_on_self_skipped") is True


def test_define_style_keeps_unaddressed_children_and_reports(tmp_path):
    """Redefining a style addressing only character formatting keeps the
    paragraph formatting, uiPriority and next-style it already had."""
    from word_mcp.ops import structure as sx

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    sx.define_style(
        pkg, style_id="Body", name="Body Copy", style_type="paragraph",
        paragraph_formatting={"line_spacing": 2.0, "space_after_pt": 0},
        character_formatting={"size_pt": 12},
    )
    s = _style_el(pkg, "Body")
    _ordered_sub(s, "uiPriority", _STYLE_ORDER).set(qn("w:val"), "9")
    pkg.save(do_backup=False)

    pkg = DocxPackage(path)
    out = sx.define_style(
        pkg, style_id="Body", name="Body Copy", style_type="paragraph",
        character_formatting={"bold": True},
    )
    pkg.save(do_backup=False)
    assert out["replaced"] is True
    s = _style_el(DocxPackage(path), "Body")
    assert s.find(qn("w:uiPriority")).get(qn("w:val")) == "9"
    sp = s.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    assert sp is not None and sp.get(qn("w:line")) == "480"
    assert s.find(f"{qn('w:rPr')}/{qn('w:b')}") is not None
    # the size the earlier call set is still there (not addressed now)
    assert s.find(f"{qn('w:rPr')}/{qn('w:sz')}").get(qn("w:val")) == "24"


def test_define_style_refuses_changing_an_existing_style_type(tmp_path):
    from word_mcp.ops import structure as sx
    from word_mcp.core.errors import WordMcpError
    import pytest

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    with pytest.raises(WordMcpError, match="type"):
        sx.define_style(
            pkg, style_id="Normal", name="Normal", style_type="character",
        )


# ==================================================================
# #862: com_refresh_fields must reach header and footer stories
# ==================================================================


def _word_available():
    import sys

    if sys.platform != "win32":
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    import winreg

    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application")
        return True
    except OSError:
        return False


def test_story_counts_never_sum_the_same_fields_twice():
    """Pure-Python half: two update passes over the same fields report the
    larger per-story count, not the sum."""
    from word_mcp.com import bridge

    merged = bridge._merge_counts(
        {"body": 106, "primary_header": 2}, {"body": 106, "primary_footer": 2}
    )
    assert merged == {"body": 106, "primary_header": 2, "primary_footer": 2}
    assert bridge._STORY_NAMES[7] == "primary_header"
    assert bridge._STORY_NAMES[9] == "primary_footer"


def _two_section_doc_with_stale_headers(path):
    """Two sections, each with a header field whose cached result is stale."""
    from word_mcp.ops import furniture as fu
    from word_mcp.ops import structure as sx

    d = Document()
    for i in range(40):
        d.add_paragraph(f"Section one paragraph {i} " + "lorem ipsum " * 8)
    d.add_section()
    for i in range(40):
        d.add_paragraph(f"Section two paragraph {i} " + "lorem ipsum " * 8)
    d.save(str(path))

    pkg = DocxPackage(path)
    sx.set_document_properties(pkg, title="Fresh Title")
    fu.set_header_footer(pkg, "header", "sec one", section=0)
    fu.set_header_footer(pkg, "header", "sec two", section=1)
    pkg.save(do_backup=False)

    pkg = DocxPackage(path)
    for i, part in enumerate(("word/header1.xml", "word/header2.xml")):
        root = pkg.root(part)
        p = root.find(qn("w:p"))
        for r in list(p.findall(qn("w:r"))):
            p.remove(r)
        begin = etree.SubElement(p, qn("w:r"))
        etree.SubElement(begin, qn("w:fldChar")).set(
            qn("w:fldCharType"), "begin"
        )
        instr_run = etree.SubElement(p, qn("w:r"))
        it = etree.SubElement(instr_run, qn("w:instrText"))
        it.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        it.text = " DOCPROPERTY Title \\* MERGEFORMAT "
        sep = etree.SubElement(p, qn("w:r"))
        etree.SubElement(sep, qn("w:fldChar")).set(
            qn("w:fldCharType"), "separate"
        )
        cached = etree.SubElement(p, qn("w:r"))
        etree.SubElement(cached, qn("w:t")).text = f"STALE{i + 1}"
        end = etree.SubElement(p, qn("w:r"))
        etree.SubElement(end, qn("w:fldChar")).set(qn("w:fldCharType"), "end")
        pkg.mark_dirty(part)
    pkg.save(do_backup=False)


def _header_texts(path, part):
    pkg = DocxPackage(path)
    return [t.text for t in pkg.root(part).iter(qn("w:t"))]


def test_com_refresh_fields_updates_every_section_header(tmp_path):
    """Iterating doc.StoryRanges reaches only the FIRST story of each type,
    so section two's header kept its stale cache. The NextStoryRange walk
    reaches every one, and the result reports per-story counts."""
    import pytest

    if not _word_available():
        pytest.skip("Word/pywin32 not available on this machine")
    from word_mcp.com import bridge

    path = tmp_path / "sections.docx"
    _two_section_doc_with_stale_headers(path)
    assert _header_texts(path, "word/header2.xml") == ["STALE2"]

    out = bridge.refresh_fields(str(path))
    assert out["fields_refreshed"] is True
    assert "fields_by_story" in out
    assert any(
        "header" in story for story in out["fields_by_story"]
    ), out["fields_by_story"]
    assert _header_texts(path, "word/header1.xml") == ["Fresh Title"]
    assert _header_texts(path, "word/header2.xml") == ["Fresh Title"], (
        "section two's header field was never updated"
    )


# ==================================================================
# #865: first_line_indent_pt must not leave a residual w:hanging
# ==================================================================


def _ind_of(p):
    ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    if ind is None:
        return None
    return {
        "left": ind.get(qn("w:left")),
        "firstLine": ind.get(qn("w:firstLine")),
        "hanging": ind.get(qn("w:hanging")),
    }


def test_first_line_indent_zero_clears_a_hanging_indent(tmp_path):
    """A heading that inherited left=720 hanging=720 from a reference-list
    style came out left=0 hanging=720 firstLine=0: a half-inch negative
    first line on a centred heading."""
    from word_mcp.ops import text as tx

    path = _build(tmp_path / "d.docx", ("References",))
    pkg = DocxPackage(path)
    tx.set_paragraph_format(
        pkg, [0], {"indent_left_pt": 36, "first_line_indent_pt": -36}
    )
    live = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"][0]
    assert _ind_of(live)["hanging"] == "720"
    tx.set_paragraph_format(
        pkg, [0], {"indent_left_pt": 0, "first_line_indent_pt": 0}
    )
    pkg.save(do_backup=False)

    ind = _ind_of(_body_paras(path)[0][0])
    assert ind["firstLine"] == "0"
    assert ind["hanging"] is None, "residual hanging indent left in place"


def test_negative_first_line_indent_clears_a_first_line_indent(tmp_path):
    """The mirror: switching to a hanging indent must drop firstLine."""
    from word_mcp.ops import text as tx

    path = _build(tmp_path / "d.docx", ("Entry",))
    pkg = DocxPackage(path)
    tx.set_paragraph_format(pkg, [0], {"first_line_indent_pt": 18})
    tx.set_paragraph_format(pkg, [0], {"first_line_indent_pt": -36})
    pkg.save(do_backup=False)
    ind = _ind_of(_body_paras(path)[0][0])
    assert ind["hanging"] == "720"
    assert ind["firstLine"] is None


def test_define_style_first_line_indent_clears_hanging(tmp_path):
    """Same writer, style edition: redefining a hanging-indent style with a
    positive first line must not keep the hanging attribute."""
    from word_mcp.ops import structure as sx

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    sx.define_style(
        pkg, style_id="RefEntry", name="Ref Entry",
        paragraph_formatting={"indent_left_pt": 36, "first_line_indent_pt": -36},
    )
    sx.define_style(
        pkg, style_id="RefEntry", name="Ref Entry",
        paragraph_formatting={"first_line_indent_pt": 0},
    )
    pkg.save(do_backup=False)
    s = _style_el(DocxPackage(path), "RefEntry")
    ind = s.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert ind.get(qn("w:firstLine")) == "0"
    assert ind.get(qn("w:hanging")) in (None, "0"), "hanging indent survived"
    assert ind.get(qn("w:left")) == "720", "the left indent was not addressed"


# ==================================================================
# #866: rPr children must be written in CT_RPr schema order
# ==================================================================


def _rpr_children(p, run_index=0):
    runs = [r for r in p.iter(qn("w:r"))]
    rpr = runs[run_index].find(qn("w:rPr"))
    return [etree.QName(c).localname for c in rpr]


def _in_schema_order(names):
    from word_mcp.ops.text import _RPR_ORDER

    ranks = [_RPR_ORDER.index(n) for n in names if n in _RPR_ORDER]
    return ranks == sorted(ranks)


def test_insert_paragraphs_writes_rpr_in_schema_order(tmp_path):
    """The blue-insertion run produced <w:u/><w:color/>, which is backwards:
    CT_RPr sequences color before u."""
    path = _build(tmp_path / "d.docx", ("Anchor",))
    srv.insert_paragraphs(
        str(path),
        [{
            "text": "Inserted blue underlined text",
            "formatting": {"underline": True, "color": "0000FF"},
        }],
        location={"paragraph": 0},
    )
    paras, _pkg = _body_paras(path)
    names = _rpr_children(paras[1])
    assert set(names) >= {"u", "color"}
    assert _in_schema_order(names), names


def test_format_text_keeps_rpr_in_schema_order(tmp_path):
    """Adding emphasis to an existing rPr appended b and i after u, which
    is also invalid: b and i come first."""
    path = _build(tmp_path / "d.docx", ("Anchor",))
    srv.insert_paragraphs(
        str(path),
        [{
            "text": "Inserted blue underlined text",
            "formatting": {"underline": True, "color": "0000FF"},
        }],
        location={"paragraph": 0},
    )
    srv.format_text(
        str(path),
        find="Inserted blue underlined text",
        formatting={"bold": True, "italic": True},
    )
    paras, _pkg = _body_paras(path)
    names = _rpr_children(paras[1])
    assert set(names) >= {"b", "i", "u", "color"}
    assert _in_schema_order(names), names


def test_format_text_repairs_an_out_of_order_rpr(tmp_path):
    """A run whose rPr an older version wrote out of order is sorted when
    the tool next touches it."""
    from word_mcp.ops import text as tx

    path = _build(tmp_path / "d.docx", ("Target text",))
    pkg = DocxPackage(path)
    p = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"][0]
    r = p.find(qn("w:r"))
    rpr = etree.Element(qn("w:rPr"))
    r.insert(0, rpr)
    etree.SubElement(rpr, qn("w:u")).set(qn("w:val"), "single")
    etree.SubElement(rpr, qn("w:color")).set(qn("w:val"), "0000FF")
    tx.format_text(pkg, find="Target text", formatting={"bold": True})
    pkg.save(do_backup=False)
    paras, _pkg = _body_paras(path)
    names = _rpr_children(paras[0])
    assert _in_schema_order(names), names


def test_define_style_rpr_is_in_schema_order(tmp_path):
    """The same writer feeds style definitions."""
    from word_mcp.ops import structure as sx

    path = _build(tmp_path / "d.docx", ("Body.",))
    pkg = DocxPackage(path)
    sx.define_style(
        pkg, style_id="Blue", name="Blue",
        character_formatting={
            "underline": True, "color": "0000FF", "bold": True,
            "size_pt": 12, "font": "Times New Roman",
        },
    )
    pkg.save(do_backup=False)
    s = _style_el(DocxPackage(path), "Blue")
    names = [etree.QName(c).localname for c in s.find(qn("w:rPr"))]
    assert _in_schema_order(names), names


# ==================================================================
# #867: citation_parity must not drown a real bibliography in
# organizational-author false positives
# ==================================================================


def _parity_doc(tmp_path, body, refs):
    from docx import Document as _D

    path = tmp_path / "parity.docx"
    d = _D()
    for para in body:
        d.add_paragraph(para)
    d.add_heading("References", 1)
    for r in refs:
        d.add_paragraph(r)
    d.save(str(path))
    return path


def test_organizational_authors_are_matched_not_reported_missing(tmp_path):
    """The reported false positives: 'Archives (2003)' for 'National
    Archives. (2003).', 'War (2026)' for 'U.S. Department of War. (2026)',
    'Assembly (1977a)', and a possessive read as a surname."""
    from word_mcp.ops import citecheck

    path = _parity_doc(
        tmp_path,
        [
            "The National Archives (2003) holds the cable traffic.",
            "See also the U.S. Department of War (2026, March) report, and "
            "the National Assembly of the Republic of Korea (1977a).",
            "Bandura's (1977) account of self-efficacy is the origin.",
            "The dictionary entry is unambiguous (조선말 대사전, 1992).",
        ],
        [
            "National Archives. (2003). Cable traffic of the period.",
            "U.S. Department of War. (2026, March). Posture statement.",
            "National Assembly of the Republic of Korea. (1977a, June). "
            "Plenary record.",
            "Bandura, A. (1977). Self-efficacy. Psychological Review.",
            "조선말 대사전. (1992). 평양: 사회과학출판사.",
        ],
    )
    r = citecheck.check_citation_parity(DocxPackage(path))
    assert r["missing_references"] == [], r["missing_references"]
    assert r["missing_references_unparsed"] == [], (
        r["missing_references_unparsed"]
    )
    assert r["uncited_references"] == [], r["uncited_references"]
    assert r["parity_ok"] is True


def test_reference_entries_counts_every_entry(tmp_path):
    """reference_entries reported 197 for a list of 217 because unparsed
    entries were left out of the count."""
    from word_mcp.ops import citecheck

    path = _parity_doc(
        tmp_path,
        ["A claim (Smith, 2020)."],
        [
            "Smith, J. (2020). A title. Journal.",
            "An entry with no year at all, which cannot be keyed.",
        ],
    )
    r = citecheck.check_citation_parity(DocxPackage(path))
    assert r["reference_entries"] == 2
    assert r["reference_entries_parsed"] == 1
    assert len(r["unparsed_reference_entries"]) == 1


def test_a_genuinely_missing_surname_still_reports(tmp_path):
    """The check must stay useful: a real gap is still flagged, and in the
    confident list."""
    from word_mcp.ops import citecheck

    path = _parity_doc(
        tmp_path,
        ["Framed by Bordin (1979), extended later (Bordin, 1994)."],
        ["Bordin, E. S. (1979). The generalizability of the concept."],
    )
    r = citecheck.check_citation_parity(DocxPackage(path))
    assert r["missing_references"] == ["Bordin (1994)"]
    assert r["parity_ok"] is False


def test_prose_lead_words_are_not_read_as_authors(tmp_path):
    """'As Muller (2019) shows' must key on Muller, not on 'As Muller'."""
    from word_mcp.ops import citecheck

    path = _parity_doc(
        tmp_path,
        ["As Muller (2019) shows, the claim holds."],
        ["Muller, K. (2019). Der Titel. Zeitschrift."],
    )
    r = citecheck.check_citation_parity(DocxPackage(path))
    assert r["parity_ok"] is True
    assert r["unique_cited_works"] == 1, "'As Muller' counted as its own work"


# ==================================================================
# #864: whole-paragraph character formatting by index / anchor
# ==================================================================


def _italic_state(p):
    """(run italics, paragraph-mark italic) for one paragraph."""
    ppr = p.find(qn("w:pPr"))
    runs = []
    for r in p.iter(qn("w:r")):
        if r.getparent() is ppr:
            continue
        runs.append(r.find(f"{qn('w:rPr')}/{qn('w:i')}") is not None)
    mark = ppr is not None and ppr.find(
        f"{qn('w:rPr')}/{qn('w:i')}"
    ) is not None
    return runs, mark


def _make_italic(pkg, indices):
    from word_mcp.ops import text as tx

    tx.format_paragraphs(pkg, indices, {"italic": True})


def test_format_text_range_deitalicises_whole_paragraphs(tmp_path):
    """De-italicising 19 headings whose text also occurs in running prose:
    a find string cannot target them, and the old range form took one
    paragraph at a time."""
    path = _build(
        tmp_path / "d.docx",
        (
            "The Combined Forces Command",
            "The Combined Forces Command was established in 1978.",
            "The Combined Forces Command",
        ),
    )
    pkg = DocxPackage(path)
    _make_italic(pkg, [0, 2])
    pkg.save(do_backup=False)
    paras, _pkg = _body_paras(path)
    assert _italic_state(paras[0]) == ([True], True)

    srv.format_text(
        str(path), range={"start": 0, "end": 0}, formatting={"italic": False}
    )
    srv.format_text(
        str(path), range={"start": 2, "end": 2}, formatting={"italic": False}
    )
    paras, _pkg = _body_paras(path)
    assert _italic_state(paras[0]) == ([False], False), "paragraph mark kept"
    assert _italic_state(paras[2]) == ([False], False)
    # the body sentence, which shares the heading's text, is untouched
    assert _italic_state(paras[1]) == ([False], False)


def test_format_text_multi_paragraph_range(tmp_path):
    """A range of paragraphs formats in ONE call."""
    path = _build(tmp_path / "d.docx", ("H1", "H2", "H3", "Body"))
    pkg = DocxPackage(path)
    _make_italic(pkg, [0, 1, 2])
    pkg.save(do_backup=False)

    out = srv.format_text(
        str(path), range={"start": 0, "end": 2}, formatting={"italic": False}
    )
    assert out["formatted_paragraphs"] == [0, 1, 2]
    paras, _pkg = _body_paras(path)
    for p in paras[:3]:
        assert _italic_state(p) == ([False], False)


def test_format_text_multi_paragraph_range_refuses_find(tmp_path):
    import pytest
    from word_mcp.core.errors import WordMcpError

    path = _build(tmp_path / "d.docx", ("Alpha", "Bravo"))
    with pytest.raises(WordMcpError, match="find"):
        srv.format_text(
            str(path), range={"start": 0, "end": 1}, find="Alpha",
            formatting={"bold": True},
        )


def test_apply_edits_format_op_covers_the_paragraph_mark(tmp_path):
    """The anchor-addressed route gets the same whole-paragraph semantics."""
    path = _build(tmp_path / "d.docx", ("Chapter Four", "Body text."))
    pkg = DocxPackage(path)
    _make_italic(pkg, [0])
    pkg.save(do_backup=False)

    view = srv.get_document_view(str(path))
    line = next(
        ln for ln in view["view"].splitlines() if "Chapter Four" in ln
    )
    anchor = line.split("]")[0].lstrip("[")
    srv.apply_edits(
        str(path),
        [{"op": "format", "anchor": anchor, "formatting": {"italic": False}}],
    )
    paras, _pkg = _body_paras(path)
    assert _italic_state(paras[0]) == ([False], False)
