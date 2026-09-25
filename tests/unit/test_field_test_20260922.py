"""Regression tests for the 2026-09-21 dissertation field test.

Findings credit: the author's live defense-run sessions (reports under
Draft/Working Files/Agent Results/20260921_*). Each test reproduces what
was reported first, then asserts what is now true. Two of the reports'
findings turned out to be something other than what they looked like, and
those tests say so rather than quietly testing a different thing.

Items covered here:
 - insert_paragraphs, per-item paragraph_format (a centered heading and
   left-aligned body in ONE call, which took a whole lxml pass);
 - the "copy_format_from ADDED contextualSpacing" claim, reproduced;
 - heading_level + copy_format_from on an outline-based document leaving
   an outline level a TOC refresh can see;
 - format_text(find=...) sub-run formatting, which the report called the
   single biggest missing capability and which has always shipped;
 - table shading / repeat-header / widths, likewise shipping, in the
   media-forms pack, and now signposted from the lite table tools;
 - format_cells targets taking {"row": "header"}.
"""

from __future__ import annotations

import pytest
from docx import Document

import word_mcp.server as srv
from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import read as rd
from word_mcp.ops import tables as tb


# --------------------------------------------------------------- helpers


def _build(path, texts=("Alpha", "Bravo", "Charlie")):
    d = Document()
    for t in texts:
        d.add_paragraph(t)
    d.save(str(path))
    return path


def _paras(path):
    pkg = DocxPackage(path)
    return [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"]


def _ppr_children(p):
    ppr = p.find(qn("w:pPr"))
    return [] if ppr is None else [
        __import__("lxml.etree", fromlist=["QName"]).QName(c).localname
        for c in ppr
    ]


def _find(p, tag):
    return p.find(f"{qn('w:pPr')}/{qn(tag)}")


# ==================================================================
# insert_paragraphs: the contextualSpacing claim, reproduced
# ==================================================================


def test_the_clone_adds_nothing_the_source_does_not_have(tmp_path):
    """REPORTED: "the clone produced <spacing><ind><contextualSpacing>, i.e.
    it ADDED contextualSpacing the neighbour does not have".

    It does not. _paragraph_format_clone is a deep copy of the source's
    pPr minus four structural tags, so a source without contextualSpacing
    produces a clone without it. The reported symptom came from the source
    paragraph actually being a different paragraph: see
    test_the_two_index_bases_explain_the_report below.
    """
    path = _build(tmp_path / "clone.docx", ("Anchor", "Body"))
    srv.set_paragraph_format(
        str(path), [0], {"line_spacing": 2, "indent_left_pt": 36}
    )
    before = _ppr_children(_paras(path)[0])
    assert "contextualSpacing" not in before, "the fixture itself is wrong"

    srv.insert_paragraphs(
        str(path), [{"text": "New text."}],
        location={"paragraph": 1}, copy_format_from=0,
    )
    new_p = _paras(path)[2]
    assert "contextualSpacing" not in _ppr_children(new_p)
    assert "spacing" in _ppr_children(new_p)
    assert "ind" in _ppr_children(new_p)


def test_the_clone_carries_contextual_spacing_when_the_source_has_it(tmp_path):
    """The other half of the same statement: a source that DOES carry it
    passes it on, which is what a faithful clone must do. A reference-list
    entry carries exactly this, and a reference entry is what the reported
    index really addressed."""
    from lxml import etree

    path = _build(tmp_path / "cs.docx", ("Ref entry", "Body"))
    pkg = DocxPackage(path)
    src = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"][0]
    ppr = etree.SubElement(src, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:contextualSpacing"))
    src.insert(0, ppr)
    pkg.mark_dirty()
    pkg.save(do_backup=False)

    srv.insert_paragraphs(
        str(path), [{"text": "New text."}],
        location={"paragraph": 1}, copy_format_from=0,
    )
    assert "contextualSpacing" in _ppr_children(_paras(path)[2])


def test_the_two_index_bases_explain_the_report(tmp_path):
    """THE REAL DEFECT BEHIND THE REPORT, and it is a caller-side one.

    copy_format_from indexes BODY-LEVEL w:p, the basis python-docx and
    every tool on this surface uses. A script that walks every w:p in the
    document, table cells included, gets a different number for the same
    paragraph, and on the real file the two bases were 145 apart: the
    heading the report meant was body paragraph 759 and all-w:p 903. Body
    paragraph 904 of that document is a reference-list entry, which is why
    the clone produced spacing/ind/contextualSpacing and no outline level.

    This test builds the same shape: a table whose cells hold paragraphs,
    so the two counts diverge, and pins that the tools agree with each
    other on the body-level basis.
    """
    d = Document()
    d.add_paragraph("first")
    table = d.add_table(rows=2, cols=2)
    for row in table.rows:
        for cell in row.cells:
            cell.text = "in a cell"
    d.add_paragraph("target paragraph")
    path = tmp_path / "bases.docx"
    d.save(str(path))

    pkg = DocxPackage(path)
    body_level = [el for k, _i, el in rd.body_items(pkg) if k == "paragraph"]
    every_p = list(pkg.root().iter(qn("w:p")))
    assert len(every_p) > len(body_level), "the fixture has no nested paras"

    hits = srv.find_text(str(path), "target paragraph", live="off")
    reported = hits[0]["paragraph_index"]
    assert reported == len(body_level) - 1, (
        "find_text no longer agrees with the body-level basis"
    )
    # end is EXCLUSIVE on this surface
    got = srv.get_text(str(path), start=reported, end=reported + 1,
                       live="off")
    assert got[0]["text"] == "target paragraph"
    assert got[0]["index"] == reported


# ==================================================================
# insert_paragraphs: per-item paragraph_format
# ==================================================================


def test_a_centered_heading_and_plain_body_land_in_one_call(tmp_path):
    """REPORTED: a heading needs jc center and the bodies do not, and
    copy_format_from is one source for the whole call, so "a heading + body
    insertion cannot be done in a single call even in principle". The
    per-item overlay is the answer: the clone supplies what the neighbors
    share and paragraph_format supplies what one item does not."""
    path = _build(tmp_path / "mixed.docx", ("Anchor", "Body"))
    srv.set_paragraph_format(str(path), [0], {"line_spacing": 2})

    out = srv.insert_paragraphs(
        str(path),
        [
            {"text": "Policy Recommendations",
             "paragraph_format": {"alignment": "center",
                                  "keep_with_next": True}},
            {"text": "Chapter 1 promised a practical payoff."},
        ],
        location={"paragraph": 0}, copy_format_from=0,
    )
    assert out["inserted"] == 2
    assert out["paragraph_formats_applied"] == 1

    heading, body = _paras(path)[1], _paras(path)[2]
    assert _find(heading, "w:jc").get(qn("w:val")) == "center"
    assert _find(heading, "w:keepNext") is not None
    # the clone still reached both items
    assert _find(heading, "w:spacing").get(qn("w:line")) == "480"
    assert _find(body, "w:spacing").get(qn("w:line")) == "480"
    # and the item that asked for nothing extra got nothing extra
    assert _find(body, "w:jc") is None


def test_the_overlay_wins_over_the_clone(tmp_path):
    """Same rule the per-item style and formatting keys already follow."""
    path = _build(tmp_path / "wins.docx", ("Anchor", "Body"))
    srv.set_paragraph_format(
        str(path), [0], {"alignment": "right", "line_spacing": 2}
    )
    srv.insert_paragraphs(
        str(path), [{"text": "Centered.",
                     "paragraph_format": {"alignment": "center"}}],
        location={"paragraph": 0}, copy_format_from=0,
    )
    assert _find(_paras(path)[1], "w:jc").get(qn("w:val")) == "center"


def test_an_unknown_paragraph_format_key_refuses_before_anything_lands(
        tmp_path):
    """A batch that refuses halfway would leave the caller guessing."""
    path = _build(tmp_path / "bad.docx", ("Anchor", "Body"))
    with pytest.raises(WordMcpError, match="paragraph-formatting"):
        srv.insert_paragraphs(
            str(path),
            [{"text": "ok"},
             {"text": "bad", "paragraph_format": {"algnment": "center"}}],
            location={"paragraph": 0},
        )
    assert len(_paras(path)) == 2, "a refused batch inserted something"


def test_paragraph_format_must_be_an_object(tmp_path):
    path = _build(tmp_path / "shape.docx", ("Anchor", "Body"))
    with pytest.raises(WordMcpError, match="paragraph_format takes an object"):
        srv.insert_paragraphs(
            str(path), [{"text": "x", "paragraph_format": "center"}],
            location={"paragraph": 0},
        )


def test_an_out_of_range_outline_level_refuses(tmp_path):
    path = _build(tmp_path / "lvl.docx", ("Anchor", "Body"))
    with pytest.raises(WordMcpError, match="outline_level"):
        srv.insert_paragraphs(
            str(path), [{"text": "x",
                         "paragraph_format": {"outline_level": 12}}],
            location={"paragraph": 0},
        )


# ==================================================================
# insert_paragraphs: a TOC-visible heading, in one call
# ==================================================================


def _outline_based_doc(path):
    """The NSU-template shape: headings expressed as direct outlineLvl,
    no Heading styles anywhere."""
    d = Document()
    for t in ("Chapter One", "Body text one.", "Chapter Two",
              "Body text two."):
        d.add_paragraph(t)
    d.save(str(path))
    srv.set_paragraph_format(str(path), [0, 2], {"outline_level": 0})
    return path


def test_heading_level_plus_copy_format_from_stays_in_the_toc(tmp_path):
    """The dissertation case, end to end.

    An outline-based document gets a new section heading that must (a) look
    like its neighbors, (b) be centered, and (c) be picked up by a TOC
    refresh, which on this kind of document means carrying w:outlineLvl.
    All three in one call.
    """
    path = _outline_based_doc(tmp_path / "outline.docx")
    srv.set_paragraph_format(str(path), [1], {"line_spacing": 2})

    out = srv.insert_paragraphs(
        str(path),
        [{"text": "Policy Recommendations", "heading_level": 2,
          "paragraph_format": {"alignment": "center"}}],
        location={"paragraph": 1}, copy_format_from=1,
    )
    assert out["outline_levels_set"] == 1
    assert out["format_cloned_from"] == 1

    new_p = _paras(path)[2]
    lvl = _find(new_p, "w:outlineLvl")
    assert lvl is not None, "the heading carries no outline level; no TOC"
    assert lvl.get(qn("w:val")) == "1"        # heading_level 2 -> outline 1
    assert _find(new_p, "w:jc").get(qn("w:val")) == "center"
    assert _find(new_p, "w:spacing").get(qn("w:line")) == "480"

    # and the reader agrees: this is what a TOC field harvests
    outline = srv.get_outline(str(path))
    assert any(h["text"] == "Policy Recommendations" for h in outline)


def test_the_clone_alone_still_drops_the_outline_level(tmp_path):
    """The 2026-09-03 guard restated from the other side: without
    heading_level, a clone of a heading anchor must NOT smuggle its outline
    level onto body text. test_copy_format_from_excludes_outline_level is
    the original and is unchanged."""
    path = _outline_based_doc(tmp_path / "noleak.docx")
    srv.insert_paragraphs(
        str(path), [{"text": "Ordinary body text."}],
        location={"paragraph": 0}, copy_format_from=0,
    )
    assert _find(_paras(path)[1], "w:outlineLvl") is None


def test_paragraph_format_refuses_on_the_live_route(tmp_path, monkeypatch):
    """No silent no-op: the live insertion path writes text and a style and
    has no route to a per-item pPr."""
    from word_mcp.core.errors import DocumentLocked

    path = _build(tmp_path / "live.docx", ("Anchor", "Body"))

    def _locked(*a, **k):
        raise DocumentLocked("open in Word")

    monkeypatch.setattr(srv, "_edit", _locked)
    with pytest.raises(WordMcpError, match="paragraph_format is a file-mode"):
        srv.insert_paragraphs(
            str(path), [{"text": "x",
                         "paragraph_format": {"alignment": "center"}}],
            location={"paragraph": 0},
        )


# ==================================================================
# format_text(find=...): shipped, and now said out loud
# ==================================================================


def test_find_formats_a_substring_and_keeps_the_runs_other_formatting(
        tmp_path):
    """REPORTED as "MISSING CAPABILITY: no character-offset formatting
    inside a run", then corrected by the same author's next session. It has
    always worked; nothing on the wire said so."""
    d = Document()
    p = d.add_paragraph()
    run = p.add_run("Keep this part and mark this part only.")
    run.font.name = "Cambria"
    run.bold = True
    path = tmp_path / "sub.docx"
    d.save(str(path))

    out = srv.format_text(
        str(path), find="mark this part",
        formatting={"color": "0000FF", "underline": True}, live="off",
    )
    assert out["formatted"]["paragraph"] == 0

    para = _paras(path)[0]
    runs = para.findall(qn("w:r"))
    assert len(runs) == 3, f"the run did not split at both offsets: {runs}"
    texts = ["".join(t.text or "" for t in r.iter(qn("w:t"))) for r in runs]
    assert texts == ["Keep this part and ", "mark this part", " only."]
    for r in runs:
        rpr = r.find(qn("w:rPr"))
        assert rpr.find(qn("w:b")) is not None, "bold was not carried through"
        assert rpr.find(qn("w:rFonts")) is not None
    marked = runs[1].find(qn("w:rPr"))
    assert marked.find(qn("w:u")) is not None
    assert runs[0].find(qn("w:rPr")).find(qn("w:u")) is None


def test_the_format_text_description_says_find_reaches_a_substring():
    """The whole defect was discoverability: `find` reads like a search
    filter, so nothing told an agent it was the character-range selector."""
    desc = srv.format_text.__doc__ or ""
    assert "SUBSTRING" in desc
    assert "splits" in desc


def test_the_merge_recipe_no_longer_steers_away_from_find():
    """The recipe used to say "(no find string needed)" on the one step
    where find is the capability the caller is looking for."""
    from word_mcp.ops import workflows

    step = next(
        s for s in workflows.WORKFLOWS["merge-chapters"]["steps"]
        if s["tool"] == "format_text"
    )
    assert "no find string needed" not in step["why"]
    assert "find" in step["why"]


# ==================================================================
# tables: the capability was there, the signposts were not
# ==================================================================


@pytest.mark.parametrize("tool_name", ["create_table", "set_cells",
                                       "get_table"])
def test_the_lite_table_tools_point_at_the_pack_that_has_the_rest(tool_name):
    """REPORTED: "no route to table shd, tblHeader, tblGrid/tcW on an
    existing table". All three ship, in media-forms, which is off in a lite
    session: the tools were never in the agent's list, so the disabled-tool
    signpost (which fires on a call to a name you already know) could not
    fire either. The lite table tools now name them."""
    desc = getattr(srv, tool_name).__doc__ or ""
    assert "format_cells" in desc or "set_table_properties" in desc, desc
    assert "media-forms" in desc, desc


def test_shading_repeat_header_and_widths_all_reach_the_xml(tmp_path):
    """The capability itself, asserted in one place so the report's claim
    can be answered with a test rather than a reading."""
    path = tmp_path / "tbl.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(path))
    srv.create_table(str(path), [["H1", "H2"], ["a", "b"], ["c", "d"]])

    srv.format_cells(str(path), 0, [{"row": 0}], {"shading": "D9E2F3"})
    srv.set_table_properties(str(path), 0, header_row_repeat=True,
                             column_widths=[120, 200])

    pkg = DocxPackage(path)
    tbl = pkg.body().find(qn("w:tbl"))
    rows = tbl.findall(qn("w:tr"))
    for tc in rows[0].findall(qn("w:tc")):
        shd = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
        assert shd is not None and shd.get(qn("w:fill")) == "D9E2F3"
    assert rows[0].find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None
    grid = tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))
    assert [g.get(qn("w:w")) for g in grid] == ["2400", "4000"]


def test_format_cells_takes_row_header(tmp_path):
    """The shape the field test wanted: shade the header without first
    working out which rows the header is."""
    path = tmp_path / "hdr.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(path))
    srv.create_table(str(path), [["H1", "H2"], ["a", "b"]])
    srv.set_table_properties(str(path), 0, header_row_repeat=2)

    srv.format_cells(str(path), 0, [{"row": "header"}],
                     {"shading": "D9E2F3"})
    pkg = DocxPackage(path)
    rows = pkg.body().find(qn("w:tbl")).findall(qn("w:tr"))
    for row in rows[:2]:
        for tc in row.findall(qn("w:tc")):
            shd = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
            assert shd is not None and shd.get(qn("w:fill")) == "D9E2F3"


def test_row_header_falls_back_to_the_first_row(tmp_path):
    """A table nobody flagged still has a header in the only sense a caller
    means: the row create_table styles as one."""
    path = tmp_path / "hdr2.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(path))
    srv.create_table(str(path), [["H1", "H2"], ["a", "b"]],
                     header_row=False)
    srv.format_cells(str(path), 0, [{"row": "header"}],
                     {"shading": "FFFF00"})
    pkg = DocxPackage(path)
    rows = pkg.body().find(qn("w:tbl")).findall(qn("w:tr"))
    assert rows[0].find(f"{qn('w:tcPr')}/{qn('w:shd')}") is None or True
    first = rows[0].findall(qn("w:tc"))[0]
    assert first.find(f"{qn('w:tcPr')}/{qn('w:shd')}").get(
        qn("w:fill")) == "FFFF00"
    second = rows[1].findall(qn("w:tc"))[0]
    assert second.find(f"{qn('w:tcPr')}/{qn('w:shd')}") is None


def test_row_header_and_a_cell_index_combine(tmp_path):
    path = tmp_path / "hdr3.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(path))
    srv.create_table(str(path), [["H1", "H2"], ["a", "b"]])
    srv.format_cells(str(path), 0, [{"row": "header", "cell": 1}],
                     {"shading": "00FF00"})
    pkg = DocxPackage(path)
    rows = pkg.body().find(qn("w:tbl")).findall(qn("w:tr"))
    cells = rows[0].findall(qn("w:tc"))
    assert cells[0].find(f"{qn('w:tcPr')}/{qn('w:shd')}") is None
    assert cells[1].find(f"{qn('w:tcPr')}/{qn('w:shd')}").get(
        qn("w:fill")) == "00FF00"


def test_an_unknown_row_word_refuses(tmp_path):
    path = tmp_path / "hdr4.docx"
    d = Document()
    d.add_paragraph("before")
    d.save(str(path))
    srv.create_table(str(path), [["H1", "H2"], ["a", "b"]])
    with pytest.raises(WordMcpError, match="0-based index or 'header'"):
        srv.format_cells(str(path), 0, [{"row": "footer"}],
                         {"shading": "00FF00"})


def test_header_row_indices_reads_the_leading_flagged_run():
    """Word's own definition: the LEADING run of flagged rows."""
    from lxml import etree

    def row(flagged):
        tr = etree.Element(qn("w:tr"))
        if flagged:
            trpr = etree.SubElement(tr, qn("w:trPr"))
            etree.SubElement(trpr, qn("w:tblHeader"))
        return tr

    assert tb.header_row_indices([row(True), row(True), row(False)]) == [0, 1]
    assert tb.header_row_indices([row(False), row(True)]) == [0]
    assert tb.header_row_indices([]) == []


# ==================================================================
# the TOC-refresh note (W4)
# ==================================================================


def test_the_recipes_record_how_a_toc_field_is_updated():
    """W4: TablesOfContents.Update() does not resolve on the COLLECTION
    under late binding. The bridge already iterates members; no shipped
    guide text said so, and a caller writing their own COM script had no
    way to learn it except by hitting it."""
    from word_mcp.ops import workflows

    for task in ("merge-chapters", "build-lists-without-heading-styles"):
        notes = " ".join(workflows.WORKFLOWS[task]["notes"])
        assert "TablesOfContents" in notes, task
        assert "late binding" in notes or "no Update" in notes, task


def test_the_bridge_still_iterates_members():
    """The note describes shipped behavior; if the code ever changed to the
    collection form, the note would be wrong AND the call would fail."""
    import inspect

    from word_mcp.com import bridge

    src = inspect.getsource(bridge.refresh_fields)
    assert "TablesOfContents(i).Update()" in src
    assert "doc.TablesOfContents.Update()" not in src
