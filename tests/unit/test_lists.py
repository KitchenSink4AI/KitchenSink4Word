"""Lists patch (v1.0.1): real bullet/number glyphs, nesting, restart semantics."""

import shutil
from pathlib import Path

import pytest
from docx import Document

from word_mcp.core.errors import WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import lists as ls, read

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture
def tables_doc(tmp_path):
    dst = tmp_path / "report_tables.docx"
    shutil.copy(CORPUS / "report_tables.docx", dst)
    return dst


def test_bullet_list_real_numbering(tables_doc):
    pkg = DocxPackage(tables_doc)
    result = ls.add_list(
        pkg, ["First point", "Second point", "Third point"], kind="bullet",
        at_end=True,
    )
    pkg.save(do_backup=False)

    pkg2 = DocxPackage(tables_doc)
    assert pkg2.has_part("word/numbering.xml")
    # Content type + relationship declared.
    assert "numbering+xml" in pkg2.raw_part("[Content_Types].xml").decode()
    assert "numbering" in pkg2.raw_part("word/_rels/document.xml.rels").decode()
    # Paragraphs actually reference the numbering (the add_bullet bug: style
    # without numPr renders no glyph).
    lists_found = ls.get_lists(pkg2)
    mine = [x for x in lists_found if x["num_id"] == result["num_id"]]
    assert len(mine) == 1
    assert [i["text"] for i in mine[0]["items"]] == [
        "First point", "Second point", "Third point",
    ]
    # numbering.xml maps the numId to a bullet abstractNum.
    nroot = pkg2.root("word/numbering.xml")
    num = next(
        n for n in nroot.findall(qn("w:num"))
        if n.get(qn("w:numId")) == str(result["num_id"])
    )
    abs_id = num.find(qn("w:abstractNumId")).get(qn("w:val"))
    abstract = next(
        a for a in nroot.findall(qn("w:abstractNum"))
        if a.get(qn("w:abstractNumId")) == abs_id
    )
    fmt = abstract.find(qn("w:lvl")).find(qn("w:numFmt")).get(qn("w:val"))
    assert fmt == "bullet"
    Document(str(tables_doc))


def test_numbered_list_and_nesting(tables_doc):
    pkg = DocxPackage(tables_doc)
    ls.add_list(
        pkg,
        [
            {"text": "Top item", "level": 0},
            {"text": "Nested a", "level": 1},
            {"text": "Nested b", "level": 1},
            {"text": "Second top", "level": 0},
        ],
        kind="number",
        at_end=True,
    )
    pkg.save(do_backup=False)
    got = ls.get_lists(DocxPackage(tables_doc))[-1]["items"]
    assert [i["level"] for i in got] == [0, 1, 1, 0]
    Document(str(tables_doc))


def test_two_numbered_lists_get_distinct_num_ids(tables_doc):
    """Separate calls -> separate numbering instances -> each restarts at 1."""
    pkg = DocxPackage(tables_doc)
    r1 = ls.add_list(pkg, ["a", "b"], kind="number", at_end=True)
    r2 = ls.add_list(pkg, ["c", "d"], kind="number", at_end=True)
    pkg.save(do_backup=False)
    assert r1["num_id"] != r2["num_id"]
    assert len(ls.get_lists(DocxPackage(tables_doc))) == 2


def test_list_at_anchor(tables_doc):
    pkg = DocxPackage(tables_doc)
    paras = read.get_paragraphs(pkg)
    lp = max(paras, key=lambda p: len(p["text"]))
    anchor = " ".join(lp["text"].split()[:6])
    ls.add_list(pkg, ["anchored item"], kind="bullet", after_anchor=anchor)
    pkg.save(do_backup=False)
    got = ls.get_lists(DocxPackage(tables_doc))[-1]["items"]
    assert got[0]["paragraph_index"] == lp["index"] + 1


def test_existing_numbering_part_reused(tables_doc):
    """tables_doc may already have numbering.xml; ids must not collide."""
    pkg = DocxPackage(tables_doc)
    had_part = pkg.has_part("word/numbering.xml")
    r1 = ls.add_list(pkg, ["x"], kind="bullet", at_end=True)
    r2 = ls.add_list(pkg, ["y"], kind="number", at_end=True)
    pkg.save(do_backup=False)
    nroot = DocxPackage(tables_doc).root("word/numbering.xml")
    num_ids = [n.get(qn("w:numId")) for n in nroot.findall(qn("w:num"))]
    assert len(num_ids) == len(set(num_ids)), f"duplicate numIds (pre-existing part: {had_part})"
    abs_ids = [
        a.get(qn("w:abstractNumId")) for a in nroot.findall(qn("w:abstractNum"))
    ]
    assert len(abs_ids) == len(set(abs_ids))


def test_bad_inputs(tables_doc):
    pkg = DocxPackage(tables_doc)
    with pytest.raises(WordMcpError):
        ls.add_list(pkg, [], kind="bullet")
    with pytest.raises(WordMcpError):
        ls.add_list(pkg, ["x"], kind="dashes")
    with pytest.raises(WordMcpError):
        ls.add_list(pkg, [{"text": "x", "level": 12}])
