"""Footnote/endnote CRUD on corpus documents (many_footnotes.docx has 171
footnotes; report_tables.docx has none, so part/style creation is exercised).
"""

import shutil
from pathlib import Path

import pytest
from docx import Document

from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import notes, read

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture
def footnotes_doc(tmp_path):
    dst = tmp_path / "many_footnotes.docx"
    shutil.copy(CORPUS / "many_footnotes.docx", dst)
    return dst


@pytest.fixture
def tables_doc(tmp_path):
    dst = tmp_path / "report_tables.docx"
    shutil.copy(CORPUS / "report_tables.docx", dst)
    return dst


def anchor_from(path):
    pkg = DocxPackage(path)
    paras = read.get_paragraphs(pkg)
    long_para = max(paras, key=lambda p: len(p["text"]))
    words = long_para["text"].split()
    return " ".join(words[3:6])


# --------------------------------------------------- add into doc WITH notes


def test_add_footnote_to_niu(footnotes_doc):
    pkg = DocxPackage(footnotes_doc)
    n_before = len(read.list_footnotes(pkg))
    anchor = anchor_from(footnotes_doc)
    result = notes.add_note(
        pkg, "footnote", anchor_text=anchor, note_text="Injected test footnote."
    )
    pkg.save(do_backup=False)

    pkg2 = DocxPackage(footnotes_doc)
    after = read.list_footnotes(pkg2)
    assert len(after) == n_before + 1
    added = [n for n in after if "Injected test footnote." in n["text"]]
    assert len(added) == 1
    assert added[0]["position"] == result["position"]
    v = notes.validate_notes(pkg2)
    assert v["footnotes"]["ok"], v


def test_new_footnote_id_is_max_plus_one(footnotes_doc):
    pkg = DocxPackage(footnotes_doc)
    existing = {
        int(n.get(qn("w:id")))
        for n in pkg.root("word/footnotes.xml").findall(qn("w:footnote"))
    }
    result = notes.add_note(
        pkg, "footnote", anchor_text=anchor_from(footnotes_doc), note_text="x"
    )
    assert result["id"] == max(existing) + 1


def test_edit_footnote_by_position(footnotes_doc):
    pkg = DocxPackage(footnotes_doc)
    target = read.list_footnotes(pkg)[4]
    notes.edit_note(
        pkg, "footnote", position=5, new_text="REWRITTEN NOTE CONTENT."
    )
    pkg.save(do_backup=False)
    after = read.list_footnotes(DocxPackage(footnotes_doc))
    assert after[4]["text"] == "REWRITTEN NOTE CONTENT."
    assert after[4]["id"] == target["id"]
    assert len(after) == 171


def test_delete_footnote_both_halves(footnotes_doc):
    pkg = DocxPackage(footnotes_doc)
    before = read.list_footnotes(pkg)
    victim = before[9]
    result = notes.delete_note(pkg, "footnote", position=10)
    pkg.save(do_backup=False)

    pkg2 = DocxPackage(footnotes_doc)
    after = read.list_footnotes(pkg2)
    assert len(after) == len(before) - 1
    assert all(n["id"] != victim["id"] for n in after)
    assert result["body_references_removed"] == 1
    # No orphaned references (the corruption case).
    v = notes.validate_notes(pkg2)
    assert v["footnotes"]["ok"], v
    # Positions renumbered continuously.
    positions = [n["position"] for n in after]
    assert positions == list(range(1, len(after) + 1))


def test_delete_separator_refused(footnotes_doc):
    pkg = DocxPackage(footnotes_doc)
    with pytest.raises((WordMcpError, TargetNotFound)):
        notes.delete_note(pkg, "footnote", note_id=-1)


# ---------------------------------------------- add into doc WITHOUT notes


def test_add_first_footnote_creates_infrastructure(tables_doc):
    pkg = DocxPackage(tables_doc)
    assert not pkg.has_part("word/footnotes.xml")
    anchor = anchor_from(tables_doc)
    notes.add_note(
        pkg, "footnote", anchor_text=anchor, note_text="First ever footnote."
    )
    pkg.save(do_backup=False)

    pkg2 = DocxPackage(tables_doc)
    assert pkg2.has_part("word/footnotes.xml")
    # Specials present.
    types = {
        n.get(qn("w:type"))
        for n in pkg2.root("word/footnotes.xml").findall(qn("w:footnote"))
    }
    assert "separator" in types and "continuationSeparator" in types
    # Content type declared.
    ct = pkg2.raw_part("[Content_Types].xml").decode("utf-8")
    assert "footnotes+xml" in ct
    # Relationship declared.
    rels = pkg2.raw_part("word/_rels/document.xml.rels").decode("utf-8")
    assert "relationships/footnotes" in rels
    # Styles injected with superscript.
    styles = pkg2.raw_part("word/styles.xml").decode("utf-8")
    assert "FootnoteReference" in styles and "superscript" in styles
    # Note readable.
    fns = read.list_footnotes(pkg2)
    assert len(fns) == 1
    assert fns[0]["text"] == "First ever footnote."
    assert notes.validate_notes(pkg2)["footnotes"]["ok"]


def test_add_first_endnote_creates_infrastructure(tables_doc):
    pkg = DocxPackage(tables_doc)
    notes.add_note(
        pkg, "endnote", anchor_text=anchor_from(tables_doc), note_text="First endnote."
    )
    pkg.save(do_backup=False)
    pkg2 = DocxPackage(tables_doc)
    assert pkg2.has_part("word/endnotes.xml")
    ens = read.list_endnotes(pkg2)
    assert len(ens) == 1
    assert ens[0]["text"] == "First endnote."
    assert notes.validate_notes(pkg2)["endnotes"]["ok"]


def test_multiple_footnotes_number_by_body_order(tables_doc):
    """Add three notes at different anchors; displayed positions must follow
    body order regardless of insertion order."""
    pkg = DocxPackage(tables_doc)
    paras = [p for p in read.get_paragraphs(pkg) if len(p["text"]) > 100]
    early, mid, late = paras[1], paras[len(paras) // 2], paras[-1]

    def words(p):
        w = p["text"].split()
        return " ".join(w[2:5])

    # Insert out of order: late first.
    notes.add_note(pkg, "footnote", anchor_text=words(late), note_text="LATE")
    notes.add_note(pkg, "footnote", anchor_text=words(early), note_text="EARLY")
    notes.add_note(pkg, "footnote", anchor_text=words(mid), note_text="MID")
    pkg.save(do_backup=False)

    fns = read.list_footnotes(DocxPackage(tables_doc))
    by_pos = {n["position"]: n["text"] for n in fns}
    assert by_pos[1] == "EARLY"
    assert by_pos[2] == "MID"
    assert by_pos[3] == "LATE"


def test_multiline_note_text(tables_doc):
    pkg = DocxPackage(tables_doc)
    notes.add_note(
        pkg,
        "footnote",
        anchor_text=anchor_from(tables_doc),
        note_text="Line one.\nLine two.",
    )
    pkg.save(do_backup=False)
    fns = read.list_footnotes(DocxPackage(tables_doc))
    assert "Line one." in fns[0]["text"] and "Line two." in fns[0]["text"]


def test_anchor_ambiguity_needs_occurrence(tables_doc):
    pkg = DocxPackage(tables_doc)
    result = notes.add_note(
        pkg, "footnote", anchor_text="the", note_text="x", occurrence=3
    )
    assert result["id"] >= 1


def test_niu_edit_survives_full_cycle(footnotes_doc):
    """Add + edit + delete on the 171-note document, then verify the untouched
    170 notes are exactly as before."""
    pkg = DocxPackage(footnotes_doc)
    before = {n["id"]: n["text"] for n in read.list_footnotes(pkg)}
    r = notes.add_note(
        pkg, "footnote", anchor_text=anchor_from(footnotes_doc), note_text="CYCLE"
    )
    pkg.save(do_backup=False)
    pkg2 = DocxPackage(footnotes_doc)
    notes.delete_note(pkg2, "footnote", note_id=r["id"])
    pkg2.save(do_backup=False)
    after = {n["id"]: n["text"] for n in read.list_footnotes(DocxPackage(footnotes_doc))}
    assert after == before
