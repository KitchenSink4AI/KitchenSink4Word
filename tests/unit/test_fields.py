"""Bookmarks, cross-references, captions, hyperlinks."""

import shutil
from pathlib import Path

import pytest
from docx import Document

from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import fields, read

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture
def tables_doc(tmp_path):
    dst = tmp_path / "report_tables.docx"
    shutil.copy(CORPUS / "report_tables.docx", dst)
    return dst


@pytest.fixture
def many_tables(tmp_path):
    dst = tmp_path / "many_tables.docx"
    shutil.copy(CORPUS / "many_tables.docx", dst)
    return dst


def anchor_from(path, word_slice=slice(3, 6)):
    pkg = DocxPackage(path)
    paras = read.get_paragraphs(pkg)
    long_para = max(paras, key=lambda p: len(p["text"]))
    return " ".join(long_para["text"].split()[word_slice])


def test_bookmark_and_cross_reference(tables_doc):
    pkg = DocxPackage(tables_doc)
    anchor = anchor_from(tables_doc)
    fields.add_bookmark(pkg, "TestTarget", anchor_text=anchor)
    paras = read.get_paragraphs(pkg)
    other = [p for p in paras if len(p["text"]) > 80 and anchor not in p["text"]]
    ref_anchor = " ".join(other[0]["text"].split()[:3])
    fields.add_cross_reference(
        pkg, after_anchor=ref_anchor, to_bookmark="TestTarget", kind="page"
    )
    pkg.save(do_backup=False)

    pkg2 = DocxPackage(tables_doc)
    assert any(b["name"] == "TestTarget" for b in fields.list_bookmarks(pkg2))
    xml = pkg2.raw_part("word/document.xml").decode("utf-8")
    assert "PAGEREF TestTarget" in xml
    Document(str(tables_doc))


def test_cross_reference_to_missing_bookmark_refused(tables_doc):
    pkg = DocxPackage(tables_doc)
    with pytest.raises(TargetNotFound):
        fields.add_cross_reference(
            pkg, after_anchor="anything", to_bookmark="NoSuchBookmark"
        )


def test_duplicate_bookmark_refused(tables_doc):
    pkg = DocxPackage(tables_doc)
    anchor = anchor_from(tables_doc)
    fields.add_bookmark(pkg, "Dup", anchor_text=anchor)
    with pytest.raises(WordMcpError):
        fields.add_bookmark(pkg, "Dup", anchor_text=anchor)


def test_invalid_bookmark_name_refused(tables_doc):
    pkg = DocxPackage(tables_doc)
    with pytest.raises(WordMcpError):
        fields.add_bookmark(pkg, "1bad name", anchor_text="x")


def test_table_caption(many_tables):
    pkg = DocxPackage(many_tables)
    fields.add_caption(
        pkg, table_index=0, label="Table", text="Coding categories overview"
    )
    pkg.save(do_backup=False)
    pkg2 = DocxPackage(many_tables)
    xml = pkg2.raw_part("word/document.xml").decode("utf-8")
    assert "SEQ Table" in xml
    assert "Coding categories overview" in xml
    Document(str(many_tables))


def test_hyperlink(tables_doc):
    pkg = DocxPackage(tables_doc)
    anchor = anchor_from(tables_doc)
    fields.add_hyperlink(
        pkg, anchor_text=anchor, url="https://example.org/paper"
    )
    pkg.save(do_backup=False)
    pkg2 = DocxPackage(tables_doc)
    rels = pkg2.raw_part("word/_rels/document.xml.rels").decode("utf-8")
    assert "https://example.org/paper" in rels
    assert "TargetMode=\"External\"" in rels
    # Anchor text still present exactly once (moved inside w:hyperlink).
    text = "\n".join(p["text"] for p in read.get_paragraphs(pkg2))
    assert anchor in text
    Document(str(tables_doc))


def test_hyperlink_inside_hyperlink_refused(tables_doc):
    pkg = DocxPackage(tables_doc)
    anchor = anchor_from(tables_doc)
    fields.add_hyperlink(pkg, anchor_text=anchor, url="https://a.example")
    with pytest.raises(WordMcpError):
        fields.add_hyperlink(pkg, anchor_text=anchor, url="https://b.example")
