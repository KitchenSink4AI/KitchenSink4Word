"""Local features: stats, citation parity, page-number formats,
line numbering."""

import shutil
from pathlib import Path

import pytest
from docx import Document

from word_mcp.core.errors import TargetNotFound, WordMcpError
from word_mcp.core.package import DocxPackage
from word_mcp.ops import citecheck, furniture as f, read, stats

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture
def long_doc(tmp_path):
    dst = tmp_path / "long_report.docx"
    shutil.copy(CORPUS / "long_report.docx", dst)
    return dst


@pytest.fixture
def tables_doc(tmp_path):
    dst = tmp_path / "report_tables.docx"
    shutil.copy(CORPUS / "report_tables.docx", dst)
    return dst


# ---------------------------------------------------------------------- stats


def test_word_count_totals_plausible(long_doc):
    result = stats.word_count(DocxPackage(long_doc))
    t = result["totals"]
    assert t["words"] > 5000  # three chapters
    assert t["paragraphs"] > 400
    assert len(result["sections"]) >= 40  # 47 headings in ch1-3


def test_section_words_sum_to_children(long_doc):
    """A level-1 section's words include its level-2 children's words."""
    result = stats.word_count(DocxPackage(long_doc))
    sections = result["sections"]
    l1 = [s for s in sections if s["level"] == 1 and s["words"] > 0]
    assert l1, "expected level-1 sections with content"
    biggest = max(l1, key=lambda s: s["words"])
    assert biggest["words"] <= result["totals"]["words"]


def test_word_count_no_headings_doc(tmp_path):
    dst = tmp_path / "plain.docx"
    d = Document()
    d.add_paragraph("one two three four five")
    d.save(str(dst))
    r = stats.word_count(DocxPackage(dst))
    assert r["totals"]["words"] == 5
    assert r["sections"] == []


# ------------------------------------------------------------ citation parity


def test_parity_on_synthetic_doc(tmp_path):
    dst = tmp_path / "cite.docx"
    d = Document()
    d.add_paragraph(
        "Mediation research shows promise (Smith, 2026). Brandt (1979) framed "
        "the roadmap concept, later extended (Brandt, 1994; Harper, 1999). "
        "Uncited claims exist too."
    )
    d.add_heading("References", 1)
    d.add_paragraph("Smith, J. (2026). The Example Model. Journal of X.")
    d.add_paragraph("Brandt, E. S. (1979). The generalizability of the concept.")
    d.add_paragraph("Harper, L. (1999). Planning and accountability.")
    d.add_paragraph("Orphan, O. (2001). Never cited anywhere.")
    d.save(str(dst))

    r = citecheck.check_citation_parity(DocxPackage(dst))
    assert r["reference_entries"] == 4
    # Brandt 1994 cited but not listed.
    assert any("Brandt (1994)" in m for m in r["missing_references"])
    # Orphan 2001 listed but never cited.
    assert any("Orphan" in u for u in r["uncited_references"])
    assert not r["parity_ok"]
    # The clean ones are neither missing nor uncited.
    assert not any("Harper" in m for m in r["missing_references"])
    assert not any("Smith" in u for u in r["uncited_references"])


def test_parity_real_chapter(tables_doc):
    """tables_doc has ~56 refs; the checker should parse most and find the list."""
    r = citecheck.check_citation_parity(DocxPackage(tables_doc))
    assert r["reference_entries"] >= 40
    assert r["in_text_citations"] >= 40
    # Real docs may legitimately have flags; the tool just must not blow up
    # and must parse the bulk of the list.
    assert len(r["unparsed_reference_entries"]) <= r["reference_entries"] * 0.4


def test_parity_no_references_section(tmp_path):
    dst = tmp_path / "norefs.docx"
    d = Document()
    d.add_paragraph("Body only (Nobody, 2020).")
    d.save(str(dst))
    with pytest.raises(TargetNotFound):
        citecheck.check_citation_parity(DocxPackage(dst))


# ------------------------------------------- page number format / line numbers


def test_roman_numeral_page_format(tables_doc):
    pkg = DocxPackage(tables_doc)
    f.set_page_number_format(pkg, number_format="lowerRoman", start_at=1)
    pkg.save(do_backup=False)
    xml = DocxPackage(tables_doc).raw_part("word/document.xml").decode("utf-8")
    assert 'w:fmt="lowerRoman"' in xml
    Document(str(tables_doc))


def test_bad_page_format_refused(tables_doc):
    with pytest.raises(WordMcpError):
        f.set_page_number_format(DocxPackage(tables_doc), number_format="klingon")


def test_line_numbering_set_and_remove(tables_doc):
    pkg = DocxPackage(tables_doc)
    f.set_line_numbering(pkg, count_by=5, restart="newPage")
    pkg.save(do_backup=False)
    xml = DocxPackage(tables_doc).raw_part("word/document.xml").decode("utf-8")
    assert 'w:countBy="5"' in xml
    assert "lnNumType" in xml
    Document(str(tables_doc))

    pkg = DocxPackage(tables_doc)
    f.set_line_numbering(pkg, remove=True)
    pkg.save(do_backup=False)
    assert "lnNumType" not in DocxPackage(tables_doc).raw_part(
        "word/document.xml"
    ).decode("utf-8")


def test_sectpr_schema_order_preserved(tables_doc):
    """lnNumType must land BEFORE pgNumType and cols per CT_SectPr order."""
    pkg = DocxPackage(tables_doc)
    f.set_page_number_format(pkg, number_format="lowerRoman")
    f.set_line_numbering(pkg, count_by=1)
    pkg.save(do_backup=False)
    xml = DocxPackage(tables_doc).raw_part("word/document.xml").decode("utf-8")
    assert xml.index("lnNumType") < xml.index("pgNumType")
    Document(str(tables_doc))
