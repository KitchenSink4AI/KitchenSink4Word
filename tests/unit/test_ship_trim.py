"""Pins for the ship trim wave: payload shape and the document-open path.

Each test states the property the trim must not lose. The token savings
are real but they are not what these guard; what they guard is that no
fact left the payload and no integrity check left the loader.
"""

from __future__ import annotations

import json
import pathlib
import struct
import zipfile

import pytest
from docx import Document

from word_mcp import server as srv
from word_mcp.core.errors import DocumentCorrupt
from word_mcp.core.package import DocxPackage
from word_mcp.ops import read as rd


# ------------------------------------------------------------- fixtures


@pytest.fixture
def plain_table(tmp_path):
    """20x6, no merged cells: the audit's measurement fixture."""
    p = tmp_path / "plain.docx"
    d = Document()
    t = d.add_table(rows=20, cols=6)
    for r in range(20):
        for c in range(6):
            t.cell(r, c).text = f"value {r}-{c}"
    d.save(str(p))
    return p


@pytest.fixture
def merged_table(tmp_path):
    """Same grid with one horizontal and one vertical merge."""
    p = tmp_path / "merged.docx"
    d = Document()
    t = d.add_table(rows=4, cols=3)
    for r in range(4):
        for c in range(3):
            t.cell(r, c).text = f"v{r}{c}"
    t.cell(0, 0).merge(t.cell(0, 1))  # gridSpan
    t.cell(2, 2).merge(t.cell(3, 2))  # vMerge
    d.save(str(p))
    return p


@pytest.fixture
def prose(tmp_path):
    p = tmp_path / "prose.docx"
    d = Document()
    for i in range(12):
        d.add_paragraph(f"Legitimacy and authority, paragraph {i}.")
    t = d.add_table(rows=1, cols=2)
    t.cell(0, 0).text = "Legitimacy in a cell"
    t.cell(0, 1).text = "nothing here"
    d.save(str(p))
    return p


# ------------------------------------------------- get_table cell envelope


def test_get_table_merge_free_rows_are_plain_strings(plain_table):
    """has_merges=false: the two per-cell defaults are stated once at the
    top of the payload instead of 120 times inside it."""
    g = srv.get_table(str(plain_table), 0)
    assert g["has_merges"] is False
    assert g["cells"][0] == [f"value 0-{c}" for c in range(6)]
    assert g["cells"][19][5] == "value 19-5"
    assert all(isinstance(cell, str) for row in g["cells"] for cell in row)


def test_get_table_merged_table_keeps_full_fidelity(merged_table):
    """The other direction: a table with merges loses nothing. Every cell
    still reports its own grid_span and vmerge."""
    g = srv.get_table(str(merged_table), 0)
    assert g["has_merges"] is True
    # python-docx joins the two cells' paragraphs when it merges them.
    assert g["cells"][0][0] == {
        "text": "v00\nv01", "grid_span": 2, "vmerge": None
    }
    vmerges = [c["vmerge"] for row in g["cells"] for c in row]
    assert "continue" in vmerges
    assert all(
        set(cell) == {"text", "grid_span", "vmerge"}
        for row in g["cells"] for cell in row
    )


def test_get_table_merge_free_payload_is_smaller(plain_table):
    """The saving itself, measured the way the project measures tokens
    (chars/4). The threshold is deliberately loose: the pin is the shape,
    not the exact number."""
    pkg = DocxPackage(plain_table)
    full = rd.get_table(pkg, 0)
    compact = rd.compact_table(full)
    before = len(json.dumps(full, indent=2)) / 4
    after = len(json.dumps(compact, indent=2)) / 4
    assert after < before / 2, f"{before:.0f} -> {after:.0f} tokens"


def test_compact_table_is_a_pure_read_shape(plain_table):
    """compact_table never mutates what it is handed; internal callers
    (stats, journalcount, reviewcycle) keep the rich dict."""
    pkg = DocxPackage(plain_table)
    full = rd.get_table(pkg, 0)
    rd.compact_table(full)
    assert full["cells"][0][0] == {
        "text": "value 0-0", "grid_span": 1, "vmerge": None
    }


def test_get_table_description_states_both_shapes():
    """A response shape the caller cannot predict is not a saving. The
    description names has_merges as the discriminator."""
    doc = srv.get_table.__doc__ or ""
    assert "has_merges=false" in doc and "has_merges=true" in doc


# ----------------------------------------------------- find_text query echo


def test_find_text_literal_hits_drop_the_query_echo(prose):
    """A literal query's every hit IS the query, so echoing it per hit
    told the caller a string it had just sent."""
    hits = srv.find_text(str(prose), "Legitimacy", live="off")
    assert len(hits) >= 13
    assert all("match" not in h for h in hits)
    assert all("context" in h for h in hits)
    assert any("Legitimacy" in h["context"] for h in hits)
    assert any("table_index" in h for h in hits)  # the cell hit survives


def test_find_text_regex_hits_keep_match(prose):
    """Regex hits genuinely differ from the query, so the field stays."""
    hits = srv.find_text(str(prose), r"paragraph \d+", regex=True, live="off")
    assert len(hits) >= 10
    assert all("match" in h for h in hits)
    assert {h["match"] for h in hits} != {r"paragraph \d+"}


def test_find_text_literal_and_regex_agree_on_locations(prose):
    """Dropping the echo changed nothing else: same hits, same places."""
    lit = srv.find_text(str(prose), "Legitimacy", live="off")
    rx = srv.find_text(str(prose), "Legitimacy", regex=True, live="off")
    strip = [{k: v for k, v in h.items() if k != "match"} for h in rx]
    assert lit == strip


# -------------------------------------------------------- open-path trims


def test_open_does_not_run_a_testzip_pre_pass(plain_table, monkeypatch):
    """testzip() decompresses and CRC-checks every entry, and the read
    loop then decompresses every entry again. One pass, same guarantee."""
    def _boom(self):
        raise AssertionError("testzip() ran: the double decompress is back")

    monkeypatch.setattr(zipfile.ZipFile, "testzip", _boom)
    pkg = DocxPackage(plain_table)
    assert "word/document.xml" in pkg.part_names()


def test_open_does_not_slurp_the_file_for_its_magic_bytes(
    plain_table, monkeypatch
):
    """read_bytes()[:8] pulled a whole document into memory to look at
    eight bytes of it."""
    target = plain_table.resolve()
    original = pathlib.Path.read_bytes

    def _guard(self):
        if self.resolve() == target:
            raise AssertionError("whole-file read on the magic-byte check")
        return original(self)

    monkeypatch.setattr(pathlib.Path, "read_bytes", _guard)
    DocxPackage(plain_table)


def test_open_still_refuses_a_bad_crc_and_names_the_part(plain_table):
    """The rigor pin, green before and after the trim: the CRC check that
    testzip() performed is performed by the read that replaced it, and the
    refusal still names the offending entry."""
    _corrupt_central_directory_crc(plain_table, "word/document.xml")
    with pytest.raises(DocumentCorrupt) as exc:
        DocxPackage(plain_table)
    assert "word/document.xml" in str(exc.value)


def test_open_still_refuses_a_non_zip(tmp_path):
    _p = tmp_path / "notazip.docx"
    _p.write_bytes(b"this is not a zip file at all, not even close")
    with pytest.raises(DocumentCorrupt):
        DocxPackage(_p)


def _corrupt_central_directory_crc(path, part: str) -> None:
    """Flip the stored CRC-32 of one entry. zipfile.read() reads the
    central-directory CRC, so this is what a real bit-rotted part looks
    like to the loader."""
    data = bytearray(path.read_bytes())
    name = part.encode()
    pos = 0
    while True:
        pos = data.find(b"PK\x01\x02", pos)
        assert pos != -1, f"{part} has no central-directory record"
        n = struct.unpack_from("<H", data, pos + 28)[0]
        if bytes(data[pos + 46: pos + 46 + n]) == name:
            break
        pos += 4
    data[pos + 16] ^= 0xFF
    path.write_bytes(bytes(data))
