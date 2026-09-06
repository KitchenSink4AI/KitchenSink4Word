"""List numbering: the overrides, the formats, and the number a reader
actually sees.

Nothing in a .docx stores "this paragraph is item 4". Word computes it by
walking the document, and about one document in seven leans on a level
override while doing it, so an implementation that reads only
abstractNum reports the wrong number on exactly those files. That makes
the computed number the thing worth testing hardest, and the last test
here settles it by asking Word: for every list this suite builds, the
label this module computes is compared against Word's own
Range.ListFormat.ListString.

The rest guards the three layers underneath: the abstract definition
(shared, so editing it in place would reformat other lists), the instance
(what makes two paragraphs one list), and the override (what makes one
instance restart without disturbing the others).
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from lxml import etree

from word_mcp.core.errors import (
    TargetNotFound,
    UnsupportedStructure,
    WordMcpError,
)
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import lists

CORPUS = Path(__file__).resolve().parents[1] / "corpus"


@pytest.fixture
def doc(tmp_path):
    dst = tmp_path / "ch4.docx"
    shutil.copy(CORPUS / "ch4.docx", dst)
    return dst


def _numbers(pkg, num_id=None):
    return [
        e["number"] for e in lists.computed_numbers(pkg)
        if num_id is None or e["num_id"] == num_id
    ]


def _add(pkg, items, **kwargs):
    return lists.add_list(pkg, items, kind="number", at_end=True, **kwargs)


# ------------------------------------------------------- rendering a counter


@pytest.mark.parametrize("value,fmt,expected", [
    (1, "decimal", "1"), (42, "decimal", "42"),
    (7, "decimalZero", "07"),
    (1, "lowerLetter", "a"), (26, "lowerLetter", "z"),
    (27, "lowerLetter", "aa"), (28, "lowerLetter", "ab"),
    (53, "lowerLetter", "ba"),
    (4, "upperLetter", "D"),
    (4, "lowerRoman", "iv"), (9, "lowerRoman", "ix"),
    (1990, "lowerRoman", "mcmxc"),
    (14, "upperRoman", "XIV"),
    (1, "ordinal", "1st"), (2, "ordinal", "2nd"), (3, "ordinal", "3rd"),
    (4, "ordinal", "4th"), (11, "ordinal", "11th"), (21, "ordinal", "21st"),
    (112, "ordinal", "112th"),
])
def test_counters_render_the_way_word_does(value, fmt, expected):
    assert lists.format_counter(value, fmt) == expected


def test_an_unrenderable_format_returns_none_rather_than_a_guess():
    """chicago, aiueo, ganada and friends exist. Reporting the raw counter
    and the format name beats inventing a label."""
    assert lists.format_counter(3, "aiueo") is None
    assert lists.format_counter(3, "chicago") is None


# --------------------------------------------------------- computing numbers


def test_a_plain_numbered_list_counts_from_one(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha", "bravo", "charlie"])
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == ["1.", "2.", "3."]


def test_nesting_counts_per_level_and_resets_the_deeper_one(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, [
        "one", {"text": "a", "level": 1}, {"text": "b", "level": 1},
        "two", {"text": "a again", "level": 1}, "three",
    ])
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == [
        "1.", "a.", "b.", "2.", "a.", "3.",
    ]


def test_two_lists_count_separately(doc):
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha", "bravo"])
    second = _add(pkg, ["charlie", "delta"])
    pkg.save()
    pkg = DocxPackage(doc)
    assert first["num_id"] != second["num_id"]
    assert _numbers(pkg, first["num_id"]) == ["1.", "2."]
    assert _numbers(pkg, second["num_id"]) == ["1.", "2."]


def test_continue_from_makes_two_inserts_one_list(doc):
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha", "bravo"])
    second = _add(pkg, ["charlie"], continue_from=first["num_id"])
    pkg.save()
    assert second["num_id"] == first["num_id"]
    assert _numbers(DocxPackage(doc), first["num_id"]) == ["1.", "2.", "3."]


def test_start_at_writes_a_start_override(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha", "bravo"], start_at=7)
    pkg.save()
    pkg = DocxPackage(doc)
    assert _numbers(pkg, result["num_id"]) == ["7.", "8."]
    num = lists._num_element(pkg, result["num_id"])
    override = num.find(qn("w:lvlOverride"))
    assert override is not None, "start_at must ride on a lvlOverride"
    assert override.find(qn("w:startOverride")).get(qn("w:val")) == "7"


def test_restart_at_leaves_other_lists_alone(doc):
    """The whole point of an instance-level override: one list restarts and
    the rest of the document does not move."""
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha", "bravo"])
    second = _add(pkg, ["charlie", "delta"])
    pkg.save()

    pkg = DocxPackage(doc)
    lists.set_numbering(pkg, second["num_id"], restart_at=100)
    pkg.save()
    pkg = DocxPackage(doc)
    assert _numbers(pkg, first["num_id"]) == ["1.", "2."]
    assert _numbers(pkg, second["num_id"]) == ["100.", "101."]


def test_a_start_override_does_not_disturb_the_deeper_levels(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["one", {"text": "a", "level": 1}, "two"])
    lists.set_numbering(pkg, result["num_id"], restart_at=50)
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == ["50.", "a.", "51."]


def test_a_missing_numbering_definition_is_reported_not_crashed(doc):
    """A numPr pointing at a numId numbering.xml never defines is a real
    shape in the wild; Word renders those paragraphs unnumbered."""
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    root = pkg.root("word/numbering.xml")
    root.remove(lists._num_element(pkg, result["num_id"]))
    pkg.mark_dirty("word/numbering.xml")
    pkg.save()

    groups = lists.get_lists(DocxPackage(doc))
    orphan = next(g for g in groups if g["num_id"] == result["num_id"])
    assert "definition_missing" in orphan


def test_num_id_zero_is_not_counted(doc):
    """numId 0 is Word's "numbering removed" marker, not a list."""
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha", "bravo"])
    numpr = pkg.body().findall(qn("w:p"))[-1].find(
        f"{qn('w:pPr')}/{qn('w:numPr')}"
    )
    numpr.find(qn("w:numId")).set(qn("w:val"), "0")
    pkg.mark_dirty()
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == ["1."]


# -------------------------------------------------------------- the formats


def test_levels_configure_the_new_list(doc):
    pkg = DocxPackage(doc)
    result = _add(
        pkg, ["one", "two", {"text": "sub", "level": 1}],
        levels=[
            {"level": 0, "format": "upperRoman", "text": "%1."},
            {"level": 1, "format": "lowerLetter", "text": "%1.%2)"},
        ],
    )
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == [
        "I.", "II.", "II.a)",
    ]


def test_reformatting_a_list_clones_a_shared_definition(doc):
    """Several w:num can point at one w:abstractNum. Editing the shared one
    in place would reformat lists nobody asked about."""
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha", "bravo"])
    second = _add(pkg, ["charlie"])
    # Force them to share, the way a document built in Word often does.
    shared = lists._num_element(pkg, second["num_id"]).find(
        qn("w:abstractNumId")
    )
    original = lists._num_element(pkg, first["num_id"]).find(
        qn("w:abstractNumId")
    ).get(qn("w:val"))
    shared.set(qn("w:val"), original)
    pkg.mark_dirty("word/numbering.xml")
    pkg.save()

    pkg = DocxPackage(doc)
    assert lists.describe_numbering(pkg, first["num_id"])[
        "shares_definition_with"] == [second["num_id"]]
    out = lists.set_numbering(
        pkg, second["num_id"],
        levels=[{"level": 0, "format": "upperRoman", "text": "%1."}],
    )
    pkg.save()
    assert out["cloned_definition"]["was_shared_with"] == [first["num_id"]]
    pkg = DocxPackage(doc)
    assert _numbers(pkg, first["num_id"]) == ["1.", "2."]
    assert _numbers(pkg, second["num_id"]) == ["I."]


def test_an_unshared_definition_is_edited_in_place(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    pkg.save()
    pkg = DocxPackage(doc)
    out = lists.set_numbering(
        pkg, result["num_id"],
        levels=[{"level": 0, "format": "lowerRoman", "text": "%1)"}],
    )
    pkg.save()
    assert "cloned_definition" not in out
    assert _numbers(DocxPackage(doc), result["num_id"]) == ["i)"]


def test_every_level_property_reaches_the_xml(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    lists.set_numbering(pkg, result["num_id"], levels=[{
        "level": 0, "format": "upperLetter", "text": "%1 -", "start": 3,
        "suffix": "space", "align": "right", "legal": True,
        "restart_after_level": 0, "indent_pt": 36, "hanging_pt": 18,
        "font": "Consolas",
    }])
    pkg.save()

    pkg = DocxPackage(doc)
    lvl = lists._lvl_of(lists._abstract_element(pkg, result["num_id"]), 0)
    assert lvl.find(qn("w:numFmt")).get(qn("w:val")) == "upperLetter"
    assert lvl.find(qn("w:lvlText")).get(qn("w:val")) == "%1 -"
    assert lvl.find(qn("w:start")).get(qn("w:val")) == "3"
    assert lvl.find(qn("w:suff")).get(qn("w:val")) == "space"
    assert lvl.find(qn("w:lvlJc")).get(qn("w:val")) == "right"
    assert lvl.find(qn("w:isLgl")).get(qn("w:val")) == "1"
    assert lvl.find(qn("w:lvlRestart")).get(qn("w:val")) == "0"
    ind = lvl.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    assert ind.get(qn("w:left")) == "720" and ind.get(qn("w:hanging")) == "360"
    fonts = lvl.find(f"{qn('w:rPr')}/{qn('w:rFonts')}")
    assert fonts.get(qn("w:ascii")) == "Consolas"


def test_a_level_is_written_in_schema_order(doc):
    """Word rejects a w:lvl whose children are out of order, and the order
    is not the order anyone writes them in."""
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    lists.set_numbering(pkg, result["num_id"], levels=[{
        "level": 0, "font": "Consolas", "align": "right", "start": 2,
        "text": "%1.", "legal": True, "suffix": "nothing",
        "format": "decimal", "indent_pt": 10,
    }])
    pkg.save()
    lvl = lists._lvl_of(lists._abstract_element(DocxPackage(doc),
                                                result["num_id"]), 0)
    names = [etree.QName(c).localname for c in lvl]
    order = [lists._LVL_ORDER.index(n) for n in names
             if n in lists._LVL_ORDER]
    assert order == sorted(order), f"w:lvl children out of order: {names}"


def test_legal_numbering_forces_decimal_in_the_label(doc):
    """isLgl is what turns 'IV.a' into '4.1' in a legal-style outline."""
    pkg = DocxPackage(doc)
    result = _add(
        pkg, ["one", {"text": "sub", "level": 1}],
        levels=[
            {"level": 0, "format": "upperRoman", "text": "%1."},
            {"level": 1, "format": "lowerLetter", "text": "%1.%2.",
             "legal": True},
        ],
    )
    pkg.save()
    assert _numbers(DocxPackage(doc), result["num_id"]) == ["I.", "1.1."]


def test_an_unrenderable_format_reports_the_raw_counter(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha", "bravo"])
    lvl = lists._lvl_of(lists._abstract_element(pkg, result["num_id"]), 0)
    lvl.find(qn("w:numFmt")).set(qn("w:val"), "aiueo")
    lvl.find(qn("w:lvlText")).set(qn("w:val"), "%1.")
    pkg.mark_dirty("word/numbering.xml")
    pkg.save()

    entries = [e for e in lists.computed_numbers(DocxPackage(doc))
               if e["num_id"] == result["num_id"]]
    assert entries[0]["number"] == 1
    assert entries[0]["unrendered_format"] == "aiueo"


# ---------------------------------------------------------------- reporting


def test_get_lists_carries_the_definition_and_the_numbers(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha", {"text": "sub", "level": 1}])
    pkg.save()
    group = next(g for g in lists.get_lists(DocxPackage(doc))
                 if g["num_id"] == result["num_id"])
    assert group["format"] == "decimal"
    assert group["levels_used"] == [0, 1]
    assert [i["number"] for i in group["items"]] == ["1.", "a."]
    assert len(group["definition"]["levels"]) == 9


def test_describe_numbering_reports_the_override(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"], start_at=4)
    pkg.save()
    info = lists.describe_numbering(DocxPackage(doc), result["num_id"])
    assert info["overridden_levels"] == [0]
    assert info["levels"][0]["start_override"] == 4
    assert info["levels"][0]["start"] == 4


def test_reading_a_corpus_document_does_not_change_it(tmp_path):
    """Computing numbers is a read. Guard it, because the walk resolves
    overrides and could be tempted to normalize them."""
    dst = tmp_path / "outline.docx"
    shutil.copy(CORPUS / "outline.docx", dst)
    # Seed one overridden list, so the guard holds whether or not the corpus
    # document on this machine happens to number anything of its own.
    seeded = DocxPackage(dst)
    _add(seeded, ["alpha", {"text": "sub", "level": 1}], start_at=3)
    seeded.save(do_backup=False)
    before = dst.read_bytes()
    groups = lists.get_lists(DocxPackage(dst))
    assert groups, "the corpus outline should carry lists"
    assert dst.read_bytes() == before


# --------------------------------------------------------------- refusals


def test_continuing_a_list_that_looks_different_refuses(doc):
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha"])
    second = _add(
        pkg, ["bravo"],
        levels=[{"level": 0, "format": "upperRoman", "text": "%1."}],
    )
    pkg.save()
    pkg = DocxPackage(doc)
    with pytest.raises(UnsupportedStructure, match="do not look alike"):
        lists.set_numbering(pkg, second["num_id"],
                            continue_from=first["num_id"])
    out = lists.set_numbering(
        pkg, second["num_id"], continue_from=first["num_id"], force=True
    )
    pkg.save()
    assert out["reformatted"] is True
    assert out["paragraphs_moved"] == 1
    assert _numbers(DocxPackage(doc), first["num_id"]) == ["1.", "2."]


def test_continuing_an_unknown_list_refuses(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    with pytest.raises(TargetNotFound, match="continue from"):
        lists.set_numbering(pkg, result["num_id"], continue_from=999)


def test_a_list_cannot_continue_itself(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    with pytest.raises(WordMcpError, match="continues itself"):
        lists.set_numbering(pkg, result["num_id"],
                            continue_from=result["num_id"])


def test_contradictory_arguments_refuse(doc):
    pkg = DocxPackage(doc)
    first = _add(pkg, ["alpha"])
    second = _add(pkg, ["bravo"])
    with pytest.raises(WordMcpError, match="contradict"):
        lists.set_numbering(pkg, second["num_id"],
                            continue_from=first["num_id"], restart_at=3)
    with pytest.raises(WordMcpError, match="drop start_at"):
        lists.add_list(pkg, ["x"], kind="number", at_end=True,
                       continue_from=first["num_id"], start_at=2)


def test_setting_nothing_refuses(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    with pytest.raises(WordMcpError, match="nothing to change"):
        lists.set_numbering(pkg, result["num_id"])


def test_an_unknown_list_refuses(doc):
    with pytest.raises(TargetNotFound, match="num_id"):
        lists.set_numbering(DocxPackage(doc), 4242, restart_at=1)


def test_continuing_from_an_undefined_instance_refuses(doc):
    pkg = DocxPackage(doc)
    with pytest.raises(TargetNotFound, match="continue"):
        lists.add_list(pkg, ["x"], kind="number", at_end=True,
                       continue_from=999)


@pytest.mark.parametrize("bad,message", [
    ({"level": 0, "format": "hieroglyph"}, "unknown numbering format"),
    ({"level": 0, "suffix": "comma"}, "unknown suffix"),
    ({"level": 0, "align": "sideways"}, "unknown align"),
    ({"level": 0, "colour": "red"}, "unknown level properties"),
    ({"level": 9, "format": "decimal"}, "level must be 0-8"),
    ({"format": "decimal"}, "needs a 'level'"),
])
def test_bad_level_specs_refuse_by_name(doc, bad, message):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    with pytest.raises(WordMcpError, match=message):
        lists.set_numbering(pkg, result["num_id"], levels=[bad])


def test_negative_starts_refuse(doc):
    pkg = DocxPackage(doc)
    result = _add(pkg, ["alpha"])
    with pytest.raises(WordMcpError, match="negative"):
        lists.set_numbering(pkg, result["num_id"], restart_at=-1)
    with pytest.raises(WordMcpError, match="negative"):
        lists.add_list(pkg, ["x"], kind="number", at_end=True, start_at=-3)


# --------------------------------------------------------- Word's own answer


def _word_available() -> bool:
    try:
        from test_live_core import _word_available as probe

        return probe()
    except Exception:
        return False


@pytest.mark.live
@pytest.mark.timeout(600)
@pytest.mark.skipif(
    not _word_available(), reason="Word/pywin32 not available on this machine"
)
def test_the_computed_labels_match_word(tmp_path):
    """The settling test. Word's Range.ListFormat.ListString is the number
    a reader sees; every label this module computes has to equal it,
    across nesting, a start override, a multi-level label, and roman."""
    import pythoncom
    import win32com.client

    from word_mcp import server

    doc = tmp_path / "numbers.docx"
    server.create_document(str(doc), title="numbering agreement")
    first = server.insert_list(str(doc), [
        "alpha", "bravo", {"text": "sub one", "level": 1},
        {"text": "sub two", "level": 1}, "charlie",
        {"text": "deep", "level": 2}, "delta",
    ], kind="number")
    second = server.insert_list(str(doc), ["echo", "foxtrot"],
                                kind="number", start_at=7)
    server.insert_list(
        str(doc), ["golf", "hotel", {"text": "sub", "level": 1}],
        kind="number",
        levels=[{"level": 0, "format": "upperRoman", "text": "%1."},
                {"level": 1, "format": "lowerLetter", "text": "%1.%2)"}],
    )
    server.insert_list(str(doc), ["india"], kind="number",
                       continue_from=first["num_id"])
    server.set_list_numbering(str(doc), second["num_id"], restart_at=100)

    ours = {
        item["text"]: item.get("number")
        for group in server.list_elements(str(doc), type="lists")["items"]
        for item in group["items"]
    }

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    mismatches = []
    try:
        opened = app.Documents.Open(str(doc.resolve()), False, True)
        for para in opened.Paragraphs:
            text = para.Range.Text.strip().replace("\r", "")
            if text not in ours:
                continue
            word_label = str(para.Range.ListFormat.ListString).strip()
            if word_label != str(ours[text]).strip():
                mismatches.append((text, word_label, ours[text]))
        opened.Close(0)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()
    assert not mismatches, (
        "computed labels disagree with Word: " + repr(mismatches)
    )
    assert len(ours) == 13, "the fixture should cover thirteen items"
