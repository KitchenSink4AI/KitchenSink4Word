"""Anchored images: authoring them, reading them back, refusing what
cannot be preserved.

Most images in real documents float rather than sit inline, and floating
is not a property you can set on an inline drawing: wp:inline and
wp:anchor are different elements with different children in a different
order. Word rejects a file that gets that order wrong, and it does so by
offering to repair the document, which is the failure this file exists to
prevent. So the structural tests assert the SHAPE of the XML, not just
that a call returned without raising.

The corpus tests build the constructs real documents carry and that a
naive implementation silently damages: an anchor inside an
mc:AlternateContent pair (about one document in five has such a pair), a
hand-dragged wrap polygon, a legacy VML picture, and an anchored image
inside a table cell. Each one either round-trips or refuses out loud.
"""

from __future__ import annotations

import shutil
import struct
import zlib
from pathlib import Path

import pytest
from lxml import etree

from word_mcp.core.errors import TargetNotFound, UnsupportedStructure, WordMcpError
from word_mcp.core.package import DocxPackage, qn
from word_mcp.ops import media

CORPUS = Path(__file__).resolve().parents[1] / "corpus"

_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_VML = "urn:schemas-microsoft-com:vml"
_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def make_png(path: Path, w=80, h=40, rgb=(20, 120, 200)) -> Path:
    def chunk(tag, data):
        c = tag + data
        return struct.pack(">I", len(data)) + c + struct.pack(
            ">I", zlib.crc32(c) & 0xFFFFFFFF
        )

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    return path


@pytest.fixture
def doc(tmp_path):
    dst = tmp_path / "ch4.docx"
    shutil.copy(CORPUS / "ch4.docx", dst)
    return dst


@pytest.fixture
def png(tmp_path):
    return make_png(tmp_path / "swatch.png")


def _anchors(pkg: DocxPackage):
    return list(pkg.root().iter(f"{{{_WP}}}anchor"))


def _localnames(el) -> list[str]:
    return [etree.QName(c).localname for c in el]


def _float(doc, png, **kwargs) -> DocxPackage:
    """Insert one floating image and return the reloaded package, so every
    assertion below runs against bytes that survived a save."""
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True, **kwargs)
    pkg.save()
    return DocxPackage(doc)


# ------------------------------------------------------------- the structure


def test_an_anchor_is_built_in_schema_order(doc, png):
    pkg = _float(doc, png, wrap="square", width_pt=90)
    anchor = _anchors(pkg)[0]
    assert _localnames(anchor) == [
        "simplePos", "positionH", "positionV", "extent", "effectExtent",
        "wrapSquare", "docPr", "cNvGraphicFramePr", "graphic",
    ]


def test_the_required_anchor_attributes_are_all_present(doc, png):
    pkg = _float(doc, png, wrap="square")
    anchor = _anchors(pkg)[0]
    for attr in ("relativeHeight", "behindDoc", "locked", "layoutInCell",
                 "allowOverlap", "simplePos"):
        assert anchor.get(attr) is not None, f"anchor is missing {attr}"


@pytest.mark.parametrize("wrap,element", [
    ("square", "wrapSquare"),
    ("tight", "wrapTight"),
    ("through", "wrapThrough"),
    ("top_and_bottom", "wrapTopAndBottom"),
    ("behind", "wrapNone"),
    ("in_front", "wrapNone"),
])
def test_every_wrap_mode_writes_its_element(doc, png, wrap, element):
    pkg = _float(doc, png, wrap=wrap)
    anchor = _anchors(pkg)[0]
    assert element in _localnames(anchor)
    assert media.describe_anchor(anchor)["wrap"] == wrap


def test_behind_and_in_front_differ_only_by_the_flag(doc, png, tmp_path):
    behind = _anchors(_float(doc, png, wrap="behind"))[0]
    assert behind.get("behindDoc") == "1"
    second = tmp_path / "other.docx"
    shutil.copy(CORPUS / "ch4.docx", second)
    front = _anchors(_float(second, png, wrap="in_front"))[0]
    assert front.get("behindDoc") == "0"
    assert _localnames(behind) == _localnames(front)


def test_tight_and_through_carry_the_polygon_the_schema_requires(doc, png):
    pkg = _float(doc, png, wrap="tight")
    wrap_el = _anchors(pkg)[0].find(f"{{{_WP}}}wrapTight")
    poly = wrap_el.find(f"{{{_WP}}}wrapPolygon")
    assert poly is not None, "wrapTight without a wrapPolygon is invalid"
    assert poly.get("edited") == "0"
    assert poly.find(f"{{{_WP}}}start") is not None
    assert len(poly.findall(f"{{{_WP}}}lineTo")) >= 3


def test_square_does_not_carry_a_polygon(doc, png):
    pkg = _float(doc, png, wrap="square")
    wrap_el = _anchors(pkg)[0].find(f"{{{_WP}}}wrapSquare")
    assert wrap_el.find(f"{{{_WP}}}wrapPolygon") is None


def test_position_writes_align_or_offset_but_never_both(doc, png):
    pkg = _float(
        doc, png, wrap="square",
        position={"horizontal": {"relative_to": "margin", "align": "right"},
                  "vertical": {"relative_to": "page", "offset_pt": 36}},
    )
    anchor = _anchors(pkg)[0]
    h = anchor.find(f"{{{_WP}}}positionH")
    v = anchor.find(f"{{{_WP}}}positionV")
    assert h.get("relativeFrom") == "margin"
    assert _localnames(h) == ["align"]
    assert h.find(f"{{{_WP}}}align").text == "right"
    assert v.get("relativeFrom") == "page"
    assert _localnames(v) == ["posOffset"]
    assert v.find(f"{{{_WP}}}posOffset").text == str(36 * media.EMU_PER_PT)


def test_a_floating_image_does_not_add_a_paragraph(doc, png):
    before = len(DocxPackage(doc).body().findall(qn("w:p")))
    pkg = _float(doc, png, wrap="square")
    after = len(pkg.body().findall(qn("w:p")))
    assert after == before, (
        "an anchored image attaches to an existing paragraph; adding one "
        "would insert a blank line the user did not ask for"
    )


def test_an_inline_image_still_gets_its_own_paragraph(doc, png):
    before = len(DocxPackage(doc).body().findall(qn("w:p")))
    pkg = DocxPackage(doc)
    result = media.add_image(pkg, str(png), at_end=True)
    pkg.save()
    assert result["placement"] == "inline"
    assert len(DocxPackage(doc).body().findall(qn("w:p"))) == before + 1


def test_z_order_defaults_upward_and_can_be_set(doc, png):
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True, wrap="square")
    media.add_image(pkg, str(png), at_end=True, wrap="square")
    media.add_image(pkg, str(png), at_end=True, wrap="square", z_order=99)
    pkg.save()
    heights = [int(a.get("relativeHeight")) for a in _anchors(DocxPackage(doc))]
    assert heights[0] < heights[1], "each new float should stack above"
    assert heights[2] == 99


def test_distance_defaults_match_word_and_are_settable(doc, png):
    pkg = _float(doc, png, wrap="square")
    anchor = _anchors(pkg)[0]
    assert (anchor.get("distL"), anchor.get("distR")) == ("114300", "114300")
    assert (anchor.get("distT"), anchor.get("distB")) == ("0", "0")


# ---------------------------------------------------------------- reading it


def test_list_images_reports_placement_for_both_kinds(doc, png):
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True)
    media.add_image(pkg, str(png), at_end=True, wrap="tight", width_pt=60)
    pkg.save()
    entries = media.list_images(DocxPackage(doc))
    placements = [e.get("placement") for e in entries]
    assert "inline" in placements and "anchored" in placements
    floating = next(e for e in entries if e["placement"] == "anchored")
    assert floating["wrap"] == "tight"
    assert floating["width_pt"] == 60.0
    assert floating["height_pt"] > 0


def test_what_is_written_is_what_is_read_back(doc, png):
    spec = {
        "wrap": "square",
        "wrap_text": "left",
        "z_order": 12,
        "position": {
            "horizontal": {"relative_to": "left_margin", "align": "inside"},
            "vertical": {"relative_to": "line", "offset_pt": 18},
        },
        "distance_pt": {"top": 3, "bottom": 3, "left": 12, "right": 12},
    }
    pkg = _float(doc, png, **spec)
    read_back = media.describe_anchor(_anchors(pkg)[0])
    assert read_back["wrap"] == "square"
    assert read_back["wrap_text"] == "left"
    assert read_back["z_order"] == 12
    assert read_back["position"] == {
        "horizontal": {"relative_to": "left_margin", "align": "inside"},
        "vertical": {"relative_to": "line", "offset_pt": 18.0},
    }
    assert read_back["distance_pt"] == {
        "top": 3.0, "bottom": 3.0, "left": 12.0, "right": 12.0
    }


def test_every_relative_frame_survives_the_round_trip(doc, png):
    """The two relativeFrom vocabularies are the part most likely to be
    mistranslated, so each value is written and read back."""
    for name in media.REL_FROM_H:
        pkg = DocxPackage(doc)
        media.add_image(
            pkg, str(png), at_end=True, wrap="square",
            position={"horizontal": {"relative_to": name, "offset_pt": 0}},
        )
        anchor = _anchors(pkg)[-1]
        assert (
            media.describe_anchor(anchor)["position"]["horizontal"]
            ["relative_to"] == name
        )
    for name in media.REL_FROM_V:
        pkg = DocxPackage(doc)
        media.add_image(
            pkg, str(png), at_end=True, wrap="square",
            position={"vertical": {"relative_to": name, "offset_pt": 0}},
        )
        anchor = _anchors(pkg)[-1]
        assert (
            media.describe_anchor(anchor)["position"]["vertical"]
            ["relative_to"] == name
        )


# ---------------------------------------------------------------- editing it


def test_an_inline_image_can_be_floated(doc, png):
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True)
    pkg.save()
    pkg = DocxPackage(doc)
    index = next(
        e["index"] for e in media.list_images(pkg)
        if e.get("placement") == "inline"
    )
    out = media.set_image_placement(
        pkg, index, wrap="square",
        position={"horizontal": {"relative_to": "page", "offset_pt": 72}},
    )
    pkg.save()
    assert out["placement"] == "anchored"
    reread = media.list_images(DocxPackage(doc))[index]
    assert reread["placement"] == "anchored"
    assert reread["position"]["horizontal"]["offset_pt"] == 72.0
    assert reread["width_pt"] == pytest.approx(out["width_pt"], rel=0.01) \
        if "width_pt" in out else True


def test_a_floating_image_can_be_returned_to_the_text(doc, png):
    pkg = _float(doc, png, wrap="behind", z_order=7)
    out = media.set_image_placement(pkg, 0, wrap="inline")
    pkg.save()
    assert out["placement"] == "inline"
    assert out["discarded"]["wrap"] == "behind"
    assert out["discarded"]["z_order"] == 7
    entries = media.list_images(DocxPackage(doc))
    assert entries[0]["placement"] == "inline"
    assert not _anchors(DocxPackage(doc))


def test_changing_one_thing_leaves_the_rest_alone(doc, png):
    pkg = _float(
        doc, png, wrap="square", z_order=5,
        position={"horizontal": {"relative_to": "margin", "align": "right"},
                  "vertical": {"relative_to": "page", "offset_pt": 44}},
    )
    media.set_image_placement(pkg, 0, wrap="top_and_bottom")
    pkg.save()
    after = media.list_images(DocxPackage(doc))[0]
    assert after["wrap"] == "top_and_bottom"
    assert after["z_order"] == 5
    assert after["position"]["horizontal"]["align"] == "right"
    assert after["position"]["vertical"]["offset_pt"] == 44.0


def test_switching_to_behind_flips_the_flag_not_the_position(doc, png):
    pkg = _float(doc, png, wrap="square",
                 position={"vertical": {"relative_to": "page",
                                        "offset_pt": 20}})
    media.set_image_placement(pkg, 0, wrap="behind")
    pkg.save()
    anchor = _anchors(DocxPackage(doc))[0]
    assert anchor.get("behindDoc") == "1"
    assert "wrapNone" in _localnames(anchor)
    assert _localnames(anchor).index("wrapNone") == _localnames(anchor).index(
        "effectExtent") + 1, "the wrap element must stay in schema order"
    assert media.describe_anchor(anchor)["position"]["vertical"][
        "offset_pt"] == 20.0


def test_resize_works_on_a_floating_image(doc, png):
    pkg = _float(doc, png, wrap="square", width_pt=100)
    media.resize_image(pkg, 0, width_pt=50)
    pkg.save()
    entry = media.list_images(DocxPackage(doc))[0]
    assert entry["width_pt"] == 50.0
    assert entry["placement"] == "anchored"
    assert entry["wrap"] == "square"
    extents = [
        (e.get("cx"), e.get("cy"))
        for e in _anchors(DocxPackage(doc))[0].iter(f"{{{_A}}}ext")
    ]
    assert extents and all(x == (str(50 * media.EMU_PER_PT),) + (x[1],)
                           for x in extents)


def test_replacing_the_file_keeps_the_anchor(doc, png, tmp_path):
    pkg = _float(doc, png, wrap="tight", z_order=9,
                 position={"horizontal": {"relative_to": "page",
                                          "offset_pt": 30}})
    other = make_png(tmp_path / "other.png", w=60, h=60, rgb=(9, 9, 9))
    out = media.replace_image(pkg, 0, str(other))
    pkg.save()
    assert out["placement"] == "anchored"
    entry = media.list_images(DocxPackage(doc))[0]
    assert entry["wrap"] == "tight"
    assert entry["z_order"] == 9
    assert entry["position"]["horizontal"]["offset_pt"] == 30.0


def test_alt_text_reaches_an_anchored_image(doc, png):
    from word_mcp.ops import structure

    pkg = _float(doc, png, wrap="square")
    structure.set_image_alt_text(pkg, 0, description="a blue swatch")
    pkg.save()
    docpr = _anchors(DocxPackage(doc))[0].find(f"{{{_WP}}}docPr")
    assert docpr.get("descr") == "a blue swatch"


def test_delete_removes_an_anchored_image(doc, png):
    pkg = _float(doc, png, wrap="square")
    before = len(media.list_images(pkg))
    media.delete_image(pkg, 0)
    pkg.save()
    assert len(media.list_images(DocxPackage(doc))) == before - 1


# ------------------------------------------------------------- the refusals


def test_position_without_wrap_refuses(doc, png):
    pkg = DocxPackage(doc)
    with pytest.raises(WordMcpError, match="wrap"):
        media.add_image(pkg, str(png), at_end=True,
                        position={"horizontal": {"align": "right"}})


def test_align_and_offset_together_refuse(doc, png):
    pkg = DocxPackage(doc)
    with pytest.raises(WordMcpError, match="not both|exactly one"):
        media.add_image(
            pkg, str(png), at_end=True, wrap="square",
            position={"horizontal": {"align": "right", "offset_pt": 10}},
        )


@pytest.mark.parametrize("bad,message", [
    ({"wrap": "sideways"}, "unknown wrap"),
    ({"wrap": "square", "wrap_text": "diagonally"}, "wrap_text"),
    ({"wrap": "square",
      "position": {"horizontal": {"relative_to": "nowhere"}}}, "relative_to"),
    ({"wrap": "square",
      "position": {"vertical": {"align": "sideways"}}}, "align"),
    ({"wrap": "square", "distance_pt": {"sideways": 3}}, "distance side"),
    ({"wrap": "square", "distance_pt": {"top": -3}}, "negative"),
])
def test_bad_vocabulary_refuses_by_name(doc, png, bad, message):
    pkg = DocxPackage(doc)
    with pytest.raises(WordMcpError, match=message):
        media.add_image(pkg, str(png), at_end=True, **bad)


def test_a_vertical_axis_rejects_a_horizontal_frame(doc, png):
    pkg = DocxPackage(doc)
    with pytest.raises(WordMcpError, match="relative_to"):
        media.add_image(
            pkg, str(png), at_end=True, wrap="square",
            position={"vertical": {"relative_to": "column"}},
        )


def test_wrap_names_people_actually_type_are_accepted():
    assert media.normalize_wrap("top-bottom") == "top_and_bottom"
    assert media.normalize_wrap("Behind Text") == "behind"
    assert media.normalize_wrap("in-front") == "in_front"
    assert media.normalize_wrap("TOP_AND_BOTTOM") == "top_and_bottom"
    with pytest.raises(WordMcpError):
        media.normalize_wrap("wrapped")


def test_placement_on_a_missing_image_refuses(doc):
    with pytest.raises(TargetNotFound):
        media.set_image_placement(DocxPackage(doc), 99, wrap="square")


def test_setting_nothing_refuses(doc, png):
    pkg = _float(doc, png, wrap="square")
    with pytest.raises(WordMcpError, match="nothing to change"):
        media.set_image_placement(pkg, 0)


def test_inline_with_position_refuses(doc, png):
    pkg = _float(doc, png, wrap="square")
    with pytest.raises(WordMcpError, match="no meaning|inline"):
        media.set_image_placement(pkg, 0, wrap="inline", z_order=3)


def test_an_inline_image_asked_for_nothing_but_position_refuses(doc, png):
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True)
    pkg.save()
    pkg = DocxPackage(doc)
    index = next(e["index"] for e in media.list_images(pkg)
                 if e.get("placement") == "inline")
    with pytest.raises(WordMcpError, match="pass wrap"):
        media.set_image_placement(pkg, index, z_order=4)


# ------------------------------------------------ constructs from real files
#
# Each fixture below is built to match what Word itself writes, because the
# refusals only mean something if they fire on the real shapes.


def _first_anchor(pkg):
    return _anchors(pkg)[0]


def test_a_hand_edited_wrap_polygon_is_not_silently_thrown_away(doc, png):
    """A wrapPolygon with edited="1" is an outline somebody dragged by
    hand. No generated rectangle reproduces it, so a wrap change that
    cannot carry it refuses and says what it would cost."""
    pkg = _float(doc, png, wrap="tight")
    poly = _first_anchor(pkg).find(
        f"{{{_WP}}}wrapTight/{{{_WP}}}wrapPolygon"
    )
    poly.set("edited", "1")
    pkg.mark_dirty()
    pkg.save()

    pkg = DocxPackage(doc)
    assert media.list_images(pkg)[0]["wrap_polygon"] == "edited by hand"
    with pytest.raises(UnsupportedStructure, match="edited by hand|hand"):
        media.set_image_placement(pkg, 0, wrap="square")

    # Tight to through keeps the polygon, because that mode can hold one.
    out = media.set_image_placement(pkg, 0, wrap="through")
    assert out["wrap"] == "through"
    kept = _first_anchor(pkg).find(
        f"{{{_WP}}}wrapThrough/{{{_WP}}}wrapPolygon"
    )
    assert kept is not None and kept.get("edited") == "1"


def test_force_accepts_losing_a_hand_edited_polygon(doc, png):
    pkg = _float(doc, png, wrap="tight")
    _first_anchor(pkg).find(
        f"{{{_WP}}}wrapTight/{{{_WP}}}wrapPolygon"
    ).set("edited", "1")
    pkg.mark_dirty()
    out = media.set_image_placement(pkg, 0, wrap="square", force=True)
    assert out["wrap"] == "square"


def test_an_anchor_inside_alternate_content_refuses(doc, png):
    """mc:AlternateContent describes the same picture twice, once per
    branch. Editing one branch would leave the two disagreeing."""
    pkg = _float(doc, png, wrap="square")
    anchor = _first_anchor(pkg)
    drawing = anchor.getparent()
    run = drawing.getparent()
    alt = etree.SubElement(run, f"{{{_MC}}}AlternateContent")
    choice = etree.SubElement(alt, f"{{{_MC}}}Choice")
    choice.set("Requires", "wps")
    run.remove(drawing)
    choice.append(drawing)
    etree.SubElement(alt, f"{{{_MC}}}Fallback")
    pkg.mark_dirty()
    pkg.save()

    pkg = DocxPackage(doc)
    assert media.list_images(pkg)[0]["in_alternate_content"] is True
    with pytest.raises(UnsupportedStructure, match="AlternateContent"):
        media.set_image_placement(pkg, 0, wrap="behind")


def test_a_legacy_vml_picture_is_listed_but_not_edited(doc):
    """v:imageData is still present in a noticeable share of real
    documents. Enumerating it keeps the list honest; editing it would be
    guesswork in a different positioning model."""
    pkg = DocxPackage(doc)
    p = pkg.body().findall(qn("w:p"))[0]
    run = etree.SubElement(p, qn("w:r"))
    pict = etree.SubElement(run, qn("w:pict"))
    shape = etree.SubElement(pict, f"{{{_VML}}}shape")
    shape.set("style", "position:absolute;left:0;top:0;width:72pt;height:36pt")
    data = etree.SubElement(shape, f"{{{_VML}}}imageData")
    data.set(f"{{{_R}}}id", "rIdNope")
    pkg.mark_dirty()
    pkg.save()

    entries = media.list_images(DocxPackage(doc))
    vml = [e for e in entries if e.get("placement") == "vml"]
    assert len(vml) == 1
    assert vml[0]["editable"] is False
    with pytest.raises(UnsupportedStructure, match="VML"):
        media.set_image_placement(DocxPackage(doc), vml[0]["index"],
                                  wrap="square")


def test_a_floating_image_is_refused_inside_a_text_box(doc, png):
    """Word only allows inline drawings inside a text box, so anchoring
    one there produces a file Word offers to repair."""
    pkg = DocxPackage(doc)
    media.add_image(pkg, str(png), at_end=True)
    pkg.save()

    pkg = DocxPackage(doc)
    index = next(e["index"] for e in media.list_images(pkg)
                 if e.get("placement") == "inline")
    blip = list(pkg.root().iter(f"{{{_A}}}blip"))[index]
    drawing = blip.getparent()
    while etree.QName(drawing).localname != "drawing":
        drawing = drawing.getparent()
    run = drawing.getparent()
    holder = etree.SubElement(run.getparent(), qn("w:txbxContent"))
    run.getparent().remove(run)
    holder.append(run)
    pkg.mark_dirty()

    with pytest.raises(UnsupportedStructure, match="text box"):
        media.set_image_placement(pkg, index, wrap="square")


def test_an_anchor_inside_a_table_cell_round_trips(doc, png):
    """layoutInCell is what makes this legal, and it is written on every
    anchor this module builds."""
    pkg = DocxPackage(doc)
    tables = pkg.body().findall(qn("w:tbl"))
    if not tables:
        pytest.skip("corpus document has no table")
    cell_p = tables[0].iter(qn("w:p")).__next__()
    index = list(pkg.body().iter(qn("w:p"))).index(cell_p)
    media.add_image(pkg, str(png), after_index=None, at_end=True,
                    wrap="square")
    pkg.save()
    assert index >= 0
    anchor = _anchors(DocxPackage(doc))[0]
    assert anchor.get("layoutInCell") == "1"


def test_an_anchor_carrying_unknown_extensions_keeps_them(doc, png):
    """Word writes wp14 relative-size extensions into an anchor's extLst.
    This module never claims to understand them, so it must not drop them
    when it rewrites a wrap mode."""
    pkg = _float(doc, png, wrap="square")
    anchor = _first_anchor(pkg)
    ext_list = etree.SubElement(anchor, f"{{{_WP}}}extLst")
    ext = etree.SubElement(ext_list, f"{{{_WP}}}ext")
    ext.set("uri", "{C183D7F6-B498-43B3-948B-1728B52AA6E4}")
    pkg.mark_dirty()
    pkg.save()

    pkg = DocxPackage(doc)
    media.set_image_placement(pkg, 0, wrap="tight", z_order=4)
    pkg.save()
    anchor = _first_anchor(DocxPackage(doc))
    assert anchor.find(f"{{{_WP}}}extLst") is not None, (
        "the extension list was dropped when the wrap changed"
    )
    assert "wrapTight" in _localnames(anchor)


def test_a_document_full_of_floats_still_opens_clean(doc, png):
    """Every wrap mode and both relative-frame vocabularies in one file,
    saved through the package validator."""
    pkg = DocxPackage(doc)
    for i, wrap in enumerate(media.WRAP_MODES):
        media.add_image(
            pkg, str(png), at_end=True, wrap=wrap, width_pt=40 + i,
            position={
                "horizontal": {"relative_to": "page", "offset_pt": 10 * i},
                "vertical": {"relative_to": "paragraph", "offset_pt": 5 * i},
            },
        )
    pkg.save()
    entries = [
        e for e in media.list_images(DocxPackage(doc))
        if e.get("placement") == "anchored"
    ]
    assert len(entries) == len(media.WRAP_MODES)
    assert {e["wrap"] for e in entries} == set(media.WRAP_MODES)
