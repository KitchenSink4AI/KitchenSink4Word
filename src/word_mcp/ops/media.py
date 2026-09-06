"""Images: insert, list, replace, resize, and place.

Two shapes of image live in a Word document and the difference is
structural, not cosmetic. An INLINE image (wp:inline) sits in the text
like a very large character. An ANCHORED image (wp:anchor) is attached
to a paragraph and positioned independently of it, with text wrapping
around, above and below, behind, or in front of it. Anchored images turn
out to be the more common of the two in real documents, so an inline-only
model cannot author most of what people actually make.

python-docx has no anchor support at all, so everything below the inline
section is built directly on the OOXML, following the repo's raw-lxml
pattern: construct the element tree in the order the schema requires, let
DocxPackage validate the payload before it replaces the file, and refuse
rather than guess whenever an existing construct cannot be preserved.

The ordering rules that matter, because Word rejects a file that breaks
them: CT_Anchor's children are simplePos, positionH, positionV, extent,
effectExtent, ONE wrap element, docPr, cNvGraphicFramePr, graphic, in
that sequence. positionH and positionV each carry exactly one of align or
posOffset. wrapTight and wrapThrough each REQUIRE a wrapPolygon; the
other wrap modes must not carry one.
"""

from __future__ import annotations

import struct
from pathlib import Path

from lxml import etree

from ..core.errors import TargetNotFound, UnsupportedStructure, WordMcpError
from ..core.package import DocxPackage, qn
from ..core.sandbox import check_path

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

_EXT_TYPES = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".gif": "image/gif",
    ".bmp": "image/bmp",
    ".tiff": "image/tiff",
    ".emf": "image/x-emf",
    ".wmf": "image/x-wmf",
}

EMU_PER_INCH = 914400
EMU_PER_PT = 12700


def _image_size_px(data: bytes, ext: str) -> tuple[int, int] | None:
    """Native pixel size for PNG/JPEG/GIF without external deps."""
    try:
        if ext == ".png" and data[:8] == b"\x89PNG\r\n\x1a\n":
            w, h = struct.unpack(">II", data[16:24])
            return w, h
        if ext == ".gif":
            w, h = struct.unpack("<HH", data[6:10])
            return w, h
        if ext in (".jpg", ".jpeg"):
            i = 2
            while i < len(data) - 9:
                if data[i] != 0xFF:
                    i += 1
                    continue
                marker = data[i + 1]
                if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                    h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                    return w, h
                seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
                i += 2 + seg_len
    except Exception:
        pass
    return None


def _next_docpr_id(pkg: DocxPackage) -> int:
    ids = [
        int(el.get("id", "0"))
        for el in pkg.root().iter(f"{{{_WP}}}docPr")
    ]
    return max(ids, default=0) + 1


def _embed_media(pkg: DocxPackage, data: bytes, ext: str) -> tuple[str, str]:
    """Write the image bytes into the package and wire them up.

    Adds the media part, the [Content_Types].xml default for the
    extension when it is not already declared, and the document
    relationship. Returns (relationship id, media part name).
    """
    n = 1
    while pkg.has_part(f"word/media/image{n}{ext}") or any(
        name.startswith(f"word/media/image{n}.") for name in pkg.part_names()
    ):
        n += 1
    media_part = f"word/media/image{n}{ext}"
    pkg.set_raw_part(media_part, data)

    ct_root = pkg.root("[Content_Types].xml")
    ext_name = ext.lstrip(".")
    if not any(
        d.get("Extension") == ext_name
        for d in ct_root.findall(f"{{{_CT_NS}}}Default")
    ):
        default = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
        default.set("Extension", ext_name)
        default.set("ContentType", _EXT_TYPES[ext])
        pkg.mark_dirty("[Content_Types].xml")

    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    existing = {r.get("Id") for r in rels_root}
    i = 1
    while f"rId{i}" in existing:
        i += 1
    rid = f"rId{i}"
    rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
    rel.set("Id", rid)
    rel.set(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )
    rel.set("Target", "media/" + media_part.rsplit("/", 1)[1])
    pkg.mark_dirty(rels_part)
    return rid, media_part


def _display_extent(
    data: bytes, ext: str, width_pt: float | None
) -> tuple[int, int]:
    """Display size in EMU: the native size capped at 6.5in unless a width
    is given, with the aspect ratio kept either way."""
    px = _image_size_px(data, ext)
    if px:
        native_w_emu = int(px[0] / 96 * EMU_PER_INCH)
        native_h_emu = int(px[1] / 96 * EMU_PER_INCH)
    else:
        native_w_emu = native_h_emu = int(3 * EMU_PER_INCH)
    if width_pt:
        cx = int(width_pt * EMU_PER_PT)
    else:
        cx = min(native_w_emu, int(6.5 * EMU_PER_INCH))
    cy = int(cx * native_h_emu / native_w_emu)
    return cx, cy


def _picture_graphic(
    parent: etree._Element, *, rid: str, name: str, pic_id: int,
    cx: int, cy: int,
) -> etree._Element:
    """The a:graphic subtree both wp:inline and wp:anchor carry: identical
    in each, which is why the two placements can share one builder."""
    graphic = etree.SubElement(parent, f"{{{_A}}}graphic")
    gdata = etree.SubElement(graphic, f"{{{_A}}}graphicData")
    gdata.set("uri", _PIC)
    pic = etree.SubElement(gdata, f"{{{_PIC}}}pic")
    nvpr = etree.SubElement(pic, f"{{{_PIC}}}nvPicPr")
    cnv = etree.SubElement(nvpr, f"{{{_PIC}}}cNvPr")
    cnv.set("id", str(pic_id))
    cnv.set("name", name)
    etree.SubElement(nvpr, f"{{{_PIC}}}cNvPicPr")
    blipfill = etree.SubElement(pic, f"{{{_PIC}}}blipFill")
    blip = etree.SubElement(blipfill, f"{{{_A}}}blip")
    blip.set(f"{{{_R_NS}}}embed", rid)
    stretch = etree.SubElement(blipfill, f"{{{_A}}}stretch")
    etree.SubElement(stretch, f"{{{_A}}}fillRect")
    sppr = etree.SubElement(pic, f"{{{_PIC}}}spPr")
    xfrm = etree.SubElement(sppr, f"{{{_A}}}xfrm")
    off = etree.SubElement(xfrm, f"{{{_A}}}off")
    off.set("x", "0")
    off.set("y", "0")
    ext_el = etree.SubElement(xfrm, f"{{{_A}}}ext")
    ext_el.set("cx", str(cx))
    ext_el.set("cy", str(cy))
    geom = etree.SubElement(sppr, f"{{{_A}}}prstGeom")
    geom.set("prst", "rect")
    etree.SubElement(geom, f"{{{_A}}}avLst")
    return graphic


def add_image(
    pkg: DocxPackage,
    image_path: str,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    width_pt: float | None = None,
    alignment: str = "center",
    wrap: str | None = None,
    position: dict | None = None,
    z_order: int | None = None,
    allow_overlap: bool = True,
    distance_pt: dict | None = None,
    wrap_text: str | None = None,
) -> dict:
    """Insert an image, inline by default and anchored when wrap is given.

    Inline: its own paragraph, aligned by alignment. Anchored: a run
    appended to the located paragraph, positioned by position and wrapped
    by wrap, so the paragraph's own text flows around it. Width defaults
    to the native size capped at 6.5in; height keeps aspect ratio.
    """
    check_path(image_path, "read image file")
    src = Path(image_path)
    if not src.exists():
        raise TargetNotFound(f"image file not found: {image_path}")
    ext = src.suffix.lower()
    if ext not in _EXT_TYPES:
        raise WordMcpError(
            f"unsupported image type {ext}; use {sorted(_EXT_TYPES)}"
        )
    if width_pt is not None and width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    if wrap is None and (position is not None or z_order is not None):
        raise WordMcpError(
            "position and z_order describe a floating image, so they need a "
            f"wrap mode too; pass wrap as one of {sorted(WRAP_MODES)}"
        )
    data = src.read_bytes()
    rid, media_part = _embed_media(pkg, data, ext)
    cx, cy = _display_extent(data, ext, width_pt)
    docpr_id = _next_docpr_id(pkg)

    if wrap is None:
        return _add_inline_image(
            pkg, src, rid=rid, media_part=media_part, cx=cx, cy=cy,
            docpr_id=docpr_id, alignment=alignment, at_end=at_end,
            after_index=after_index, after_anchor=after_anchor,
        )
    return _add_anchored_image(
        pkg, src, rid=rid, media_part=media_part, cx=cx, cy=cy,
        docpr_id=docpr_id, wrap=wrap, position=position, z_order=z_order,
        allow_overlap=allow_overlap, distance_pt=distance_pt,
        at_end=at_end, after_index=after_index, after_anchor=after_anchor,
        wrap_text=wrap_text,
    )


def _add_inline_image(
    pkg: DocxPackage, src: Path, *, rid: str, media_part: str, cx: int,
    cy: int, docpr_id: int, alignment: str, at_end: bool,
    after_index: int | None, after_anchor: str | None,
) -> dict:
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    jc = etree.SubElement(ppr, qn("w:jc"))
    jc.set(
        qn("w:val"),
        {"left": "left", "center": "center", "right": "right"}[alignment],
    )
    r = etree.SubElement(p, qn("w:r"))
    drawing = etree.SubElement(r, qn("w:drawing"))
    inline = etree.SubElement(drawing, f"{{{_WP}}}inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, "0")
    extent = etree.SubElement(inline, f"{{{_WP}}}extent")
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    docpr = etree.SubElement(inline, f"{{{_WP}}}docPr")
    docpr.set("id", str(docpr_id))
    docpr.set("name", f"Picture {docpr_id}")
    _picture_graphic(
        inline, rid=rid, name=src.name, pic_id=docpr_id, cx=cx, cy=cy
    )

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        if sectpr is not None:
            sectpr.addprevious(p)
        else:
            body.append(p)
    elif after_anchor is not None:
        _resolve_anchor(pkg, after_anchor).addnext(p)
    else:
        _body_paragraph(pkg, after_index).addnext(p)
    pkg.mark_dirty()
    return {
        "image_added": media_part,
        "placement": "inline",
        "width_pt": round(cx / EMU_PER_PT, 1),
        "height_pt": round(cy / EMU_PER_PT, 1),
    }


# ------------------------------------------------------- anchored placement
#
# The vocabulary is spelled the way a person describes the picture, and
# translated here into the schema's names. Two of the six wrap modes share
# one element (wrapNone) and differ only by the behindDoc flag, which is
# why "behind" and "in front" are separate names on the way in.

#: wrap name -> (wrap element localname, behindDoc value)
WRAP_MODES: dict[str, tuple[str, str]] = {
    "square": ("wrapSquare", "0"),
    "tight": ("wrapTight", "0"),
    "through": ("wrapThrough", "0"),
    "top_and_bottom": ("wrapTopAndBottom", "0"),
    "behind": ("wrapNone", "1"),
    "in_front": ("wrapNone", "0"),
}

#: Spellings people reach for, mapped onto the six names above.
_WRAP_ALIASES = {
    "top-bottom": "top_and_bottom", "top_bottom": "top_and_bottom",
    "topandbottom": "top_and_bottom", "top and bottom": "top_and_bottom",
    "behind_text": "behind", "behind text": "behind",
    "in-front": "in_front", "in front": "in_front",
    "in_front_of_text": "in_front", "in front of text": "in_front",
    "none": "in_front", "float": "square", "tight_wrap": "tight",
}

#: Which sides text may flow down, for the modes that wrap text.
WRAP_TEXT_SIDES = {
    "both_sides": "bothSides", "bothsides": "bothSides",
    "both": "bothSides", "left": "left", "right": "right",
    "largest": "largest",
}

#: ST_RelFromH: what a horizontal position is measured from.
REL_FROM_H = {
    "margin": "margin", "page": "page", "column": "column",
    "character": "character", "left_margin": "leftMargin",
    "right_margin": "rightMargin", "inside_margin": "insideMargin",
    "outside_margin": "outsideMargin",
}

#: ST_RelFromV: what a vertical position is measured from.
REL_FROM_V = {
    "margin": "margin", "page": "page", "paragraph": "paragraph",
    "line": "line", "top_margin": "topMargin",
    "bottom_margin": "bottomMargin", "inside_margin": "insideMargin",
    "outside_margin": "outsideMargin",
}

ALIGN_H = {"left", "center", "right", "inside", "outside"}
ALIGN_V = {"top", "center", "bottom", "inside", "outside"}

#: CT_Anchor's child sequence. Word refuses a file that reorders these,
#: so every insertion goes through _place_child rather than append().
_ANCHOR_ORDER = [
    "simplePos", "positionH", "positionV", "extent", "effectExtent",
    "wrapNone", "wrapSquare", "wrapTight", "wrapThrough",
    "wrapTopAndBottom", "docPr", "cNvGraphicFramePr", "graphic",
    # extLst is last in the schema and holds extensions this module does
    # not model (wp14 relative sizing, most commonly). Naming it here is
    # what keeps an inserted child from landing AFTER it.
    "extLst",
]

_WRAP_ELEMENTS = {
    "wrapNone", "wrapSquare", "wrapTight", "wrapThrough", "wrapTopAndBottom",
}

#: Word's own default gap beside a wrapped picture: 0.1in left and right,
#: nothing above or below.
_DEFAULT_DIST = {"top": 0, "bottom": 0, "left": 114300, "right": 114300}

_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_VML_NS = "urn:schemas-microsoft-com:vml"


def _local(el: etree._Element) -> str:
    return etree.QName(el).localname


def _place_child(anchor: etree._Element, child: etree._Element) -> None:
    """Insert a child of wp:anchor at its schema-mandated position."""
    index = _ANCHOR_ORDER.index(_local(child))
    for existing in anchor:
        name = _local(existing)
        if name in _ANCHOR_ORDER and _ANCHOR_ORDER.index(name) > index:
            existing.addprevious(child)
            return
    anchor.append(child)


def normalize_wrap(wrap: str) -> str:
    """One of the six wrap names, from whatever the caller spelled."""
    key = str(wrap).strip().lower().replace("-", "_")
    key = _WRAP_ALIASES.get(str(wrap).strip().lower(), _WRAP_ALIASES.get(key, key))
    if key not in WRAP_MODES:
        raise WordMcpError(
            f"unknown wrap {wrap!r}; use one of {sorted(WRAP_MODES)} "
            f"('behind' and 'in_front' both float free of the text, one "
            f"under it and one over it)"
        )
    return key


def _default_polygon(parent: etree._Element) -> None:
    """wrapTight and wrapThrough REQUIRE a wrap polygon, so an unedited
    rectangle is written in the shape coordinate space (0 to 21600 on both
    axes). Word recomputes a real outline the moment a human drags the
    wrap points; edited="0" says nobody has yet."""
    poly = etree.SubElement(parent, f"{{{_WP}}}wrapPolygon")
    poly.set("edited", "0")
    start = etree.SubElement(poly, f"{{{_WP}}}start")
    start.set("x", "0")
    start.set("y", "0")
    for x, y in ((0, 21600), (21600, 21600), (21600, 0), (0, 0)):
        line = etree.SubElement(poly, f"{{{_WP}}}lineTo")
        line.set("x", str(x))
        line.set("y", str(y))


def _build_wrap(wrap: str, wrap_text: str | None) -> etree._Element:
    name, _behind = WRAP_MODES[wrap]
    el = etree.Element(f"{{{_WP}}}{name}")
    if name in ("wrapSquare", "wrapTight", "wrapThrough"):
        side = WRAP_TEXT_SIDES.get(
            str(wrap_text or "both_sides").strip().lower().replace("-", "_")
        )
        if side is None:
            raise WordMcpError(
                f"unknown wrap_text {wrap_text!r}; use one of "
                f"{sorted(set(WRAP_TEXT_SIDES.values()))}"
            )
        el.set("wrapText", side)
    if name in ("wrapTight", "wrapThrough"):
        _default_polygon(el)
    return el


def _build_position(
    axis: str, spec: dict | None
) -> etree._Element:
    """One wp:positionH or wp:positionV. Exactly one of align or posOffset
    is written, because the schema allows exactly one."""
    horizontal = axis == "h"
    table = REL_FROM_H if horizontal else REL_FROM_V
    aligns = ALIGN_H if horizontal else ALIGN_V
    spec = dict(spec or {})
    rel_key = str(
        spec.get("relative_to", "column" if horizontal else "paragraph")
    ).strip().lower().replace("-", "_")
    if rel_key not in table:
        raise WordMcpError(
            f"unknown {'horizontal' if horizontal else 'vertical'} "
            f"relative_to {spec.get('relative_to')!r}; use one of "
            f"{sorted(table)}"
        )
    align = spec.get("align")
    offset = spec.get("offset_pt")
    if align is not None and offset is not None:
        raise WordMcpError(
            "give align or offset_pt for one axis, not both: Word stores "
            "exactly one of them"
        )
    if align is None and offset is None:
        align = "center" if horizontal else None
        offset = None if horizontal else 0
    el = etree.Element(f"{{{_WP}}}position{'H' if horizontal else 'V'}")
    el.set("relativeFrom", table[rel_key])
    if align is not None:
        value = str(align).strip().lower()
        if value not in aligns:
            raise WordMcpError(
                f"unknown {'horizontal' if horizontal else 'vertical'} "
                f"align {align!r}; use one of {sorted(aligns)}"
            )
        etree.SubElement(el, f"{{{_WP}}}align").text = value
    else:
        etree.SubElement(el, f"{{{_WP}}}posOffset").text = str(
            int(round(float(offset) * EMU_PER_PT))
        )
    return el


def _next_relative_height(pkg: DocxPackage) -> int:
    used = [
        int(a.get("relativeHeight", "0") or 0)
        for a in pkg.root().iter(f"{{{_WP}}}anchor")
    ]
    return max(used, default=1) + 1


def _dist_attrs(distance_pt: dict | None) -> dict[str, str]:
    out = dict(_DEFAULT_DIST)
    for key, value in (distance_pt or {}).items():
        name = str(key).strip().lower()
        if name not in out:
            raise WordMcpError(
                f"unknown distance side {key!r}; use top, bottom, left, right"
            )
        if float(value) < 0:
            raise WordMcpError("distance_pt values cannot be negative")
        out[name] = int(round(float(value) * EMU_PER_PT))
    return {
        "distT": str(out["top"]), "distB": str(out["bottom"]),
        "distL": str(out["left"]), "distR": str(out["right"]),
    }


def _build_anchor(
    pkg: DocxPackage, *, cx: int, cy: int, docpr_id: int, wrap: str,
    position: dict | None, z_order: int | None, allow_overlap: bool,
    distance_pt: dict | None, wrap_text: str | None, name: str,
) -> etree._Element:
    """A complete wp:anchor, children written in schema order."""
    wrap = normalize_wrap(wrap)
    _wrap_name, behind = WRAP_MODES[wrap]
    anchor = etree.Element(f"{{{_WP}}}anchor")
    for key, value in _dist_attrs(distance_pt).items():
        anchor.set(key, value)
    anchor.set("simplePos", "0")
    anchor.set(
        "relativeHeight",
        str(int(z_order) if z_order is not None
            else _next_relative_height(pkg)),
    )
    anchor.set("behindDoc", behind)
    anchor.set("locked", "0")
    anchor.set("layoutInCell", "1")
    anchor.set("allowOverlap", "1" if allow_overlap else "0")

    simple = etree.SubElement(anchor, f"{{{_WP}}}simplePos")
    simple.set("x", "0")
    simple.set("y", "0")
    position = position or {}
    anchor.append(_build_position("h", position.get("horizontal")))
    anchor.append(_build_position("v", position.get("vertical")))
    extent = etree.SubElement(anchor, f"{{{_WP}}}extent")
    extent.set("cx", str(cx))
    extent.set("cy", str(cy))
    effect = etree.SubElement(anchor, f"{{{_WP}}}effectExtent")
    for side in ("l", "t", "r", "b"):
        effect.set(side, "0")
    anchor.append(_build_wrap(wrap, wrap_text))
    docpr = etree.SubElement(anchor, f"{{{_WP}}}docPr")
    docpr.set("id", str(docpr_id))
    docpr.set("name", f"Picture {docpr_id}")
    frame = etree.SubElement(anchor, f"{{{_WP}}}cNvGraphicFramePr")
    locks = etree.SubElement(frame, f"{{{_A}}}graphicFrameLocks")
    locks.set("noChangeAspect", "1")
    return anchor


def _anchor_paragraph(
    pkg: DocxPackage, *, at_end: bool, after_index: int | None,
    after_anchor: str | None,
):
    """The paragraph a floating image attaches to.

    An anchored drawing lives in a run inside a paragraph and floats
    relative to it, so unlike the inline path this does NOT create a
    paragraph of its own: it returns the located one, and only makes an
    empty paragraph when the body has none to attach to.
    """
    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if after_anchor is not None:
        return _resolve_anchor(pkg, after_anchor)
    if not at_end and after_index is not None:
        return _body_paragraph(pkg, after_index)
    paragraphs = body.findall(qn("w:p"))
    if paragraphs:
        return paragraphs[-1]
    p = etree.Element(qn("w:p"))
    sectpr = body.find(qn("w:sectPr"))
    if sectpr is not None:
        sectpr.addprevious(p)
    else:
        body.append(p)
    return p


def _add_anchored_image(
    pkg: DocxPackage, src: Path, *, rid: str, media_part: str, cx: int,
    cy: int, docpr_id: int, wrap: str, position: dict | None,
    z_order: int | None, allow_overlap: bool, distance_pt: dict | None,
    at_end: bool, after_index: int | None, after_anchor: str | None,
    wrap_text: str | None = None,
) -> dict:
    paragraph = _anchor_paragraph(
        pkg, at_end=at_end, after_index=after_index,
        after_anchor=after_anchor,
    )
    if _inside_textbox(paragraph):
        raise UnsupportedStructure(
            "a floating image cannot be anchored inside a text box: Word "
            "only allows inline drawings there. Anchor it to a paragraph "
            "in the document body instead, or insert it inline."
        )
    anchor = _build_anchor(
        pkg, cx=cx, cy=cy, docpr_id=docpr_id, wrap=wrap, position=position,
        z_order=z_order, allow_overlap=allow_overlap,
        distance_pt=distance_pt, wrap_text=wrap_text, name=src.name,
    )
    _picture_graphic(
        anchor, rid=rid, name=src.name, pic_id=docpr_id, cx=cx, cy=cy
    )
    r = etree.SubElement(paragraph, qn("w:r"))
    drawing = etree.SubElement(r, qn("w:drawing"))
    drawing.append(anchor)
    pkg.mark_dirty()
    out = {
        "image_added": media_part,
        "placement": "anchored",
        "width_pt": round(cx / EMU_PER_PT, 1),
        "height_pt": round(cy / EMU_PER_PT, 1),
    }
    out.update(describe_anchor(anchor))
    return out


def _inside_textbox(el: etree._Element) -> bool:
    node = el
    while node is not None:
        if _local(node) in ("txbxContent", "textbox"):
            return True
        node = node.getparent()
    return False


def _inside_alternate_content(el: etree._Element) -> bool:
    node = el
    while node is not None:
        if node.tag == f"{{{_MC_NS}}}AlternateContent":
            return True
        node = node.getparent()
    return False


def describe_anchor(anchor: etree._Element) -> dict:
    """Read one wp:anchor back out in the same vocabulary that writes it."""
    out: dict = {}
    wrap_el = next(
        (c for c in anchor if _local(c) in _WRAP_ELEMENTS), None
    )
    behind = anchor.get("behindDoc") in ("1", "true")
    if wrap_el is None:
        out["wrap"] = "unknown"
    elif _local(wrap_el) == "wrapNone":
        out["wrap"] = "behind" if behind else "in_front"
    else:
        out["wrap"] = next(
            key for key, (name, _b) in WRAP_MODES.items()
            if name == _local(wrap_el) and key not in ("behind", "in_front")
        )
        if wrap_el.get("wrapText"):
            out["wrap_text"] = wrap_el.get("wrapText")
        poly = wrap_el.find(f"{{{_WP}}}wrapPolygon")
        if poly is not None and poly.get("edited") in ("1", "true"):
            out["wrap_polygon"] = "edited by hand"
    out["behind_text"] = behind
    out["z_order"] = int(anchor.get("relativeHeight", "0") or 0)
    out["allow_overlap"] = anchor.get("allowOverlap") in ("1", "true")
    reverse_h = {v: k for k, v in REL_FROM_H.items()}
    reverse_v = {v: k for k, v in REL_FROM_V.items()}
    for axis, tag, reverse in (
        ("horizontal", "positionH", reverse_h),
        ("vertical", "positionV", reverse_v),
    ):
        el = anchor.find(f"{{{_WP}}}{tag}")
        if el is None:
            continue
        entry = {
            "relative_to": reverse.get(
                el.get("relativeFrom"), el.get("relativeFrom")
            )
        }
        align = el.find(f"{{{_WP}}}align")
        offset = el.find(f"{{{_WP}}}posOffset")
        if align is not None and align.text:
            entry["align"] = align.text.strip()
        elif offset is not None and offset.text:
            entry["offset_pt"] = round(
                int(offset.text.strip()) / EMU_PER_PT, 1
            )
        out.setdefault("position", {})[axis] = entry
    dist = {}
    for side, attr in (("top", "distT"), ("bottom", "distB"),
                       ("left", "distL"), ("right", "distR")):
        value = anchor.get(attr)
        if value:
            dist[side] = round(int(value) / EMU_PER_PT, 1)
    if dist:
        out["distance_pt"] = dist
    return out


def _drawing_container(blip: etree._Element) -> etree._Element | None:
    """The wp:inline or wp:anchor a blip sits in, whichever it is."""
    node = blip.getparent()
    while node is not None:
        if node.tag in (f"{{{_WP}}}inline", f"{{{_WP}}}anchor"):
            return node
        node = node.getparent()
    return None


_VML_REFUSAL = (
    "image {index} is a legacy VML picture (v:imageData), not a DrawingML "
    "one. VML positions and sizes itself through a CSS-like style string "
    "rather than an anchor, so this server reads it and refuses to rewrite "
    "it rather than guessing at a second layout model. Edit it in Word, or "
    "delete it and insert the picture again."
)


def _blip_at(pkg: DocxPackage, image_index: int) -> etree._Element:
    """The blip for an index from list_images.

    list_images reports legacy VML pictures after the DrawingML ones, so an
    index can legitimately land past the end of the blip list. That case
    gets the VML refusal, which names the real reason, rather than an
    out-of-range error that would read like the image does not exist.
    """
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if 0 <= image_index < len(blips):
        return blips[image_index]
    vml = list(pkg.root().iter(f"{{{_VML_NS}}}imageData"))
    if len(blips) <= image_index < len(blips) + len(vml):
        raise UnsupportedStructure(_VML_REFUSAL.format(index=image_index))
    raise TargetNotFound(
        f"image index {image_index} out of range "
        f"({len(blips) + len(vml)} images)"
    )


def list_images(pkg: DocxPackage) -> list[dict]:
    """Every image in the document part, inline and anchored alike.

    An anchored entry carries its full placement (wrap mode, position on
    both axes, z-order, whether it sits behind the text), so a caller can
    read a floating image, change one thing, and write it back without
    having to reconstruct the rest.
    """
    rels_root = (
        pkg.root("word/_rels/document.xml.rels")
        if pkg.has_part("word/_rels/document.xml.rels")
        else None
    )
    rid_target = (
        {r.get("Id"): r.get("Target") for r in rels_root}
        if rels_root is not None
        else {}
    )
    out = []
    for i, blip in enumerate(pkg.root().iter(f"{{{_A}}}blip")):
        rid = blip.get(f"{{{_R_NS}}}embed")
        entry = {"index": i, "rel_id": rid, "target": rid_target.get(rid)}
        container = _drawing_container(blip)
        if container is not None:
            extent = container.find(f"{{{_WP}}}extent")
            if extent is not None:
                entry["width_pt"] = round(int(extent.get("cx")) / EMU_PER_PT, 1)
                entry["height_pt"] = round(int(extent.get("cy")) / EMU_PER_PT, 1)
            if _local(container) == "anchor":
                entry["placement"] = "anchored"
                entry.update(describe_anchor(container))
            else:
                entry["placement"] = "inline"
            if _inside_alternate_content(container):
                entry["in_alternate_content"] = True
            if _inside_textbox(container):
                entry["in_textbox"] = True
        out.append(entry)
    out.extend(_list_vml_images(pkg, rid_target, start=len(out)))
    return out


def _list_vml_images(
    pkg: DocxPackage, rid_target: dict, *, start: int
) -> list[dict]:
    """Legacy VML pictures (v:imageData), which are still present in a
    noticeable share of real documents.

    They are REPORTED so an enumeration is not silently short, and they are
    listed after the DrawingML images so existing indices do not shift.
    Nothing here edits them: VML positioning is a different model and this
    module would be guessing.
    """
    out = []
    seen: set[int] = set()
    for i, data in enumerate(pkg.root().iter(f"{{{_VML_NS}}}imageData")):
        if id(data) in seen:
            continue
        seen.add(id(data))
        rid = data.get(f"{{{_R_NS}}}id") or data.get(f"{{{_R_NS}}}embed")
        out.append({
            "index": start + i,
            "rel_id": rid,
            "target": rid_target.get(rid),
            "placement": "vml",
            "editable": False,
            "note": (
                "legacy VML picture: readable here, but placement and size "
                "are not editable through this server"
            ),
        })
    return out


def resize_image(pkg: DocxPackage, image_index: int, *, width_pt: float) -> dict:
    """Resize by index (as reported by list_images), keeping aspect ratio.

    Works on both placements: the extent lives in the same place in
    wp:inline and wp:anchor, so a floating image resizes like an inline
    one and keeps its position and wrap.
    """
    if width_pt <= 0:
        raise WordMcpError("width_pt must be positive")
    blip = _blip_at(pkg, image_index)
    container = _drawing_container(blip)
    if container is None:
        raise UnsupportedStructure(
            "this image is not in a DrawingML inline or anchor container "
            "(a legacy VML picture, most likely); resize it in Word"
        )
    extent = container.find(f"{{{_WP}}}extent")
    old_cx = int(extent.get("cx"))
    old_cy = int(extent.get("cy"))
    new_cx = int(width_pt * EMU_PER_PT)
    new_cy = int(new_cx * old_cy / old_cx)
    extent.set("cx", str(new_cx))
    extent.set("cy", str(new_cy))
    for xfrm_ext in container.iter(f"{{{_A}}}ext"):
        xfrm_ext.set("cx", str(new_cx))
        xfrm_ext.set("cy", str(new_cy))
    pkg.mark_dirty()
    return {
        "resized": image_index,
        "width_pt": width_pt,
        "height_pt": round(new_cy / EMU_PER_PT, 1),
    }


def set_image_placement(
    pkg: DocxPackage,
    image_index: int,
    *,
    wrap: str | None = None,
    position: dict | None = None,
    z_order: int | None = None,
    allow_overlap: bool | None = None,
    distance_pt: dict | None = None,
    wrap_text: str | None = None,
    force: bool = False,
) -> dict:
    """Change how an existing image sits on the page.

    wrap='inline' returns a floating image to the text flow. Any other
    wrap floats an inline image, or re-wraps one that already floats.
    Everything the call does not mention is left exactly as it was, which
    is the whole point: a caller can change the wrap of a picture somebody
    positioned by hand without losing where they put it.

    Two constructs are refused rather than rewritten. An image inside an
    mc:AlternateContent pair exists twice, once per branch, and editing
    one branch would leave the two disagreeing about the same picture. A
    wrap polygon carrying edited="1" is a hand-dragged outline, and no
    generated rectangle can reproduce it, so changing away from that wrap
    mode needs force and says what it will cost.
    """
    blip = _blip_at(pkg, image_index)
    container = _drawing_container(blip)
    if container is None:
        raise UnsupportedStructure(
            "this image is not a DrawingML picture (a legacy VML picture, "
            "most likely). Its placement model is different and this "
            "server does not rewrite it; reposition it in Word."
        )
    if _inside_alternate_content(container):
        raise UnsupportedStructure(
            "this image sits inside an mc:AlternateContent pair, so the "
            "same picture is described twice, once for modern Word and "
            "once as a fallback. Rewriting one branch would leave the two "
            "disagreeing, so this call refuses rather than half-editing it."
        )
    drawing = container.getparent()
    kind = _local(container)
    wants_inline = wrap is not None and str(wrap).strip().lower() == "inline"

    if wants_inline:
        if position is not None or z_order is not None:
            raise WordMcpError(
                "wrap='inline' puts the image back in the text flow, where "
                "position and z_order have no meaning; drop them or choose "
                "a floating wrap mode"
            )
        if kind == "inline":
            return {"placement": "inline", "changed": False,
                    "note": "the image was already inline"}
        lost = describe_anchor(container)
        new = _to_inline(container)
        drawing.replace(container, new)
        pkg.mark_dirty()
        return {
            "placement": "inline", "changed": True,
            "discarded": {
                key: lost[key] for key in ("wrap", "position", "z_order")
                if key in lost
            },
        }

    if kind == "inline":
        if wrap is None:
            raise WordMcpError(
                "this image is inline, so there is no placement to change "
                f"yet; pass wrap (one of {sorted(WRAP_MODES)}) to float it"
            )
        if _inside_textbox(drawing):
            raise UnsupportedStructure(
                "a floating image cannot live inside a text box: Word only "
                "allows inline drawings there. This one stays inline."
            )
        extent = container.find(f"{{{_WP}}}extent")
        graphic = container.find(f"{{{_A}}}graphic")
        if extent is None or graphic is None:
            raise UnsupportedStructure(
                "this inline drawing is missing its wp:extent or its "
                "a:graphic, so there is nothing complete to float"
            )
        anchor = _build_anchor(
            pkg, cx=int(extent.get("cx")), cy=int(extent.get("cy")),
            docpr_id=_docpr_id_of(container, pkg), wrap=wrap,
            position=position, z_order=z_order,
            allow_overlap=True if allow_overlap is None else allow_overlap,
            distance_pt=distance_pt, wrap_text=wrap_text, name="",
        )
        anchor.append(graphic)
        drawing.replace(container, anchor)
        pkg.mark_dirty()
        return {"placement": "anchored", "changed": True,
                **describe_anchor(anchor)}

    # Already anchored: edit in place, touching only what was asked for.
    changed: list[str] = []
    if wrap is not None:
        name = normalize_wrap(wrap)
        old = next((c for c in container if _local(c) in _WRAP_ELEMENTS), None)
        if old is not None:
            poly = old.find(f"{{{_WP}}}wrapPolygon")
            edited = poly is not None and poly.get("edited") in ("1", "true")
            keeps_polygon = WRAP_MODES[name][0] in ("wrapTight", "wrapThrough")
            if edited and not keeps_polygon and not force:
                raise UnsupportedStructure(
                    "this image's wrap outline was edited by hand (the "
                    "wrap polygon carries edited=\"1\"), and the requested "
                    "wrap mode cannot hold a polygon, so the outline would "
                    "be lost. Pass force to accept that, or keep a tight "
                    "or through wrap."
                )
            if edited and keeps_polygon:
                # Move the hand-edited outline into the new wrap element
                # rather than replacing it with a generated rectangle.
                new = etree.Element(f"{{{_WP}}}{WRAP_MODES[name][0]}")
                new.set("wrapText", WRAP_TEXT_SIDES.get(
                    str(wrap_text or old.get("wrapText") or "both_sides")
                    .strip().lower().replace("-", "_"), "bothSides"))
                new.append(poly)
            else:
                new = _build_wrap(name, wrap_text or (
                    old.get("wrapText") if old is not None else None))
            container.remove(old)
        else:
            new = _build_wrap(name, wrap_text)
        _place_child(container, new)
        container.set("behindDoc", WRAP_MODES[name][1])
        changed.append("wrap")
    elif wrap_text is not None:
        old = next((c for c in container if _local(c) in _WRAP_ELEMENTS), None)
        if old is None or _local(old) == "wrapTopAndBottom" or (
            _local(old) == "wrapNone"
        ):
            raise WordMcpError(
                "wrap_text only applies to the wrap modes that flow text "
                "beside the image (square, tight, through)"
            )
        side = WRAP_TEXT_SIDES.get(
            str(wrap_text).strip().lower().replace("-", "_")
        )
        if side is None:
            raise WordMcpError(f"unknown wrap_text {wrap_text!r}")
        old.set("wrapText", side)
        changed.append("wrap_text")

    if position is not None:
        for axis, tag in (("horizontal", "positionH"),
                          ("vertical", "positionV")):
            if axis not in position:
                continue
            new = _build_position("h" if axis == "horizontal" else "v",
                                  position[axis])
            old = container.find(f"{{{_WP}}}{tag}")
            if old is not None:
                container.remove(old)
            _place_child(container, new)
            changed.append(axis)
    if z_order is not None:
        container.set("relativeHeight", str(int(z_order)))
        changed.append("z_order")
    if allow_overlap is not None:
        container.set("allowOverlap", "1" if allow_overlap else "0")
        changed.append("allow_overlap")
    if distance_pt is not None:
        for key, value in _dist_attrs(distance_pt).items():
            container.set(key, value)
        changed.append("distance_pt")
    if not changed:
        raise WordMcpError(
            "nothing to change: give wrap, position, z_order, "
            "allow_overlap, wrap_text, or distance_pt"
        )
    pkg.mark_dirty()
    return {"placement": "anchored", "changed": changed,
            **describe_anchor(container)}


def _docpr_id_of(container: etree._Element, pkg: DocxPackage) -> int:
    docpr = container.find(f"{{{_WP}}}docPr")
    if docpr is not None and docpr.get("id"):
        try:
            return int(docpr.get("id"))
        except ValueError:
            pass
    return _next_docpr_id(pkg)


def _to_inline(anchor: etree._Element) -> etree._Element:
    """A wp:inline carrying the anchor's size, identity, and picture.

    Everything anchor-only (position, wrap, z-order, the extension list)
    is dropped, because inline placement has nowhere to keep it. The
    caller reports what went.
    """
    inline = etree.Element(f"{{{_WP}}}inline")
    for attr in ("distT", "distB", "distL", "distR"):
        inline.set(attr, anchor.get(attr, "0"))
    extent = anchor.find(f"{{{_WP}}}extent")
    if extent is None:
        raise UnsupportedStructure(
            "this anchor carries no wp:extent, so its size is unknown and "
            "an inline drawing cannot be built from it"
        )
    inline.append(extent)
    effect = anchor.find(f"{{{_WP}}}effectExtent")
    if effect is not None:
        inline.append(effect)
    docpr = anchor.find(f"{{{_WP}}}docPr")
    if docpr is not None:
        inline.append(docpr)
    frame = anchor.find(f"{{{_WP}}}cNvGraphicFramePr")
    if frame is not None:
        inline.append(frame)
    graphic = anchor.find(f"{{{_A}}}graphic")
    if graphic is None:
        raise UnsupportedStructure(
            "this anchor carries no a:graphic, so there is no picture to "
            "carry into an inline drawing"
        )
    inline.append(graphic)
    return inline


def replace_image(pkg: DocxPackage, image_index: int, new_image_path: str) -> dict:
    """Swap an image's bytes, keeping placement and display size.

    Nothing in the document markup is touched, only the media part behind
    it, so an anchored image keeps its wrap, position, and z-order across
    the swap by construction rather than by careful copying.
    """
    check_path(new_image_path, "read image file")
    src = Path(new_image_path)
    if not src.exists():
        raise TargetNotFound(f"image file not found: {new_image_path}")
    ext = src.suffix.lower()
    if ext not in _EXT_TYPES:
        raise WordMcpError(f"unsupported image type {ext}")
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if not 0 <= image_index < len(blips):
        raise TargetNotFound(f"image index {image_index} out of range")
    rid = blips[image_index].get(f"{{{_R_NS}}}embed")
    rels_root = pkg.root("word/_rels/document.xml.rels")
    target = next((r.get("Target") for r in rels_root if r.get("Id") == rid), None)
    if target is None:
        raise TargetNotFound(f"no relationship for image {image_index}")
    part = "word/" + target.lstrip("/")
    old_ext = "." + part.rsplit(".", 1)[1].lower()
    if old_ext != ext:
        raise WordMcpError(
            f"replacement must be the same type as the original ({old_ext}); "
            "or add a new image and delete this one"
        )
    pkg.set_raw_part(part, src.read_bytes())
    container = _drawing_container(blips[image_index])
    placement = (
        "anchored" if container is not None and _local(container) == "anchor"
        else "inline"
    )
    out = {"replaced": part, "placement": placement}
    if placement == "anchored":
        out.update(describe_anchor(container))
    return out


# ----------------------------------------------------------- delete (v2 ops)
#
# v1 had no image deletion path (a parity gap the V2_DESIGN Section 3.2
# delete_element multiplex closes). These helpers are additive: nothing
# above this line changed.

_INERT_PARA_CHILDREN = {"pPr", "proofErr", "bookmarkStart", "bookmarkEnd"}


def _remove_pkg_part(pkg: DocxPackage, name: str) -> bool:
    """Drop a part from the package plus its [Content_Types].xml override.

    DocxPackage has no public remove-part API; save() serializes strictly
    from _order/_raw/_dirty, so removing the entry from all of them removes
    the part from the output (the same reach-in ops/cleanup.py and
    ops/dataio.py already use; candidate for promotion to
    DocxPackage.remove_part())."""
    if not pkg.has_part(name):
        return False
    pkg._raw.pop(name)
    pkg._order.remove(name)
    pkg._trees.pop(name, None)
    pkg._dirty.discard(name)
    ct_root = pkg.root("[Content_Types].xml")
    for o in list(ct_root.findall(f"{{{_CT_NS}}}Override")):
        if o.get("PartName") == "/" + name:
            ct_root.remove(o)
            pkg.mark_dirty("[Content_Types].xml")
    return True


def _target_referenced(pkg: DocxPackage, part: str) -> bool:
    """True when ANY remaining .rels part in the package still targets
    `part` (header/footer/notes rels keep shared media alive)."""
    import posixpath

    for name in pkg.part_names():
        if not name.endswith(".rels"):
            continue
        base = name.split("/_rels/", 1)[0] if "/_rels/" in name else ""
        for rel in pkg.root(name):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target") or ""
            if target.startswith("/"):
                resolved = posixpath.normpath(target.lstrip("/"))
            elif base:
                resolved = posixpath.normpath(posixpath.join(base, target))
            else:
                resolved = posixpath.normpath(target)
            if resolved == part:
                return True
    return False


def _paragraph_is_empty(p: etree._Element) -> bool:
    """No content besides inert markers and runs holding only rPr."""
    for child in p:
        name = etree.QName(child).localname
        if name in _INERT_PARA_CHILDREN:
            continue
        if name == "r":
            if any(etree.QName(rc).localname != "rPr" for rc in child):
                return False
            continue
        return False
    return True


def _detach_drawing(drawing: etree._Element) -> bool:
    """Remove a w:drawing, prune the emptied run, and remove the emptied
    paragraph when it is a body-level block and the body keeps at least one
    other block. Returns True when the whole paragraph was removed."""
    run = drawing.getparent()
    run.remove(drawing)
    node = run
    if run.tag == qn("w:r") and all(
        etree.QName(c).localname == "rPr" for c in run
    ):
        node = run.getparent()
        node.remove(run)
    p = node
    while p is not None and p.tag != qn("w:p"):
        p = p.getparent()
    if p is None or not _paragraph_is_empty(p):
        return False
    body = p.getparent()
    if body is None or body.tag != qn("w:body"):
        # Inside a table cell or text box: cells must keep a paragraph.
        return False
    keeps_others = any(
        el is not p and etree.QName(el).localname in ("p", "tbl")
        for el in body
    )
    if not keeps_others:
        return False
    body.remove(p)
    return True


def delete_image(pkg: DocxPackage, image_index: int) -> dict:
    """Delete an image by its list_images index: the drawing goes (inline or
    floating), the emptied paragraph goes when it held nothing else, and
    the relationship plus media part go once nothing else references them
    (shared media referenced from headers/footers/notes survives)."""
    blips = list(pkg.root().iter(f"{{{_A}}}blip"))
    if not 0 <= image_index < len(blips):
        raise TargetNotFound(
            f"image index {image_index} out of range ({len(blips)} images; "
            "see list_images)"
        )
    blip = blips[image_index]
    rid = blip.get(f"{{{_R_NS}}}embed")
    drawing = blip.getparent()
    while drawing is not None and drawing.tag != qn("w:drawing"):
        drawing = drawing.getparent()
    if drawing is None:
        raise WordMcpError(
            "image is not inside a w:drawing (legacy VML picture); "
            "deletion is unsupported for this shape"
        )
    removed_paragraph = _detach_drawing(drawing)
    pkg.mark_dirty()

    rels_part = "word/_rels/document.xml.rels"
    media_part = None
    part_removed = False
    if rid is not None and pkg.has_part(rels_part):
        rels_root = pkg.root(rels_part)
        target = next(
            (r.get("Target") for r in rels_root if r.get("Id") == rid), None
        )
        if target is not None and not target.startswith(".."):
            media_part = "word/" + target.lstrip("/")
        still_used = any(
            b.get(f"{{{_R_NS}}}embed") == rid
            for b in pkg.root().iter(f"{{{_A}}}blip")
        )
        if not still_used:
            for r in list(rels_root):
                if r.get("Id") == rid:
                    rels_root.remove(r)
                    pkg.mark_dirty(rels_part)
            if media_part and not _target_referenced(pkg, media_part):
                part_removed = _remove_pkg_part(pkg, media_part)
    return {
        "deleted_image": image_index,
        "removed_paragraph": removed_paragraph,
        "media_part": media_part,
        "media_part_removed": part_removed,
    }
