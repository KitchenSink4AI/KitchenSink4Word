"""Whole-document positional insertion: insert one .docx's entire body into
another at an exact position, with full resource reconciliation.

Built for the chapter-merge workflow (2026-08-28 merge test): merging a
100+-paragraph chapter with tables into a dissertation at a specific position
was previously only possible with a hand-written lxml script, because
com_merge_documents concatenates end-to-end and insert_paragraphs cannot carry
tables or large payloads. This module is that script, made safe.

What "safe" means here — a naive element transplant corrupts, so every
resource class the transplanted content references is reconciled:

- STYLES: matched to target styles BY NAME (the apply_template model, ops/
  template.py). Matching names remap styleIds to the target's id; the target's
  formatting governs (merge-don't-replace). Unmatched styles are cloned into
  the target with their basedOn/link/next dependency chains, under fresh ids
  when the id is taken.
- NUMBERING: every source list instance (numId) gets its own freshly-cloned
  abstractNum + num pair in the target with non-colliding ids, so numbering
  restart semantics are preserved exactly and no source list ever attaches to
  a target list.
- FOOTNOTES/ENDNOTES: definitions are copied with new non-colliding ids and
  the transplanted references retagged; note content goes through the same
  style/numbering/relationship reconciliation as body content.
- IMAGES: media parts copied under fresh names, new relationship ids, content
  types ensured, docPr ids renumbered for uniqueness.
- CHARTS: the chart part and its private subtree (embedded workbook, colors/
  style parts) are copied part-for-part with rewritten relationship targets
  and content-type overrides (the ops/charts.py plumbing pattern).
- HYPERLINKS and other external relationships: re-registered with fresh rIds.
- BOOKMARKS: ids remapped to fresh values; name collisions renamed with a
  suffix (reported), and w:anchor hyperlinks / REF-family fields inside the
  inserted content retargeted to the renamed bookmarks.
- COMMENTS: comment transplant is OUT OF SCOPE. Comment references and range
  markers in the source content are stripped cleanly and the count reported.
- TRACKED CHANGES: body-level revision markup (w:ins/w:del/...) is carried
  as-is.
- SECTIONS: the source's trailing sectPr is NEVER carried — body content
  only; headers, footers, and page setup stay the target's. Mid-content
  section breaks (paragraph-embedded sectPr) are stripped and reported, since
  carrying them would import the source's page furniture.

Anything that cannot be carried safely (OLE objects, ActiveX controls,
subdocuments, altChunks, SmartArt/unknown embedded parts) REFUSES the whole
insertion, naming the blocking content — nothing is ever half-applied.

insert_document's `formatting` parameter mirrors Word's paste options:
"source" (default — direct formatting preserved, Word InsertFile behavior),
"merge" (semantic emphasis kept, font/size/color/spacing/indent direct
overrides stripped so the target's styles size the carried text), and
"destination" (all direct rPr/pPr formatting stripped except structural
properties — numPr, outlineLvl, tab stops — so the target's styles govern
rendering entirely). Stripping happens on the COPIED elements only; the
source file is never modified.

copy_table() transplants a single top-level table through the exact same
reconciliation pipeline (styles by name, numbering, rels/images, notes,
bookmarks) scoped to that one element, with the same positioning contract
and refusal classes as insert_document.
"""

from __future__ import annotations

import colorsys
import copy
import math
import posixpath
import re

from lxml import etree

from ..core.errors import (
    AmbiguousTarget,
    TargetNotFound,
    UnsupportedStructure,
    ValidationFailed,
    WordMcpError,
)
from ..core.package import DocxPackage, qn
from . import lists as _lists
from . import media as _media
from . import notes as _notes
from .read import paragraph_text
from .template import _style_maps

_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

_REL_IMAGE = _R_NS + "/image"
_REL_HYPERLINK = _R_NS + "/hyperlink"
_REL_CHART = _R_NS + "/chart"
_REL_CHARTEX = "http://schemas.microsoft.com/office/2014/relationships/chartEx"

# Internal relationship types we know how to transplant. Everything else that
# the inserted content references via r:id is a refusal (conservative mode).
_SUBTREE_REL_TYPES = {_REL_CHART, _REL_CHARTEX}

# Elements whose presence in the source content blocks the whole insertion.
_BLOCKED_ELEMENTS = {
    qn("w:object"): "embedded OLE object (w:object)",
    qn("w:control"): "ActiveX control (w:control)",
    qn("w:subDoc"): "subdocument reference (w:subDoc)",
    qn("w:altChunk"): "altChunk (embedded alternative-format content)",
    qn("w:movie"): "embedded movie (w:movie)",
    "{urn:schemas-microsoft-com:office:office}OLEObject": (
        "embedded OLE object (o:OLEObject)"
    ),
}

# Tags carrying style-id references, in content AND in style/numbering defs.
_STYLE_REF_TAGS = (
    "w:pStyle",
    "w:rStyle",
    "w:tblStyle",
    "w:basedOn",
    "w:link",
    "w:next",
    "w:styleLink",
    "w:numStyleLink",
)

# story -> (destination part, its rels part) — origin and destination match.
_STORY_PARTS = {
    "document": ("word/document.xml", "word/_rels/document.xml.rels"),
    "footnote": ("word/footnotes.xml", "word/_rels/footnotes.xml.rels"),
    "endnote": ("word/endnotes.xml", "word/_rels/endnotes.xml.rels"),
}


# ------------------------------------------------------------- small helpers


def _localname(el: etree._Element) -> str:
    return etree.QName(el).localname


def _rels_part_for(part: str) -> str:
    folder, name = part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _resolve_rel_target(rels_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    base = rels_part.rsplit("_rels/", 1)[0].rstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


def _ensure_rels_part(pkg: DocxPackage, name: str) -> None:
    if pkg.has_part(name):
        return
    root = etree.Element(f"{{{_REL_NS}}}Relationships", nsmap={None: _REL_NS})
    pkg.set_raw_part(
        name,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )


def _iter_rid_attrs(el: etree._Element):
    """Yield (node, attr_key, rid) for every r:-namespace attribute below
    (and on) el."""
    prefix = "{" + _R_NS + "}"
    for node in el.iter():
        for key, val in node.attrib.items():
            if key.startswith(prefix):
                yield node, key, val


def _free_part_name(pkg: DocxPackage, like: str) -> str:
    """A part name in the same folder / same shape as `like` that does not
    exist in the target: trailing digits before the extension are replaced by
    the first free number."""
    folder, base = like.rsplit("/", 1)
    m = re.fullmatch(r"(.*?)(\d*)(\.[^./]+)?", base)
    stem, _, ext = m.group(1), m.group(2), m.group(3) or ""
    n = 1
    while True:
        cand = f"{folder}/{stem}{n}{ext}"
        if not pkg.has_part(cand):
            return cand
        n += 1


def _content_type_of(src: DocxPackage, part: str) -> tuple[str, str] | None:
    """('override'|'default', content_type) for a source part, from its
    [Content_Types].xml."""
    ct_root = src.root("[Content_Types].xml")
    part_name = "/" + part
    for o in ct_root.findall(f"{{{_CT_NS}}}Override"):
        if o.get("PartName") == part_name:
            return "override", o.get("ContentType")
    ext = part.rsplit(".", 1)[1].lower() if "." in part.rsplit("/", 1)[1] else ""
    for d in ct_root.findall(f"{{{_CT_NS}}}Default"):
        if (d.get("Extension") or "").lower() == ext:
            return "default", d.get("ContentType")
    return None


def _ensure_content_type(
    pkg: DocxPackage, src: DocxPackage, src_part: str, new_part: str
) -> None:
    """Replicate the source part's content-type declaration for its copy."""
    found = _content_type_of(src, src_part)
    ct_root = pkg.root("[Content_Types].xml")
    if found is None:
        # No declaration in the source either; fall back to a media guess.
        ext = "." + new_part.rsplit(".", 1)[1].lower() if "." in new_part else ""
        ctype = _media._EXT_TYPES.get(ext, "application/octet-stream")
        kind = "default"
    else:
        kind, ctype = found
    if kind == "override":
        part_name = "/" + new_part
        if not any(
            o.get("PartName") == part_name
            for o in ct_root.findall(f"{{{_CT_NS}}}Override")
        ):
            o = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
            o.set("PartName", part_name)
            o.set("ContentType", ctype)
            pkg.mark_dirty("[Content_Types].xml")
    else:
        ext_name = new_part.rsplit(".", 1)[1].lower()
        if not any(
            (d.get("Extension") or "").lower() == ext_name
            for d in ct_root.findall(f"{{{_CT_NS}}}Default")
        ):
            d = etree.SubElement(ct_root, f"{{{_CT_NS}}}Default")
            d.set("Extension", ext_name)
            d.set("ContentType", ctype)
            pkg.mark_dirty("[Content_Types].xml")


def _ensure_styles_part(pkg: DocxPackage) -> None:
    """Create a minimal word/styles.xml (plus content type and relationship)
    for the rare target that lacks one, so style cloning has a home."""
    part = "word/styles.xml"
    if pkg.has_part(part):
        return
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn("w:styles"), nsmap={"w": w_ns})
    pkg.set_raw_part(
        part,
        etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        ),
    )
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        o.get("PartName") == "/" + part
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        o = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        o.set("PartName", "/" + part)
        o.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.styles+xml",
        )
        pkg.mark_dirty("[Content_Types].xml")
    rels_part = "word/_rels/document.xml.rels"
    if pkg.has_part(rels_part):
        rels_root = pkg.root(rels_part)
        rel_type = _R_NS + "/styles"
        if not any(r.get("Type") == rel_type for r in rels_root):
            existing = {r.get("Id") for r in rels_root}
            n = 1
            while f"rId{n}" in existing:
                n += 1
            rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
            rel.set("Id", f"rId{n}")
            rel.set("Type", rel_type)
            rel.set("Target", "styles.xml")
            pkg.mark_dirty(rels_part)


# ------------------------------------------------------------- position logic


def _body_blocks(pkg: DocxPackage) -> list[etree._Element]:
    """Body items — paragraphs and tables as ONE document-order sequence."""
    return [
        c
        for c in pkg.body()
        if _localname(c) in ("p", "tbl")
    ]


def _resolve_position(
    pkg: DocxPackage,
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    before_first: bool = False,
) -> tuple[str, etree._Element | None, int]:
    """(mode, reference element, body-item index where inserted content
    starts). mode: 'append' (before the trailing sectPr), 'after' (after
    the reference element), or 'before' (before the reference element,
    used for insertion ahead of the first body item)."""
    blocks = _body_blocks(pkg)
    if before_first:
        if not blocks:
            return "append", None, 0
        return "before", blocks[0], 0
    if at_end:
        return "append", None, len(blocks)
    if after_index is not None:
        if not 0 <= after_index < len(blocks):
            raise TargetNotFound(
                f"after_index {after_index} out of range: the target has "
                f"{len(blocks)} body items (paragraphs + tables in document "
                f"order, valid indices 0-{len(blocks) - 1})"
            )
        return "after", blocks[after_index], after_index + 1
    # after_anchor: exact paragraph-text match, refused loudly on ambiguity —
    # heading text recurring in body prose must never resolve first-match.
    anchor = after_anchor.strip()
    matches = [
        (i, el)
        for i, el in enumerate(blocks)
        if _localname(el) == "p" and paragraph_text(el).strip() == anchor
    ]
    if not matches:
        raise TargetNotFound(
            f"no body paragraph's full text exactly matches {anchor!r} "
            "(after_anchor compares whole-paragraph text; use after_index "
            "for structural positions)"
        )
    if len(matches) > 1:
        locations = [
            {
                "body_item_index": i,
                "text": paragraph_text(el).strip()[:120],
            }
            for i, el in matches
        ]
        raise AmbiguousTarget(
            f"anchor text matches {len(matches)} paragraphs; refusing to "
            f"guess. Matches (body item indices): {locations}. Use "
            "after_index with the intended index instead."
        )
    i, el = matches[0]
    return "after", el, i + 1


# --------------------------------------------------------- style / numbering


# Style children that are identity or housekeeping, not formatting: two
# definitions that differ only in these are the same style.
_STYLE_IDENTITY_CHILDREN = frozenset({
    "name", "aliases", "rsid", "uiPriority", "semiHidden",
    "unhideWhenUsed", "qFormat", "locked", "autoRedefine", "hidden",
    "personal", "personalCompose", "personalReply",
})


def _style_signature(el: etree._Element, *, ignore_refs: bool = False) -> str:
    """Canonical serialization of a style's FORMATTING, ignoring its id,
    name and housekeeping children. ignore_refs also drops basedOn/link/
    next, whose values are style IDS and therefore differ between a source
    definition and the copy of it a previous insert already imported."""
    clone = copy.deepcopy(el)
    for attr in list(clone.attrib):
        if _localname_of_attr(attr) in ("styleId", "rsid"):
            clone.attrib.pop(attr)
    drop = _STYLE_IDENTITY_CHILDREN
    if ignore_refs:
        drop = drop | {"basedOn", "link", "next"}
    for child in list(clone):
        if _localname(child) in drop:
            clone.remove(child)
    return etree.tostring(clone, method="c14n").decode()


_IMPORTED_SUFFIX = re.compile(r"^(?P<base>.+?) \(imported(?: \d+)?\)$")


def _imported_base_name(name: str | None) -> str | None:
    """'Table Grid (imported 2)' -> 'Table Grid'."""
    if not name:
        return None
    m = _IMPORTED_SUFFIX.match(name)
    return m.group("base") if m else name


def _localname_of_attr(attr: str) -> str:
    return attr.rsplit("}", 1)[-1]


class _StyleResolver:
    """By-name style reconciliation (the apply_template model): matching
    names remap ids to the target's; unmatched styles are cloned with their
    dependency chains under fresh ids.

    TABLE styles are the exception (adversarial review 2026-09-20, M3): a
    table style carries conditional formatting per region (tblStylePr), so
    a same-named definition that differs cannot be baked onto the carried
    cells the way a paragraph or character style can. Such a style is
    imported under a fresh id AND a fresh name, and the carried tables are
    re-pointed at it. Paragraph and character styles are NEVER renamed:
    that would break heading-based TOCs and the by-name contract."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        self.pkg = pkg
        self.src_id2name, _ = _style_maps(src)
        self.tgt_id2name, self.tgt_name2id = _style_maps(pkg)
        self.src_defs: dict[str, etree._Element] = {}
        if src.has_part("word/styles.xml"):
            for s in src.root("word/styles.xml").findall(qn("w:style")):
                sid = s.get(qn("w:styleId"))
                if sid:
                    self.src_defs[sid] = s
        self.tgt_defs: dict[str, etree._Element] = {}
        if pkg.has_part("word/styles.xml"):
            for s in pkg.root("word/styles.xml").findall(qn("w:style")):
                sid = s.get(qn("w:styleId"))
                if sid:
                    self.tgt_defs[sid] = s
        self.remap: dict[str, str] = {}  # src id -> different target id
        self.matched: list[dict] = []
        self.cloned: list[dict] = []
        self.imported_renamed: list[dict] = []
        self.cloned_defs: list[etree._Element] = []
        self.unresolved: list[str] = []
        self._done: set[str] = set()

    def _free_name(self, name: str) -> str:
        candidate = f"{name} (imported)"
        n = 2
        while candidate in self.tgt_name2id:
            candidate = f"{name} (imported {n})"
            n += 1
        return candidate

    def _already_imported(
        self, name: str, src_def: etree._Element
    ) -> str | None:
        """The id of an "X (imported…)" style a PREVIOUS insert created from
        this same definition, or None. Identity is the formatting signature
        (ids and names stripped) plus the parent's name, so a style is
        imported once however many times its source is inserted."""
        want = _style_signature(src_def, ignore_refs=True)
        want_parent = self._parent_name(src_def, self.src_id2name)
        for cand_id, cand_name in self.tgt_id2name.items():
            if cand_id == name or cand_name == name:
                continue
            if _imported_base_name(cand_name) != name:
                continue
            cand = self.tgt_defs.get(cand_id)
            if cand is None or cand.get(qn("w:type")) != src_def.get(
                qn("w:type")
            ):
                continue
            if _style_signature(cand, ignore_refs=True) != want:
                continue
            if self._parent_name(cand, self.tgt_id2name) != want_parent:
                continue
            return cand_id
        return None

    @staticmethod
    def _parent_name(el: etree._Element, id2name: dict) -> str | None:
        base = el.find(qn("w:basedOn"))
        if base is None or not base.get(qn("w:val")):
            return None
        return _imported_base_name(id2name.get(base.get(qn("w:val"))))

    def resolve(self, sid: str) -> None:
        if not sid or sid in self._done:
            return
        self._done.add(sid)
        name = self.src_id2name.get(sid)
        if name is None:
            # No definition in the source. If the target defines the id the
            # reference lands on that style; otherwise it dangles exactly as
            # it dangled in the source (Word falls back to Normal). Reported.
            if sid not in self.tgt_id2name:
                self.unresolved.append(sid)
            return
        src_def = self.src_defs[sid]
        is_table = src_def.get(qn("w:type")) == "table"
        tgt_id = self.tgt_name2id.get(name)
        rename_to = None
        if tgt_id is not None and is_table:
            tgt_def = self.tgt_defs.get(tgt_id)
            if tgt_def is not None and _style_signature(
                tgt_def
            ) != _style_signature(src_def):
                already = self._already_imported(name, src_def)
                if already is not None:
                    # A previous insert of this source already imported it;
                    # re-point at that copy instead of stacking another
                    # "(imported N)" in the user's gallery (review m8).
                    if already != sid:
                        self.remap[sid] = already
                    self.matched.append({
                        "name": self.tgt_id2name.get(already, name),
                        "source_id": sid, "target_id": already,
                        "reused_import": True,
                    })
                    return
                rename_to = self._free_name(name)
                tgt_id = None  # import it instead of matching by name
        if tgt_id is not None:
            # Name match: the target's definition (formatting) governs.
            if tgt_id != sid:
                self.remap[sid] = tgt_id
            self.matched.append(
                {"name": name, "source_id": sid, "target_id": tgt_id}
            )
            return
        # Clone, keeping the source id when free.
        new_id = sid
        n = 1
        while new_id in self.tgt_id2name:
            new_id = f"{sid}Ins{n}"
            n += 1
        d = copy.deepcopy(self.src_defs[sid])
        d.set(qn("w:styleId"), new_id)
        new_name = rename_to or name
        if rename_to:
            name_el = d.find(qn("w:name"))
            if name_el is None:
                name_el = etree.Element(qn("w:name"))
                d.insert(0, name_el)
            name_el.set(qn("w:val"), rename_to)
        _ensure_styles_part(self.pkg)
        root = self.pkg.root("word/styles.xml")
        root.append(d)
        self.pkg.mark_dirty("word/styles.xml")
        self.tgt_id2name[new_id] = new_name
        self.tgt_name2id[new_name] = new_id
        self.tgt_defs[new_id] = d
        if new_id != sid:
            self.remap[sid] = new_id
        if rename_to:
            self.imported_renamed.append({
                "source_name": name, "source_id": sid,
                "imported_as": rename_to, "style_id": new_id,
                "type": "table",
            })
        else:
            self.cloned.append({"id": new_id, "name": name, "source_id": sid})
        self.cloned_defs.append(d)
        for dep_tag in ("w:basedOn", "w:link", "w:next"):
            dep = d.find(qn(dep_tag))
            if dep is not None:
                self.resolve(dep.get(qn("w:val")))


class _NumberingResolver:
    """Clone source abstractNum/num pairs with fresh non-colliding ids.
    Every source numId gets its OWN target instance (never merged with a
    target list), so restart semantics and lvlOverride/startOverride are
    preserved exactly."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        self.src = src
        self.pkg = pkg
        self.map: dict[str, str] = {}  # src numId -> tgt numId
        self.abs_map: dict[str, str] = {}
        self.cloned_abstracts: list[etree._Element] = []
        self.unresolved: list[str] = []
        self._done: set[str] = set()

    def _tgt_root(self) -> etree._Element:
        _lists._ensure_numbering_part(self.pkg)
        return self.pkg.root("word/numbering.xml")

    def resolve(self, num_id: str) -> None:
        if not num_id or num_id in self._done:
            return
        self._done.add(num_id)
        if num_id == "0":  # numId 0 = "numbering removed" marker; keep as-is
            self.map[num_id] = num_id
            return
        if not self.src.has_part("word/numbering.xml"):
            self.unresolved.append(num_id)
            return
        src_root = self.src.root("word/numbering.xml")
        num = next(
            (
                n
                for n in src_root.findall(qn("w:num"))
                if n.get(qn("w:numId")) == num_id
            ),
            None,
        )
        if num is None:
            self.unresolved.append(num_id)
            return
        tgt_root = self._tgt_root()
        abs_ref = num.find(qn("w:abstractNumId"))
        abs_id = abs_ref.get(qn("w:val")) if abs_ref is not None else None
        new_abs_id = self.abs_map.get(abs_id)
        if abs_id is not None and new_abs_id is None:
            abstract = next(
                (
                    a
                    for a in src_root.findall(qn("w:abstractNum"))
                    if a.get(qn("w:abstractNumId")) == abs_id
                ),
                None,
            )
            if abstract is not None:
                existing_abs = [
                    int(a.get(qn("w:abstractNumId"), "0") or 0)
                    for a in tgt_root.findall(qn("w:abstractNum"))
                ]
                new_abs_id = str(max(existing_abs, default=-1) + 1)
                clone = copy.deepcopy(abstract)
                clone.set(qn("w:abstractNumId"), new_abs_id)
                # w:nsid identifies a list ACROSS documents: two files made
                # from the same template share it, and Word treats two
                # abstractNums with one nsid as one list, which merges their
                # numbering. Drop a colliding nsid so the carried list keeps
                # its own sequence (adversarial review 2026-09-20).
                nsid = clone.find(qn("w:nsid"))
                if nsid is not None:
                    taken = {
                        n.get(qn("w:val"))
                        for a in tgt_root.findall(qn("w:abstractNum"))
                        if (n := a.find(qn("w:nsid"))) is not None
                    }
                    if nsid.get(qn("w:val")) in taken:
                        clone.remove(nsid)
                nums = tgt_root.findall(qn("w:num"))
                if nums:  # schema order: abstractNum before num
                    nums[0].addprevious(clone)
                else:
                    tgt_root.append(clone)
                self.abs_map[abs_id] = new_abs_id
                self.cloned_abstracts.append(clone)
        existing_nums = [
            int(n.get(qn("w:numId"), "0") or 0)
            for n in tgt_root.findall(qn("w:num"))
        ]
        new_num_id = str(max(existing_nums, default=0) + 1)
        num_clone = copy.deepcopy(num)
        num_clone.set(qn("w:numId"), new_num_id)
        clone_abs_ref = num_clone.find(qn("w:abstractNumId"))
        if clone_abs_ref is not None and new_abs_id is not None:
            clone_abs_ref.set(qn("w:val"), new_abs_id)
        tgt_root.append(num_clone)
        self.pkg.mark_dirty("word/numbering.xml")
        self.map[num_id] = new_num_id


def _collect_style_refs(elements) -> set[str]:
    refs: set[str] = set()
    for el in elements:
        for tag in _STYLE_REF_TAGS:
            for node in el.iter(qn(tag)):
                val = node.get(qn("w:val"))
                if val:
                    refs.add(val)
    return refs


def _collect_num_refs(elements) -> set[str]:
    refs: set[str] = set()
    for el in elements:
        for node in el.iter(qn("w:numId")):
            val = node.get(qn("w:val"))
            if val:
                refs.add(val)
    return refs


def _apply_style_remap(elements, remap: dict[str, str]) -> None:
    if not remap:
        return
    for el in elements:
        for tag in _STYLE_REF_TAGS:
            for node in el.iter(qn(tag)):
                val = node.get(qn("w:val"))
                if val in remap:
                    node.set(qn("w:val"), remap[val])


def _apply_num_remap(elements, num_map: dict[str, str], unresolved: set[str]):
    """Remap numIds; numPr pointing at a numbering instance the SOURCE never
    defined is stripped (it rendered plain in the source too) and counted."""
    stripped = 0
    for el in elements:
        for node in list(el.iter(qn("w:numId"))):
            val = node.get(qn("w:val"))
            if val in num_map:
                node.set(qn("w:val"), num_map[val])
            elif val in unresolved:
                numpr = node.getparent()  # w:numPr
                if numpr is not None and _localname(numpr) == "numPr":
                    numpr.getparent().remove(numpr)
                    stripped += 1
    return stripped


# ------------------------------- inherited-formatting reconciliation (source)
#
# The half-single-spaced-dissertation bug (field test, 2026-09-03): with
# formatting="source", paragraphs that carried NO explicit line_spacing or
# space_after relied on their SOURCE file's docDefaults for those values.
# After transplant they resolved against the TARGET's docDefaults instead,
# silently changing the rendered spacing.
#
# The same bug one layer up (field test, 2026-09-20, punchlist #860):
# python-docx sources define their body look on the Normal STYLE, not on
# docDefaults. Styles reconcile BY NAME with the target's definition
# governing, so a carried paragraph that inherited 12pt/double from its own
# Normal silently inherited the target's Normal instead, and the
# docDefaults-only check reported nothing wrong.
#
# So the whole inheritance chain is resolved on both sides — run/paragraph
# direct values, the style basedOn chain as it will resolve AFTER by-name
# reconciliation, then docDefaults — and every tracked property whose
# resolved value would change is baked explicit onto the carried copy.
# Direct values are left alone (they already win and already carry).

# --------------------------------------------------------------------------
# Tracked properties. Round 1 tracked six attributes, so a same-named style
# that differed in bold, italic or colour changed the rendered text with
# nothing baked and nothing reported (adversarial review 2026-09-20, B3).
# The sets below cover the run and paragraph properties that change how
# text looks. Every one of them is DETECTED; the ones with a safe explicit
# form are baked, the rest are reported under not_baked with a reason.
#
# Toggles: presence = on, w:val="0"/"false"/"off" = off, absent = inherit.
_TOGGLE_RPR: dict[str, str] = {  # toggle -> built-in value
    "b": "0", "bCs": "0", "i": "0", "iCs": "0", "caps": "0",
    "smallCaps": "0", "strike": "0", "dstrike": "0", "outline": "0",
    "shadow": "0", "emboss": "0", "imprint": "0", "vanish": "0",
    "webHidden": "0", "specVanish": "0", "oMath": "0", "rtl": "0",
}
# Attribute-carrying run properties: attr -> built-in ("" = no safe
# built-in, so a source that defines nothing is reported, never guessed).
_ATTR_RPR: dict[str, dict[str, str]] = {
    "rFonts": {
        "ascii": "", "hAnsi": "", "eastAsia": "", "cs": "",
        "asciiTheme": "", "hAnsiTheme": "", "eastAsiaTheme": "",
        "cstheme": "", "hint": "",
    },
    "sz": {"val": ""},
    "szCs": {"val": ""},
    "color": {"val": "auto", "themeColor": "", "themeShade": "",
              "themeTint": ""},
    "highlight": {"val": "none"},
    "u": {"val": "none", "color": ""},
    "vertAlign": {"val": "baseline"},
    "position": {"val": "0"},
    "spacing": {"val": "0"},
    "kern": {"val": "0"},
    "w": {"val": "100"},
    "em": {"val": "none"},
    "effect": {"val": "none"},
    "shd": {"val": "clear", "color": "auto", "fill": "auto"},
    "lang": {"val": "", "eastAsia": "", "bidi": ""},
    # A visible box around text used to vanish in silence (round-2 review).
    "bdr": {"val": "none", "sz": "", "space": "", "color": "auto",
            "frame": "", "shadow": ""},
    "fitText": {"val": "", "id": ""},
    "eastAsianLayout": {
        "id": "", "combine": "", "combineBrackets": "", "vert": "",
        "vertCompress": "",
    },
}
_TOGGLE_PPR: dict[str, str] = {
    "keepNext": "0", "keepLines": "0", "pageBreakBefore": "0",
    "contextualSpacing": "0", "suppressLineNumbers": "0",
    "suppressAutoHyphens": "0", "mirrorIndents": "0", "topLinePunct": "0",
    "widowControl": "1", "kinsoku": "1", "wordWrap": "1",
    "overflowPunct": "1", "autoSpaceDE": "1", "autoSpaceDN": "1",
    "adjustRightInd": "1", "snapToGrid": "1",
}
_ATTR_PPR: dict[str, dict[str, str]] = {
    "jc": {"val": "left"},
    "ind": {"left": "0", "start": "0", "right": "0", "end": "0",
            "firstLine": "0", "hanging": "0"},
    "spacing": {"after": "0", "before": "0", "line": "240",
                "lineRule": "auto", "afterAutospacing": "0",
                "beforeAutospacing": "0", "afterLines": "0",
                "beforeLines": "0"},
    "textAlignment": {"val": "auto"},
    "shd": {"val": "clear", "color": "auto", "fill": "auto"},
    "textDirection": {"val": "lrTb"},
    # Detected but never baked: outlineLvl is structure (TOC membership),
    # not appearance, so a difference is reported instead of written.
    "outlineLvl": {"val": ""},
    # Table conditional-formatting marker: meaningless outside a table and
    # carried with the cell anyway, so it is reported, never written.
    "cnfStyle": {"val": ""},
}
_NEVER_BAKE_PPR = {("outlineLvl", "val"), ("cnfStyle", "val")}
# Multi-child properties compared and carried as whole elements (an
# attribute-wise diff is meaningless for them).
_COMPLEX_PPR = ("pBdr", "tabs", "framePr")

_REASON_NO_BUILTIN = (
    "the source defines no value anywhere in its chain, so its rendered "
    "value comes from the theme or from Word's own default and cannot be "
    "written explicitly"
)
_REASON_STRUCTURAL = (
    "the carried heading keeps the TARGET style's outline level, so its "
    "level in the table of contents and its position in the navigation "
    "pane follow the target, not the source. Outline level is structure "
    "rather than appearance: writing the source's level would put two "
    "competing heading hierarchies in one outline. Set it deliberately "
    "with set_paragraph_format(outline_level=...) if the source's level "
    "is the one you want."
)
_REASON_NUMBERED = (
    "the paragraph's indent comes from its numbering level, and a direct "
    "indent would override the list geometry"
)
_REASON_COMPLEX = (
    "the target defines this multi-part property and the source does not; "
    "writing an empty one to cancel it is not safe"
)
_REASON_CONDITIONAL = (
    "a table conditional-formatting marker has no meaning outside the "
    "table it belongs to, and the carried cells keep their own"
)

_RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath",
]


def _ordered_get_or_add(parent: etree._Element, local: str, order: list[str]):
    existing = parent.find(qn(f"w:{local}"))
    if existing is not None:
        return existing
    el = etree.Element(qn(f"w:{local}"))
    my_rank = order.index(local)
    for child in parent:
        name = _localname(child)
        if name in order and order.index(name) > my_rank:
            child.addprevious(el)
            return el
    parent.append(el)
    return el


def _toggle_value(el: etree._Element) -> str:
    return "0" if el.get(qn("w:val")) in ("0", "false", "off") else "1"


def _complex_signature(
    el: etree._Element, themes: tuple | None = None
) -> str:
    """A whole-element property as a comparable, writable string.

    Detached from its package first: inclusive c14n of an element still in
    its tree emits every namespace its ANCESTORS declare, so the same
    w:pBdr serialised out of two packages whose roots declare different
    namespace sets compared unequal and baked a border nobody asked for.
    Only the namespaces the element itself uses survive here.
    """
    clone = copy.deepcopy(el)
    if themes is not None:
        _freeze_theme_colors([clone], themes[0], themes[1])
    clone = etree.fromstring(etree.tostring(clone))
    etree.cleanup_namespaces(clone)
    return etree.tostring(clone, method="c14n").decode()


def _direct_props(
    holder: etree._Element | None, which: str, themes: tuple | None = None,
) -> dict[tuple[str, str], str]:
    """Tracked (element, attribute) -> value pairs explicitly present on a
    pPr or rPr. Toggles report under the pseudo-attribute "val"; complex
    properties report their canonical serialization under "_xml".

    `themes` is (this package's theme, the other package's theme). When
    given, theme references inside a complex property are resolved against
    this package's theme wherever the two disagree about the slot, so two
    byte-identical w:pBdr elements that the two themes paint different
    colours no longer compare equal (round-3 review, m10). Slots the two
    themes agree on are left as live references, so an in-template merge
    still follows the merged document's theme.
    """
    props: dict[tuple[str, str], str] = {}
    if holder is None:
        return props
    toggles = _TOGGLE_PPR if which == "pPr" else _TOGGLE_RPR
    attrs_table = _ATTR_PPR if which == "pPr" else _ATTR_RPR
    for name in toggles:
        el = holder.find(qn(f"w:{name}"))
        if el is not None:
            props[(name, "val")] = _toggle_value(el)
    for elem_name, attrs in attrs_table.items():
        el = holder.find(qn(f"w:{elem_name}"))
        if el is None:
            continue
        for a in attrs:
            v = el.get(qn(f"w:{a}"))
            if v is not None:
                props[(elem_name, a)] = v
        for base, trio in _THEME_ASPECTS.get(elem_name, {}).items():
            packed = _pack_theme_ref(el, base, trio)
            if packed is not None:
                props[(elem_name, _THEME_KEY + base)] = packed
    if which == "pPr":
        for name in _COMPLEX_PPR:
            el = holder.find(qn(f"w:{name}"))
            if el is not None:
                props[(name, "_xml")] = _complex_signature(el, themes)
    return props


def _docdefaults_props(
    pkg: DocxPackage, which: str, themes: tuple | None = None,
) -> dict[tuple[str, str], str]:
    """Tracked property values from styles.xml docDefaults."""
    if not pkg.has_part("word/styles.xml"):
        return {}
    dd = pkg.root("word/styles.xml").find(qn("w:docDefaults"))
    if dd is None:
        return {}
    if which == "pPr":
        holder = dd.find(f"{qn('w:pPrDefault')}/{qn('w:pPr')}")
    else:
        holder = dd.find(f"{qn('w:rPrDefault')}/{qn('w:rPr')}")
    return _direct_props(holder, which, themes)


def _builtin(key: tuple[str, str], which: str) -> str:
    """Word's value when no layer defines the property; "" when there is
    no safe one."""
    elem, attr = key
    toggles = _TOGGLE_PPR if which == "pPr" else _TOGGLE_RPR
    if elem in toggles:
        return toggles[elem]
    table = _ATTR_PPR if which == "pPr" else _ATTR_RPR
    return table.get(elem, {}).get(attr, "")


# ------------------------------------------------------------- theme colours
#
# Two packages can carry byte-identical <w:color w:val="4472C4"
# w:themeColor="accent1"/> and still render different colours, because the
# colour lives in word/theme/theme1.xml and the theme is never transplanted:
# the carried text silently takes the TARGET's accent (adversarial review
# round 2, M6). Two chapters from two Word templates is the ordinary case.
#
# So the two colour schemes are compared slot by slot, and wherever carried
# content references a slot that resolves differently, the SOURCE's resolved
# RGB is baked into w:val and the themeColor/themeTint/themeShade attributes
# are dropped, so Word cannot re-resolve it against its own theme.

_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
_THEME_SLOTS = (
    "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
    "accent4", "accent5", "accent6", "hlink", "folHlink",
)
# ST_ThemeColor (w:themeColor) -> clrScheme slot. text1/text2/background1/
# background2 go through settings.xml's w:clrSchemeMapping first.
_THEME_NAME_TO_SLOT = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followedHyperlink": "folHlink",
}
_MAPPED_NAMES = {
    "text1": ("t1", "dk1"), "text2": ("t2", "dk2"),
    "background1": ("bg1", "lt1"), "background2": ("bg2", "lt2"),
}
_SYS_COLOR_FALLBACK = {"windowText": "000000", "window": "FFFFFF"}

# Colour-bearing aspects: element -> {base attribute: (theme, tint, shade)}.
# Each one becomes its own tracked pseudo-property.
_THEME_ASPECTS: dict[str, dict[str, tuple[str, str, str]]] = {
    "color": {"val": ("themeColor", "themeTint", "themeShade")},
    "u": {"color": ("themeColor", "themeTint", "themeShade")},
    "bdr": {"color": ("themeColor", "themeTint", "themeShade")},
    "shd": {
        "fill": ("themeFill", "themeFillTint", "themeFillShade"),
        "color": ("themeColor", "themeTint", "themeShade"),
    },
}
_THEME_KEY = "_theme:"


def _apply_tint_shade(hex_rgb: str, tint: str | None, shade: str | None) -> str:
    """Word's themeTint/themeShade, applied to HSL LUMINANCE.

    A per-channel RGB blend toward white is the obvious reading and it is
    wrong: Word converts the theme colour to HSL, scales the luminance and
    converts back, so the hue and the saturation survive the tint. The
    difference is invisible on a fully saturated hue and large everywhere
    else (adversarial review round 3, M7). Measured against Word's own
    rendering: accent1 = C00000 with themeFillTint="99" renders as
    (255, 64, 64), not the (217, 102, 102) an RGB blend produces.

    With the hex byte read as f = value / 255:

        themeTint   L' = L * f + (1 - f)     lighter as f falls
        themeShade  L' = L * f               darker as f falls

    ECMA-376 is self-contradictory here: the w:color subclause (17.3.2.6)
    describes a per-channel blend while w:u, w:bdr and w:shd describe the
    HSL conversion. [MS-OI29500] 2.1.72 names that as a defect and states
    that Word uses the HSL luminance algorithm for w:color too, which is
    what the measurements show.

    When both attributes are present Word renders the TINT and ignores the
    shade: [MS-OI29500] 2.1.72 and 2.1.144 both say "the standard does not
    state which setting is applied when both the themeShade and the
    themeTint attributes are present. Word applies the themeTint setting",
    and 12 of 12 measured fixtures agree.

    Rounding is round-half-away-from-zero on the way back to 8-bit RGB,
    which makes f = 1.0 an exact identity; every measured value agrees
    with Word's render to within the 2-unit tolerance the PDF rasteriser
    itself carries on an untinted control.
    """
    try:
        rgb = tuple(int(hex_rgb[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except (ValueError, IndexError):
        return hex_rgb
    attr = tint or shade
    if not attr:
        return hex_rgb.upper()
    try:
        f = int(attr, 16) / 255.0
    except ValueError:
        return hex_rgb
    h, lum, sat = colorsys.rgb_to_hls(*rgb)
    lum = lum * f + (1.0 - f) if tint else lum * f
    out = colorsys.hls_to_rgb(h, max(0.0, min(1.0, lum)), sat)
    return "".join(
        f"{max(0, min(255, int(math.floor(c * 255.0 + 0.5)))):02X}"
        for c in out
    )


class _ThemeColors:
    """One package's theme colour scheme, with its settings.xml mapping."""

    def __init__(self, pkg: DocxPackage):
        self.slots: dict[str, str] = {}
        self.mapping: dict[str, str] = {}
        self.part: str | None = None
        for name in pkg.part_names():
            if re.fullmatch(r"word/theme/theme\d+\.xml", name):
                self.part = name
                break
        if self.part:
            scheme = pkg.root(self.part).find(
                f"{{{_A}}}themeElements/{{{_A}}}clrScheme"
            )
            if scheme is not None:
                for slot in _THEME_SLOTS:
                    el = scheme.find(f"{{{_A}}}{slot}")
                    if el is None:
                        continue
                    srgb = el.find(f"{{{_A}}}srgbClr")
                    if srgb is not None and srgb.get("val"):
                        self.slots[slot] = srgb.get("val").upper()
                        continue
                    sys_clr = el.find(f"{{{_A}}}sysClr")
                    if sys_clr is not None:
                        last = sys_clr.get("lastClr") or _SYS_COLOR_FALLBACK.get(
                            sys_clr.get("val") or "", ""
                        )
                        if last:
                            self.slots[slot] = last.upper()
        if pkg.has_part("word/settings.xml"):
            m = pkg.root("word/settings.xml").find(qn("w:clrSchemeMapping"))
            if m is not None:
                for key in ("t1", "t2", "bg1", "bg2"):
                    v = m.get(qn(f"w:{key}"))
                    if v:
                        self.mapping[key] = v

    @property
    def defined(self) -> bool:
        return bool(self.slots)

    def slot_for(self, theme_color: str) -> str | None:
        if theme_color in _MAPPED_NAMES:
            key, default = _MAPPED_NAMES[theme_color]
            mapped = self.mapping.get(key)
            return _THEME_NAME_TO_SLOT.get(mapped, default) if mapped else default
        return _THEME_NAME_TO_SLOT.get(theme_color)

    def resolve(
        self, theme_color: str, tint: str | None, shade: str | None
    ) -> str | None:
        slot = self.slot_for(theme_color)
        if slot is None:
            return None
        base = self.slots.get(slot)
        if not base:
            return None
        return _apply_tint_shade(base, tint, shade)

    def differing_slots(self, other: "_ThemeColors") -> list[str]:
        if not self.defined or not other.defined:
            return []
        return sorted(
            slot for slot in _THEME_SLOTS
            if self.slots.get(slot) != other.slots.get(slot)
        )

    def differing_mappings(self, other: "_ThemeColors") -> list[str]:
        """The ST_ThemeColor names that the two packages send to DIFFERENT
        slots. text1/text2/background1/background2 go through
        settings.xml's w:clrSchemeMapping, which does not travel either, so
        a dark-mode source and a light-mode target resolve text1 to
        opposite ends of the same colour scheme. Reported separately
        because the two colour schemes can be identical while the rendered
        colour still changes (round-3 review, m11)."""
        if not self.defined or not other.defined:
            return []
        return sorted(
            name for name in _MAPPED_NAMES
            if self.slot_for(name) != other.slot_for(name)
        )


# The two theme trios, by the attribute that names the slot. Every
# colour-bearing WordprocessingML element uses one or both of them, so a
# descent keyed on these finds the ones no aspect table lists: the CT_Border
# children of w:pBdr, w:tblBorders, w:tcBorders and w:pgBorders, and
# w:background (round-3 review, m10).
_THEME_TRIOS = (
    ("themeColor", "themeTint", "themeShade"),
    ("themeFill", "themeFillTint", "themeFillShade"),
)


def _theme_base_attr(el: etree._Element, theme_attr: str) -> str:
    """Which attribute the theme reference resolves INTO. CT_Color keeps
    its colour in w:val; every other colour-bearing type keeps it in
    w:color, and a fill always lands in w:fill."""
    if theme_attr == "themeFill":
        return "fill"
    return "val" if el.tag == qn("w:color") else "color"


def _pack_theme_ref(el: etree._Element, base: str, trio: tuple) -> str | None:
    """The raw colour reference on one element, as a packed string, or None
    when it names no theme colour (a plain hex needs no resolution)."""
    theme = el.get(qn(f"w:{trio[0]}"))
    if not theme or theme == "none":
        return None
    return "|".join(
        el.get(qn(f"w:{a}")) or ""
        for a in (base, trio[0], trio[1], trio[2])
    )


def _direct_skip(direct: dict) -> set:
    """Keys a paragraph or run already carries directly. A colour aspect
    counts as carried when EITHER its plain attribute or its theme
    reference is present, so the two halves never fight each other."""
    keys = set(direct)
    for elem, aspects in _THEME_ASPECTS.items():
        for base in aspects:
            if (elem, base) in keys or (elem, _THEME_KEY + base) in keys:
                keys.add((elem, base))
                keys.add((elem, _THEME_KEY + base))
    return keys


def _freeze_theme_colors(
    elements, src_theme: _ThemeColors, tgt_theme: _ThemeColors
) -> int:
    """Resolve theme colour references in CARRIED xml against the source
    theme and write the result as a plain value, dropping the theme
    attributes so Word cannot re-resolve them against its own theme.

    The theme part is never transplanted, so every reference in carried
    content (direct formatting, and the definitions of cloned or imported
    styles) would otherwise take the target's colours (round-2 review, M6).
    Only references whose slot actually resolves differently are touched.

    The descent is keyed on the theme ATTRIBUTES rather than on a list of
    element names, so it reaches the colour carriers no aspect table names:
    a themed bottom rule under a heading (w:pBdr/w:bottom), the borders of
    a carried table (w:tblBorders, w:tcBorders), page borders, and
    w:background. Those used to pass through untouched and silently
    re-theme (round-3 review, m10).
    """
    if not src_theme.defined or not tgt_theme.defined:
        return 0
    frozen = 0
    for root in elements:
        for el in root.iter():
            if not isinstance(el.tag, str):
                continue
            for trio in _THEME_TRIOS:
                base = _theme_base_attr(el, trio[0])
                packed = _pack_theme_ref(el, base, trio)
                if packed is None:
                    continue
                src_hex = _resolve_theme_ref(packed, src_theme)
                tgt_hex = _resolve_theme_ref(packed, tgt_theme)
                if src_hex is None or src_hex == tgt_hex:
                    continue
                el.set(qn(f"w:{base}"), src_hex)
                for gone in trio:
                    el.attrib.pop(qn(f"w:{gone}"), None)
                frozen += 1
    return frozen


def _resolve_theme_ref(packed: str | None, theme: _ThemeColors) -> str | None:
    """A packed reference resolved against one package's theme."""
    if not packed:
        return None
    val, name, tint, shade = (packed.split("|") + ["", "", "", ""])[:4]
    hit = theme.resolve(name, tint or None, shade or None)
    if hit:
        return hit
    return (val or "").upper() or None


# ----------------------------------------------------------- numbering layer
#
# A list paragraph's indent comes from numbering.xml (abstractNum -> lvl ->
# pPr -> ind), NOT from its style, and a direct w:ind overrides it. Round 1
# resolved the style chain only, so it baked the style's indent onto list
# paragraphs and pushed their bullets out into the left margin (adversarial
# review 2026-09-20, B2). Numbering is transplanted definition-for-
# definition, so the layer resolves identically on both sides and its
# properties never differ; resolving it is what keeps them out of the bake.


class _NumberingIndex:
    """numId/ilvl -> the level's pPr and rPr, honouring w:lvlOverride."""

    def __init__(self, pkg: DocxPackage):
        self.abstract: dict[str, dict[str, etree._Element]] = {}
        self.num_to_abstract: dict[str, str] = {}
        self.overrides: dict[tuple[str, str], etree._Element] = {}
        if not pkg.has_part("word/numbering.xml"):
            return
        root = pkg.root("word/numbering.xml")
        for a in root.findall(qn("w:abstractNum")):
            aid = a.get(qn("w:abstractNumId"))
            if aid is None:
                continue
            self.abstract[aid] = {
                lvl.get(qn("w:ilvl")): lvl for lvl in a.findall(qn("w:lvl"))
            }
        for n in root.findall(qn("w:num")):
            nid = n.get(qn("w:numId"))
            ref = n.find(qn("w:abstractNumId"))
            if nid is None or ref is None:
                continue
            self.num_to_abstract[nid] = ref.get(qn("w:val"))
            for ov in n.findall(qn("w:lvlOverride")):
                lvl = ov.find(qn("w:lvl"))
                if lvl is not None:
                    self.overrides[(nid, ov.get(qn("w:ilvl")))] = lvl

    def level(self, num_id: str | None, ilvl: str) -> etree._Element | None:
        if not num_id or num_id == "0":
            return None
        hit = self.overrides.get((num_id, ilvl))
        if hit is not None:
            return hit
        aid = self.num_to_abstract.get(num_id)
        if aid is None:
            return None
        return self.abstract.get(aid, {}).get(ilvl)

    def props(self, num_id: str | None, ilvl: str, which: str) -> dict:
        lvl = self.level(num_id, ilvl)
        if lvl is None:
            return {}
        return _direct_props(lvl.find(qn(f"w:{which}")), which)


def _style_index(pkg: DocxPackage):
    """styleId -> element, styleId -> basedOn, default paragraph styleId."""
    els: dict[str, etree._Element] = {}
    based: dict[str, str] = {}
    default_para: str | None = None
    if pkg.has_part("word/styles.xml"):
        for s in pkg.root("word/styles.xml").findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            if not sid:
                continue
            els[sid] = s
            b = s.find(qn("w:basedOn"))
            if b is not None and b.get(qn("w:val")):
                based[sid] = b.get(qn("w:val"))
            if (
                s.get(qn("w:type")) == "paragraph"
                and s.get(qn("w:default")) in ("1", "true", "on")
            ):
                default_para = sid
    if default_para is None and "Normal" in els:
        # No style is marked w:default: Word falls back to Normal, and so
        # must the resolver, or a paragraph with no pStyle resolves against
        # docDefaults alone.
        default_para = "Normal"
    return els, based, default_para


class _DefaultsBaker:
    """Bakes inherited formatting onto carried copies (formatting='source'
    only) wherever the source's and the post-transplant target's resolved
    values differ. Layers, in Word's own precedence: docDefaults, the style
    basedOn chain (as it will resolve AFTER by-name reconciliation), the
    numbering level, then direct formatting, which is never overwritten.
    See the section comment above."""

    def __init__(self, src: DocxPackage, pkg: DocxPackage):
        # Built first: the complex-property comparison needs both themes to
        # tell a themed border apart from an identical one (review m10).
        src_theme, tgt_theme = _ThemeColors(src), _ThemeColors(pkg)
        self.src_themes = (src_theme, tgt_theme)
        self.tgt_themes = (tgt_theme, src_theme)
        self.src_dd = {
            "pPr": _docdefaults_props(src, "pPr", self.src_themes),
            "rPr": _docdefaults_props(src, "rPr", self.src_themes),
        }
        self.tgt_dd = {
            "pPr": _docdefaults_props(pkg, "pPr", self.tgt_themes),
            "rPr": _docdefaults_props(pkg, "rPr", self.tgt_themes),
        }
        self.src_el, self.src_based, self.default_para_style = _style_index(src)
        self.tgt_el, self.tgt_based, self.tgt_default_para = _style_index(pkg)
        self.src_id2name, _ = _style_maps(src)
        _tgt_id2name, self.tgt_name2id = _style_maps(pkg)
        # Numbering travels definition-for-definition, so the SOURCE's
        # levels describe both sides of the diff.
        self.numbering = _NumberingIndex(src)
        self.src_theme, self.tgt_theme = src_theme, tgt_theme
        self.theme_slots_differ = self.src_theme.differing_slots(self.tgt_theme)
        self.theme_colors_baked = 0
        self._src_chain: dict[tuple[str, str], dict] = {}
        # Two DIFFERENT mappings, so two memos: _tgt_own is "values along
        # the TARGET's own basedOn chain", _tgt_chain is "values a SOURCE
        # style id resolves to after transplant". Sharing one dict made the
        # post-transplant cycle guard shadow the target's own chain
        # whenever the two files use the same styleId, which is the normal
        # case for Heading1.
        self._tgt_own: dict[tuple[str, str], dict] = {}
        self._tgt_chain: dict[tuple[str, str], dict] = {}
        self._style_numpr: dict[str, tuple[str | None, str]] = {}
        self._para_diff: dict[tuple, tuple[dict, set]] = {}
        self._run_diff: dict[tuple, dict] = {}
        self.paragraphs_baked = 0
        self.runs_baked = 0
        self.marks_baked = 0
        # Reporting: what actually differed, which shared style names were
        # responsible, and what could not be written explicitly.
        self.differing: set[tuple[str, tuple[str, str]]] = set()
        self.style_diffs: dict[str, set[tuple[str, tuple[str, str]]]] = {}
        self.not_baked: dict[tuple[str, tuple[str, str]], str] = {}

    # ---- chain resolution

    def _chain_props(
        self,
        sid: str | None,
        which: str,
        els: dict[str, etree._Element],
        based: dict[str, str],
        memo: dict,
        themes: tuple | None = None,
    ) -> dict[tuple[str, str], str]:
        """Values DEFINED along a style's basedOn chain (ancestors first, so
        the style's own values win). docDefaults are NOT included."""
        if not sid:
            return {}
        hit = memo.get((sid, which))
        if hit is not None:
            return hit
        chain: list[str] = []
        cur: str | None = sid
        seen: set[str] = set()
        while cur and cur not in seen and cur in els:
            seen.add(cur)
            chain.append(cur)
            cur = based.get(cur)
        props: dict[tuple[str, str], str] = {}
        for s in reversed(chain):
            props.update(
                _direct_props(els[s].find(qn(f"w:{which}")), which, themes)
            )
        memo[(sid, which)] = props
        return props

    def _src_style_props(self, sid: str | None, which: str) -> dict:
        return self._chain_props(
            sid, which, self.src_el, self.src_based, self._src_chain,
            self.src_themes,
        )

    def _tgt_style_props(self, sid: str | None, which: str) -> dict:
        """Values a SOURCE style id will resolve to after transplant.

        A name match hands the TARGET's own chain over (the documented
        by-name contract). A source style whose name does NOT match is
        cloned, so its own definition survives and its basedOn resolves the
        same way, recursively. An id with no source definition at all falls
        through to the target's style of that id, which is what the
        reference will land on (adversarial review, m2: the id fallback
        used to fire for name-unmatched styles too, and reported
        differences that did not exist)."""
        if not sid:
            return {}
        hit = self._tgt_chain.get((sid, which))
        if hit is not None:
            return hit
        self._tgt_chain[(sid, which)] = {}  # cycle guard
        name = self.src_id2name.get(sid)
        if name is None:
            tgt_id = sid if sid in self.tgt_el else None
        else:
            tgt_id = self.tgt_name2id.get(name)
        if tgt_id is not None:
            props = self._chain_props(
                tgt_id, which, self.tgt_el, self.tgt_based, self._tgt_own,
                self.tgt_themes,
            )
        else:
            # A name-unmatched source style is CLONED, and the clone's own
            # theme references are frozen against the source theme, so the
            # source's resolution is what the carried content will see.
            props = dict(self._tgt_style_props(self.src_based.get(sid), which))
            el = self.src_el.get(sid)
            if el is not None:
                props.update(
                    _direct_props(
                        el.find(qn(f"w:{which}")), which, self.src_themes
                    )
                )
        self._tgt_chain[(sid, which)] = props
        return props

    # ---- numbering

    def _style_numbering(self, sid: str | None) -> tuple[str | None, str]:
        """(numId, ilvl) a style chain supplies, nearest definition first."""
        if not sid:
            return None, "0"
        hit = self._style_numpr.get(sid)
        if hit is not None:
            return hit
        num_id: str | None = None
        ilvl = "0"
        cur: str | None = sid
        seen: set[str] = set()
        while cur and cur not in seen and cur in self.src_el:
            seen.add(cur)
            numpr = self.src_el[cur].find(f"{qn('w:pPr')}/{qn('w:numPr')}")
            if numpr is not None:
                nid = numpr.find(qn("w:numId"))
                lvl = numpr.find(qn("w:ilvl"))
                if num_id is None and nid is not None:
                    num_id = nid.get(qn("w:val"))
                if lvl is not None and lvl.get(qn("w:val")):
                    ilvl = lvl.get(qn("w:val"))
                if num_id is not None:
                    break
            cur = self.src_based.get(cur)
        self._style_numpr[sid] = (num_id, ilvl)
        return num_id, ilvl

    def _paragraph_numbering(
        self, ppr: etree._Element | None, pstyle: str | None
    ) -> tuple[str | None, str]:
        numpr = ppr.find(qn("w:numPr")) if ppr is not None else None
        if numpr is not None:
            nid = numpr.find(qn("w:numId"))
            lvl = numpr.find(qn("w:ilvl"))
            num_id = nid.get(qn("w:val")) if nid is not None else None
            ilvl = lvl.get(qn("w:val")) if lvl is not None else "0"
            if num_id is not None:
                return num_id, ilvl or "0"
        return self._style_numbering(pstyle or self.default_para_style)

    # ---- diffs

    def _paragraph_diff(
        self, pstyle: str | None, num_id: str | None, ilvl: str
    ) -> tuple[dict[tuple[str, str], str], set]:
        """(properties to bake, keys the numbering level supplies) for a
        paragraph carrying this style and numbering."""
        memo_key = (pstyle, num_id, ilvl)
        hit = self._para_diff.get(memo_key)
        if hit is not None:
            return hit
        num_props = self.numbering.props(num_id, ilvl, "pPr")
        src_eff = dict(self.src_dd["pPr"])
        src_eff.update(
            self._src_style_props(pstyle or self.default_para_style, "pPr")
        )
        tgt_eff = dict(self.tgt_dd["pPr"])
        if pstyle:
            tgt_eff.update(self._tgt_style_props(pstyle, "pPr"))
        else:
            tgt_eff.update(
                self._chain_props(
                    self.tgt_default_para, "pPr", self.tgt_el,
                    self.tgt_based, self._tgt_own, self.tgt_themes,
                )
            )
        # The numbering layer sits above both style chains and travels
        # unchanged, so it resolves identically on both sides.
        src_eff.update(num_props)
        tgt_eff.update(num_props)
        out = self._diff(
            src_eff, tgt_eff, "pPr", pstyle or self.default_para_style
        )
        # Belt and braces for B2: never write an indent onto a paragraph
        # whose list geometry supplies one, even if some other layer made
        # the attribute differ.
        if any(e == "ind" for (e, _a) in num_props):
            for key in [k for k in out if k[0] == "ind"]:
                out.pop(key)
                self.not_baked.setdefault(("pPr", key), _REASON_NUMBERED)
        self._para_diff[memo_key] = (out, set(num_props))
        return self._para_diff[memo_key]

    def _run_diff_for(
        self, pstyle: str | None, rstyle: str | None
    ) -> dict[tuple[str, str], str]:
        memo_key = (pstyle, rstyle)
        hit = self._run_diff.get(memo_key)
        if hit is not None:
            return hit
        src_eff = dict(self.src_dd["rPr"])
        src_eff.update(
            self._src_style_props(pstyle or self.default_para_style, "rPr")
        )
        src_eff.update(self._src_style_props(rstyle, "rPr"))
        tgt_eff = dict(self.tgt_dd["rPr"])
        if pstyle:
            tgt_eff.update(self._tgt_style_props(pstyle, "rPr"))
        else:
            tgt_eff.update(
                self._chain_props(
                    self.tgt_default_para, "rPr", self.tgt_el,
                    self.tgt_based, self._tgt_own, self.tgt_themes,
                )
            )
        tgt_eff.update(self._tgt_style_props(rstyle, "rPr"))
        out = self._diff(
            src_eff, tgt_eff, "rPr",
            pstyle or self.default_para_style, rstyle,
        )
        self._run_diff[memo_key] = out
        return out

    def _diff(
        self,
        src_eff: dict,
        tgt_eff: dict,
        which: str,
        *styles: str | None,
    ) -> dict[tuple[str, str], str]:
        """Keys whose resolved value changes, mapped to the value to write.
        A key the source never defines is baked from Word's built-in where
        there is a safe one; where there is not, it is recorded in
        not_baked WITH a reason rather than dropped in silence."""
        out: dict[tuple[str, str], str] = {}
        # An aspect with a theme reference on either side is decided by the
        # RESOLVED colour, so its plain attribute is not compared twice.
        themed = {
            (k[0], k[1][len(_THEME_KEY):])
            for k in set(src_eff) | set(tgt_eff)
            if k[1].startswith(_THEME_KEY)
        }
        for key in set(src_eff) | set(tgt_eff):
            sv, tv = src_eff.get(key), tgt_eff.get(key)
            if key in themed:
                continue
            if key[1].startswith(_THEME_KEY):
                base = (key[0], key[1][len(_THEME_KEY):])
                sv_hex = (
                    _resolve_theme_ref(sv, self.src_theme) if sv
                    else (src_eff.get(base) or "").upper() or None
                )
                tv_hex = (
                    _resolve_theme_ref(tv, self.tgt_theme) if tv
                    else (tgt_eff.get(base) or "").upper() or None
                )
                if sv_hex is None or sv_hex == tv_hex:
                    continue
                self.differing.add((which, key))
                self._attribute(key, which, styles)
                out[key] = sv_hex
                self.theme_colors_baked += 1
                continue
            if sv == tv:
                continue
            self.differing.add((which, key))
            self._attribute(key, which, styles)
            if which == "pPr" and key in _NEVER_BAKE_PPR:
                self.not_baked.setdefault(
                    (which, key),
                    _REASON_CONDITIONAL if key[0] == "cnfStyle"
                    else _REASON_STRUCTURAL,
                )
                continue
            if key[1] == "_xml":
                if sv is None:
                    self.not_baked.setdefault((which, key), _REASON_COMPLEX)
                    continue
                out[key] = sv
                continue
            value = sv if sv is not None else _builtin(key, which)
            if not value:
                self.not_baked.setdefault((which, key), _REASON_NO_BUILTIN)
                continue
            out[key] = value
        return out

    def _attribute(self, key, which: str, styles) -> None:
        """Name the shared style responsible for a difference, but only when
        that style's chain actually defines the property (otherwise it came
        from docDefaults and naming a style would mislead)."""
        for sid in styles:
            name = self.src_id2name.get(sid) if sid else None
            if not name or not self.tgt_name2id.get(name):
                continue
            if key in self._src_style_props(
                sid, which
            ) or key in self._tgt_style_props(sid, which):
                self.style_diffs.setdefault(name, set()).add((which, key))

    # ---- baking

    def _write(
        self, holder: etree._Element, key: tuple[str, str], value: str,
        which: str, source_el: etree._Element | None = None,
    ) -> None:
        order = _PPR_BAKE_ORDER if which == "pPr" else _RPR_ORDER
        toggles = _TOGGLE_PPR if which == "pPr" else _TOGGLE_RPR
        elem, attr = key
        if attr.startswith(_THEME_KEY):
            base = attr[len(_THEME_KEY):]
            el = _ordered_get_or_add(holder, elem, order)
            el.set(qn(f"w:{base}"), value)
            for gone in _THEME_ASPECTS.get(elem, {}).get(base, ()):
                el.attrib.pop(qn(f"w:{gone}"), None)
            return
        if attr == "_xml":
            existing = holder.find(qn(f"w:{elem}"))
            if existing is not None:
                holder.remove(existing)
            new = etree.fromstring(value.encode())
            ref = _ordered_get_or_add(holder, elem, order)
            holder.replace(ref, new)
            return
        el = _ordered_get_or_add(holder, elem, order)
        if elem in toggles:
            if value == "0":
                el.set(qn("w:val"), "0")
            else:
                el.attrib.pop(qn("w:val"), None)
            return
        el.set(qn(f"w:{attr}"), value)

    def _bake_run_props(
        self, holder_owner: etree._Element, rpr: etree._Element | None,
        pstyle: str | None, insert_at: int,
    ) -> bool:
        """Bake run properties onto one rPr (a run's, or a paragraph
        mark's). Returns True when anything was written."""
        rstyle = None
        if rpr is not None:
            rs = rpr.find(qn("w:rStyle"))
            if rs is not None:
                rstyle = rs.get(qn("w:val"))
        run_bake = self._run_diff_for(pstyle, rstyle)
        if not run_bake:
            return False
        direct = _direct_skip(_direct_props(rpr, "rPr"))
        changed = False
        for key in sorted(run_bake):
            if key in direct:
                continue  # direct formatting already carries it
            if rpr is None:
                rpr = etree.Element(qn("w:rPr"))
                holder_owner.insert(insert_at, rpr)
            self._write(rpr, key, run_bake[key], "rPr")
            changed = True
        return changed

    def bake_paragraph(self, p: etree._Element) -> None:
        ppr = p.find(qn("w:pPr"))
        pstyle = None
        if ppr is not None:
            ps = ppr.find(qn("w:pStyle"))
            if ps is not None:
                pstyle = ps.get(qn("w:val"))
        num_id, ilvl = self._paragraph_numbering(ppr, pstyle)
        # ---- paragraph properties
        to_bake, _num_keys = self._paragraph_diff(pstyle, num_id, ilvl)
        if to_bake:
            direct = _direct_skip(_direct_props(ppr, "pPr"))
            changed = False
            for key in sorted(to_bake):
                if key in direct:
                    continue  # direct formatting already carries it
                if ppr is None:
                    ppr = etree.Element(qn("w:pPr"))
                    p.insert(0, ppr)
                self._write(ppr, key, to_bake[key], "pPr")
                changed = True
            if changed:
                self.paragraphs_baked += 1
        # ---- run properties, runs first
        for r in p.iter(qn("w:r")):
            if r.getparent() is ppr:
                continue  # the paragraph-mark rPr holder is not a run
            if self._bake_run_props(r, r.find(qn("w:rPr")), pstyle, 0):
                self.runs_baked += 1
        # ---- the paragraph MARK's own run properties: an empty paragraph
        # is rendered entirely by its mark, so a spacer took the target's
        # line height (adversarial review, M4).
        if ppr is None and self._run_diff_for(pstyle, None):
            ppr = etree.Element(qn("w:pPr"))
            p.insert(0, ppr)
        if ppr is not None:
            mark_rpr = ppr.find(qn("w:rPr"))
            owner_index = len(ppr)
            if mark_rpr is None:
                mark_rpr = _ordered_get_or_add(ppr, "rPr", _PPR_BAKE_ORDER)
            if self._bake_run_props(ppr, mark_rpr, pstyle, owner_index):
                self.marks_baked += 1
            if len(mark_rpr) == 0:
                ppr.remove(mark_rpr)
            if len(ppr) == 0:
                p.remove(ppr)

    # ---- reporting

    @property
    def reportable(self) -> bool:
        return bool(self.differing)

    @staticmethod
    def _fmt(items) -> list[str]:
        return sorted(
            f"{which}.{e}" if a == "_xml"
            else f"{which}.{e}.{a[len(_THEME_KEY):]}(theme)"
            if a.startswith(_THEME_KEY)
            else f"{which}.{e}.{a}"
            for (which, (e, a)) in items
        )

    def report(self) -> dict:
        baked = self.paragraphs_baked + self.runs_baked + self.marks_baked
        # The note is a function of what actually happened: round 1 claimed
        # the source appearance was kept even when nothing was baked and
        # differences were left unresolved (adversarial review, B3, m3).
        if self.not_baked and baked:
            note = (
                "the source and target resolve these properties differently "
                "(document defaults, a style of the same name, or both). "
                "Most were given explicit values on the carried content so "
                "it keeps the source appearance; the ones under not_baked "
                "were NOT, and will render the target's way."
            )
        elif self.not_baked:
            note = (
                "the source and target resolve these properties differently, "
                "and NONE of them could be written explicitly (see "
                "not_baked); the carried content will render the target's "
                "way for them."
            )
        else:
            note = (
                "the source and target resolve these properties differently "
                "(document defaults, a style of the same name, or both); "
                "carried paragraphs, runs and paragraph marks that inherited "
                "them were given explicit values so they keep the source "
                "appearance. The target's own content is untouched."
            )
        out = {
            "differ": True,
            "differing_properties": self._fmt(self.differing),
            **(
                {
                    "theme_colors_baked": self.theme_colors_baked,
                    "theme_slots_differing": self.theme_slots_differ,
                }
                if self.theme_colors_baked or self.theme_slots_differ
                else {}
            ),
            "paragraphs_baked": self.paragraphs_baked,
            "runs_baked": self.runs_baked,
            "paragraph_marks_baked": self.marks_baked,
            "note": note,
        }
        if self.style_diffs:
            out["styles_reconciled"] = [
                {
                    "style_name": name,
                    "differing_properties": self._fmt(keys),
                }
                for name, keys in sorted(self.style_diffs.items())
            ]
        if self.not_baked:
            out["not_baked"] = self._fmt(self.not_baked)
            reasons = sorted({r for r in self.not_baked.values()})
            out["not_baked_reasons"] = reasons
        return out


# pPr insertion order for baked elements (CT_PPr child sequence, abridged).
_PPR_BAKE_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]


# ------------------------------------------------- formatting strip (modes)

_FORMATTING_MODES = ("source", "merge", "destination")

# `formatting="merge"` (Word's Merge Formatting paste mode): semantic
# emphasis stays direct (bold, italic, underline, strike, sub/superscript,
# highlight, and anything else not listed below), while direct font-family/
# size/color and character-spacing run overrides plus paragraph spacing/
# line-spacing/indent overrides are stripped so the TARGET's styles govern
# the look of the carried text.
_MERGE_STRIP_RPR = frozenset({"rFonts", "sz", "szCs", "color", "spacing"})
_MERGE_STRIP_PPR = frozenset({"spacing", "ind"})

# `formatting="destination"`: strip ALL direct rPr/pPr formatting except
# what is structurally required or non-visual. Kept on runs: style refs,
# proofing language / noProof, RTL and complex-script structure, hidden-text
# flags (stripping those would REVEAL content, not restyle it), math, and
# revision marks. Kept on paragraphs: style ref, numbering, outline level,
# tab stops, text direction, table-conditional/div plumbing, the (recursed)
# paragraph-mark rPr, and revision records.
_DEST_KEEP_RPR = frozenset(
    {
        "rStyle", "lang", "noProof", "rtl", "cs",
        "vanish", "webHidden", "specVanish", "oMath",
        "ins", "del", "rPrChange",
    }
)
_DEST_KEEP_PPR = frozenset(
    {
        "pStyle", "numPr", "outlineLvl", "tabs",
        "bidi", "textDirection", "divId", "cnfStyle",
        "rPr", "sectPr", "pPrChange",
    }
)


def _strip_rpr_direct(rpr: etree._Element, mode: str) -> bool:
    changed = False
    for child in list(rpr):
        name = _localname(child)
        drop = (
            name in _MERGE_STRIP_RPR
            if mode == "merge"
            else name not in _DEST_KEEP_RPR
        )
        if drop:
            rpr.remove(child)
            changed = True
    return changed


def _strip_direct_formatting(elements, mode: str) -> tuple[int, int]:
    """Strip direct formatting from COPIED content per `formatting` mode
    ("merge" | "destination"). Operates only on the deep copies headed into
    the target — the source file is never touched. Returns counts of runs
    and paragraphs that lost at least one direct property."""
    runs_stripped = 0
    paras_stripped = 0
    for el in elements:
        for p in el.iter(qn("w:p")):
            changed = False
            ppr = p.find(qn("w:pPr"))
            if ppr is not None:
                for child in list(ppr):
                    name = _localname(child)
                    if name == "rPr":  # paragraph-mark run properties
                        if _strip_rpr_direct(child, mode):
                            changed = True
                        if len(child) == 0:
                            ppr.remove(child)
                        continue
                    drop = (
                        name in _MERGE_STRIP_PPR
                        if mode == "merge"
                        else name not in _DEST_KEEP_PPR
                    )
                    if drop:
                        ppr.remove(child)
                        changed = True
                if len(ppr) == 0:
                    p.remove(ppr)
            if changed:
                paras_stripped += 1
        for r in el.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            if _strip_rpr_direct(rpr, mode):
                runs_stripped += 1
                if len(rpr) == 0:
                    r.remove(rpr)
    return runs_stripped, paras_stripped


# ---------------------------------------------------------------- main tools


def _check_positioners(
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    before_first: bool = False,
) -> None:
    positioners = (
        (after_index is not None) + (after_anchor is not None)
        + bool(at_end) + bool(before_first)
    )
    if positioners != 1:
        raise WordMcpError(
            "give exactly one positioner: after_index, after_anchor, "
            "at_end=True, or before_first=True"
        )


def _open_source(pkg: DocxPackage, source_path: str) -> DocxPackage:
    src = DocxPackage(source_path)
    try:
        same = src.path.resolve() == pkg.path.resolve()
    except OSError:  # pragma: no cover
        same = str(src.path) == str(pkg.path)
    if same:
        raise WordMcpError(
            "source and target are the same file; refusing to insert a "
            "document into itself"
        )
    return src


def insert_document(
    pkg: DocxPackage,
    source_path: str,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    before_first: bool = False,
    formatting: str = "source",
) -> dict:
    """Insert the ENTIRE body content of the document at `source_path` into
    this document at one position.

    Positioning (exactly one required):
    - after_index: body ITEM index — paragraphs and tables counted together
      in document order (item 0 is the first block, whether paragraph or
      table). Insertion happens after that item. NOTE: this differs from the
      paragraph-only indices of get_text/insert_paragraphs when the target
      contains tables.
    - after_anchor: a paragraph whose FULL text exactly matches (whitespace-
      trimmed). More than one match refuses with every match location listed.
    - at_end: after the last body item, before the trailing sectPr.

    Carried: paragraphs, tables (incl. merged cells and nested tables),
    images, hyperlinks, numbered/bulleted lists, footnote/endnote references
    with their definitions, bookmarks, charts, tracked changes, equations.
    Styles reconcile by name (target formatting wins on a name match;
    unmatched styles are cloned). The source's trailing sectPr is never
    carried; mid-content section breaks are stripped and reported. Comment
    references are stripped and counted — comment transplant is out of scope.
    OLE objects, ActiveX controls, subdocuments, altChunks, and unknown
    embedded parts refuse the whole insertion (strip them from the source
    first); nothing is ever half-applied.

    formatting (Word paste modes, applied to the carried COPIES only):
    - "source" (default): direct formatting preserved — Word's InsertFile.
    - "merge": semantic emphasis direct formatting kept (bold, italic,
      underline, strike, sub/superscript, highlight) but direct font-family/
      size/color/character-spacing run overrides and paragraph spacing/
      line-spacing/indent overrides stripped — Word's Merge Formatting.
    - "destination": ALL direct rPr/pPr formatting stripped except what is
      structurally required (numPr, outlineLvl, tab stops stay); the
      target's styles govern rendering entirely.
    With "merge"/"destination" the result reports runs/paragraphs stripped.
    With "source", properties the source paragraphs inherited from their
    file's document defaults (docDefaults) are resolved to explicit values
    when the two files' defaults differ, so the carried content keeps the
    source's rendered spacing/indent/font; the result reports what was
    reconciled under "document_defaults".

    The source file is never modified."""
    _check_positioners(after_index, after_anchor, at_end, before_first)
    if formatting not in _FORMATTING_MODES:
        raise WordMcpError(
            f"formatting must be one of {list(_FORMATTING_MODES)}, "
            f"got {formatting!r}"
        )
    src = _open_source(pkg, source_path)
    copied: list[etree._Element] = []
    for c in src.body():
        if _localname(c) == "sectPr":
            continue
        if _localname(c) in ("commentRangeStart", "commentRangeEnd"):
            continue  # body-level comment markers; counted via references
        copied.append(copy.deepcopy(c))
    if not copied:
        raise TargetNotFound(
            f"source document has no body content to insert: {source_path}"
        )
    return _transplant(
        pkg,
        src,
        copied,
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
        before_first=before_first,
        formatting=formatting,
    )


def copy_table(
    pkg: DocxPackage,
    source_path: str,
    table_index: int,
    *,
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
) -> dict:
    """Transplant ONE top-level table from the document at `source_path` into
    this document, through the same reconciliation pipeline as
    insert_document (styles matched by name with the target's formatting
    governing, fresh numbering instances, images/hyperlinks re-registered,
    footnote/endnote definitions carried under new ids, bookmark ids
    remapped) scoped to the single table element.

    table_index counts the source's top-level body tables in document order
    (0-based; tables nested inside cells travel with their parent and are
    not separately addressable). Positioning follows the insert_document
    contract exactly: exactly one of after_index (body ITEM index of the
    TARGET, paragraphs and tables counted together), after_anchor (exact
    whole-paragraph text match, ambiguity refuses with every match listed),
    or at_end. Uncarryable content inside the table (OLE objects, ActiveX,
    subdocuments, altChunks, unknown embedded parts) refuses the whole copy;
    nothing is ever half-applied. The source file is never modified."""
    _check_positioners(after_index, after_anchor, at_end)
    if not isinstance(table_index, int) or isinstance(table_index, bool):
        raise WordMcpError(f"table_index must be an integer, got {table_index!r}")
    src = _open_source(pkg, source_path)
    tables = [el for el in _body_blocks(src) if _localname(el) == "tbl"]
    if not tables:
        raise TargetNotFound(
            f"source document has no top-level body tables: {source_path}"
        )
    if not 0 <= table_index < len(tables):
        raise TargetNotFound(
            f"table_index {table_index} out of range: the source has "
            f"{len(tables)} top-level body table(s) (valid indices "
            f"0-{len(tables) - 1})"
        )
    tbl = tables[table_index]
    rows = len(tbl.findall(qn("w:tr")))
    grid = tbl.find(qn("w:tblGrid"))
    cols = len(grid.findall(qn("w:gridCol"))) if grid is not None else None
    result = _transplant(
        pkg,
        src,
        [copy.deepcopy(tbl)],
        after_index=after_index,
        after_anchor=after_anchor,
        at_end=at_end,
        formatting="source",
    )
    result["source_table_index"] = table_index
    result["rows"] = rows
    if cols is not None:
        result["columns"] = cols
    return result


def _transplant(
    pkg: DocxPackage,
    src: DocxPackage,
    copied: list[etree._Element],
    *,
    after_index: int | None,
    after_anchor: str | None,
    at_end: bool,
    formatting: str,
    before_first: bool = False,
) -> dict:
    """Shared reconciliation-and-insert pipeline over already-deep-copied
    source elements. Phase A only scans (any refusal leaves the target
    untouched); phase B mutates the in-memory target, which is saved only by
    the caller."""
    pos_mode, ref_el, start_item = _resolve_position(
        pkg, after_index, after_anchor, at_end, before_first
    )

    # ---- phase A: prepare and SCAN (no target mutation on refusal)

    # Mid-content section breaks: strip BEFORE the relationship scan so their
    # header/footer references never enter the transplant set.
    section_breaks_stripped = 0
    for el in copied:
        for sp in list(el.iter(qn("w:sectPr"))):
            parent = sp.getparent()
            if parent is not None:
                parent.remove(sp)
                section_breaks_stripped += 1

    # Referenced note definitions, copied now so they join every later pass.
    note_plan: dict[str, list[tuple[str, etree._Element]]] = {}
    for kind, cfg in _notes._KINDS.items():
        ref_ids: list[str] = []
        for el in copied:
            for ref in el.iter(qn(cfg["body_ref"])):
                rid = ref.get(qn("w:id"))
                if rid is not None and rid not in ref_ids:
                    ref_ids.append(rid)
        if not ref_ids:
            continue
        if not src.has_part(cfg["part"]):
            raise UnsupportedStructure(
                f"source references {kind} ids {ref_ids} but has no "
                f"{cfg['part']}; the source document is corrupt"
            )
        defs = {
            n.get(qn("w:id")): n
            for n in src.root(cfg["part"]).findall(qn(cfg["note"]))
        }
        missing = [i for i in ref_ids if i not in defs]
        if missing:
            raise UnsupportedStructure(
                f"source references {kind} ids {missing} that have no "
                f"definition in {cfg['part']}; the source document is corrupt"
            )
        note_plan[kind] = [(i, copy.deepcopy(defs[i])) for i in ref_ids]

    units: list[tuple[etree._Element, str]] = [(el, "document") for el in copied]
    for kind, pairs in note_plan.items():
        units += [(d, kind) for _, d in pairs]

    # Comment references: stripped cleanly, counted, reported.
    comments_stripped = 0
    for el, _story in units:
        for node in list(el.iter(qn("w:commentReference"))):
            run = node.getparent()
            if run is not None and _localname(run) == "r":
                container = run.getparent()
                if container is not None:
                    container.remove(run)
            else:  # pragma: no cover - malformed but strip anyway
                node.getparent().remove(node)
            comments_stripped += 1
        for tag in ("w:commentRangeStart", "w:commentRangeEnd"):
            for node in list(el.iter(qn(tag))):
                if node.getparent() is not None:
                    node.getparent().remove(node)

    # Blocking content scan.
    blocked: dict[str, int] = {}
    for el, _story in units:
        for tag, label in _BLOCKED_ELEMENTS.items():
            hits = len(list(el.iter(tag)))
            if hits:
                blocked[label] = blocked.get(label, 0) + hits
    if blocked:
        raise UnsupportedStructure(
            "source contains content that cannot be transplanted safely: "
            + "; ".join(f"{v}x {k}" for k, v in sorted(blocked.items()))
            + ". Nothing was inserted. Strip these from the source first "
            "(e.g., delete the embedded objects in Word) and retry."
        )

    # Relationship scan and classification.
    src_rels: dict[str, dict[str, etree._Element]] = {}
    for story, (_dst, rels_name) in _STORY_PARTS.items():
        rels: dict[str, etree._Element] = {}
        if src.has_part(rels_name):
            for rel in src.root(rels_name):
                rels[rel.get("Id")] = rel
        src_rels[story] = rels

    rel_uses: list[tuple[str, etree._Element, str, str]] = []
    rel_refusals: list[str] = []
    for el, story in units:
        for node, key, rid in _iter_rid_attrs(el):
            rel = src_rels[story].get(rid)
            if rel is None:
                rel_refusals.append(
                    f"relationship {rid} (referenced from {story} content) "
                    "has no definition in the source"
                )
                continue
            rtype = rel.get("Type")
            external = rel.get("TargetMode") == "External"
            if not external:
                target_part = _resolve_rel_target(
                    _STORY_PARTS[story][1], rel.get("Target", "")
                )
                if rtype == _REL_IMAGE or rtype in _SUBTREE_REL_TYPES:
                    if not src.has_part(target_part):
                        rel_refusals.append(
                            f"relationship {rid} targets missing source part "
                            f"{target_part}"
                        )
                else:
                    rel_refusals.append(
                        f"unsupported embedded content: relationship {rid} of "
                        f"type {rtype} -> {target_part}"
                    )
            rel_uses.append((story, node, key, rid))
    if rel_refusals:
        raise UnsupportedStructure(
            "source content references parts this tool cannot transplant "
            "safely: " + "; ".join(sorted(set(rel_refusals)))
            + ". Nothing was inserted. Strip that content from the source "
            "first and retry."
        )

    # Pre-existing target note problems (never blamed on this insertion).
    pre_notes = _notes.validate_notes(pkg)

    # ---- phase B: mutate the in-memory target (saved only by the caller,
    # so any exception below still leaves the file untouched)

    # B0. Formatting mode: strip direct formatting from the carried COPIES
    # (body content and note definitions alike; the source is untouched).
    runs_stripped = paras_stripped = 0
    if formatting != "source":
        runs_stripped, paras_stripped = _strip_direct_formatting(
            [el for el, _ in units], formatting
        )

    # B0b. formatting="source": reconcile document defaults. Properties the
    # source paragraphs inherited from their file's docDefaults are baked
    # as explicit values wherever the two files' defaults differ, so the
    # carried content keeps the source's rendered look instead of silently
    # resolving against the target's defaults (field test, 2026-09-03).
    dd_baker = None
    if formatting == "source":
        dd_baker = _DefaultsBaker(src, pkg)
        for el, _story in units:
            if el.tag == qn("w:p"):
                dd_baker.bake_paragraph(el)
            for p in el.iterdescendants(qn("w:p")):
                dd_baker.bake_paragraph(p)
        if not dd_baker.reportable:
            dd_baker = None

    # B1+B2. Styles and numbering, to a fixpoint (cloned styles can reference
    # numbering; cloned numbering can reference styles via styleLink/pStyle).
    styles = _StyleResolver(src, pkg)
    numbering = _NumberingResolver(src, pkg)
    unit_els = [el for el, _ in units]
    pending_sids = _collect_style_refs(unit_els)
    pending_nids = _collect_num_refs(unit_els)
    for _round in range(10):
        if not pending_sids and not pending_nids:
            break
        for sid in sorted(pending_sids):
            styles.resolve(sid)
        for nid in sorted(pending_nids):
            numbering.resolve(nid)
        pending_sids = {
            s
            for s in _collect_style_refs(numbering.cloned_abstracts)
            | _collect_style_refs(styles.cloned_defs)
            if s not in styles._done
        }
        pending_nids = {
            n
            for n in _collect_num_refs(styles.cloned_defs)
            if n not in numbering._done
        }
    remap_targets = unit_els + styles.cloned_defs + numbering.cloned_abstracts
    _apply_style_remap(remap_targets, styles.remap)
    lists_stripped = _apply_num_remap(
        remap_targets, numbering.map, set(numbering.unresolved)
    )

    # B2b. formatting="source": theme colour references in the carried xml
    # (direct formatting and the definitions of cloned or imported styles)
    # are resolved against the SOURCE theme and frozen, because the theme
    # part never travels and the target's colours would take over (M6).
    theme_refs_frozen = 0
    theme_slots_differ: list[str] = []
    theme_mappings_differ: list[str] = []
    if formatting == "source":
        src_theme, tgt_theme = _ThemeColors(src), _ThemeColors(pkg)
        theme_slots_differ = src_theme.differing_slots(tgt_theme)
        theme_mappings_differ = src_theme.differing_mappings(tgt_theme)
        theme_refs_frozen = _freeze_theme_colors(
            remap_targets, src_theme, tgt_theme
        )

    # B3. Notes: ensure parts/styles exist, assign fresh ids, retag, append.
    notes_carried = {"footnote": 0, "endnote": 0}
    for kind, pairs in note_plan.items():
        cfg = _notes._KINDS[kind]
        _notes._ensure_part(pkg, kind)
        _notes._ensure_styles(pkg, kind)
        notes_root = pkg.root(cfg["part"])
        id_map: dict[str, str] = {}
        for old_id, def_copy in pairs:
            new_id = str(_notes._next_id(pkg, kind))
            def_copy.set(qn("w:id"), new_id)
            notes_root.append(def_copy)
            id_map[old_id] = new_id
        pkg.mark_dirty(cfg["part"])
        for el in copied:
            for ref in el.iter(qn(cfg["body_ref"])):
                old = ref.get(qn("w:id"))
                if old in id_map:
                    ref.set(qn("w:id"), id_map[old])
        notes_carried[kind] = len(pairs)

    # B4. Relationships and parts.
    rid_memo: dict[tuple[str, str], str] = {}
    media_memo: dict[str, str] = {}
    subtree_memo: dict[str, str] = {}
    chart_parts: set[str] = set()
    rels_ids: dict[str, set[str]] = {}
    hyperlinks_carried = 0
    external_rels_carried = 0

    def _dst_rels_root(story: str) -> tuple[str, etree._Element]:
        name = _STORY_PARTS[story][1]
        _ensure_rels_part(pkg, name)
        root = pkg.root(name)
        if name not in rels_ids:
            rels_ids[name] = {r.get("Id") for r in root}
        return name, root

    def _add_rel(
        story: str, rtype: str, target: str, *, external: bool
    ) -> str:
        name, root = _dst_rels_root(story)
        n = 1
        while f"rId{n}" in rels_ids[name]:
            n += 1
        rid = f"rId{n}"
        rels_ids[name].add(rid)
        rel = etree.SubElement(root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", rid)
        rel.set("Type", rtype)
        rel.set("Target", target)
        if external:
            rel.set("TargetMode", "External")
        pkg.mark_dirty(name)
        return rid

    def _copy_media(src_part: str) -> str:
        if src_part in media_memo:
            return media_memo[src_part]
        ext = (
            "." + src_part.rsplit(".", 1)[1].lower()
            if "." in src_part.rsplit("/", 1)[1]
            else ""
        )
        n = 1
        while pkg.has_part(f"word/media/image{n}{ext}") or any(
            name.startswith(f"word/media/image{n}.")
            for name in pkg.part_names()
        ):
            n += 1
        new_part = f"word/media/image{n}{ext}"
        pkg.set_raw_part(new_part, src.raw_part(src_part))
        _ensure_content_type(pkg, src, src_part, new_part)
        media_memo[src_part] = new_part
        return new_part

    def _copy_subtree(src_part: str) -> str:
        """Copy a part plus its private dependency subtree (chart ->
        embedded workbook / colors / style parts), keeping the part's
        internal rIds valid by rewriting only relationship Targets."""
        if src_part in subtree_memo:
            return subtree_memo[src_part]
        new_part = _free_part_name(pkg, src_part)
        pkg.set_raw_part(new_part, b"")  # reserve the name before recursing
        subtree_memo[src_part] = new_part
        src_rels_name = _rels_part_for(src_part)
        if src.has_part(src_rels_name):
            rels_root = copy.deepcopy(src.root(src_rels_name))
            for rel in rels_root:
                if rel.get("TargetMode") == "External":
                    continue
                child_src = _resolve_rel_target(
                    src_rels_name, rel.get("Target", "")
                )
                if not src.has_part(child_src):
                    raise UnsupportedStructure(
                        f"source part {src_part} references missing part "
                        f"{child_src}; the source document is corrupt; "
                        "nothing was inserted"
                    )
                child_new = _copy_subtree(child_src)
                rel.set(
                    "Target",
                    posixpath.relpath(
                        child_new, start=new_part.rsplit("/", 1)[0]
                    ),
                )
            pkg.set_raw_part(
                _rels_part_for(new_part),
                etree.tostring(
                    rels_root,
                    xml_declaration=True,
                    encoding="UTF-8",
                    standalone=True,
                ),
            )
        pkg.set_raw_part(new_part, src.raw_part(src_part))
        _ensure_content_type(pkg, src, src_part, new_part)
        return new_part

    for story, node, key, rid in rel_uses:
        memo_key = (story, rid)
        new_rid = rid_memo.get(memo_key)
        if new_rid is None:
            rel = src_rels[story][rid]
            rtype = rel.get("Type")
            if rel.get("TargetMode") == "External":
                new_rid = _add_rel(
                    story, rtype, rel.get("Target", ""), external=True
                )
                external_rels_carried += 1
                if rtype == _REL_HYPERLINK:
                    hyperlinks_carried += 1
            else:
                src_part = _resolve_rel_target(
                    _STORY_PARTS[story][1], rel.get("Target", "")
                )
                if rtype == _REL_IMAGE:
                    new_part = _copy_media(src_part)
                else:  # chart / chartEx (phase A allowed nothing else)
                    new_part = _copy_subtree(src_part)
                    chart_parts.add(src_part)
                # Both document.xml and the notes parts live in word/, so the
                # relative target is the part name minus the word/ prefix.
                new_rid = _add_rel(
                    story, rtype, new_part.split("word/", 1)[1], external=False
                )
            rid_memo[memo_key] = new_rid
        node.set(key, new_rid)

    # B5. Bookmarks: fresh ids; name collisions renamed and retargeted.
    existing_names: set[str] = set()
    max_bm_id = 0
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if not pkg.has_part(part):
            continue
        for bs in pkg.root(part).iter(qn("w:bookmarkStart")):
            name = bs.get(qn("w:name"))
            if name:
                existing_names.add(name)
            try:
                max_bm_id = max(max_bm_id, int(bs.get(qn("w:id"), "0") or 0))
            except ValueError:  # pragma: no cover
                pass
    bm_id_map: dict[str, str] = {}
    bm_renames: list[dict] = []
    name_map: dict[str, str] = {}
    bookmarks_carried = 0
    for el, _story in units:
        for bs in el.iter(qn("w:bookmarkStart")):
            old_id = bs.get(qn("w:id"))
            name = bs.get(qn("w:name")) or ""
            max_bm_id += 1
            new_id = str(max_bm_id)
            if old_id is not None:
                bm_id_map[old_id] = new_id
            bs.set(qn("w:id"), new_id)
            if name:
                if name in existing_names:
                    n = 1
                    new_name = f"{name[:34]}_ins{n}"
                    while new_name in existing_names:
                        n += 1
                        new_name = f"{name[:34]}_ins{n}"
                    bs.set(qn("w:name"), new_name)
                    name_map[name] = new_name
                    bm_renames.append({"from": name, "to": new_name})
                    existing_names.add(new_name)
                else:
                    existing_names.add(name)
            bookmarks_carried += 1
    for el, _story in units:
        for be in el.iter(qn("w:bookmarkEnd")):
            old_id = be.get(qn("w:id"))
            if old_id in bm_id_map:
                be.set(qn("w:id"), bm_id_map[old_id])
    # Retarget internal links and REF-family fields to renamed bookmarks —
    # only within the inserted content (target content is never rewritten).
    refs_retargeted = 0
    if name_map:
        for el, _story in units:
            for link in el.iter(qn("w:hyperlink")):
                anchor = link.get(qn("w:anchor"))
                if anchor in name_map:
                    link.set(qn("w:anchor"), name_map[anchor])
                    refs_retargeted += 1
            for instr in el.iter(qn("w:instrText")):
                text = instr.text or ""
                for old, new in name_map.items():
                    pattern = (
                        r"\b(REF|PAGEREF|NOTEREF|HYPERLINK\s+\\l)(\s+\"?)"
                        + re.escape(old)
                        + r"\b"
                    )
                    new_text = re.sub(pattern, r"\g<1>\g<2>" + new, text)
                    if new_text != text:
                        text = new_text
                        refs_retargeted += 1
                instr.text = text

    # B6. docPr ids unique across the whole document (body + notes).
    max_docpr = 0
    for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
        if not pkg.has_part(part):
            continue
        for d in pkg.root(part).iter(f"{{{_WP}}}docPr"):
            try:
                max_docpr = max(max_docpr, int(d.get("id", "0") or 0))
            except ValueError:  # pragma: no cover
                pass
    for el, _story in units:
        for d in el.iter(f"{{{_WP}}}docPr"):
            max_docpr += 1
            d.set("id", str(max_docpr))

    # B7. Insert at the resolved position.
    body = pkg.body()
    if ref_el is None:
        sectpr = body.find(qn("w:sectPr"))
        for el in copied:
            if sectpr is not None:
                sectpr.addprevious(el)
            else:
                body.append(el)
    elif pos_mode == "before":
        for el in copied:
            ref_el.addprevious(el)
    else:
        for el in reversed(copied):
            ref_el.addnext(el)
    pkg.mark_dirty()

    # B8. Closure validation on everything just inserted; a failure here
    # aborts before the caller saves, leaving the file untouched.
    for el, story in units:
        rels_name, rels_root = _dst_rels_root(story)
        rel_map = {r.get("Id"): r for r in rels_root}
        for _node, _key, rid in _iter_rid_attrs(el):
            rel = rel_map.get(rid)
            if rel is None:
                raise ValidationFailed(
                    f"internal error: inserted content references undefined "
                    f"relationship {rid} in {rels_name}; document not saved"
                )
            if rel.get("TargetMode") != "External":
                resolved = _resolve_rel_target(
                    rels_name, rel.get("Target", "")
                )
                if not pkg.has_part(resolved):
                    raise ValidationFailed(
                        f"internal error: relationship {rid} targets missing "
                        f"part {resolved}; document not saved"
                    )
    post_notes = _notes.validate_notes(pkg)
    for kind_key, report in post_notes.items():
        pre = pre_notes.get(kind_key, {})
        new_missing = set(report["missing_definitions"]) - set(
            pre.get("missing_definitions", [])
        )
        new_dups = set(report["duplicate_references"]) - set(
            pre.get("duplicate_references", [])
        )
        if new_missing or new_dups:
            raise ValidationFailed(
                f"internal error: note integrity broke during insertion "
                f"({kind_key}: missing {sorted(new_missing)}, duplicates "
                f"{sorted(new_dups)}); document not saved"
            )

    n_paras = sum(1 for el in copied if _localname(el) == "p")
    n_tables = sum(1 for el in copied if _localname(el) == "tbl")
    n_items = n_paras + n_tables
    result = {
        "inserted_from": str(src.path),
        "position": {
            "mode": (
                "at_end"
                if at_end
                else "before_first"
                if before_first
                else ("after_index" if after_index is not None else "after_anchor")
            ),
            "starts_at_body_item": start_item,
        },
        "body_item_range": (
            [start_item, start_item + n_items - 1] if n_items else None
        ),
        "paragraphs": n_paras,
        "tables": n_tables,
        "images_carried": len(media_memo),
        "charts_carried": len(chart_parts),
        "hyperlinks_carried": hyperlinks_carried,
        "external_relationships_carried": external_rels_carried,
        "footnotes_carried": notes_carried["footnote"],
        "endnotes_carried": notes_carried["endnote"],
        "lists_carried": len(
            [k for k, v in numbering.map.items() if k != "0"]
        ),
        "styles": {
            "matched_by_name": len(styles.matched),
            "remapped_ids": dict(sorted(styles.remap.items())),
            "cloned": styles.cloned,
            **(
                {
                    "reused_imports": [
                        {k: v for k, v in m.items() if k != "reused_import"}
                        for m in styles.matched if m.get("reused_import")
                    ],
                    "reused_imports_note": (
                        "an earlier insert already imported this table "
                        "style under a new name because the target defines "
                        "a different style of the same name; the tables "
                        "carried this time point at that existing copy "
                        "rather than adding another one to the gallery"
                    ),
                }
                if any(m.get("reused_import") for m in styles.matched)
                else {}
            ),
            **(
                {
                    "imported_renamed": styles.imported_renamed,
                    "imported_renamed_note": (
                        "a table style of this name exists in the target with "
                        "a different definition; its conditional formatting "
                        "cannot be baked onto the carried cells, so the "
                        "source's table style was imported under a new name "
                        "and the carried tables point at it"
                    ),
                }
                if styles.imported_renamed
                else {}
            ),
        },
        "bookmarks_carried": bookmarks_carried,
        "bookmarks_renamed": bm_renames,
        "bookmark_refs_retargeted": refs_retargeted,
        "comments_stripped": comments_stripped,
        "section_breaks_stripped": section_breaks_stripped,
        "formatting_mode": formatting,
    }
    if theme_refs_frozen or (
        dd_baker is not None and dd_baker.theme_colors_baked
    ):
        result["theme_colors"] = {
            "differing_slots": theme_slots_differ,
            **(
                {"differing_mappings": theme_mappings_differ}
                if theme_mappings_differ
                else {}
            ),
            "references_frozen": theme_refs_frozen,
            "resolved_from_styles": (
                dd_baker.theme_colors_baked if dd_baker is not None else 0
            ),
            "note": (
                "the two files resolve these theme colours differently "
                "(a different colour scheme in the theme part, a different "
                "clrSchemeMapping in settings.xml, or both) and neither the "
                "theme part nor the mapping travels, so theme colour "
                "references in the carried content were resolved against "
                "the SOURCE theme and written as plain values; they will "
                "not follow the target's theme from here on"
            ),
        }
    if formatting != "source":
        result["runs_stripped"] = runs_stripped
        result["paragraphs_stripped"] = paras_stripped
    if dd_baker is not None:
        result["document_defaults"] = dd_baker.report()
    if styles.unresolved:
        result["style_refs_unresolved_in_source"] = sorted(styles.unresolved)
    if numbering.unresolved:
        result["numbering_refs_unresolved_in_source"] = sorted(
            numbering.unresolved
        )
        result["numbering_refs_stripped"] = lists_stripped
    return result
