"""Bulleted and numbered lists, and the numbering behind them.

Real bullets/numbers are NOT a style: each list paragraph carries
w:pPr/w:numPr/(w:ilvl, w:numId), where numId points into word/numbering.xml
(w:num -> w:abstractNum with per-level formats). ListParagraph style alone
renders indented plain text, the bug this module exists to prevent.

Three layers, and knowing which one a change belongs to is most of the
work:

- w:abstractNum holds the FORMAT of each of the nine levels: what numeral
  to use, what the label looks like, where it starts, how far it indents.
  Several lists can share one abstract definition.
- w:num is one INSTANCE of an abstract definition. Two paragraphs sharing
  a numId are one list and count together; two numIds are two lists, each
  counting from its own start. That is why a plain add_list call restarts
  at 1: it creates its own instance, which is what Word does for a new
  list.
- w:lvlOverride, inside a w:num, changes one level for that instance only.
  A bare w:startOverride is the "restart numbering at N" a person means,
  and it is the mechanism naive implementations skip: about one document
  in seven carries a level override, so a tool that ignores them reads the
  wrong number off the page.

Which is also why the numbers this module reports are COMPUTED rather
than read. Nothing in the file stores "this paragraph is item 4"; Word
works it out by walking the document, and so does computed_numbers.
"""

from __future__ import annotations

from lxml import etree

from ..core.errors import TargetNotFound, UnsupportedStructure, WordMcpError
from ..core.package import DocxPackage, qn

_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

# (numFmt, lvlText, font) cycling across levels 0-8.
_BULLET_LEVELS = [
    ("bullet", "", "Symbol"),
    ("bullet", "o", "Courier New"),
    ("bullet", "", "Wingdings"),
]
_NUMBER_LEVELS = [
    ("decimal", "%{n}.", None),
    ("lowerLetter", "%{n}.", None),
    ("lowerRoman", "%{n}.", None),
]


def _ensure_numbering_part(pkg: DocxPackage) -> None:
    part = "word/numbering.xml"
    if pkg.has_part(part):
        return
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    root = etree.Element(qn("w:numbering"), nsmap={"w": w_ns})
    pkg.set_raw_part(
        part,
        etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True),
    )
    ct_root = pkg.root("[Content_Types].xml")
    if not any(
        o.get("PartName") == "/" + part
        for o in ct_root.findall(f"{{{_CT_NS}}}Override")
    ):
        override = etree.SubElement(ct_root, f"{{{_CT_NS}}}Override")
        override.set("PartName", "/" + part)
        override.set(
            "ContentType",
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.numbering+xml",
        )
        pkg.mark_dirty("[Content_Types].xml")
    rels_part = "word/_rels/document.xml.rels"
    rels_root = pkg.root(rels_part)
    rel_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/"
        "relationships/numbering"
    )
    if not any(r.get("Type") == rel_type for r in rels_root):
        existing = {r.get("Id") for r in rels_root}
        i = 1
        while f"rId{i}" in existing:
            i += 1
        rel = etree.SubElement(rels_root, f"{{{_REL_NS}}}Relationship")
        rel.set("Id", f"rId{i}")
        rel.set("Type", rel_type)
        rel.set("Target", "numbering.xml")
        pkg.mark_dirty(rels_part)


def _new_numbering(pkg: DocxPackage, kind: str) -> int:
    """Create an abstractNum + num pair; return the numId to reference."""
    root = pkg.root("word/numbering.xml")
    abstract_ids = [
        int(a.get(qn("w:abstractNumId"), "0"))
        for a in root.findall(qn("w:abstractNum"))
    ]
    num_ids = [
        int(n.get(qn("w:numId"), "0")) for n in root.findall(qn("w:num"))
    ]
    abs_id = max(abstract_ids, default=-1) + 1
    num_id = max(num_ids, default=0) + 1

    levels = _BULLET_LEVELS if kind == "bullet" else _NUMBER_LEVELS
    abstract = etree.Element(qn("w:abstractNum"))
    abstract.set(qn("w:abstractNumId"), str(abs_id))
    ml = etree.SubElement(abstract, qn("w:multiLevelType"))
    ml.set(qn("w:val"), "hybridMultilevel")
    for ilvl in range(9):
        num_fmt, lvl_text, font = levels[ilvl % len(levels)]
        lvl = etree.SubElement(abstract, qn("w:lvl"))
        lvl.set(qn("w:ilvl"), str(ilvl))
        start = etree.SubElement(lvl, qn("w:start"))
        start.set(qn("w:val"), "1")
        fmt = etree.SubElement(lvl, qn("w:numFmt"))
        fmt.set(qn("w:val"), num_fmt)
        text_el = etree.SubElement(lvl, qn("w:lvlText"))
        text_el.set(qn("w:val"), lvl_text.replace("{n}", str(ilvl + 1)))
        jc = etree.SubElement(lvl, qn("w:lvlJc"))
        jc.set(qn("w:val"), "left")
        ppr = etree.SubElement(lvl, qn("w:pPr"))
        ind = etree.SubElement(ppr, qn("w:ind"))
        ind.set(qn("w:left"), str(720 * (ilvl + 1)))
        ind.set(qn("w:hanging"), "360")
        if font:
            rpr = etree.SubElement(lvl, qn("w:rPr"))
            rfonts = etree.SubElement(rpr, qn("w:rFonts"))
            for attr in ("w:ascii", "w:hAnsi", "w:hint"):
                rfonts.set(
                    qn(attr), font if attr != "w:hint" else "default"
                )

    # abstractNum elements must precede num elements in numbering.xml.
    nums = root.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(abstract)
    else:
        root.append(abstract)
    num = etree.SubElement(root, qn("w:num"))
    num.set(qn("w:numId"), str(num_id))
    ref = etree.SubElement(num, qn("w:abstractNumId"))
    ref.set(qn("w:val"), str(abs_id))
    pkg.mark_dirty("word/numbering.xml")
    return num_id


# ------------------------------------------------------- the numbering model
#
# Everything below reads and edits word/numbering.xml. It is deliberately
# separate from the paragraph-building code above: a list's numbering can
# be changed without touching a single paragraph, and usually should be.

#: numFmt values this module can render into an actual label. Anything
#: else is reported by name with the raw counter, never guessed at.
_RENDERABLE = {
    "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman",
    "ordinal", "decimalZero", "none", "bullet",
}

_ROMAN = [
    (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
    (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"), (5, "v"),
    (4, "iv"), (1, "i"),
]


def _roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for value, letter in _ROMAN:
        while n >= value:
            out.append(letter)
            n -= value
    return "".join(out)


def _letters(n: int) -> str:
    """1 -> a, 26 -> z, 27 -> aa, the way Word continues past the alphabet."""
    if n <= 0:
        return str(n)
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("a") + rem) + out
    return out


def _ordinal(n: int) -> str:
    if 11 <= (n % 100) <= 13:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }".replace(
        " ", ""
    )


def format_counter(value: int, num_fmt: str) -> str | None:
    """One level's counter as its label, or None when the format is one
    this module does not render (a caller reports the raw number and the
    format name instead of inventing a label)."""
    if num_fmt == "decimal":
        return str(value)
    if num_fmt == "decimalZero":
        return f"{value:02d}"
    if num_fmt == "lowerLetter":
        return _letters(value)
    if num_fmt == "upperLetter":
        return _letters(value).upper()
    if num_fmt == "lowerRoman":
        return _roman(value)
    if num_fmt == "upperRoman":
        return _roman(value).upper()
    if num_fmt == "ordinal":
        return _ordinal(value)
    if num_fmt in ("none", "bullet"):
        return ""
    return None


def _numbering_root(pkg: DocxPackage):
    if not pkg.has_part("word/numbering.xml"):
        return None
    return pkg.root("word/numbering.xml")


def _num_element(pkg: DocxPackage, num_id: int):
    root = _numbering_root(pkg)
    if root is None:
        return None
    for num in root.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(num_id):
            return num
    return None


def _abstract_element(pkg: DocxPackage, num_id: int):
    """The w:abstractNum a numId resolves to, or None."""
    num = _num_element(pkg, num_id)
    if num is None:
        return None
    ref = num.find(qn("w:abstractNumId"))
    if ref is None:
        return None
    abs_id = ref.get(qn("w:val"))
    root = _numbering_root(pkg)
    for abstract in root.findall(qn("w:abstractNum")):
        if abstract.get(qn("w:abstractNumId")) == abs_id:
            return abstract
    return None


def _lvl_of(container, level: int):
    if container is None:
        return None
    for lvl in container.findall(qn("w:lvl")):
        if lvl.get(qn("w:ilvl")) == str(level):
            return lvl
    return None


def _override_of(num, level: int):
    if num is None:
        return None
    for ov in num.findall(qn("w:lvlOverride")):
        if ov.get(qn("w:ilvl")) == str(level):
            return ov
    return None


def _val(el, tag: str, default=None):
    if el is None:
        return default
    child = el.find(qn(tag))
    if child is None:
        return default
    return child.get(qn("w:val"), default)


def effective_level(pkg: DocxPackage, num_id: int, level: int) -> dict:
    """One level as Word sees it: the abstract definition, with the
    instance's own lvlOverride laid over the top.

    Both halves matter. The abstract definition is shared, so a document
    can have twenty lists that all look the same, and the override is what
    makes list seventeen restart at 5 without disturbing the other
    nineteen.
    """
    num = _num_element(pkg, num_id)
    abstract = _abstract_element(pkg, num_id)
    base = _lvl_of(abstract, level)
    override = _override_of(num, level)
    over_lvl = _lvl_of(override, level) if override is not None else None
    if over_lvl is None and override is not None:
        over_lvl = override.find(qn("w:lvl"))

    out = {
        "level": level,
        "format": _val(over_lvl, "w:numFmt") or _val(base, "w:numFmt")
        or "decimal",
        "text": _val(over_lvl, "w:lvlText") or _val(base, "w:lvlText") or "",
        "start": int(_val(over_lvl, "w:start") or _val(base, "w:start") or 1),
        "restart_after_level": _val(over_lvl, "w:lvlRestart")
        or _val(base, "w:lvlRestart"),
        "legal": (_val(over_lvl, "w:isLgl") or _val(base, "w:isLgl")) not in
        (None, "0", "false"),
        "suffix": _val(over_lvl, "w:suff") or _val(base, "w:suff") or "tab",
    }
    if out["restart_after_level"] is not None:
        out["restart_after_level"] = int(out["restart_after_level"])
    start_override = None
    if override is not None:
        so = override.find(qn("w:startOverride"))
        if so is not None:
            start_override = int(so.get(qn("w:val"), "1"))
    if start_override is not None:
        out["start_override"] = start_override
        out["start"] = start_override
    return out


def describe_numbering(pkg: DocxPackage, num_id: int) -> dict:
    """A numbering instance: which abstract definition it uses, every
    level's effective format, and which levels the instance overrides."""
    num = _num_element(pkg, num_id)
    if num is None:
        raise TargetNotFound(
            f"no numbering instance with num_id {num_id}; read the ids from "
            f"list_elements(type='lists')"
        )
    abstract = _abstract_element(pkg, num_id)
    ref = num.find(qn("w:abstractNumId"))
    levels = [effective_level(pkg, num_id, i) for i in range(9)]
    overridden = sorted(
        int(ov.get(qn("w:ilvl"), "0"))
        for ov in num.findall(qn("w:lvlOverride"))
    )
    root = _numbering_root(pkg)
    abs_val = ref.get(qn("w:val")) if ref is not None else None
    shared_with = sorted(
        int(other.get(qn("w:numId"), "0"))
        for other in (root.findall(qn("w:num")) if root is not None else [])
        if other is not num
        and other.find(qn("w:abstractNumId")) is not None
        and other.find(qn("w:abstractNumId")).get(qn("w:val")) == abs_val
    )
    out = {
        "num_id": num_id,
        "abstract_num_id": int(ref.get(qn("w:val"))) if ref is not None else None,
        "multi_level_type": _val(abstract, "w:multiLevelType"),
        "levels": levels,
        "overridden_levels": overridden,
    }
    if shared_with:
        out["shares_definition_with"] = shared_with
    return out


def computed_numbers(pkg: DocxPackage) -> list[dict]:
    """Every numbered paragraph with the label a reader actually sees.

    Nothing in the file stores it. Word computes it by walking the
    document in order, and the counters depend on the start value, any
    startOverride, and the restart rule for each level, so this walk has
    to do the same thing. A level's counter resets when a level at or
    above its lvlRestart increments; the default is "any shallower level",
    which is what the numberless majority of documents rely on.
    """
    from .read import body_items, paragraph_text

    counters: dict[tuple[int, int], int] = {}
    seen: dict[tuple[int, int], bool] = {}
    out: list[dict] = []
    for kind, idx, el in body_items(pkg):
        if kind != "paragraph":
            continue
        numpr = el.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
        if numpr is None:
            continue
        nid_el = numpr.find(qn("w:numId"))
        if nid_el is None:
            continue
        num_id = int(nid_el.get(qn("w:val"), "0"))
        if num_id == 0:  # numId 0 means "numbering removed"
            continue
        ilvl_el = numpr.find(qn("w:ilvl"))
        level = int(ilvl_el.get(qn("w:val"), "0")) if ilvl_el is not None else 0
        config = effective_level(pkg, num_id, level)

        key = (num_id, level)
        if seen.get(key):
            counters[key] = counters.get(key, config["start"]) + 1
        else:
            counters[key] = config["start"]
            seen[key] = True
        # Deeper levels start over. lvlRestart says how shallow the
        # increment has to be before a level resets; 0 means never.
        for (other_num, other_level) in list(seen):
            if other_num != num_id or other_level <= level:
                continue
            restart = effective_level(
                pkg, num_id, other_level
            )["restart_after_level"]
            if restart == 0:
                continue
            if restart is None or level < restart:
                seen[(other_num, other_level)] = False

        entry = {
            "paragraph_index": idx,
            "num_id": num_id,
            "level": level,
            "format": config["format"],
            "text": paragraph_text(el),
        }
        label = _render_label(pkg, num_id, level, counters, config)
        if label is None:
            entry["number"] = counters[key]
            entry["unrendered_format"] = config["format"]
        else:
            entry["number"] = label
        out.append(entry)
    return out


def _render_label(
    pkg: DocxPackage, num_id: int, level: int, counters: dict, config: dict
) -> str | None:
    """Substitute the counters into a level's lvlText ("%1.%2." and so on).

    Returns None when any level involved uses a format this module does
    not render, so the caller can report the raw counter instead of a
    wrong label.
    """
    if config["format"] == "bullet":
        return config["text"]
    text = config["text"] or ""
    if not text:
        rendered = format_counter(counters[(num_id, level)], config["format"])
        return rendered
    out = text
    for placeholder_level in range(9):
        token = f"%{placeholder_level + 1}"
        if token not in out:
            continue
        value = counters.get((num_id, placeholder_level))
        sub_config = (
            config if placeholder_level == level
            else effective_level(pkg, num_id, placeholder_level)
        )
        if value is None:
            value = sub_config["start"]
        # isLgl forces every level of the label to plain decimal, which is
        # how legal numbering turns "IV.a" into "4.1".
        fmt = "decimal" if config["legal"] else sub_config["format"]
        rendered = format_counter(value, fmt)
        if rendered is None:
            return None
        out = out.replace(token, rendered)
    return out


# ------------------------------------------------------- editing the numbers

#: CT_Lvl's child sequence. Word rejects a file that reorders these, so
#: every level edit rebuilds the element through _write_level.
_LVL_ORDER = [
    "start", "numFmt", "lvlRestart", "pStyle", "isLgl", "suff", "lvlText",
    "lvlPicBulletId", "legacy", "lvlJc", "pPr", "rPr",
]

#: The numFmt values Word accepts that this module will write. Deliberately
#: shorter than the full ECMA list: a format nobody can see rendered is a
#: format nobody should be able to set by typo.
WRITABLE_FORMATS = {
    "decimal", "decimalZero", "lowerLetter", "upperLetter", "lowerRoman",
    "upperRoman", "ordinal", "bullet", "none",
}

_SUFFIXES = {"tab", "space", "nothing"}
_JUSTIFY = {"left", "center", "right"}


def _place_lvl_child(lvl, child) -> None:
    index = _LVL_ORDER.index(etree.QName(child).localname)
    for existing in lvl:
        name = etree.QName(existing).localname
        if name in _LVL_ORDER and _LVL_ORDER.index(name) > index:
            existing.addprevious(child)
            return
    lvl.append(child)


def _set_simple(lvl, tag: str, value) -> None:
    """Replace one w:val-carrying child, keeping the schema order."""
    existing = lvl.find(qn(tag))
    if existing is not None:
        lvl.remove(existing)
    el = etree.Element(qn(tag))
    el.set(qn("w:val"), str(value))
    _place_lvl_child(lvl, el)


def _write_level(lvl, spec: dict) -> list[str]:
    """Apply one level spec to a w:lvl element. Returns what changed."""
    changed = []
    if "format" in spec:
        fmt = str(spec["format"])
        if fmt not in WRITABLE_FORMATS:
            raise WordMcpError(
                f"unknown numbering format {fmt!r}; use one of "
                f"{sorted(WRITABLE_FORMATS)}"
            )
        _set_simple(lvl, "w:numFmt", fmt)
        changed.append("format")
    if "text" in spec:
        _set_simple(lvl, "w:lvlText", str(spec["text"]))
        changed.append("text")
    if "start" in spec:
        start = int(spec["start"])
        if start < 0:
            raise WordMcpError("start cannot be negative")
        _set_simple(lvl, "w:start", start)
        changed.append("start")
    if "suffix" in spec:
        suffix = str(spec["suffix"])
        if suffix not in _SUFFIXES:
            raise WordMcpError(
                f"unknown suffix {suffix!r}; use one of {sorted(_SUFFIXES)} "
                f"(what sits between the label and the text)"
            )
        _set_simple(lvl, "w:suff", suffix)
        changed.append("suffix")
    if "align" in spec:
        align = str(spec["align"])
        if align not in _JUSTIFY:
            raise WordMcpError(
                f"unknown align {align!r}; use one of {sorted(_JUSTIFY)}"
            )
        _set_simple(lvl, "w:lvlJc", align)
        changed.append("align")
    if "legal" in spec:
        _set_simple(lvl, "w:isLgl", "1" if spec["legal"] else "0")
        changed.append("legal")
    if "restart_after_level" in spec:
        value = spec["restart_after_level"]
        if value is None:
            existing = lvl.find(qn("w:lvlRestart"))
            if existing is not None:
                lvl.remove(existing)
        else:
            _set_simple(lvl, "w:lvlRestart", int(value))
        changed.append("restart_after_level")
    if "indent_pt" in spec or "hanging_pt" in spec:
        ppr = lvl.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            _place_lvl_child(lvl, ppr)
        ind = ppr.find(qn("w:ind"))
        if ind is None:
            ind = etree.SubElement(ppr, qn("w:ind"))
        if "indent_pt" in spec:
            ind.set(qn("w:left"), str(int(float(spec["indent_pt"]) * 20)))
            changed.append("indent_pt")
        if "hanging_pt" in spec:
            ind.set(qn("w:hanging"), str(int(float(spec["hanging_pt"]) * 20)))
            changed.append("hanging_pt")
    if "font" in spec:
        rpr = lvl.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("w:rPr"))
            _place_lvl_child(lvl, rpr)
        rfonts = rpr.find(qn("w:rFonts"))
        if rfonts is None:
            rfonts = etree.SubElement(rpr, qn("w:rFonts"))
        for attr in ("w:ascii", "w:hAnsi"):
            rfonts.set(qn(attr), str(spec["font"]))
        rfonts.set(qn("w:hint"), "default")
        changed.append("font")
    return changed


def _normalize_levels(levels) -> dict[int, dict]:
    if not isinstance(levels, list) or not levels:
        raise WordMcpError("levels must be a non-empty list of level specs")
    out: dict[int, dict] = {}
    known = {
        "level", "format", "text", "start", "suffix", "align", "legal",
        "restart_after_level", "indent_pt", "hanging_pt", "font",
    }
    for spec in levels:
        if not isinstance(spec, dict) or "level" not in spec:
            raise WordMcpError(
                "each level spec needs a 'level' (0-8) plus the properties "
                "to set"
            )
        unknown = set(spec) - known
        if unknown:
            raise WordMcpError(
                f"unknown level properties {sorted(unknown)}; use "
                f"{sorted(known - {'level'})}"
            )
        level = int(spec["level"])
        if not 0 <= level <= 8:
            raise WordMcpError("level must be 0-8")
        out[level] = {k: v for k, v in spec.items() if k != "level"}
    return out


def _clone_abstract(pkg: DocxPackage, num_id: int) -> int:
    """Give this instance an abstract definition of its own.

    Several w:num can point at one w:abstractNum, so editing the shared
    definition would silently reformat every other list using it. Cloning
    first is what keeps a per-list change per-list.
    """
    import copy

    root = _numbering_root(pkg)
    abstract = _abstract_element(pkg, num_id)
    num = _num_element(pkg, num_id)
    clone = copy.deepcopy(abstract)
    new_id = max(
        (int(a.get(qn("w:abstractNumId"), "0") or 0)
         for a in root.findall(qn("w:abstractNum"))),
        default=-1,
    ) + 1
    clone.set(qn("w:abstractNumId"), str(new_id))
    for nsid in clone.findall(qn("w:nsid")):
        clone.remove(nsid)  # a fresh definition, not the same one twice
    nums = root.findall(qn("w:num"))
    if nums:
        nums[0].addprevious(clone)
    else:
        root.append(clone)
    num.find(qn("w:abstractNumId")).set(qn("w:val"), str(new_id))
    return new_id


def _shared_abstract(pkg: DocxPackage, num_id: int) -> list[int]:
    info = describe_numbering(pkg, num_id)
    return info.get("shares_definition_with", [])


def _set_start_override(pkg: DocxPackage, num_id: int, level: int, start: int):
    """The restart mechanism: a w:lvlOverride carrying a w:startOverride.

    This is the construct that makes "restart at 1" work without a second
    abstract definition, and the one a numbering implementation that reads
    only abstractNum gets wrong.
    """
    num = _num_element(pkg, num_id)
    override = _override_of(num, level)
    if override is None:
        override = etree.Element(qn("w:lvlOverride"))
        override.set(qn("w:ilvl"), str(level))
        ref = num.find(qn("w:abstractNumId"))
        if ref is not None:
            ref.addnext(override)
        else:
            num.insert(0, override)
    existing = override.find(qn("w:startOverride"))
    if existing is not None:
        override.remove(existing)
    el = etree.Element(qn("w:startOverride"))
    el.set(qn("w:val"), str(start))
    override.insert(0, el)  # startOverride precedes lvl in CT_NumLvl
    return override


def set_numbering(
    pkg: DocxPackage,
    num_id: int,
    *,
    restart_at: int | None = None,
    level: int = 0,
    continue_from: int | None = None,
    levels: list | None = None,
    force: bool = False,
) -> dict:
    """Change how one list numbers itself.

    restart_at writes a startOverride on that level of this instance, so
    the list starts at N without disturbing any other list sharing its
    definition. continue_from merges this list into another instance, so
    the two count as one run. levels rewrites the per-level formats, and
    clones the abstract definition first when other lists share it.
    """
    if _num_element(pkg, num_id) is None:
        raise TargetNotFound(
            f"no numbering instance with num_id {num_id}; read the ids from "
            f"list_elements(type='lists')"
        )
    if not 0 <= level <= 8:
        raise WordMcpError("level must be 0-8")
    if continue_from is not None and (restart_at is not None or levels):
        raise WordMcpError(
            "continue_from hands this list over to another instance, so a "
            "restart or a format change in the same call would contradict "
            "it; make them two calls"
        )
    if restart_at is None and continue_from is None and not levels:
        raise WordMcpError(
            "nothing to change: give restart_at, continue_from, or levels"
        )

    out: dict = {"num_id": num_id, "changed": []}
    if continue_from is not None:
        return _continue_numbering(pkg, num_id, continue_from, force=force)
    if restart_at is not None:
        if int(restart_at) < 0:
            raise WordMcpError("restart_at cannot be negative")
        _set_start_override(pkg, num_id, level, int(restart_at))
        out["changed"].append("restart_at")
        out["restart_at"] = int(restart_at)
        out["level"] = level
    if levels:
        specs = _normalize_levels(levels)
        shared = _shared_abstract(pkg, num_id)
        if shared:
            new_abs = _clone_abstract(pkg, num_id)
            out["cloned_definition"] = {
                "abstract_num_id": new_abs,
                "was_shared_with": shared,
                "why": (
                    "the definition was shared, so editing it would have "
                    "reformatted those lists too"
                ),
            }
        abstract = _abstract_element(pkg, num_id)
        if abstract is None:
            raise UnsupportedStructure(
                f"num_id {num_id} does not resolve to an abstractNum, so "
                f"there is no level definition to edit"
            )
        touched = []
        for lvl_index, spec in sorted(specs.items()):
            lvl = _lvl_of(abstract, lvl_index)
            if lvl is None:
                raise UnsupportedStructure(
                    f"this list's definition has no level {lvl_index}; Word "
                    f"writes all nine, so this file's numbering is unusual "
                    f"and editing it would be guesswork"
                )
            touched.append({"level": lvl_index,
                            "set": _write_level(lvl, spec)})
        out["changed"].append("levels")
        out["levels"] = touched
    pkg.mark_dirty("word/numbering.xml")
    out["definition"] = describe_numbering(pkg, num_id)
    return out


def _continue_numbering(
    pkg: DocxPackage, num_id: int, target: int, *, force: bool
) -> dict:
    """Fold one list into another so they count as a single run.

    Two paragraphs number together exactly when they share a numId, so
    this rewrites the numPr of every paragraph in the source list. When
    the two lists look different, that difference is visible and the call
    refuses unless the caller says it is intended.
    """
    if _num_element(pkg, target) is None:
        raise TargetNotFound(
            f"no numbering instance with num_id {target} to continue from"
        )
    if target == num_id:
        raise WordMcpError("a list already continues itself")
    source_fmt = effective_level(pkg, num_id, 0)
    target_fmt = effective_level(pkg, target, 0)
    differs = [
        key for key in ("format", "text")
        if source_fmt[key] != target_fmt[key]
    ]
    if differs and not force:
        raise UnsupportedStructure(
            f"list {num_id} and list {target} do not look alike "
            f"({', '.join(differs)} differ: "
            f"{ {k: source_fmt[k] for k in differs} } against "
            f"{ {k: target_fmt[k] for k in differs} }), so continuing one "
            f"into the other would change how it renders. Pass force if "
            f"that is intended."
        )
    moved = 0
    for numpr in pkg.root().iter(qn("w:numPr")):
        nid = numpr.find(qn("w:numId"))
        if nid is not None and nid.get(qn("w:val")) == str(num_id):
            nid.set(qn("w:val"), str(target))
            moved += 1
    if not moved:
        raise TargetNotFound(
            f"no paragraph in the document body carries num_id {num_id}, "
            f"so there is nothing to continue"
        )
    pkg.mark_dirty()
    return {
        "num_id": target,
        "changed": ["continue_from"],
        "paragraphs_moved": moved,
        "continued_from": num_id,
        "reformatted": bool(differs),
        "definition": describe_numbering(pkg, target),
    }


def _ensure_list_style(pkg: DocxPackage) -> None:
    root = pkg.root("word/styles.xml")
    have = {s.get(qn("w:styleId")) for s in root.findall(qn("w:style"))}
    if "ListParagraph" in have:
        return
    s = etree.SubElement(root, qn("w:style"))
    s.set(qn("w:type"), "paragraph")
    s.set(qn("w:styleId"), "ListParagraph")
    etree.SubElement(s, qn("w:name")).set(qn("w:val"), "List Paragraph")
    etree.SubElement(s, qn("w:basedOn")).set(qn("w:val"), "Normal")
    ppr = etree.SubElement(s, qn("w:pPr"))
    ind = etree.SubElement(ppr, qn("w:ind"))
    ind.set(qn("w:left"), "720")
    cs = etree.SubElement(ppr, qn("w:contextualSpacing"))
    pkg.mark_dirty("word/styles.xml")


def add_list(
    pkg: DocxPackage,
    items: list,
    *,
    kind: str = "bullet",
    after_index: int | None = None,
    after_anchor: str | None = None,
    at_end: bool = False,
    continue_from: int | None = None,
    start_at: int | None = None,
    levels: list | None = None,
) -> dict:
    """Insert a bulleted or numbered list. Items: strings, or dicts
    {text, level} with level 0-8 for nesting.

    A call creates its own numbering instance by default, so the list
    restarts at 1, which is what Word does for a new list.
    continue_from carries on an existing instance instead; start_at
    restarts the new one at N; levels configures the per-level formats.
    """
    if kind not in ("bullet", "number"):
        raise WordMcpError("kind must be 'bullet' or 'number'")
    if not items:
        raise WordMcpError("items must be a non-empty list")
    if continue_from is not None and (start_at is not None or levels):
        raise WordMcpError(
            "continue_from joins an existing list, so its numbering and "
            "formats come from that list; drop start_at and levels, or "
            "drop continue_from"
        )

    norm: list[tuple[str, int]] = []
    for item in items:
        if isinstance(item, dict):
            level = int(item.get("level", 0))
            if not 0 <= level <= 8:
                raise WordMcpError("level must be 0-8")
            norm.append((str(item["text"]), level))
        else:
            norm.append((str(item), 0))

    paragraphs, num_id = build_list_paragraphs(
        pkg, norm, kind, continue_from=continue_from, start_at=start_at,
        levels=levels,
    )

    from .text import _body_paragraph, _resolve_anchor

    body = pkg.body()
    if at_end or (after_index is None and after_anchor is None):
        sectpr = body.find(qn("w:sectPr"))
        for p in paragraphs:
            if sectpr is not None:
                sectpr.addprevious(p)
            else:
                body.append(p)
    elif after_anchor is not None:
        ref = _resolve_anchor(pkg, after_anchor)
        for p in reversed(paragraphs):
            ref.addnext(p)
    else:
        ref = _body_paragraph(pkg, after_index)
        for p in reversed(paragraphs):
            ref.addnext(p)
    pkg.mark_dirty()
    out = {"list_added": kind, "items": len(norm), "num_id": num_id}
    if continue_from is not None:
        out["continued"] = continue_from
    if start_at is not None:
        out["start_at"] = start_at
    return out


def build_list_paragraphs(
    pkg: DocxPackage, norm: list[tuple[str, int]], kind: str,
    *, continue_from: int | None = None, start_at: int | None = None,
    levels: list | None = None,
) -> tuple[list[etree._Element], int]:
    """Build (without inserting) the list's w:p elements, registering a
    fresh numbering instance in the package unless continue_from names one
    to reuse. norm: [(text, level 0-8)].
    Extracted from add_list so the batch layer's markdown list inserts
    share one construction path. Mutates numbering/styles parts only; the
    caller splices the paragraphs and calls mark_dirty()."""
    _ensure_numbering_part(pkg)
    _ensure_list_style(pkg)
    if continue_from is not None:
        if _num_element(pkg, continue_from) is None:
            raise TargetNotFound(
                f"no numbering instance with num_id {continue_from} to "
                f"continue; read the ids from list_elements(type='lists')"
            )
        num_id = int(continue_from)
    else:
        num_id = _new_numbering(pkg, kind)
        if levels:
            for lvl_index, spec in sorted(_normalize_levels(levels).items()):
                lvl = _lvl_of(_abstract_element(pkg, num_id), lvl_index)
                if lvl is not None:
                    _write_level(lvl, spec)
        if start_at is not None:
            if int(start_at) < 0:
                raise WordMcpError("start_at cannot be negative")
            _set_start_override(pkg, num_id, 0, int(start_at))
        pkg.mark_dirty("word/numbering.xml")

    from . import _runmap

    paragraphs = []
    for text, level in norm:
        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        pstyle = etree.SubElement(ppr, qn("w:pStyle"))
        pstyle.set(qn("w:val"), "ListParagraph")
        numpr = etree.SubElement(ppr, qn("w:numPr"))
        ilvl = etree.SubElement(numpr, qn("w:ilvl"))
        ilvl.set(qn("w:val"), str(level))
        nid = etree.SubElement(numpr, qn("w:numId"))
        nid.set(qn("w:val"), str(num_id))
        r = etree.SubElement(p, qn("w:r"))
        if text:
            t = etree.SubElement(r, qn("w:t"))
            t.text = text
            _runmap._preserve_space(t)
        paragraphs.append(p)
    return paragraphs, num_id


def get_lists(pkg: DocxPackage) -> list[dict]:
    """List paragraphs grouped by numbering instance (numId).

    Each item carries the number a reader SEES, computed the way Word
    computes it, and each list carries the definition behind it: the
    per-level formats, which levels the instance overrides, and which
    other lists share its definition (editing a shared one reformats them
    all, which is why set_list_numbering clones first).
    """
    from .read import body_items, paragraph_text

    numbers = {
        entry["paragraph_index"]: entry for entry in computed_numbers(pkg)
    }
    out: dict[int, list] = {}
    for kind_, idx, el in body_items(pkg):
        if kind_ != "paragraph":
            continue
        numpr = el.find(f"{qn('w:pPr')}/{qn('w:numPr')}")
        if numpr is None:
            continue
        nid_el = numpr.find(qn("w:numId"))
        ilvl_el = numpr.find(qn("w:ilvl"))
        if nid_el is None:
            continue
        nid = int(nid_el.get(qn("w:val"), "0"))
        item = {
            "paragraph_index": idx,
            "level": int(ilvl_el.get(qn("w:val"), "0")) if ilvl_el is not None else 0,
            "text": paragraph_text(el),
        }
        computed = numbers.get(idx)
        if computed is not None:
            item["number"] = computed["number"]
            if "unrendered_format" in computed:
                item["unrendered_format"] = computed["unrendered_format"]
        out.setdefault(nid, []).append(item)
    result = []
    for nid, items in sorted(out.items()):
        entry = {"num_id": nid, "items": items}
        try:
            definition = describe_numbering(pkg, nid)
        except TargetNotFound:
            entry["definition_missing"] = (
                "these paragraphs reference a numbering instance that "
                "word/numbering.xml does not define, so Word renders them "
                "with no number at all"
            )
        else:
            entry["format"] = definition["levels"][0]["format"]
            entry["levels_used"] = sorted({i["level"] for i in items})
            entry["definition"] = definition
        result.append(entry)
    return result
