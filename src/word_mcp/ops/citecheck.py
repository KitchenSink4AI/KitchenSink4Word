"""Citation-reference parity checking (APA-style, heuristic).

Cross-checks in-text citations against the reference list in both directions:
- citations whose (surname, year) has no reference entry  -> missing_references
- reference entries never cited in the body              -> uncited_references

This is a FLAGGING tool, not a fixer: APA parsing is heuristic (organizational
authors, legal citations, and unusual formats can evade the patterns), so
results are review candidates, not verdicts. Parenthetical (Author, 2020;
Other, 2021), narrative Author (2020), et al., &, and year-letters (2020a)
are handled.
"""

from __future__ import annotations

import re

from ..core.errors import TargetNotFound
from ..core.package import DocxPackage
from .read import get_outline, get_paragraphs

# Reference-list heading, matched by TEXT (localized): English plus Korean
# (참고문헌 / 참고 문헌), German (Literaturverzeichnis), French
# (Bibliographie), Spanish (Bibliografía), Japanese+Chinese (参考文献),
# Portuguese (Referências), Italian/Portuguese (Bibliografia). Shared by
# journalcount, styleconvert, and anonymize — extending this extends them all.
_REF_HEADINGS = re.compile(
    r"^\s*(references|bibliography|works cited|reference list"
    r"|참고\s*문헌"          # ko
    r"|literaturverzeichnis"  # de
    r"|bibliographie"         # fr (also de alternative)
    r"|bibliograf[íi]a"       # es / it+pt
    r"|参考文献"              # ja + zh
    r"|refer[êe]ncias"        # pt (accented and plain)
    r")\s*$",
    re.I,
)

# "n.d." (no date) is a legal APA year; hangul surnames are legal authors —
# both were invisible to the Latin-only patterns (v1.5 adversarial F4).
_YEAR = r"(?:(?:1[89]\d\d|20\d\d)[a-z]?|n\.d\.)"
# one author token: a capitalized Latin word (accents included — Müller,
# García, Łukasz; \w continues with any Unicode letter) OR a hangul word
_AUT = r"(?:[A-ZÀ-ÞĀ-Žƀ-Ƀ][\w'’\-.]*[\w'’.]|[가-힣]+)"
# Organizational authors are multi-word ("National Archives", "U.S.
# Department of War", "National Assembly of the Republic of Korea"), so an
# author is a PHRASE: capitalized words plus the lowercase connectors that
# hold them together (field test 2026-09-20, punchlist #867).
_CONNECT = r"(?:of|the|for|and|de|del|la|van|von|der|di|du|el|los|las)"
_PHRASE = rf"{_AUT}(?:\s+(?:{_AUT}|{_CONNECT}))*"
# Narrative: Smith (2026) / Smith and Jones (2026) / Smith et al. (2026) /
# National Archives (2003)
_NARRATIVE = re.compile(
    rf"(?<![\w가-힣])({_PHRASE})"
    rf"(?:\s+et al\.?)?"
    # the parenthesis may carry a page locator or the rest of an APA date
    rf"\s*\(({_YEAR})(?:,[^()]{{0,40}})?\)"
)
# Parenthetical chunk: everything ahead of the year in one semicolon-
# separated piece is the author field (Smith / Smith & Jones / Smith et
# al. / Memorandum of conversation), normalized afterwards.
_PAREN_CHUNK = re.compile(rf"^(.{{2,150}}?),\s*({_YEAR})")
# Reference entries: Surname, I. (2026). | Carter, J. (1977a, July 21). |
# National Archives. (2003). | 조선말 대사전. (1992).
# The author field is everything ahead of the first paren holding a year.
_REF_YEAR = re.compile(rf"\(({_YEAR})\b[^)]*\)")
# "Surname, A. B." — the personal-name pattern, which is what makes a
# last-token fallback safe to skip.
_PERSONAL = re.compile(rf"^\s*{_AUT}\s*,\s*(?:[A-ZÀ-ÞĀ-Ž]\.\s*)+")
_POSSESSIVE = re.compile(r"['’][sS]$")
# Leading words that are prose, not part of the author's name.
_LEAD_STOP = {
    "as", "in", "by", "see", "also", "e.g.", "i.e.", "cf.", "per", "from",
    "with", "and", "but", "however", "the", "a", "an", "for", "to", "of",
    "at", "on", "after", "before", "while", "since", "though", "although",
    "when", "where", "both", "either", "neither", "compare", "following",
    "according", "note", "notes", "such", "like", "unlike", "via",
}


def _norm_token(tok: str) -> str:
    tok = tok.strip(" .,;:()[]“”\"'’")
    tok = _POSSESSIVE.sub("", tok)
    return tok.lower()


def _norm_phrase(phrase: str) -> list[str]:
    """An author phrase as normalized tokens, with leading prose words
    ('as Smith (2020)') dropped. A paragraph boundary ends the phrase: it
    swallowed the preceding heading otherwise (adversarial review, m6)."""
    phrase = phrase.rsplit("\n", 1)[-1]
    toks = [t for t in (_norm_token(t) for t in phrase.split()) if t]
    while len(toks) > 1 and toks[0] in _LEAD_STOP:
        toks.pop(0)
    return toks


def _citation_keys(phrase: str, year: str) -> tuple[str, list[tuple[str, str]]]:
    """The canonical key for counting, plus every key worth matching a
    reference entry on: the whole phrase, its last word (organizational
    authors whose entry was parsed on another token), and its first."""
    toks = _norm_phrase(phrase)
    if not toks:
        return "", []
    whole = " ".join(toks)
    keys = [(whole, year)]
    if len(toks) > 1:
        keys.append((toks[-1], year))
        keys.append((toks[0], year))
    return whole, keys


def _reference_keys(author_field: str, year: str) -> list[tuple[str, str]]:
    toks = _norm_phrase(author_field)
    if not toks:
        return []
    keys = [(" ".join(toks), year), (toks[0], year)]
    if len(toks) > 1 and not _PERSONAL.match(author_field):
        # Organizational: a citation may name it by its last word.
        keys.append((toks[-1], year))
    return keys


def check_citation_parity(pkg: DocxPackage) -> dict:
    paras = get_paragraphs(pkg)
    outline = get_outline(pkg)

    ref_heading = next(
        (h for h in outline if _REF_HEADINGS.match(h["text"])), None
    )
    if ref_heading is None:
        # Fall back: any paragraph that IS exactly a reference heading.
        candidates = [
            p for p in paras if _REF_HEADINGS.match(p["text"] or "")
        ]
        if not candidates:
            raise TargetNotFound(
                "no References/Bibliography heading found; cannot locate the "
                "reference list"
            )
        ref_start = candidates[-1]["index"]
    else:
        ref_start = ref_heading["paragraph_index"]

    # End of reference list: next heading after ref_start, or document end.
    next_headings = [
        h["paragraph_index"]
        for h in outline
        if h["paragraph_index"] > ref_start
    ]
    ref_end = min(next_headings) if next_headings else None

    body_text = "\n".join(
        p["text"] for p in paras if p["index"] < ref_start
    )
    ref_paras = [
        p["text"]
        for p in paras
        if p["index"] > ref_start
        and (ref_end is None or p["index"] < ref_end)
        and p["text"].strip()
    ]

    # ---- collect in-text citations: canonical key -> {count, keys}
    cited: dict[tuple[str, str], dict] = {}

    def record(phrase: str, year: str) -> None:
        year = year.lower()
        whole, keys = _citation_keys(phrase, year)
        if not whole:
            return
        entry = cited.setdefault(
            (whole, year),
            {
                "count": 0,
                "keys": keys,
                "phrase": phrase.rsplit("\n", 1)[-1].strip(),
            },
        )
        entry["count"] += 1

    for m in _NARRATIVE.finditer(body_text):
        record(m.group(1), m.group(2))
    for paren in re.finditer(r"\(([^()]{4,300}?)\)", body_text):
        inner = paren.group(1)
        if not re.search(_YEAR, inner):
            continue
        for piece in inner.split(";"):
            m = _PAREN_CHUNK.match(piece.strip())
            if m:
                record(m.group(1), m.group(2))

    # ---- collect reference entries
    listed: dict[tuple[str, str], int] = {}  # key -> entry index
    entries: list[dict] = []
    unparsed: list[str] = []
    for entry in ref_paras:
        year = _REF_YEAR.search(entry)
        if not year:
            unparsed.append(entry[:120])
            continue
        author_field = entry[: year.start()]
        keys = _reference_keys(author_field, year.group(1).lower())
        if not keys:
            unparsed.append(entry[:120])
            continue
        idx = len(entries)
        entries.append({"text": entry[:120], "cited": False})
        for key in keys:
            listed.setdefault(key, idx)

    # ---- match, marking every entry a citation resolves to
    missing: list[str] = []
    missing_unparsed: list[str] = []
    for (whole, year), info in cited.items():
        hit = next((listed[k] for k in info["keys"] if k in listed), None)
        if hit is not None:
            entries[hit]["cited"] = True
            continue
        label = f"{info['phrase']} ({year})"
        # A single-token author is the personal-surname case the heuristic
        # handles well; a phrase is an organizational or unparsed author,
        # kept apart so the actionable list stays actionable (#867).
        if len(whole.split()) == 1:
            missing.append(f"{whole.title()} ({year})")
        else:
            missing_unparsed.append(label)

    uncited = sorted(e["text"] for e in entries if not e["cited"])
    missing = sorted(set(missing))
    missing_unparsed = sorted(set(missing_unparsed))

    out = {
        "in_text_citations": sum(i["count"] for i in cited.values()),
        "unique_cited_works": len(cited),
        "reference_entries": len(entries) + len(unparsed),
        "reference_entries_parsed": len(entries),
        "missing_references": missing,  # cited but not listed — serious
        "missing_references_unparsed": missing_unparsed,  # author unclear
        "uncited_references": uncited,  # listed but never cited — review
        "unparsed_reference_entries": unparsed,  # no year, no key
        "parity_ok": not missing and not missing_unparsed and not uncited,
        "note": (
            "Heuristic APA matching on (author, year), where the author is "
            "a surname or an organizational name phrase. missing_references "
            "holds single-surname citations with no entry; "
            "missing_references_unparsed holds citations whose author the "
            "heuristic could not resolve to an entry, which is where "
            "organizational authors and unusual formats land. Unparsed "
            "entries were not checked."
        ),
    }
    if missing_unparsed:
        # A citation the heuristic could not resolve is UNCHECKED, not
        # clean, so the check must not pass on it (round-4 MAJOR-1).
        out["missing_references_unparsed_reason"] = (
            "author not resolved; these citations were not checked against "
            "the reference list"
        )
    return out
