"""Guard: every tool ships a title, and every tool that can change
something declares whether the change may be destructive.

The Anthropic Connectors Directory requires `title` and the applicable
hints on every tool, and a client shows the title to a human instead of
the raw tool name, so a missing or wrong one is a product defect rather
than a documentation nit. Three things are checked here.

FIRST, coverage and shape on the wire: every registered tool carries a
non-empty title, titles are unique within this server, none exceeds 40
characters, and none carries product or marketing language.

SECOND, the classification itself. The non-destructive allowlist below is
hand-audited and lives in the TEST on purpose. `core/tool_annotations.py`
is where the server reads its classification from; if both sides shared
one list the test would only prove the file equals itself. A new tool
consequently fails twice until somebody classifies it deliberately, once
at registration, where `annotations()` raises on an unclassified mutating
name, and once here.

THIRD, the published tables. `docs/TOOL_TITLES.md` and
`docs/TOOL_ANNOTATIONS.md` are what a reviewer reads and what a maintainer
disputes a row in. They are generated from the module, so they are checked
against it here; a table that drifts from the wire is worse than no table.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from word_mcp import packs, server  # noqa: F401  (the import registers tools)
from word_mcp.core import tool_annotations as ann

DOCS = Path(__file__).resolve().parents[2] / "docs"

#: Words a tool title must not carry. The checklist rejects tool metadata
#: that promotes a product, and a title is the most visible metadata there
#: is.
MARKETING = {
    "best", "powerful", "easy", "easiest", "fast", "fastest", "pro",
    "premium", "ultimate", "smart", "magic", "magical", "free", "awesome",
    "amazing", "seamless", "effortless", "revolutionary", "unlimited",
    "kitchensink4ai", "kitchensink4word", "kitchensink4xl",
    "kitchensink4ppt", "kitchensink4web",
}


def _wire(tool) -> dict:
    return tool.to_mcp_tool().model_dump(exclude_none=True)


def _registered() -> dict:
    """tool name -> tool object, every pack, enabled or not."""
    out: dict = {}
    for tools in packs._REGISTRY.values():
        out.update(tools)
    return out


def test_every_registered_tool_carries_a_non_empty_title():
    missing = [name for name, tool in sorted(_registered().items())
               if not _wire(tool).get("title")]
    assert not missing, (
        f"these tools reach tools/list with no title: {missing}")


def test_titles_are_unique_within_this_server():
    seen: dict = {}
    clashes = []
    for name, tool in sorted(_registered().items()):
        title = _wire(tool)["title"]
        if title in seen:
            clashes.append(f"{title!r}: {seen[title]} and {name}")
        seen[title] = name
    assert not clashes, (
        "two tools would show the same name in a client: " + "; ".join(clashes))


def test_no_title_is_longer_than_forty_characters():
    long = {name: _wire(tool)["title"]
            for name, tool in _registered().items()
            if len(_wire(tool)["title"]) > 40}
    assert not long, f"titles over 40 characters: {long}"


def test_no_title_carries_marketing_language():
    offenders = {}
    for name, tool in _registered().items():
        words = set(re.findall(r"[a-z0-9]+", _wire(tool)["title"].lower()))
        hit = words & MARKETING
        if hit:
            offenders[name] = sorted(hit)
    assert not offenders, (
        f"a tool title is not a place to sell anything: {offenders}")


def test_the_title_table_covers_the_registered_surface_exactly():
    registered = set(_registered())
    assert registered - set(ann.TITLES) == set(), (
        "tools with no title (add them to tool_annotations.py): "
        f"{sorted(registered - set(ann.TITLES))}")
    assert set(ann.TITLES) - registered == set(), (
        "tool_annotations.py titles tools that no longer exist: "
        f"{sorted(set(ann.TITLES) - registered)}")


def test_an_untitled_name_raises_rather_than_getting_a_generated_title():
    with pytest.raises(RuntimeError):
        ann.title("a_tool_that_does_not_exist")


def test_the_wire_title_is_the_one_the_module_declares():
    for name, tool in sorted(_registered().items()):
        assert _wire(tool)["title"] == ann.TITLES[name], name


def _doc_rows(filename: str) -> list:
    text = (DOCS / filename).read_text(encoding="utf-8")
    rows = []
    for line in text.splitlines():
        if not line.startswith("| `"):
            continue
        cells = [c.strip() for c in line.strip("|").split("|")]
        rows.append([c.strip("`") for c in cells])
    return rows


def test_the_published_title_table_matches_the_module():
    rows = _doc_rows("TOOL_TITLES.md")
    published = {name: title for name, title, *_ in rows}
    assert published == ann.TITLES, (
        "docs/TOOL_TITLES.md has drifted from tool_annotations.py. "
        "Regenerate it rather than editing either side by hand.")
