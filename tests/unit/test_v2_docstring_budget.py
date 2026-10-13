"""v2 docstring budgets.

MULTIPLEX is the finalized v2 multiplex set. Tools that register later
(apply_edits, get_document_view, enable_tools) are pre-listed so they
inherit their cap without touching this file.
"""

from __future__ import annotations

import asyncio

from word_mcp import server


def _tools():
    """The shipped surface only. Some test files re-register a handful of
    tools on the shared FastMCP instance, so under a full-suite run some
    registry entries carry test text. Filter to tools whose function lives
    in a word_mcp module; the budgets police shipped code."""
    tools = asyncio.run(server.mcp.list_tools())
    return [
        t for t in tools
        if getattr(getattr(t, "fn", None), "__module__", "").startswith(
            "word_mcp"
        )
    ]


# The v2 multiplex/high-traffic tools that earn the ~350-token cap
# (delete_element is included).
MULTIPLEX = {
    "list_elements", "validate", "manage_note", "manage_comment",
    "manage_source", "modify_table_structure", "com_multi_document",
    "insert_reference_list", "apply_edits", "get_document_view",
    "search_and_replace", "resolve_revisions", "manage_backups",
    "enable_tools", "get_workflows", "diagnose_document",
    "insert_document", "convert_citation_style", "delete_element",
    # Both grew a placement grammar with the anchored-image work: six wrap
    # modes, two relative-frame vocabularies, and align-or-offset per axis.
    # An agent that cannot see the vocabulary cannot use the feature.
    "insert_image", "set_image",
    # And both list tools, for the same reason: the numbering suite gave
    # them a per-level format grammar (nine levels, nine numeral styles,
    # a label template) that an agent cannot use unseen.
    "insert_list", "set_list_numbering",
}


def test_no_em_dashes_in_descriptions():
    """Standing rule: no em dashes anywhere public. Keep it that way."""
    for tool in _tools():
        desc = tool.description or ""
        assert "—" not in desc, f"{tool.name} description has an em dash"


def test_docstring_budget():
    """Every tool description inside 60-130 tokens (chars/4; multiplex
    tools get headroom to ~350)."""
    for tool in _tools():
        desc = tool.description or ""
        assert desc, f"{tool.name} has no description"
        tokens = len(desc) / 4
        cap = 350 if tool.name in MULTIPLEX else 130
        assert 60 <= tokens <= cap, (
            f"{tool.name} description is ~{tokens:.0f} tokens, "
            f"outside [60, {cap}]"
        )
