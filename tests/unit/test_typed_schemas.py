"""Guard: the schemas a client actually receives say what they accept.

The word equivalent of KitchenSink4XL's tests/unit/test_typed_schemas.py,
written for the defect the 2026-09-21 field test hit: ``live`` was declared
fourteen times as a bare ``live: str = "auto"``, so the advertised schema
was a string with no vocabulary, and the only record of the three legal
values was the refusal a caller reads after a failed call. Two calls were
spent guessing (``live="false"``, then ``live="file"``).

Two checks, deliberately different in kind. The first walks every ``live``
parameter on the whole surface and asserts the enum and a description,
which catches a new tool declaring the bare str the day it is written. The
second asserts that ``location`` did NOT grow a body, because that is the
ruling that keeps the session bill down and a silent regression in it costs
every session tokens without failing anything else.

The whole surface is exercised, not the lite subset: a pack a session
enables later ships the same schemas.
"""

from __future__ import annotations

import asyncio
import json

import pytest

from word_mcp import packs, server  # noqa: F401  (import registers tools)
from word_mcp.core import locate, schemas

#: Keys that carry no type information on their own.
_NON_TYPING_KEYS = {"title", "description", "default"}


def _tools() -> dict:
    return {t.name: t for t in asyncio.run(server.mcp.list_tools())}


def _params(tool) -> dict:
    return (getattr(tool, "parameters", None) or {}).get("properties", {}) or {}


def _named(param: str) -> list[tuple[str, dict]]:
    return sorted(
        (name, _params(tool)[param])
        for name, tool in _tools().items()
        if param in _params(tool)
    )


def test_every_live_parameter_advertises_the_three_routes():
    """The field-test defect, as a standing guard.

    A bare ``str`` annotation is valid and useless: it tells a client the
    parameter takes text and leaves ``auto | force | off`` discoverable
    only by getting it wrong. Every tool with a live route says the set.
    """
    found = _named("live")
    assert len(found) >= 14, (
        f"only {len(found)} live parameters walked; the walk broke"
    )
    for name, schema in found:
        assert schema.get("enum") == ["auto", "force", "off"], (
            f"{name}.live does not advertise the route vocabulary: {schema!r}"
        )
        assert schema.get("description"), (
            f"{name}.live has an enum but no sentence saying what it selects"
        )
        assert schema.get("type") == "string", (
            f"{name}.live stopped being a string: {schema!r}"
        )


def test_the_live_schema_matches_the_validator():
    """The schema and server._route_live's refusal are two copies of one
    vocabulary. Nothing stops them drifting except this."""
    import inspect

    source = inspect.getsource(server._route_live)
    assert 'live not in ("auto", "force", "off")' in source, (
        "the live validator changed shape; re-check LIVE_SCHEMA against it"
    )
    assert schemas.LIVE_SCHEMA["enum"] == ["auto", "force", "off"]


def test_live_still_accepts_every_documented_route():
    """The schema describes; the validator still decides. Typing must not
    narrow what the tools take."""
    for value in ("auto", "force", "off"):
        assert server._route_live(
            value, lambda: {"route": "file"}, lambda: {"route": "live"}
        ) in ({"route": "file"}, {"route": "live"})
    with pytest.raises(ValueError, match="auto|force|off"):
        server._route_live("file", lambda: {}, lambda: {})


def test_the_addressing_schema_stays_types_only():
    """The cost ruling, kept from drifting back by hand.

    Eighteen tools take a ``location``. Enumerating the selector grammar on
    each of them was measured in the sibling repo at about 270 tokens per
    parameter, and the doctrine is pay on error, not on load. If a future
    release wants the grammar in the schema, this test is where the
    decision gets re-made rather than re-happening.
    """
    found = _named("location")
    assert len(found) >= 18, f"only {len(found)} location parameters walked"
    for name, schema in found:
        branches = schema.get("anyOf") or [schema]
        kinds = {b.get("type") for b in branches}
        assert "object" in kinds, (
            f"{name}.location is no longer typed as an object: {schema!r}"
        )
        for branch in branches:
            assert not branch.get("properties"), (
                f"{name}.location grew a body again; it is billed on every "
                f"parameter that takes one: {branch!r}"
            )


def test_the_selector_vocabulary_is_documented_where_a_client_reads_it():
    """The grammar left the schema, so it must be somewhere a caller meets
    it: the server instructions at handshake, and the resolver's own
    refusals at the moment a caller gets the shape wrong."""
    text = server.mcp.instructions or ""
    for selector in locate.SELECTORS:
        assert selector in text, (
            f"the {selector!r} selector is not named in the server "
            "instructions, so nothing tells a client it exists"
        )
    grammar = locate.vocabulary()
    for selector in locate.SELECTORS:
        assert f'"{selector}"' in grammar, (
            f"{selector!r} has no worked form in the refusal vocabulary"
        )
    for position in locate.POSITIONS:
        assert position in grammar, (
            f"the {position!r} position is not named in the refusals"
        )


def test_a_bad_selector_shape_teaches_the_whole_grammar():
    """One failed call, not one per selector. The field test sent
    {"search": "text"} and got a refusal that named the search form only."""
    import tempfile
    from pathlib import Path

    from docx import Document

    from word_mcp.core.errors import WordMcpError
    from word_mcp.core.package import DocxPackage

    path = Path(tempfile.mkdtemp()) / "grammar.docx"
    d = Document()
    d.add_paragraph("plain body text")
    d.save(path)
    pkg = DocxPackage(str(path))

    with pytest.raises(WordMcpError) as exc:
        locate.resolve_location(pkg, {"search": "plain string"})
    message = str(exc.value)
    for selector in locate.SELECTORS:
        assert f'"{selector}"' in message, (
            f"the refusal does not teach the {selector!r} form: {message!r}"
        )
    assert "position" in message

    with pytest.raises(WordMcpError) as exc:
        locate.resolve_location(pkg, {"pargraph": 3})
    assert "paragraph" in str(exc.value)

    with pytest.raises(WordMcpError) as exc:
        locate.resolve_location(pkg, {"paragraph": 0, "outline": "1"})
    assert '"outline"' in str(exc.value)


def test_the_schema_is_inlined_rather_than_referenced():
    """Dereferenced on purpose. Gemini's function-calling schema subset does
    not resolve $ref, so a shared $defs entry would make the weakest clients
    skip the tool. The cost is repetition, paid knowingly."""
    for name, tool in sorted(_tools().items()):
        blob = json.dumps(getattr(tool, "parameters", None) or {})
        assert "$ref" not in blob, f"{name} ships a $ref"
        assert "$defs" not in blob, f"{name} ships a $defs"


def test_no_tool_parameter_emits_an_empty_schema():
    """A parameter annotated ``Any`` serializes to ``{}``, which is what
    made two clients delete the sibling server's tools outright (OpenCode:
    "could not understand the instance {}"; Gemini CLI: skipped as missing
    types). Nothing on this surface may reach a client that way."""
    offenders = []
    checked = 0
    for name, tool in sorted(_tools().items()):
        for param, schema in _params(tool).items():
            checked += 1
            body = {k: v for k, v in schema.items()
                    if k not in _NON_TYPING_KEYS}
            if not body:
                offenders.append(f"{name}.{param}")
    assert checked > 400, f"only {checked} parameters walked; the walk broke"
    assert not offenders, (
        "these parameters reach the client as an empty {} schema: "
        + ", ".join(offenders)
    )
