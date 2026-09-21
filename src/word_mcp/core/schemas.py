"""core/schemas.py: shared JSON Schema for parameters several tools take.

THE DEFECT THIS CLOSES. ``live`` is declared fourteen times as a bare
``live: str = "auto"``. FastMCP derives the advertised schema from the bare
annotation, so every one of those parameters reaches the client as
``{"type": "string", "default": "auto"}``: a string, with no vocabulary and
no sentence saying what it selects. The only place the three legal values
are written down is the ValueError in server._route_live, which a caller
reads after a failed call. The 2026-09-21 field test spent two calls on it
(``live="false"`` then ``live="file"``, both refused with
``BAD_PARAMS: live must be auto|force|off``) before the refusal taught what
the schema should have.

THE APPROACH, PORTED FROM KitchenSink4XL. Validation behavior does not
change. ``WithJsonSchema`` replaces the ADVERTISED schema while pydantic
keeps coercing the underlying Python type, so ``live`` still arrives as the
plain str ``_route_live`` has always taken and ``_route_live`` keeps
raising its own refusal for a value outside the set. An enum is worth its
tokens here because the set is closed, three items long, and impossible to
guess.

WHAT IS DELIBERATELY NOT HERE: ``location``. Eighteen tools take one, and
the selector vocabulary was measured in the sibling repo at about 270
tokens on each parameter that carries it, roughly 9.5k tokens billed to
every session to describe keys most sessions never use
(xlsx_mcp/core/schemas.py LOCATION_SCHEMA, ruling 2026-09-06). The doctrine
is pay on error, not on load, so the vocabulary lives in the server
instructions, each tool's own description, core/locate.py's module
docstring, and above all core/locate.py's refusals, which name every
selector, every position and every candidate they found at the moment a
caller gets one wrong. A validator could never say that much. This repo
keeps the same ruling, and ``location`` stays annotated ``dict | None``.

Schemas are DEREFERENCED (FastMCP's default), not written with ``$ref``:
the clients this kind of typing exists for have the weakest schema parsers,
and Gemini's function-calling subset does not resolve ``$ref`` at all, so a
shared definition would reintroduce the same class of skip in a new form.
The schema below is therefore repeated inline on each parameter that takes
one, which is why its text is kept to one line.
"""

from __future__ import annotations

from typing import Annotated, Any

from pydantic.json_schema import WithJsonSchema

# ------------------------------------------------------------------- live

#: The dual-mode route selector every tool with a live path takes. The
#: values are the ones server._route_live enforces; keep the two in step.
#:
#: FORTY-FOUR CHARACTERS, and that is the budget, not a style choice. This
#: is billed on fourteen parameters, twelve of them on the lite surface,
#: so every character here costs about twelve times what it looks like.
#: The enum alone takes lite from 9,909 to about 10,011 tokens against the
#: 10,200 ceiling test_pack_bills_cost_aware holds; what is left buys
#: roughly one short clause. The enum is the part that closes the defect
#: (a caller can no longer invent a value), so the sentence only has to
#: say which way each value routes. What each route MEANS in full is in
#: the server instructions, which a client receives once at handshake
#: instead of once per parameter: the same pay-on-load arithmetic that
#: keeps the location grammar out of the schema.
LIVE_SCHEMA: dict[str, Any] = {
    "type": "string",
    "enum": ["auto", "force", "off"],
    "description": "auto: file or live; force: live; off: refuse",
}

Live = Annotated[str, WithJsonSchema(LIVE_SCHEMA)]

__all__ = ["LIVE_SCHEMA", "Live"]
