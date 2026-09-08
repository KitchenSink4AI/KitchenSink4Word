"""THE PRE-SHIP GREP GATE: a check that reads the artifact.

Every publishing defect the 2026-09-08 audit found was catchable by one
grep, and not one of them was caught, because what got checked was the
report about the file rather than the file. An order was given, an agent
reported compliance, and nobody opened the manifest. This is the check that
opens it.

Five patterns, all of them things that mean "somebody was still working
here" and none of them things a stranger should ever read:

- ``COPY PENDING`` and ``FACTS TO CONVEY``, the build's own instructions to
  a future writer. Six of these shipped on a sibling's install screen.
- ``PLACEHOLDER``, the marker this product was printing at users out of its
  update-check strings until the 2026-09-09 copy fill wave replaced them.
- The em dash, banned across every public surface in every language.
- A ``user_config`` entry typed anything but ``boolean``. The install screen
  is checkboxes; a free-text box is a typo waiting to break somebody's
  install, and the manifest schema has no enum type that would make one
  safe. ``EXEMPT_TYPED_FIELDS`` names the one field the author ruled an
  exception, and naming it here is what keeps the exception from spreading.
- An absolute ``C:\\Users\\`` path, which is a build that only runs on the
  machine that built it.

SCOPE. Everything in ``ENFORCED`` is clean of all five as of the commit
that adds this file, so the gate is green the day it lands and every hit
after that is a regression. ``docs/index.html`` joined that list on
2026-09-09, after four landing pages shipped figures their own READMEs
contradicted; the published-figure guard at the bottom of this file
covers the same class from the other direction. The ``src/`` trees are
deliberately NOT in scope: the em dashes in their refusal strings are an
author decision, not a gate decision. This is the KS4Web gate (``tests/unit/
test_ship_gate.py`` there), ported by the 2026-09-09 copy fill wave so the
family is checked by one rule rather than by one product's discipline.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[2]

#: Text surfaces scanned. Every one is clean at the commit that adds this
#: gate. A surface a stranger reads belongs here; add it rather than
#: trusting that nobody edits it.
ENFORCED = (
    "bundle/manifest.json",
    "README.md",
    "docs/llms.txt",
    "docs/QUICKSTART.md",
    "docs/index.html",
)

#: (needle, what it means) for the plain-text patterns. Written as literals
#: rather than a regex so the failure message can quote the needle back.
FORBIDDEN_TEXT = (
    ("COPY PENDING", "a copy marker aimed at a future writer"),
    ("FACTS TO CONVEY", "build instructions, not product copy"),
    ("PLACEHOLDER", "an unfilled string"),
    ("\u2014", "an em dash, banned on every public surface"),
)

#: The absolute-path pattern, in both the spellings a JSON file can hold it.
HOME_PATHS = ("C:\\Users\\", "C:\\\\Users\\\\")

#: The typed install-screen fields the author ruled exceptions, with the
#: type each is allowed to be. Empty for every product but KS4XL, whose
#: folder picker was ruled in by name because a directory chooser is a
#: picker rather than a free-text box.
EXEMPT_TYPED_FIELDS: dict[str, str] = {}


def _read(rel: str) -> str:
    return (ROOT / rel).read_text(encoding="utf-8")


def _sources() -> dict[str, str]:
    return {rel: _read(rel) for rel in ENFORCED}


@pytest.mark.parametrize("needle,why", FORBIDDEN_TEXT)
def test_no_working_state_reaches_a_shipping_surface(needle, why):
    for where, text in _sources().items():
        assert needle not in text, (
            f"{where} carries {needle!r}: {why}. This is the pattern the "
            f"2026-09-08 audit found on a sibling's release candidate; fix "
            f"the surface, do not loosen the gate.")


def test_no_absolute_home_path_reaches_a_shipping_surface():
    """A hardcoded user directory means the artifact runs on exactly one
    machine. A sibling's first release candidate launched
    ``C:\\Users\\...\\venv\\Scripts\\...exe`` and could not have installed
    for anybody else."""
    for where, text in _sources().items():
        for pattern in HOME_PATHS:
            assert pattern not in text, (
                f"{where} hardcodes {pattern!r}. The bundle has to launch on "
                f"a machine nobody here owns.")


def test_every_install_screen_field_is_a_checkbox():
    """No typed field on the install screen except the ones ruled in by
    name. The rule is stated as a scan rather than as a parity assertion,
    so it holds for the manifest a human installs from."""
    for where in ENFORCED:
        if not where.endswith(".json"):
            continue
        config = json.loads(_read(where)).get("user_config", {})
        typed = sorted(
            key for key, entry in config.items()
            if entry.get("type") != "boolean"
            and entry.get("type") != EXEMPT_TYPED_FIELDS.get(key))
        assert not typed, (
            f"{where} ships typed install-screen fields {typed}. The install "
            f"screen is checkboxes; a value a user has to spell correctly "
            f"belongs in a launch file, not on this screen. An exception is "
            f"added to EXEMPT_TYPED_FIELDS by an author ruling, never by a "
            f"wave that found the gate inconvenient.")


def test_every_enforced_surface_is_still_a_real_file():
    """A row pointing at a renamed file would quietly turn into no gate at
    all, which is how a total scope stops being total."""
    for rel in ENFORCED:
        assert (ROOT / rel).is_file(), rel


#: Surfaces cut from a measured source. Every figure below is published on
#: all three, so a wave that restamps one and forgets another turns this
#: red. The 2026-09-09 stale landing pages lived in exactly that gap: the
#: README said one number and docs/index.html said another, and nothing
#: checked either against the other.
FIGURE_SURFACES = (
    "README.md",
    "docs/index.html",
    "docs/llms.txt",
)

#: (figure, what it measures, the script that produces it). Values are the
#: measurement at the commit that adds this guard. Change one only together
#: with a re-run of the named script.
PUBLISHED_FIGURES = (
    ('221', 'document operations', 'scripts/count_operations.py'),
    ('110', 'tools', 'scripts/measure_surface.py'),
    ('1,839', 'tests', 'pytest --collect-only'),
    ('8,000', 'tokens at lite start', 'scripts/measure_surface.py'),
    ('27,800', 'tokens at full surface', 'scripts/measure_surface.py'),
)

#: (string, what it used to mean). Absent from every FIGURE_SURFACES file at
#: this commit. A hit means a surface went backwards.
SUPERSEDED_FIGURES = (
    ('1,324', 'the pre-2.1.0 test count'),
    ('1.324', 'the same count with a german or spanish separator'),
    ('1 324', 'the same count with a french separator'),
    ('219 document ops', 'the pre-consolidation operation count'),
)


@pytest.mark.parametrize("figure,what,source", PUBLISHED_FIGURES)
def test_every_published_figure_reads_the_same_on_every_surface(figure, what, source):
    """A figure published on three surfaces has to be the same figure on all
    three. Checking presence rather than equality keeps the guard honest
    about locale number formatting, which differs by design."""
    for rel in FIGURE_SURFACES:
        assert figure in _read(rel), (
            f"{rel} does not carry {figure!r} ({what}, measured by {source}). "
            f"Either the surface was missed by a restamp or the figure moved "
            f"and this table was not updated. Re-run the script, do not "
            f"guess the number.")


@pytest.mark.parametrize("stale,what", SUPERSEDED_FIGURES)
def test_no_superseded_figure_survives_on_a_published_surface(stale, what):
    """The regression guard for the defect this file was widened over."""
    for rel in FIGURE_SURFACES:
        assert stale not in _read(rel), (
            f"{rel} still carries {stale!r}, {what}. Restamp it from the "
            f"measuring script rather than deleting this row.")


def test_every_figure_surface_is_still_a_real_file():
    for rel in FIGURE_SURFACES:
        assert (ROOT / rel).is_file(), rel
