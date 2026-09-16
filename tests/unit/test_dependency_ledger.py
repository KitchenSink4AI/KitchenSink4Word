"""The dependency license ledger, enforced rather than maintained by hope.

Ported from KitchenSink4Web, which carried the only ledger in the family
until the 2026-09-15 license audit (finding D-02). A dependency added to
pyproject.toml and not to DEPENDENCY_LEDGER.md fails here, which is the only
way a ledger stays true past the first week.

The check is deliberately blunt: it compares the declared dependency NAMES
against the ledger's tables. It does not resolve licenses from PyPI at test
time, because a test that needs the network is a test that fails on a plane
and gets deleted. The license column is a human judgment recorded in the
ledger; this test makes sure nothing enters the install without one.
"""

from __future__ import annotations

import re
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
LEDGER = ROOT / "DEPENDENCY_LEDGER.md"

#: License strings that may appear against a REQUIRED dependency.
PERMISSIVE = ("apache-2.0", "mit", "bsd", "psf", "isc", "python-2.0")

#: Anything matching these in the ledger is a hard failure in the required
#: install. Source-available licenses are included because "not OSI-approved
#: but you can read it" is the trap an SSPL relicense springs on a category.
FORBIDDEN_IN_REQUIRED = (
    "gpl", "sspl", "bsl", "busl", "elastic license",
    "commons clause", "source-available", "cc-by-sa",
)


def _pyproject() -> dict:
    return tomllib.loads((ROOT / "pyproject.toml").read_text(
        encoding="utf-8"))


def _base_name(spec: str) -> str:
    """'fastmcp>=3.4,<4' -> 'fastmcp'; drops environment markers too."""
    return re.split(r"[<>=!~;\[ ]", spec.strip(), maxsplit=1)[0].lower()


def _ledger_rows() -> list[tuple[str, str]]:
    """(package, license) for every table row in the ledger."""
    rows = []
    for line in LEDGER.read_text(encoding="utf-8").splitlines():
        if not line.strip().startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        for i, cell in enumerate(cells[:-1]):
            m = re.fullmatch(r"`([A-Za-z0-9._-]+)`", cell)
            if m:
                rows.append((m.group(1).lower(), cells[i + 1].lower()))
                break
    return rows


def test_ledger_exists():
    assert LEDGER.is_file(), (
        "DEPENDENCY_LEDGER.md is required: every declared dependency needs a "
        "recorded license before it ships."
    )


def test_every_required_dependency_is_in_the_ledger():
    declared = {_base_name(d)
                for d in _pyproject()["project"].get("dependencies", [])}
    listed = {name for name, _lic in _ledger_rows()}
    missing = sorted(declared - listed)
    assert not missing, (
        f"required dependencies missing from DEPENDENCY_LEDGER.md: "
        f"{missing}. Add the package, its license, and why it is required "
        f"before it ships."
    )


def test_every_optional_dependency_is_in_the_ledger():
    optional = _pyproject()["project"].get("optional-dependencies", {})
    declared = {_base_name(d) for group in optional.values() for d in group}
    listed = {name for name, _lic in _ledger_rows()}
    missing = sorted(declared - listed)
    assert not missing, f"optional dependencies missing from the ledger: {missing}"


def test_no_copyleft_in_the_required_install():
    declared = {_base_name(d)
                for d in _pyproject()["project"].get("dependencies", [])}
    for name, license_text in _ledger_rows():
        if name not in declared:
            continue
        for bad in FORBIDDEN_IN_REQUIRED:
            assert bad not in license_text, (
                f"{name} is licensed {license_text!r} and is in the "
                f"REQUIRED install. Copyleft and source-available "
                f"dependencies live behind an optional extra so a "
                f"dependency's license never becomes the server's problem."
            )


def test_required_licenses_are_recognizably_permissive():
    declared = {_base_name(d)
                for d in _pyproject()["project"].get("dependencies", [])}
    for name, license_text in _ledger_rows():
        if name not in declared:
            continue
        assert any(p in license_text for p in PERMISSIVE), (
            f"{name} has license {license_text!r}, which is not on the "
            f"permissive list. Either it is fine and belongs on the list, "
            f"or it does not belong in the required install."
        )


def test_the_ledger_records_the_font_attribution():
    """The fonts on the docs page are the one thing this project actually
    redistributes, so they are the one real attribution obligation. The OFL
    files must exist and the ledger must say why."""
    text = LEDGER.read_text(encoding="utf-8")
    assert "OFL" in text and "SIL Open Font License" in text
    fonts = ROOT / "docs" / "fonts"
    for name in ("OFL-Fraunces.txt", "OFL-IBMPlexMono.txt"):
        assert (fonts / name).is_file(), (
            f"{name} is missing from docs/fonts/. The SIL OFL requires the "
            f"copyright notice and license to travel with the font software."
        )
