"""Source-wide em-dash guard, ported from KitchenSink4Web's
tests/unit/test_copy_guards.py (test_no_em_dashes_anywhere_public).

Web checks the strings a session actually serves: tool descriptions, the
server instructions, the projection payload, and the published pages. This
repo already carries the runtime half of that (test_v2_docstring_budget
checks every tool description, test_v2_envelope checks every refusal hint,
test_v2_public_copy checks README/docs). What none of them could see is a
hint that lives in ops/ or com/ and only reaches a user on the failure path
it belongs to, which is exactly where the 2026-09-14 sweep found about a
hundred em dashes. This guard reads the SOURCE instead of the surface, so a
string is covered whether or not a test ever triggers it.

SCOPE, stated so the gap is deliberate rather than forgotten:

- IN: every string literal under src/ and scripts/ that is not a docstring.
  That is the runtime copy: refusal hints, warnings, notes, report headers,
  and the workflow prose get_workflows serves verbatim.
- IN: every docstring in server.py, because that file is where the served
  tool descriptions are written.
- OUT: comments and the docstrings of internal modules under ops/ and com/.
  Those are developer text; no user or model reads them. They are also
  where roughly a hundred em dashes still live, and rewriting engineering
  notes is not what the copy rule is for.

The allowlist below is the mapping's functional-exclusion ruling, entry for
entry: em dashes inside regex patterns and dash-matching data MATCH DASHES
IN USER DOCUMENTS. They are functionality, not copy, and removing one would
break citation conversion. Each entry is the exact literal, so a NEW em dash
in one of these files still fails; only the listed patterns pass.
"""

from __future__ import annotations

import ast
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EM_DASH = "—"
SCAN_DIRS = ("src", "scripts")

#: relative posix path -> the exact literals allowed to carry an em dash.
ALLOWLIST: dict[str, tuple[str, ...]] = {
    # Citation-locator patterns: the character classes match the en/em dash a
    # user typed in a page range ("pp. 10—12"), so the dash is input data.
    "src/word_mcp/ops/styleconvert.py": (
        r"^(?P<cont>.+?),\s*(?P<vol>\d+)\s*(?:\((?P<iss>[^)]+)\))?"
        r"(?:,\s*(?P<pp>[\dexvi\-–—,\s]+?))?\.?\s*$",
        r")(?:\s*[,:]\s*(?:pp?\.\s*)?(?P<loc>\d[\d\s,\-–—]*))?\s*$",
        r")(?:,\s*(?:pp?\.\s*)?(?P<loc>\d[\d\s,\-–—]*))?\))",
        r")(?:,\s*(?:pp?\.\s*)?(?P<pp>\d[\d\s,\-–—]*))?\)\.?\s*$",
    ),
    # Page-range normalizer and sentence-end detector: both read dashes the
    # user's document already contains.
    "src/word_mcp/ops/styleconvert_data.py": (
        r"\s*([-–—])\s*",
        r"[:.?!—-]$",
    ),
}


def _docstring_nodes(tree: ast.AST) -> set[int]:
    out: set[int] = set()
    for node in ast.walk(tree):
        body = getattr(node, "body", None)
        if isinstance(node, (ast.Module, ast.ClassDef, ast.FunctionDef,
                             ast.AsyncFunctionDef)) and body:
            first = body[0]
            if isinstance(first, ast.Expr) and \
                    isinstance(first.value, ast.Constant) and \
                    isinstance(first.value.value, str):
                out.add(id(first.value))
    return out


def _offending_strings() -> list[tuple[str, int, str]]:
    """(relative path, line, string) for every guarded literal with a dash."""
    hits: list[tuple[str, int, str]] = []
    for folder in SCAN_DIRS:
        base = ROOT / folder
        if not base.exists():
            continue
        for path in sorted(base.rglob("*.py")):
            text = path.read_text(encoding="utf-8")
            if EM_DASH not in text:
                continue
            rel = path.relative_to(ROOT).as_posix()
            allowed = ALLOWLIST.get(rel, ())
            tree = ast.parse(text)
            skip_docstrings = path.name != "server.py"
            docs = _docstring_nodes(tree) if skip_docstrings else set()
            for node in ast.walk(tree):
                if not isinstance(node, ast.Constant) or \
                        not isinstance(node.value, str):
                    continue
                if EM_DASH not in node.value or id(node) in docs:
                    continue
                if node.value in allowed:
                    continue
                hits.append((rel, node.lineno, node.value))
    return hits


def test_no_em_dashes_in_any_runtime_string():
    """Every served string in the package, in any language. A new em dash in
    a refusal hint fails here even if no test ever triggers that refusal."""
    hits = _offending_strings()
    assert not hits, "em dashes in runtime strings:\n" + "\n".join(
        f"  {rel}:{line}: {value[:90]!r}" for rel, line, value in hits
    )


def test_the_allowlist_still_matches_real_patterns():
    """An allowlist entry that no longer matches anything is a silent
    exemption waiting to cover the next mistake."""
    stale: list[str] = []
    for rel, patterns in ALLOWLIST.items():
        path = ROOT / rel
        assert path.exists(), f"allowlisted file {rel} no longer exists"
        text = path.read_text(encoding="utf-8")
        values = {
            node.value
            for node in ast.walk(ast.parse(text))
            if isinstance(node, ast.Constant) and isinstance(node.value, str)
        }
        for pattern in patterns:
            if pattern not in values:
                stale.append(f"{rel}: {pattern[:70]!r}")
    assert not stale, (
        "allowlisted literals are gone from the source; delete the entries "
        "instead of leaving a blanket exemption:\n" + "\n".join(stale)
    )
