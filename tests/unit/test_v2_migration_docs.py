"""Docs consistency: the generated rename table in docs/MIGRATION_V2.md
must be current against src/word_mcp/data/v1_to_v2.json (regenerate-and-diff,
so a stale guide fails red), and the guide's live-routed tool list must
agree with the server's actual routing."""

from __future__ import annotations

import ast
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
GUIDE = ROOT / "docs" / "MIGRATION_V2.md"
BEGIN = ("<!-- BEGIN GENERATED RENAME TABLE "
         "(scripts/generate_migration_table.py) -->")
END = "<!-- END GENERATED RENAME TABLE -->"


def _routed_tools() -> set[str]:
    """v2 tools that call _route_live, from the server AST."""
    src = (ROOT / "src" / "word_mcp" / "server.py").read_text(
        encoding="utf-8"
    )
    routed = set()
    for node in ast.walk(ast.parse(src)):
        if not isinstance(node, ast.FunctionDef):
            continue
        deco = any(
            isinstance(d, ast.Call) and isinstance(d.func, ast.Name)
            and d.func.id == "_tool"
            for d in node.decorator_list
        )
        if not deco:
            continue
        if any(
            isinstance(sub, ast.Call) and (
                (isinstance(sub.func, ast.Name)
                 and sub.func.id == "_route_live")
            )
            for sub in ast.walk(node)
        ):
            routed.add(node.name)
    return routed


def test_rename_table_is_current():
    out = subprocess.run(
        [sys.executable, "-X", "utf8",
         str(ROOT / "scripts" / "generate_migration_table.py")],
        capture_output=True, text=True, encoding="utf-8", cwd=ROOT,
    )
    assert out.returncode == 0, out.stderr
    fresh = out.stdout.replace("\r\n", "\n").strip()
    text = GUIDE.read_text(encoding="utf-8")
    start, stop = text.find(BEGIN), text.find(END)
    assert start >= 0 and stop >= 0, "generated-table markers missing"
    embedded = text[start:stop + len(END)].replace("\r\n", "\n").strip()
    assert embedded == fresh, (
        "docs/MIGRATION_V2.md rename table is stale; rerun "
        "scripts/generate_migration_table.py --write"
    )


def test_guide_has_no_em_dashes():
    assert "—" not in GUIDE.read_text(encoding="utf-8")


def test_guide_live_list_matches_server_routing():
    """The migration guide's live-routed tool list stays truthful."""
    routed = _routed_tools()
    text = GUIDE.read_text(encoding="utf-8")
    m = re.search(
        r"Live-routed v2 tools: (.*?)\.\n\n", text, re.S
    )
    assert m, "live-routed list missing from the guide"
    listed = set(re.findall(r"`(\w+)`", m.group(1)))
    listed.discard("replace_paragraph_text")  # named as the old v1 name
    assert listed == routed, (
        f"guide says {sorted(listed)}, server routes {sorted(routed)}"
    )
