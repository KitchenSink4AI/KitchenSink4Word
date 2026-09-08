"""get_server_info: the support-bundle call, and the promise it makes.

The tool exists so a user can be told "paste what this returns" without
being told to redact anything first. That promise is the thing worth
guarding, so most of this file is one test asked several ways: no value
in the payload may be a file path, a user name, or anything else that
identifies the machine beyond its OS and Python.

The rest guards the contract the other family members already keep: the
call touches no document and starts no Word, it reports the surface the
pack registry actually holds, and it survives a machine where pywin32,
Word, the update cache, and the sandbox are each absent.
"""

from __future__ import annotations

import asyncio
import json
import os
import re
from datetime import datetime, timezone
from pathlib import Path

import pytest

from fastmcp import Client

from word_mcp import __version__, packs, server
from word_mcp.core import sandbox, update_check


@pytest.fixture
def info(monkeypatch):
    """A clean environment: no mode pin, no sandbox, no update cache."""
    for name in ("KS4W_MODE", "KS4W_PACK_POLICY", "KS4W_ALL_TOOLS",
                 "KS4W_LOCK_TOOLS", sandbox.ENV_VAR,
                 update_check.OPT_OUT_ENV):
        monkeypatch.delenv(name, raising=False)
    return server.get_server_info()


# ------------------------------------------------------------ the contents


def test_reports_the_running_build(info):
    assert info["name"] == "kitchensink4word"
    assert info["version"] == __version__


def test_reports_the_surface_the_registry_holds(info):
    surface = info["surface"]
    assert surface == packs.surface_report()
    assert surface["active_tools"] >= len(packs.pack_tools("lite"))
    assert set(info["packs_available"]) == set(packs.pack_names())


def test_reports_the_startup_configuration(info):
    assert info["startup_mode"] == "lite"
    assert info["surface_locked"] is False
    assert "lite core only" in info["startup_note"]


def test_startup_configuration_follows_the_environment(monkeypatch):
    monkeypatch.setenv("KS4W_MODE", "references,review")
    monkeypatch.setenv("KS4W_PACK_POLICY", "locked")
    out = server.get_server_info()
    assert out["startup_mode"] == "references,review"
    assert out["surface_locked"] is True


def test_reports_the_host(info):
    assert info["python"] == __import__("sys").version.split()[0]
    assert info["os"]


def test_word_tier_is_reported_without_starting_word(info):
    word = info["word"]
    assert set(word) >= {"application", "com_tools"}
    assert word["com_tools"] in ("available", "unavailable")
    if word["com_tools"] == "unavailable":
        assert word["note"], "an unavailable COM tier must say why"


def test_word_tier_survives_a_machine_without_pywin32(monkeypatch):
    """A Linux CI runner and a Windows box with no pywin32 must both get
    an answer rather than an exception."""
    import builtins

    real_import = builtins.__import__

    def refuse(name, *args, **kwargs):
        if name == "pythoncom":
            raise ImportError("no pywin32 here")
        return real_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", refuse)
    word = server._word_environment()
    assert word["com_tools"] == "unavailable"
    assert word["note"]


def test_sandbox_is_reported_as_shape_not_as_paths(monkeypatch, tmp_path):
    monkeypatch.setenv(sandbox.ENV_VAR, str(tmp_path))
    out = server.get_server_info()
    assert out["sandbox"] == {"active": True, "allowed_roots": 1}
    assert str(tmp_path) not in json.dumps(out)


def test_root_count_matches_the_env(monkeypatch, tmp_path):
    monkeypatch.delenv(sandbox.ENV_VAR, raising=False)
    assert sandbox.root_count() == 0
    second = tmp_path / "b"
    second.mkdir()
    monkeypatch.setenv(
        sandbox.ENV_VAR, os.pathsep.join([str(tmp_path), str(second)])
    )
    assert sandbox.root_count() == 2
    assert sandbox.active() is True


# -------------------------------------------------------- the update check


def _enable(monkeypatch, tmp_path):
    """Switch the check on for one test (the suite turns it off globally),
    point it at a scratch cache, and make sure nothing here can reach the
    network: an unpatched fetch raises."""
    monkeypatch.delenv(update_check.OFF_ENV, raising=False)
    monkeypatch.delenv(update_check.LEGACY_OFF_ENV, raising=False)
    monkeypatch.setenv(update_check.CACHE_DIR_ENV, str(tmp_path))
    monkeypatch.setattr(update_check, "_fetch", _no_network)


def _no_network(*args, **kwargs):
    raise AssertionError("a test reached for the network")


def _fresh(latest):
    """A cache young enough that no test will trigger a fetch."""
    return {"last_check": datetime.now(timezone.utc).isoformat(),
            "last_success": datetime.now(timezone.utc).isoformat(),
            "latest_version": latest, "ok": True}


def test_update_check_off_is_reported(monkeypatch):
    monkeypatch.setenv(update_check.OFF_ENV, "off")
    out = server.get_server_info()
    assert out["update_check"]["state"] == "disabled"
    assert out["update_check"]["disabled_by"] == update_check.OFF_ENV
    assert "update" not in out


def test_update_check_legacy_opt_out_is_reported(monkeypatch):
    monkeypatch.delenv(update_check.OFF_ENV, raising=False)
    monkeypatch.setenv(update_check.LEGACY_OFF_ENV, "true")
    out = server.get_server_info()
    assert out["update_check"]["state"] == "disabled"
    assert "update" not in out


def test_update_check_without_a_cache_says_it_does_not_know(
    monkeypatch, tmp_path
):
    """No cache and no answer: the report says unknown and admits the
    index was not reached. It never goes missing and never guesses."""
    _enable(monkeypatch, tmp_path)
    out = server.get_server_info()["update_check"]
    assert out["state"] == "unknown"
    assert out["reachable"] is False
    assert out["last_successful_check"] is None
    assert "latest_version" not in out


def test_update_check_reports_a_newer_release(monkeypatch, tmp_path):
    _enable(monkeypatch, tmp_path)
    update_check.write_cache(_fresh("99.0.0"), tmp_path / "update-check.json")
    out = server.get_server_info()
    assert out["update_check"]["state"] == "update_available"
    assert out["update_check"]["latest_version"] == "99.0.0"
    assert out["update_check"]["install_note"]
    assert "99.0.0" in out["update"]


def test_update_check_never_leaks_its_cache_path(monkeypatch, tmp_path):
    _enable(monkeypatch, tmp_path)
    update_check.write_cache(_fresh("0.0.1"), tmp_path / "update-check.json")
    out = server.get_server_info()
    assert str(tmp_path) not in json.dumps(out)
    assert out["update_check"]["state"] == "current"


def test_update_check_survives_an_unreadable_cache(monkeypatch, tmp_path):
    _enable(monkeypatch, tmp_path)
    (tmp_path / "update-check.json").write_text("{not json", encoding="utf-8")
    out = server._update_check_status()
    assert out["state"] == "unknown"
    assert out["reachable"] is False


# --------------------------------------------------------- the paste-safety
# The promise: everything above can be pasted into a public bug report
# unedited. These run over the SERIALIZED payload, so a path buried in a
# nested value fails the same way a top-level one does.


def _values(payload) -> list[str]:
    """Every leaf rendered as text, so the scans below cannot be dodged by
    nesting."""
    out: list[str] = []
    if isinstance(payload, dict):
        for value in payload.values():
            out.extend(_values(value))
    elif isinstance(payload, (list, tuple)):
        for value in payload:
            out.extend(_values(value))
    else:
        out.append(str(payload))
    return out


#: Absolute-path shapes, in the two flavours that matter: a drive letter
#: followed by a separator, and a POSIX path with at least two segments.
_PATH = re.compile(r"[A-Za-z]:[\\/]|(?<![\w.])/[\w.-]+/[\w.-]+")


def test_no_value_looks_like_a_path(info):
    offenders = [v for v in _values(info) if _PATH.search(v)]
    assert not offenders, f"path-shaped values in get_server_info: {offenders}"


def test_no_value_carries_the_user_name_or_home_directory(info):
    home = str(Path.home())
    user = os.environ.get("USERNAME") or os.environ.get("USER") or ""
    blob = json.dumps(info)
    assert home not in blob
    if user and len(user) > 2:
        assert user.lower() not in blob.lower(), (
            "the payload names the current user"
        )


def test_no_open_document_is_named(info):
    """com_word_status reports open documents by full path; this call must
    not, which is why it never attaches to a running Word."""
    assert "open_documents" not in json.dumps(info)
    assert ".docx" not in json.dumps(info)


def test_the_payload_is_json_serializable(info):
    """A support bundle that cannot be serialized is not a support
    bundle."""
    assert json.loads(json.dumps(info)) == info


# ------------------------------------------------------------ registration


def test_registered_read_only_in_the_lite_core():
    assert packs.pack_of("get_server_info") == "lite"
    tool = packs._REGISTRY["lite"]["get_server_info"]
    assert tool.annotations.readOnlyHint is True


def test_reachable_over_the_wire_at_startup():
    """It is an orient call, so it has to answer before anyone enables a
    pack."""

    async def run():
        async with Client(server.mcp) as client:
            result = await client.call_tool("get_server_info", {})
            return json.loads(result.content[0].text)

    payload = asyncio.run(run())
    assert payload["name"] == "kitchensink4word"
    assert payload["ok"] is True
