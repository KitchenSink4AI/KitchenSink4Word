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
import types
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


def test_names_the_product_the_agent_is_talking_to(info):
    """A connected agent could read this payload and still not know which
    product answered: the name field is a lowercase package id and nothing
    said the brand. 2.1.1 added the identity block."""
    assert info["product"] == "KitchenSink4Word"
    assert info["package"] == "kitchensink4word"
    assert info["homepage"] == "https://kitchensink4.ai/KitchenSink4Word/"


def test_names_the_rest_of_the_family(info):
    """The suite was invisible from inside the server. Every sibling's PyPI
    name is here, and this one is not in its own family list."""
    assert info["family"] == [
        "kitchensink4xl", "kitchensink4ppt", "kitchensink4web"
    ]
    assert "kitchensink4word" not in info["family"]


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
    assert word["com_tools"] in ("available", "unavailable", "unknown")
    if word["com_tools"] != "available":
        assert word["note"], "anything but a working COM tier must say why"


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


# ------------------------------------------- the Word registration probe
#
# BUG-001. The probe used to call pythoncom.CLSIDFromProgID, an attribute
# pywin32 does not have, and one broad except turned the resulting
# AttributeError into "no Word.Application registration found on this
# machine". Users with Word installed, registered and running were told
# they had no Word. The three cases below are the whole point of the fix:
# a registered Word, an absent one, and a probe that could not answer are
# three different facts and must never collapse into two.
#
# The fake registry lets all three run on any machine, Windows or not.


class _FakeKey:
    def __init__(self, value=""):
        self._value = value

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class _FakeWinreg:
    """Just enough winreg to drive the probe, with no registry behind it."""

    HKEY_CLASSES_ROOT = 0
    KEY_READ = 0x20019
    KEY_WOW64_64KEY = 0x0100
    KEY_WOW64_32KEY = 0x0200

    def __init__(self, keys=None, errors=None):
        self.keys = dict(keys or {})
        self.errors = dict(errors or {})
        self.opened = []

    def OpenKey(self, root, sub, reserved=0, access=0):
        self.opened.append(sub)
        if sub in self.errors:
            raise self.errors[sub]
        if sub in self.keys:
            return _FakeKey(self.keys[sub])
        raise FileNotFoundError(2, "no such key")

    def QueryValueEx(self, key, name):
        return (key._value, 1)


_CLSID = "{000209FF-0000-0000-C000-000000000046}"


def _with_registry(monkeypatch, fake):
    monkeypatch.setitem(__import__("sys").modules, "winreg", fake)
    return fake


def test_probe_reports_registered_when_progid_and_server_both_resolve(monkeypatch):
    fake = _with_registry(monkeypatch, _FakeWinreg(keys={
        "Word.Application\\CLSID": _CLSID,
        "CLSID\\" + _CLSID + "\\LocalServer32": "",
    }))
    state, detail = server._word_registration()
    assert state == "registered", detail
    assert "Word.Application\\CLSID" in fake.opened


def test_probe_accepts_an_inproc_server_too(monkeypatch):
    _with_registry(monkeypatch, _FakeWinreg(keys={
        "Word.Application\\CLSID": _CLSID,
        "CLSID\\" + _CLSID + "\\InprocServer32": "",
    }))
    assert server._word_registration()[0] == "registered"


def test_probe_reports_absent_when_the_progid_is_not_there(monkeypatch):
    _with_registry(monkeypatch, _FakeWinreg())
    state, detail = server._word_registration()
    assert state == "absent"
    assert detail


def test_probe_reports_absent_when_the_class_id_has_no_server(monkeypatch):
    """A leftover class id with no server entry is not an installation."""
    _with_registry(monkeypatch, _FakeWinreg(keys={
        "Word.Application\\CLSID": _CLSID,
    }))
    state, detail = server._word_registration()
    assert state == "absent"
    assert detail


def test_a_refused_read_is_a_probe_error_not_an_absent_word(monkeypatch):
    _with_registry(monkeypatch, _FakeWinreg(errors={
        "Word.Application\\CLSID": PermissionError(5, "access denied"),
    }))
    state, detail = server._word_registration()
    assert state == "probe_error"
    assert detail == "PermissionError"


def test_an_unexpected_exception_is_a_probe_error_not_an_absent_word(monkeypatch):
    """The BUG-001 shape itself: the probe blows up in a way nobody
    planned for, and the answer must be 'I could not tell', never 'you
    have no Word'."""
    _with_registry(monkeypatch, _FakeWinreg(errors={
        "Word.Application\\CLSID": AttributeError("module has no attribute"),
    }))
    state, detail = server._word_registration()
    assert state == "probe_error"
    assert detail == "AttributeError"


def test_a_broken_server_key_read_is_a_probe_error(monkeypatch):
    _with_registry(monkeypatch, _FakeWinreg(
        keys={"Word.Application\\CLSID": _CLSID},
        errors={"CLSID\\" + _CLSID + "\\LocalServer32": OSError(1, "boom")},
    ))
    assert server._word_registration()[0] == "probe_error"


def test_an_empty_class_id_value_is_a_probe_error(monkeypatch):
    _with_registry(monkeypatch, _FakeWinreg(keys={
        "Word.Application\\CLSID": "",
    }))
    assert server._word_registration()[0] == "probe_error"


@pytest.mark.parametrize(
    "state,application,com_tools",
    [
        ("registered", "installed", "available"),
        ("absent", "not registered", "unavailable"),
        ("probe_error", "unknown", "unknown"),
    ],
)
def test_get_server_info_reports_each_probe_state_truthfully(
    monkeypatch, state, application, com_tools
):
    """The reason the bug mattered: whatever the probe says has to reach
    the user unchanged, and a probe error may never be dressed up as an
    absent Word."""
    import sys as _s

    monkeypatch.setattr(_s, "platform", "win32")
    monkeypatch.setitem(_s.modules, "pythoncom", types.ModuleType("pythoncom"))
    monkeypatch.setattr(server, "_word_registration", lambda: (state, "detail"))
    word = server._word_environment()
    assert word["application"] == application
    assert word["com_tools"] == com_tools
    assert word["registration_probe"] == "registry"
    if state == "probe_error":
        note = word["note"].lower()
        assert "not registered" not in note
        assert "could not complete" in note
    if state == "registered":
        assert "com_word_status" in word["note"], (
            "a registration is not proof that automation works, and the "
            "report has to say so"
        )


def test_the_probe_does_not_call_the_api_that_never_existed():
    """pywin32 has ProgIDFromCLSID; it has never had CLSIDFromProgID."""
    source = Path(server.__file__).read_text(encoding="utf-8")
    assert "CLSIDFromProgID" not in source
    pythoncom = pytest.importorskip("pythoncom")
    assert not hasattr(pythoncom, "CLSIDFromProgID")


def test_the_probe_starts_no_word_and_opens_no_document(monkeypatch):
    """It reads the registry and nothing else: no COM apartment, no
    Dispatch, no file."""
    import sys as _s

    if "win32com.client" in _s.modules:
        mod = _s.modules["win32com.client"]
        for name in ("Dispatch", "DispatchEx", "GetActiveObject"):
            if hasattr(mod, name):
                monkeypatch.setattr(mod, name, _refuse_com)
    if "pythoncom" in _s.modules:
        pc = _s.modules["pythoncom"]
        if hasattr(pc, "CoInitialize"):
            monkeypatch.setattr(pc, "CoInitialize", _refuse_com)
    state, _ = server._word_registration()
    assert state in ("registered", "absent", "probe_error")


def _refuse_com(*args, **kwargs):
    raise AssertionError("the registration probe must not touch COM")


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
    """A published https:// address is not a machine path and never was;
    the drive-letter half of _PATH matches its scheme (``s:/``), so URLs are
    scanned for the POSIX shape only. Anything else is checked whole."""
    offenders = []
    for value in _values(info):
        if value.startswith("https://"):
            assert _PATH.search(value[len("https://"):]) is None, value
            continue
        if _PATH.search(value):
            offenders.append(value)
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


# ------------------------------------------------------- the handshake

# get_server_info always reported the package version correctly. The
# initialize handshake did not: FastMCP was constructed with no version=,
# so it answered with its own (3.4.7 at the time), and every client log and
# registry scrape that reads serverInfo.version recorded fastmcp's number as
# this product's. The siblings all passed it; this one was missed until
# 2.1.1. Asserted against __version__ rather than a literal so a release
# bump cannot make it stale, and asserted != the fastmcp version so a future
# refactor that drops the argument turns this red instead of passing by
# coincidence.


def _handshake():
    async def run():
        async with Client(server.mcp) as client:
            return client.initialize_result

    return asyncio.run(run())


def test_the_handshake_reports_this_package_not_fastmcps():
    import fastmcp

    result = _handshake()
    assert result.serverInfo.name == "kitchensink4word"
    assert result.serverInfo.version == __version__
    assert result.serverInfo.version != fastmcp.__version__, (
        "serverInfo is reporting the framework version again; server.py "
        "needs version=__version__ on the FastMCP constructor"
    )


def test_the_instructions_name_the_product_and_the_suite():
    """The instructions string is the text injected into a connected
    agent's context. Before 2.1.1 it opened on a generic description, so an
    agent had no way to say which product it was talking to."""
    text = _handshake().instructions or ""
    assert text.startswith(
        "KitchenSink4Word (kitchensink4word on PyPI), part of the "
        "KitchenSink4AI suite. "
    ), f"instructions do not lead with the self-identification: {text[:120]!r}"
    assert "Full-featured Word (.docx) editor" in text, (
        "the self-identification replaced the instructions instead of "
        "leading them"
    )
