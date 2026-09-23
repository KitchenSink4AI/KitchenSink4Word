"""Punch-list #933: a session's packs are its own, and they do not expire.

fastmcp kept the per-session visibility rules the pre-#933 pack toggles
relied on in its session state store, with a one-day time-to-live. When a
rule expired, an enabled pack's tools dropped out of tools/list, calls
answered "Unknown tool", and enable_tools answered "already enabled"
without bringing them back (shipped in Word 2.2.0, PowerPoint 1.3.1 and
Excel 1.2.4). The fix keeps one pack record per MCP session in this
server's own middleware (packstate, packgate) and uses no fastmcp
visibility state at all.

The seven gates Codex set for the fix (X-20260924-005): two-session
isolation; survival of a shortened TTL; re-enable behaviour; a
session-local list_changed; refusal of stale calls; garbage-collection
cleanup; default and locked single-session surfaces unchanged.

Fail-first: every test here except the two gate-7 regression guards fails
on the pre-#933 code. The gate-7 tests pin the default and locked surfaces
as they were, so they pass on both, which is what "unchanged" means.

This file is the same in the Word, PowerPoint and Excel repos apart from
the product block below.
"""

from __future__ import annotations

import asyncio
import gc
import hashlib
import importlib
import json
import os
import sys
from pathlib import Path

import fastmcp.server.context as fctx
import pytest
from fastmcp import Client
from fastmcp.client.transports import StdioTransport

from word_mcp import packs, server

# ------------------------------------------------------ product block

PKG = "word_mcp"
PACK = "review"
TOOL = "get_comments"
MODE_ENV = "KS4W_MODE"
POLICY_ENV = "KS4W_PACK_POLICY"
ENVS = (MODE_ENV, POLICY_ENV, "KS4W_ALL_TOOLS", "KS4W_LOCK_TOOLS")
QUIET = {"KS4W_UPDATE_CHECK": "off", "KS4W_STAR_NUDGE": "off"}
LITE_COUNT = 29
NO_CHANGE_PREFIX = packs.NO_CHANGE_PREFIX
FAMILY_SHA256 = {
    "packstate.py": "0343a3036c6fa7ee3a58ae6ac4c4c4b66a4e5699a9b4a0ecc14299d3e3abf2e1",
    "packgate.py": "814c7354d79227b37dc4734c69adfd36b9d35505d3734e675b37efd625113680",
}


def body():
    """The module attribute TOOL's body calls to do its work."""
    from word_mcp.ops import read

    return read, "get_comments"


def tool_args(tmp_path) -> dict:
    import docx

    path = tmp_path / "comments.docx"
    docx.Document().save(str(path))
    return {"file_path": str(path), "live": "off"}


# ------------------------------------------------------------ harness

TTL = 2        # fastmcp's session-state lifetime, cut from one day
WAIT = 3.5     # comfortably past it
SETTLE = 0.2   # lets a notification land before it is counted


class Inbox:
    """Counts the tools/list_changed notifications one client receives."""

    def __init__(self):
        self.list_changed = 0

    async def __call__(self, message):
        root = getattr(message, "root", message)
        if getattr(root, "method", "") == "notifications/tools/list_changed":
            self.list_changed += 1


def names(tools) -> set[str]:
    return {t.name for t in tools}


def registered() -> set[str]:
    return {n for members in packs.tool_names().values() for n in members}


def text(result) -> str:
    return "\n".join(getattr(b, "text", "") or "" for b in result.content)


def payload(result):
    """The JSON object a call answered with, or its raw text when the
    answer is not JSON (a bare fastmcp error)."""
    raw = text(result)
    try:
        return json.loads(raw)
    except ValueError:
        return raw


async def surface(client) -> dict:
    return payload(await client.call_tool("get_server_info", {}))["surface"]


def pack_size() -> int:
    return len(packs.pack_tools(PACK))


def expected_refusal() -> dict:
    """The stale-call refusal: the normal envelope, code NOT_FOUND, the
    disabled-tool signpost's own words (#51 v2.1 erratum, S1-A shape)."""
    return {
        "ok": False,
        "error": {
            "code": "NOT_FOUND",
            "message": (
                f"tool {TOOL!r} exists but is currently disabled: it "
                f"belongs to the {PACK!r} pack."
            ),
            "hint": (
                f"call enable_tools(packs=['{PACK}']) to turn it on, then "
                "retry this call"
            ),
        },
    }


@pytest.fixture
def launch(monkeypatch):
    """Start the server the way the console entry does, by running the
    real main() with only the transport's run() stubbed out, then put the
    process-wide state back. On the pre-#933 code main() also added the
    startup visibility transform; the snapshot of the transform list
    removes it again."""
    saved_enabled = dict(packs._ENABLED)
    saved_transforms = list(server.mcp._transforms)
    monkeypatch.setattr(server.mcp, "run", lambda *a, **k: None)
    for key, value in QUIET.items():
        monkeypatch.setenv(key, value)

    def start(**env):
        for key in ENVS:
            monkeypatch.delenv(key, raising=False)
        for key, value in env.items():
            monkeypatch.setenv(key, value)
        for name in packs._ENABLED:
            packs._ENABLED[name] = packs.pack_of(name) == "lite"
        server.main()

    yield start
    server.mcp._transforms[:] = saved_transforms
    packs._ENABLED.clear()
    packs._ENABLED.update(saved_enabled)


@pytest.fixture
def body_runs(monkeypatch):
    """Counts how often TOOL's body reaches the work it does."""
    runs: list[int] = []
    module, attr = body()
    real = getattr(module, attr)

    def counted(*args, **kwargs):
        runs.append(1)
        return real(*args, **kwargs)

    monkeypatch.setattr(module, attr, counted)
    return runs


# ------------------------------------------- gate 1: two-session isolation


def test_gate1_two_sessions_keep_separate_packs(launch, tmp_path):
    """Session A enabling a pack changes A's list, calls and report, and
    nothing of session B's. Before #933 the bookkeeping was one per
    process: B's report claimed A's pack and B's own enable answered
    'already enabled' while B could not see a single tool of it."""
    launch()
    args = tool_args(tmp_path)
    n = pack_size()

    async def run():
        async with Client(server.mcp) as a, Client(server.mcp) as b:
            first = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            out = {
                "a_lists": TOOL in names(await a.list_tools()),
                "b_lists": TOOL in names(await b.list_tools()),
                "a_call": await a.call_tool(TOOL, args, raise_on_error=False),
                "b_call": await b.call_tool(TOOL, args, raise_on_error=False),
                "a_report": (await surface(a))["packs"][PACK],
                "b_report": (await surface(b))["packs"][PACK],
            }
            out["b_enable"] = payload(await b.call_tool(
                "enable_tools", {"packs": [PACK]}))
            out["b_lists_after"] = TOOL in names(await b.list_tools())
            return first, out

    first, out = asyncio.run(run())
    assert first["enabled"] == [PACK]
    assert out["a_lists"] and not out["a_call"].is_error
    assert not out["b_lists"] and out["b_call"].is_error
    assert out["a_report"] == f"{n}/{n} enabled"
    assert out["b_report"] == f"0/{n} enabled"
    assert out["b_enable"]["enabled"] == [PACK]
    assert out["b_enable"]["already_enabled"] == []
    assert out["b_lists_after"]


# ------------------------------------------ gate 2: shortened-TTL survival


def test_gate2_pack_survives_fastmcp_state_expiry(launch, tmp_path,
                                                  monkeypatch):
    """The defect itself, with fastmcp's one-day session-state lifetime cut
    to 2 s: an enabled pack must still be listed, callable and reported
    after it passes."""
    monkeypatch.setattr(fctx.Context, "_STATE_TTL_SECONDS", TTL)
    launch()
    args = tool_args(tmp_path)
    n = pack_size()

    async def run():
        async with Client(server.mcp) as a:
            await a.call_tool("enable_tools", {"packs": [PACK]})
            before = TOOL in names(await a.list_tools())
            await asyncio.sleep(WAIT)
            after = TOOL in names(await a.list_tools())
            call = await a.call_tool(TOOL, args, raise_on_error=False)
            report = (await surface(a))["packs"][PACK]
            return before, after, call, report

    before, after, call, report = asyncio.run(run())
    assert before
    assert after, "the pack left tools/list when fastmcp's state expired"
    assert not call.is_error, text(call)
    assert report == f"{n}/{n} enabled"


def test_gate2_stdio_session_survives_expiry_through_main(tmp_path):
    """The same end to end: the real console entry, main(), serving stdio
    in a child process (one session for the life of the process), with
    fastmcp's state lifetime cut to 2 s inside the child."""
    launcher = tmp_path / "launch_server.py"
    launcher.write_text(
        "import fastmcp.server.context as c\n"
        f"c.Context._STATE_TTL_SECONDS = {TTL}\n"
        f"from {PKG}.server import main\n"
        "main()\n",
        encoding="utf-8",
    )
    env = {k: v for k, v in os.environ.items() if k not in ENVS}
    env.update(QUIET)
    transport = StdioTransport(
        command=sys.executable, args=[str(launcher)], env=env,
        cwd=str(tmp_path), keep_alive=False,
    )
    args = tool_args(tmp_path)

    async def run():
        async with Client(transport) as c:
            at_start = TOOL in names(await c.list_tools())
            await c.call_tool("enable_tools", {"packs": [PACK]})
            enabled = TOOL in names(await c.list_tools())
            await asyncio.sleep(WAIT)
            after = TOOL in names(await c.list_tools())
            call = await c.call_tool(TOOL, args, raise_on_error=False)
            again = payload(await c.call_tool(
                "enable_tools", {"packs": [PACK]}))
            return at_start, enabled, after, call, again

    at_start, enabled, after, call, again = asyncio.run(run())
    assert not at_start and enabled
    assert after, "the pack left tools/list when fastmcp's state expired"
    assert not call.is_error, text(call)
    assert again["already_enabled"] == [PACK]


# ---------------------------------------------- gate 3: re-enable behaviour


def test_gate3_re_enable_after_expiry_and_after_disable(launch, monkeypatch):
    """Past the expiry, a second enable_tools is a true no-op: the pack is
    still on and still listed, the answer says no list change was sent,
    and none was. A disable followed by an enable brings it back with one
    notification each."""
    monkeypatch.setattr(fctx.Context, "_STATE_TTL_SECONDS", TTL)
    launch()

    async def run():
        inbox = Inbox()
        async with Client(server.mcp, message_handler=inbox) as a:
            first = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            await asyncio.sleep(WAIT)
            count = inbox.list_changed
            again = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            await asyncio.sleep(SETTLE)
            out = {
                "first": first,
                "again": again,
                "listed_again": TOOL in names(await a.list_tools()),
                "sent_on_again": inbox.list_changed - count,
            }
            out["off"] = payload(await a.call_tool(
                "disable_tools", {"packs": [PACK]}))
            out["gone"] = TOOL not in names(await a.list_tools())
            out["back"] = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            await asyncio.sleep(SETTLE)
            out["listed_back"] = TOOL in names(await a.list_tools())
            out["sent_after"] = inbox.list_changed - count
            return out

    out = asyncio.run(run())
    assert out["first"]["enabled"] == [PACK]
    assert out["again"]["enabled"] == []
    assert out["again"]["already_enabled"] == [PACK]
    assert out["listed_again"], "re-enable left the pack unlisted"
    assert out["sent_on_again"] == 0
    assert out["again"]["note"].startswith(NO_CHANGE_PREFIX)
    assert out["off"]["disabled"] == [PACK] and out["gone"]
    assert out["back"]["enabled"] == [PACK] and out["listed_back"]
    assert out["sent_after"] == 2


# --------------------------------------- gate 4: session-local list_changed


def test_gate4_list_changed_reaches_only_the_changing_session(launch):
    """Each change notifies the session that made it and no other; a
    second session enabling the same pack changes its own list and so
    gets its own notification."""
    launch()

    async def run():
        ia, ib = Inbox(), Inbox()
        async with Client(server.mcp, message_handler=ia) as a, \
                Client(server.mcp, message_handler=ib) as b:
            await a.call_tool("enable_tools", {"packs": [PACK]})
            await asyncio.sleep(SETTLE)
            step1 = (ia.list_changed, ib.list_changed)
            b_enable = payload(await b.call_tool(
                "enable_tools", {"packs": [PACK]}))
            await asyncio.sleep(SETTLE)
            step2 = (ia.list_changed, ib.list_changed)
            await a.call_tool("disable_tools", {"packs": [PACK]})
            await asyncio.sleep(SETTLE)
            step3 = (ia.list_changed, ib.list_changed)
            b_still = TOOL in names(await b.list_tools())
            a_gone = TOOL not in names(await a.list_tools())
            return step1, b_enable, step2, step3, b_still, a_gone

    step1, b_enable, step2, step3, b_still, a_gone = asyncio.run(run())
    assert step1 == (1, 0)
    assert b_enable["enabled"] == [PACK]
    assert step2 == (1, 1)
    assert step3 == (2, 1)
    assert b_still and a_gone


# ------------------------------------------ gate 5: stale calls are refused


def test_gate5_stale_calls_are_refused_and_never_run(launch, tmp_path,
                                                    body_runs):
    """A call to a tool whose pack is off in the calling session (never
    enabled there, or disabled since the client fetched its list) is
    refused with the product envelope, code NOT_FOUND, naming the pack and
    the enable call, and the tool's body does not run."""
    launch()
    args = tool_args(tmp_path)

    async def run():
        async with Client(server.mcp) as a, Client(server.mcp) as b:
            await a.call_tool("enable_tools", {"packs": [PACK]})
            ok = await a.call_tool(TOOL, args, raise_on_error=False)
            ran = len(body_runs)
            never_enabled = await b.call_tool(TOOL, args,
                                              raise_on_error=False)
            await a.call_tool("disable_tools", {"packs": [PACK]})
            disabled_since = await a.call_tool(TOOL, args,
                                               raise_on_error=False)
            return ok, ran, never_enabled, disabled_since

    ok, ran, never_enabled, disabled_since = asyncio.run(run())
    assert not ok.is_error and ran == 1
    for stale in (never_enabled, disabled_since):
        assert stale.is_error
        assert payload(stale) == expected_refusal()
        assert stale.structured_content == expected_refusal()
    assert len(body_runs) == 1, "a stale call ran the tool body"


def test_gate5_both_layers_give_the_same_refusal(launch, tmp_path,
                                                 body_runs, monkeypatch):
    """With the session gate middleware taken out, the call is refused by
    the tool's own boundary wrapper (the check sits inside its try block)
    with a byte-identical result, and the body still does not run. Only
    the menu differs: without the middleware nothing filters the list."""
    launch()
    args = tool_args(tmp_path)

    async def call_from_fresh_session():
        async with Client(server.mcp) as b:
            listed = TOOL in names(await b.list_tools())
            result = await b.call_tool(TOOL, args, raise_on_error=False)
            return listed, result

    listed_with, with_gate = asyncio.run(call_from_fresh_session())
    assert with_gate.structured_content == expected_refusal(), text(with_gate)
    gates = [m for m in server.mcp.middleware
             if type(m).__name__ == "SessionPackGate"]
    assert len(gates) == 1, "the session pack gate middleware is missing"
    ours = [m for m in server.mcp.middleware
            if type(m).__module__.startswith(PKG)]
    assert ours[0] is gates[0], "the gate is not this server's outermost"
    monkeypatch.setattr(server.mcp, "middleware",
                        [m for m in server.mcp.middleware if m not in gates])
    listed_without, without_gate = asyncio.run(call_from_fresh_session())

    assert not listed_with and listed_without
    assert with_gate.is_error and without_gate.is_error
    assert with_gate.structured_content == expected_refusal()
    assert without_gate.structured_content == expected_refusal()
    assert text(with_gate) == text(without_gate)
    assert body_runs == []


# ------------------------------------------------- gate 6: GC cleanup


def test_gate6_a_finished_session_leaves_nothing_behind(launch):
    """A pack one session enabled does not outlive that session: the next
    session starts on the startup surface, is told so, and gets a real
    enable of its own."""
    launch()
    n = pack_size()

    async def run():
        async with Client(server.mcp) as a:
            await a.call_tool("enable_tools", {"packs": [PACK]})
        gc.collect()
        async with Client(server.mcp) as c:
            listed = TOOL in names(await c.list_tools())
            report = (await surface(c))["packs"][PACK]
            again = payload(await c.call_tool(
                "enable_tools", {"packs": [PACK]}))
            return listed, report, again

    listed, report, again = asyncio.run(run())
    assert not listed
    assert report == f"0/{n} enabled"
    assert again["enabled"] == [PACK]


def test_gate6_records_are_collected_with_their_sessions(launch):
    """Each session's record is held weakly on its session object: two
    sessions that changed their packs hold two records, and both are gone
    once the sessions end and are collected."""
    packstate = importlib.import_module(f"{PKG}.packstate")
    launch()
    gc.collect()
    before = packstate.live_records()

    async def run():
        async with Client(server.mcp) as a, Client(server.mcp) as b:
            await a.call_tool("enable_tools", {"packs": [PACK]})
            await b.call_tool("enable_tools", {"packs": ["everything"]})
            return packstate.live_records()

    during = asyncio.run(run())
    gc.collect()
    assert during == before + 2
    assert packstate.live_records() == before


# ------------------------- gate 7: default and locked surfaces unchanged


def test_gate7_default_surface_unchanged(launch):
    """With nothing set, a session lists exactly the lite core and reports
    every pack off, as before #933. (A regression guard: it passes on the
    pre-#933 code too, which is the point.)"""
    launch()

    async def run():
        async with Client(server.mcp) as a:
            return names(await a.list_tools()), await surface(a)

    listed, report = asyncio.run(run())
    lite = set(packs.pack_tools("lite"))
    assert len(lite) == LITE_COUNT
    assert listed & registered() == lite
    assert report["active_tools"] == LITE_COUNT
    for pack in packs.pack_names():
        size = len(packs.pack_tools(pack))
        assert report["packs"][pack] == f"0/{size} enabled"


def test_gate7_locked_surface_unchanged(launch, monkeypatch):
    """A locked launch with a startup pack keeps exactly that surface for
    the life of the session, past the state expiry too: enable_tools and
    disable_tools refuse with CONFLICT and change nothing. (A regression
    guard: it passes on the pre-#933 code too.)"""
    monkeypatch.setattr(fctx.Context, "_STATE_TTL_SECONDS", TTL)
    launch(**{MODE_ENV: PACK, POLICY_ENV: "locked"})
    n = pack_size()

    async def run():
        async with Client(server.mcp) as a:
            listed = names(await a.list_tools()) & registered()
            enable = await a.call_tool(
                "enable_tools", {"packs": ["everything"]},
                raise_on_error=False)
            disable = await a.call_tool(
                "disable_tools", {"packs": [PACK]}, raise_on_error=False)
            await asyncio.sleep(WAIT)
            after = names(await a.list_tools()) & registered()
            return listed, enable, disable, after, await surface(a)

    listed, enable, disable, after, report = asyncio.run(run())
    expected = set(packs.pack_tools("lite")) | set(packs.pack_tools(PACK))
    assert listed == expected
    assert enable.is_error and payload(enable)["error"]["code"] == "CONFLICT"
    assert disable.is_error and payload(disable)["error"]["code"] == "CONFLICT"
    assert after == expected
    assert report["packs"][PACK] == f"{n}/{n} enabled"
    assert report["active_tools"] == LITE_COUNT + n


# --------------------------------------------- the shared modules stay shared


def test_shared_modules_match_the_family():
    """packstate.py and packgate.py are product-neutral and kept identical
    in KitchenSink4Word, KitchenSink4PPT and KitchenSink4XL. A change to
    either is made in all three repos, with the pins below updated in all
    three test files."""
    here = Path(packs.__file__).parent
    for name, digest in FAMILY_SHA256.items():
        source = (here / name).read_bytes().replace(b"\r\n", b"\n")
        assert hashlib.sha256(source).hexdigest() == digest, name
