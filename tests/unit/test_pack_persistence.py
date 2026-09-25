"""A pack switched on or off stays that way until a person or the AI
changes it again (owner ruling, 2026-09-26: "If someone switches it on,
they don't expect it to revert. I know I wouldn't.").

Before this, the choice lasted only as long as the MCP session that made
it: the next session (a restart of the app, a new conversation) started
again from the launch settings, and until punch-list #933 fastmcp also
dropped it a day into the same session. Now every enable_tools and
disable_tools call is saved (packstore.py) and every later session starts
from the launch settings plus the saved choices, while a launch-time lock
still wins over anything saved.

"A day later" is simulated the way the #933 gates do it: fastmcp's
one-day session-state lifetime is cut to 2 s and the test waits past it.
"A restart" is the real main() run again in the same process with every
pack record reset (in-process), or a second child process (stdio).

Fail-first: every test here fails on the #933 head, where nothing is
saved. This file is the same in the Word, PowerPoint and Excel repos apart
from the product block below.
"""

from __future__ import annotations

import asyncio
import gc
import hashlib
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
OTHER = "academic"
OTHER_TOOL = packs.pack_tools(OTHER)[0]
MODE_ENV = "KS4W_MODE"
POLICY_ENV = "KS4W_PACK_POLICY"
STORE_ENV = "KS4W_PACK_STORE_DIR"
ENVS = (MODE_ENV, POLICY_ENV, "KS4W_ALL_TOOLS", "KS4W_LOCK_TOOLS")
QUIET = {"KS4W_UPDATE_CHECK": "off", "KS4W_STAR_NUDGE": "off"}
PACKSTORE_SHA256 = "3278ef7e11aab28e908df7d9c2585cee83ca2580471cc788d4ad1f0e5b1162e9"


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


def names(tools) -> set[str]:
    return {t.name for t in tools}


def registered() -> set[str]:
    return {n for members in packs.tool_names().values() for n in members}


def text(result) -> str:
    return "\n".join(getattr(b, "text", "") or "" for b in result.content)


def payload(result):
    raw = text(result)
    try:
        return json.loads(raw)
    except ValueError:
        return raw


async def info(client) -> dict:
    return payload(await client.call_tool("get_server_info", {}))


def pack_size(pack: str = PACK) -> int:
    return len(packs.pack_tools(pack))


def store_file() -> Path:
    return Path(os.environ[STORE_ENV]) / "tool-packs.json"


@pytest.fixture
def launch(monkeypatch):
    """Start (or restart) the server the way the console entry does: the
    real main() with only the transport's run() stubbed out, from a
    process whose pack records are all back at their registered state and
    whose earlier sessions are gone."""
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
        gc.collect()
        for name in packs._ENABLED:
            packs._ENABLED[name] = packs.pack_of(name) == "lite"
        server.main()

    yield start
    server.mcp._transforms[:] = saved_transforms
    packs._ENABLED.clear()
    packs._ENABLED.update(saved_enabled)


@pytest.fixture
def body_runs(monkeypatch):
    runs: list[int] = []
    module, attr = body()
    real = getattr(module, attr)

    def counted(*args, **kwargs):
        runs.append(1)
        return real(*args, **kwargs)

    monkeypatch.setattr(module, attr, counted)
    return runs


async def _session_view(args):
    """One fresh session: is TOOL listed, what does a call answer, what
    does get_server_info report for PACK."""
    async with Client(server.mcp) as c:
        listed = TOOL in names(await c.list_tools())
        call = await c.call_tool(TOOL, args, raise_on_error=False)
        report = (await info(c))["surface"]["packs"][PACK]
        return listed, call, report


# --------------------------------------------------- on stays on


def test_switched_on_stays_on_a_day_later_and_after_a_restart(
        launch, tmp_path, monkeypatch):
    """A pack switched on is still on past fastmcp's (shortened) one-day
    lifetime, and still on in the first session after a restart: listed,
    callable, and reported on."""
    monkeypatch.setattr(fctx.Context, "_STATE_TTL_SECONDS", TTL)
    launch()
    args = tool_args(tmp_path)
    n = pack_size()

    async def first():
        async with Client(server.mcp) as a:
            result = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            await asyncio.sleep(WAIT)
            listed = TOOL in names(await a.list_tools())
            call = await a.call_tool(TOOL, args, raise_on_error=False)
            return result, listed, call

    result, listed, call = asyncio.run(first())
    assert result["enabled"] == [PACK]
    assert listed and not call.is_error, text(call)

    launch()  # restart, same launch settings
    listed, call, report = asyncio.run(_session_view(args))
    assert listed, "the pack switched on reverted at the restart"
    assert not call.is_error, text(call)
    assert report == f"{n}/{n} enabled"
    assert result.get("saved") is True


# -------------------------------------------------- off stays off


def test_switched_off_stays_off_after_a_restart(launch, tmp_path,
                                                 body_runs):
    """A pack the launch settings turn on, switched off by the session,
    stays off after a restart with the same launch settings: not listed,
    a call is refused and never runs, and the report says off."""
    launch(**{MODE_ENV: PACK})
    args = tool_args(tmp_path)
    n = pack_size()

    async def first():
        async with Client(server.mcp) as a:
            assert TOOL in names(await a.list_tools())
            return payload(await a.call_tool(
                "disable_tools", {"packs": [PACK]}))

    result = asyncio.run(first())
    assert result["disabled"] == [PACK]

    launch(**{MODE_ENV: PACK})
    listed, call, report = asyncio.run(_session_view(args))
    assert not listed, "the pack switched off came back at the restart"
    assert call.is_error
    assert payload(call)["error"]["code"] == "NOT_FOUND"
    assert report == f"0/{n} enabled"
    assert body_runs == []
    assert result.get("saved") is True


# ------------------------------- the next session, same process


def test_the_next_session_starts_from_the_choice(launch, tmp_path):
    """Within one running server (HTTP serves many sessions), a session
    that opens after the change starts with it; a session that was
    already open keeps its own surface until it changes its own packs
    (the 2026-09-02 session-scoped ruling)."""
    launch()
    n = pack_size()

    async def run():
        async with Client(server.mcp) as b:
            async with Client(server.mcp) as a:
                await a.call_tool("enable_tools", {"packs": [PACK]})
            gc.collect()
            async with Client(server.mcp) as c:
                c_listed = TOOL in names(await c.list_tools())
                c_report = (await info(c))["surface"]["packs"][PACK]
            b_listed = TOOL in names(await b.list_tools())
            b_report = (await info(b))["surface"]["packs"][PACK]
            return c_listed, c_report, b_listed, b_report

    c_listed, c_report, b_listed, b_report = asyncio.run(run())
    assert c_listed and c_report == f"{n}/{n} enabled"
    assert not b_listed and b_report == f"0/{n} enabled"


# ------------------------------------------------------ lock wins


def test_a_launch_lock_wins_over_a_saved_choice(launch, tmp_path):
    """Saved choices are not applied while the launch settings lock the
    tool set, enable_tools still refuses honestly there, and the saved
    file is left as it was. Unlocked again, the choice is back."""
    launch()
    n = pack_size()
    args = tool_args(tmp_path)

    async def enable():
        async with Client(server.mcp) as a:
            await a.call_tool("enable_tools", {"packs": [PACK]})

    asyncio.run(enable())
    before = store_file().read_bytes()

    launch(**{POLICY_ENV: "locked"})

    async def locked():
        async with Client(server.mcp) as c:
            listed = TOOL in names(await c.list_tools())
            refused = await c.call_tool(
                "enable_tools", {"packs": [PACK]}, raise_on_error=False)
            report = await info(c)
            return listed, refused, report

    listed, refused, report = asyncio.run(locked())
    assert not listed
    assert refused.is_error
    assert payload(refused)["error"]["code"] == "CONFLICT"
    assert report["surface"]["packs"][PACK] == f"0/{n} enabled"
    assert report["saved_packs"]["applied"] is False
    assert store_file().read_bytes() == before

    launch()
    listed, call, _ = asyncio.run(_session_view(args))
    assert listed and not call.is_error


def test_a_lock_keeps_a_launch_pack_that_was_saved_off(launch):
    """The other direction: a pack the launch settings turn on, saved off
    earlier, is ON under a locked launch, because the lock fixes the
    surface to the launch settings."""
    launch(**{MODE_ENV: PACK})

    async def disable():
        async with Client(server.mcp) as a:
            await a.call_tool("disable_tools", {"packs": [PACK]})

    asyncio.run(disable())
    launch(**{MODE_ENV: PACK, POLICY_ENV: "locked"})

    async def look():
        async with Client(server.mcp) as c:
            return TOOL in names(await c.list_tools())

    assert asyncio.run(look())


# ---------------------------------------- launch settings changed


def test_changed_launch_settings_are_the_newer_choice(launch):
    """Choices are kept per launch settings: a person who changes the
    launch settings gets what those settings say, and going back to the
    old settings brings back the choices made under them."""
    launch()

    async def enable():
        async with Client(server.mcp) as a:
            await a.call_tool("enable_tools", {"packs": [PACK]})

    asyncio.run(enable())
    launch(**{MODE_ENV: OTHER})

    async def look():
        async with Client(server.mcp) as c:
            listed = names(await c.list_tools())
            return TOOL in listed, OTHER_TOOL in listed

    assert asyncio.run(look()) == (False, True)
    launch()
    assert asyncio.run(look()) == (True, False)


# ---------------------------------------------- status = reality


def test_status_matches_what_is_on_after_a_restart(launch):
    """After a restart with saved choices, get_server_info reports exactly
    the packs the session lists, and names the saved choices it applied."""
    launch()

    async def change():
        async with Client(server.mcp) as a:
            await a.call_tool("enable_tools", {"packs": ["everything"]})
            await a.call_tool("disable_tools", {"packs": [PACK]})

    asyncio.run(change())
    launch()

    async def look():
        async with Client(server.mcp) as c:
            return names(await c.list_tools()), await info(c)

    listed, report = asyncio.run(look())
    for pack in packs.pack_names():
        members = set(packs.pack_tools(pack))
        on = len(members & listed)
        assert report["surface"]["packs"][pack] == (
            f"{on}/{len(members)} enabled"), pack
        assert on in (0, len(members)), pack
    assert report["surface"]["active_tools"] == len(listed & registered())
    assert TOOL not in listed
    saved = report["saved_packs"]
    assert saved["applied"] is True
    # Only choices that differ from the launch settings are kept: PACK is
    # off at a lite launch anyway, so nothing is saved for it.
    assert PACK not in saved["packs"]
    assert all(saved["packs"][p] == "on"
               for p in packs.pack_names() if p != PACK)


# ------------------------------------------------ honest failures


def test_an_unwritable_store_keeps_the_change_and_says_not_saved(
        launch, tmp_path, monkeypatch):
    """When the state directory cannot be written, the change still holds
    for the running session and the answer says it was not saved."""
    blocker = tmp_path / "not-a-directory"
    blocker.write_text("x", encoding="utf-8")
    monkeypatch.setenv(STORE_ENV, str(blocker))
    launch()

    async def run():
        async with Client(server.mcp) as a:
            result = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            return result, TOOL in names(await a.list_tools())

    result, listed = asyncio.run(run())
    assert result["enabled"] == [PACK]
    assert result["saved"] is False
    assert listed


def test_a_damaged_store_is_ignored(launch):
    """A damaged file reads as no saved choices: the server starts on its
    launch settings, and the next change writes a good file."""
    store_file().parent.mkdir(parents=True, exist_ok=True)
    store_file().write_text("{ not json", encoding="utf-8")
    launch()

    async def run():
        async with Client(server.mcp) as a:
            listed = TOOL in names(await a.list_tools())
            result = payload(await a.call_tool(
                "enable_tools", {"packs": [PACK]}))
            return listed, result

    listed, result = asyncio.run(run())
    assert not listed and result["saved"] is True
    assert json.loads(store_file().read_text(encoding="utf-8"))["format"] == 1


# -------------------------------------------- stdio, two processes


def test_stdio_restart_keeps_the_choice(tmp_path):
    """End to end: the real console entry serving stdio in a child
    process. The first process switches the pack on and exits; a second
    process started with the same launch settings lists it and reports
    it on."""
    launcher = tmp_path / "launch_server.py"
    launcher.write_text(
        f"from {PKG}.server import main\nmain()\n", encoding="utf-8")
    env = {k: v for k, v in os.environ.items() if k not in ENVS}
    env.update(QUIET)
    env[STORE_ENV] = str(tmp_path / "store")
    n = pack_size()

    def transport():
        return StdioTransport(
            command=sys.executable, args=[str(launcher)], env=env,
            cwd=str(tmp_path), keep_alive=False,
        )

    async def first():
        async with Client(transport()) as c:
            before = TOOL in names(await c.list_tools())
            await c.call_tool("enable_tools", {"packs": [PACK]})
            return before

    async def second():
        async with Client(transport()) as c:
            listed = TOOL in names(await c.list_tools())
            report = (await info(c))["surface"]["packs"][PACK]
            return listed, report

    assert asyncio.run(first()) is False
    listed, report = asyncio.run(second())
    assert listed, "the pack reverted when the process restarted"
    assert report == f"{n}/{n} enabled"


# --------------------------------------------- the shared module


def test_packstore_matches_the_family():
    """packstore.py is product-neutral and identical in KitchenSink4Word,
    KitchenSink4PPT and KitchenSink4XL."""
    source = (Path(packs.__file__).parent / "packstore.py").read_bytes()
    digest = hashlib.sha256(source.replace(b"\r\n", b"\n")).hexdigest()
    assert digest == PACKSTORE_SHA256
