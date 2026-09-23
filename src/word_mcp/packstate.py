"""Per-session pack records (punch-list #933).

Product-neutral: this file is identical in KitchenSink4Word,
KitchenSink4PPT and KitchenSink4XL, and each repo's
test_session_packs.py pins its hash, so a change here is made in all
three or in none.

The defect this fixes. Until #933 a pack switched on mid-session was
shown to the session through fastmcp's own session visibility rules
(ctx.enable_components). fastmcp keeps those rules in its session state
store with a time-to-live (Context._STATE_TTL_SECONDS, one day). When the
rule expired, the pack's tools dropped out of tools/list and a call
answered "Unknown tool", while this server's own bookkeeping still said
the pack was on, so enable_tools answered "already enabled" and never
brought the tools back. Any conversation still open a day after it
enabled a pack hit it; nobody had to do anything.

The model now. fastmcp visibility is not used at all. Each MCP session
has its own record of which tools are on, held here and owned by this
server. tools/list is filtered from that record and tools/call is
admitted from it by packgate.SessionPackGate, and again inside every
tool's own boundary wrapper by packgate.check_call. enable_tools and
disable_tools change only the calling session's record and send
tools/list_changed to that session only. Nothing here expires.

What a session is. The key is the MCP ServerSession object of the
connection, held weakly and compared by identity (ServerSession defines
no __eq__ or __hash__). stdio has one session for the life of the
process; every HTTP session is its own. When a session ends and its
object is collected, its record goes with it, so nothing needs cleaning
up by hand. A session that has not changed its packs has no record and
uses the process default, which is the startup surface main() applied
(the MODE setting and the launch toggles); its first change copies that
default.

Outside any session (in-process callers: the test suite and the
measurement scripts) there is no record to consult. The pack bookkeeping
reads and writes the process default directly, and tools/list and
tools/call are not filtered, which is what those callers saw before.
"""

from __future__ import annotations

import threading
import weakref
from typing import Any

from fastmcp.server.dependencies import get_context

#: Serializes every read-modify-write of a record. Re-entrant, because the
#: pack bookkeeping reads the record again (for its surface report) while
#: it still holds the lock after a change.
LOCK = threading.RLock()

# session -> {tool_name: enabled}. A stored mapping is replaced whole and
# never edited in place, so a reader that holds one sees a whole record.
_RECORDS: "weakref.WeakKeyDictionary[Any, dict[str, bool]]" = (
    weakref.WeakKeyDictionary()
)


def session_of(ctx: Any) -> Any | None:
    """The MCP session a fastmcp Context is serving, or None when it is not
    serving one (a Context built for an in-process call)."""
    if ctx is None:
        return None
    try:
        return ctx.session
    except RuntimeError:
        return None


def current_session() -> Any | None:
    """The MCP session of the request being handled, or None outside one.

    fastmcp publishes the active Context in a context variable, which it
    carries into the worker thread a synchronous tool runs on, so this
    answers the same inside a tool body as in the middleware."""
    try:
        ctx = get_context()
    except RuntimeError:
        return None
    return session_of(ctx)


def session_record(session: Any) -> dict[str, bool] | None:
    """The session's own record, or None when there is none: no session,
    or a session that has not changed its packs yet."""
    if session is None:
        return None
    with LOCK:
        return _RECORDS.get(session)


def keep_session_record(session: Any, enabled: dict[str, bool]) -> None:
    """Make `enabled` the session's record. The caller hands over a fresh
    mapping and does not edit it afterwards."""
    with LOCK:
        _RECORDS[session] = enabled


def live_records() -> int:
    """How many sessions hold a record right now."""
    with LOCK:
        return len(_RECORDS)
