"""The session pack gate (punch-list #933): what a session sees in
tools/list and may run through tools/call, derived from that session's
own pack record (packstate).

Product-neutral: this file is identical in KitchenSink4Word,
KitchenSink4PPT and KitchenSink4XL, and each repo's
test_session_packs.py pins its hash.

Two placements, one answer:
- SessionPackGate, registered as the first (outermost) middleware,
  filters every tools/list by the calling session's record and refuses a
  tools/call to a tool whose pack is off in that session.
- check_call() runs the same test inside every tool's boundary wrapper,
  inside its try block, so a route that skips the middleware still
  refuses.
Both raise packs.PackOff and turn it into the product's normal refusal
envelope (code NOT_FOUND), so a caller holding a stale tool list gets the
same payload from either layer. That is the shape the #51 v2.1 erratum
(S1-A) fixes for the Stage 1 work, adopted here so the refusal format
changes once.

This is a fix to the pack menu, not a security boundary. Packs organize
what a session carries; nothing here is an administrator control. Calls
made outside any MCP session are not filtered (see packstate).
"""

from __future__ import annotations

from typing import Any, Callable

from fastmcp.server.middleware import Middleware
from mcp.types import ToolListChangedNotification

from . import packs, packstate


def admitted(tool_name: str, session: Any) -> bool:
    """May this session see and call the tool? Always yes outside a
    session, for the lite core, and for a name the pack registry does not
    know (fastmcp answers for those, as it always has)."""
    if session is None:
        return True
    pack = packs.pack_of(tool_name)
    if pack is None or pack == "lite":
        return True
    return packs.is_tool_enabled(tool_name, session)


def check_session(tool_name: str, session: Any) -> None:
    """Raise PackOff when `session` has the tool's pack off."""
    if not admitted(tool_name, session):
        raise packs.PackOff(tool_name, packs.pack_of(tool_name))


def check_call(tool_name: str) -> None:
    """The boundary wrapper's check, against the session being served."""
    check_session(tool_name, packstate.current_session())


async def announce_list_changed(ctx: Any) -> bool:
    """Send tools/list_changed to ctx's session, and to no other. Returns
    False when there is no session to tell (an in-process call)."""
    if packstate.session_of(ctx) is None:
        return False
    await ctx.send_notification(ToolListChangedNotification())
    return True


class SessionPackGate(Middleware):
    """tools/list and tools/call from the calling session's pack record.

    `refuse` is the product's envelope builder (exception -> refusal
    result); it is passed in so this module stays product-neutral."""

    def __init__(self, refuse: Callable[[BaseException], Any]):
        self._refuse = refuse

    async def on_list_tools(self, context, call_next):
        tools = await call_next(context)
        session = packstate.session_of(context.fastmcp_context)
        if session is None:
            return tools
        return [t for t in tools if admitted(t.name, session)]

    async def on_call_tool(self, context, call_next):
        session = packstate.session_of(context.fastmcp_context)
        try:
            check_session(context.message.name, session)
        except packs.PackOff as exc:
            return self._refuse(exc)
        return await call_next(context)
