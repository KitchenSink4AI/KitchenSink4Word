"""The packs notice (punch-list #887), pinned word for word.

THE PROBLEM. A subagent enables a pack, gets "tools/list_changed was sent;
re-fetch the tool list if your client does not refresh automatically", and
then cannot call a single tool it just turned on: its tool list was fixed
when it was spawned and nothing it can do changes that. The 2026-09-21
field test burned a whole graphics build on this. The old sentence was true
and useless, because the clients that strand an agent are exactly the ones
that never re-fetch.

THE FIX IS COPY, so it is pinned as copy. The owner wrote the wording; the
only substitution allowed is the server's own start-up pack-list variable.
These tests exist so an edit to the sentence is a decision someone makes,
not a drift.

Three properties:
1. The note is the owner's text, with KS4W_MODE named.
2. It is returned on EVERY successful enable, including one that enabled
   nothing new. A worker that re-enables an already-on pack is precisely
   the agent that needs the instruction, and the old code attached the note
   only when something flipped.
3. The same fact reaches a client BEFORE it calls anything: the server
   instructions at handshake, and the get_workflows index.
"""

from __future__ import annotations

import pytest

from word_mcp import packs, server  # noqa: F401  (registers the tools)
from word_mcp.core.errors import WordMcpError
from word_mcp.ops import workflows

#: COPY SLOTS P1-S2-07, P1-S2-08 and P1-S2-09 (copy packet 2), landed
#: from Codex X-20260924-013 and accepted by the main thread. The owner's
#: 2026-09-22 wording pointed a Claude Desktop extension user at a launch
#: variable the extension does not offer (#930) and named an
#: administrator for a lock the user may have set. Pinned word for word:
#: an edit to any of it is a decision someone makes, not a drift.
#:
#: PREFIX + BODY: the first sentence is the only claim that depends on
#: what the call did, and saying a notification was sent when none was is
#: a lie told to exactly the caller who is trying to work out why its tool
#: list has not changed. The prefixes are code, not copy, and stay.
EXPECTED_BODY = (
    "If the new tools are not in your tool list, this client fixed its "
    "list when the session or worker started: do not retry here. In the "
    "Claude Desktop extension, ask the user to turn on 'Load every tool "
    "at startup', restart Claude Desktop, and start a new session. This "
    "loads every pack and uses more working memory. In a hand-configured "
    "client, ask the user to add the packs to KS4W_MODE as a comma list "
    "in this server's launch settings, restart the app or session, then "
    "start a new worker if needed. Claude Code only: the orchestrator can "
    "instead call enable_tools in the main session and then start a new "
    "worker. If enable_tools refuses a pack, the tool set was fixed at "
    "startup by the launch settings: do not retry until a person changes "
    "those settings and restarts the server."
)
EXPECTED_NOTE = "tools/list_changed was sent. " + EXPECTED_BODY
EXPECTED_NOOP_NOTE = (
    "These packs were already on, so no list change was sent. "
    + EXPECTED_BODY
)

EXPECTED_SENTENCE = (
    "Workers and subagents only see the tools that were on when they "
    "started: in the Claude Desktop extension, turn on 'Load every tool "
    "at startup', restart Claude Desktop, and start a new session; in a "
    "hand-configured client, start the server with KS4W_MODE set to a "
    "comma list of packs; in Claude Code, you can instead enable packs in "
    "the main session before starting workers."
)

EXPECTED_LOCKED_REFUSAL = (
    "The tool surface was fixed at startup by the host "
    "(KS4W_PACK_POLICY=locked, or the 'Lock the tool set at startup' "
    "setting). Only a human can change this launch preference: untick "
    "that setting in Claude Desktop, or restart a hand-configured server "
    "with a different KS4W_MODE or pack policy. Do not retry until then."
)


@pytest.fixture(autouse=True)
def _clean_env(monkeypatch):
    for name in (packs.ENV_MODE, packs.ENV_PACK_POLICY,
                 *packs.toggle_env_names()):
        monkeypatch.delenv(name, raising=False)


@pytest.fixture(autouse=True)
def _restore_enabled():
    saved = dict(packs._ENABLED)
    yield
    packs._ENABLED.clear()
    packs._ENABLED.update(saved)


# ------------------------------------------------------------- (a) the note


def test_the_note_is_the_pinned_slot_text():
    assert packs.LIST_CHANGED_NOTE == EXPECTED_NOTE


def test_the_note_names_this_servers_startup_variable():
    """{MODE_ENV} is substituted per repo. Naming the sibling's variable
    would send the user to a setting this server does not read."""
    assert packs.ENV_MODE == "KS4W_MODE"
    assert packs.ENV_MODE in packs.LIST_CHANGED_NOTE
    for foreign in ("KS4P_MODE", "KS4XL_MODE"):
        assert foreign not in packs.LIST_CHANGED_NOTE


def test_the_note_no_longer_tells_an_agent_to_re_fetch():
    """The retired sentence asked the agent to do the one thing it cannot."""
    assert "re-fetch the tool list" not in packs.LIST_CHANGED_NOTE


def test_a_successful_enable_carries_the_note():
    packs._ENABLED.update({n: False for n in packs.pack_tools("references")})
    result = packs.enable(["references"])
    assert result["enabled"] == ["references"]
    assert result["note"] == EXPECTED_NOTE


def test_a_no_op_re_enable_carries_the_note_too():
    """The agent that most needs the instruction is the one whose pack is
    already on and whose tool list still does not have it. Attaching the
    note only when something flipped left exactly that agent with nothing.
    """
    packs.enable(["references"])
    again = packs.enable(["references"])
    assert again["enabled"] == []
    assert again["already_enabled"] == ["references"]
    assert again["note"] == EXPECTED_NOOP_NOTE


def test_a_no_op_does_not_claim_a_notification_was_sent():
    """Nothing flipped, so nothing was emitted. Saying otherwise is a lie
    told to the one caller trying to work out why its tool list has not
    changed, and it is the sentence that would send them looking for a
    notification that never existed."""
    packs.enable(["references"])
    note = packs.enable(["references"])["note"]
    assert note.startswith(
        "These packs were already on, so no list change was sent."
    )
    assert "tools/list_changed was sent" not in note


def test_both_prefixes_carry_the_same_body():
    """Whichever prefix it gets, the agent's next move is identical."""
    packs._ENABLED.update({n: False for n in packs.pack_tools("review")})
    changed = packs.enable(["review"])["note"]
    unchanged = packs.enable(["review"])["note"]
    assert changed.endswith(EXPECTED_BODY)
    assert unchanged.endswith(EXPECTED_BODY)
    assert changed != unchanged
    assert packs.pack_note(True) == EXPECTED_NOTE
    assert packs.pack_note(False) == EXPECTED_NOOP_NOTE


def test_a_partial_enable_counts_as_changed():
    """One pack already on, one not: something WAS emitted."""
    packs.enable(["review"])
    packs._ENABLED.update({n: False for n in packs.pack_tools("assembly")})
    result = packs.enable(["review", "assembly"])
    assert result["enabled"] == ["assembly"]
    assert result["already_enabled"] == ["review"]
    assert result["note"] == EXPECTED_NOTE


def test_every_pack_and_everything_carry_it():
    for arg in (["review"], ["everything"], ["review", "assembly"]):
        note = packs.enable(arg)["note"]
        assert note.endswith(EXPECTED_BODY)
        assert note in (EXPECTED_NOTE, EXPECTED_NOOP_NOTE)


# ------------------------------------------------- (b) read before calling


def test_the_server_instructions_carry_the_worker_sentence():
    text = server.mcp.instructions or ""
    assert EXPECTED_SENTENCE in text


def test_the_get_workflows_index_carries_the_worker_sentence():
    note = workflows.get_workflows()["note"]
    assert EXPECTED_SENTENCE in note


def test_the_sentence_is_one_string_in_both_places():
    """Two copies of a sentence drift. Both surfaces read the constant."""
    assert packs.WORKER_PACK_SENTENCE == EXPECTED_SENTENCE
    assert packs.WORKER_PACK_SENTENCE in (server.mcp.instructions or "")
    assert packs.WORKER_PACK_SENTENCE in workflows.get_workflows()["note"]


# --------------------------------------------- (c) the locked refusal says so


def test_the_locked_refusal_is_the_pinned_slot_text(monkeypatch):
    """P1-S2-09. The 2.2.1 candidate opened this refusal with "An
    administrator locked the tool packs for this install.", which is false
    when a person ticked the lock box themselves."""
    monkeypatch.setenv(packs.ENV_PACK_POLICY, "locked")
    with pytest.raises(WordMcpError) as exc:
        packs.enable(["references"])
    assert str(exc.value) == EXPECTED_LOCKED_REFUSAL
    assert packs.LOCKED_REFUSAL == EXPECTED_LOCKED_REFUSAL
    assert "administrator locked" not in str(exc.value).replace(
        "never say an administrator locked it", "")
    assert getattr(exc.value, "code", None) == "CONFLICT"


# ------------------------------------------------------------- house rules


def test_no_em_dash_in_any_of_it():
    for text in (packs.pack_note(True), packs.pack_note(False),
                 packs.WORKER_PACK_SENTENCE, server.mcp.instructions or ""):
        assert "—" not in text
