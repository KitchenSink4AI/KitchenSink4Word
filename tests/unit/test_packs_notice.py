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

#: The owner's wording, 2026-09-22, with {MODE_ENV} filled in. Verbatim.
EXPECTED_NOTE = (
    "tools/list_changed was sent. If the new tools are not in your tool "
    "list, this client fixed its list when the session or worker started: "
    "do not retry here. What works in every client: ask the user to add "
    "the packs to KS4W_MODE (comma list) in this server's launch "
    "settings, restart the app or session, then start a new worker if "
    "needed. Claude Code only: the orchestrator can instead call "
    "enable_tools in the main session and then start a new worker. If "
    "enable_tools refuses a pack, an administrator locked the tool set: "
    "do not retry."
)

EXPECTED_SENTENCE = (
    "Workers and subagents only see the tools that were on when they "
    "started: start the server with KS4W_MODE set to a comma list of "
    "packs, or, in Claude Code, enable packs in the main session before "
    "starting workers."
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


def test_the_note_is_the_owners_wording():
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
    assert again["note"] == EXPECTED_NOTE


def test_every_pack_and_everything_carry_it():
    for arg in (["review"], ["everything"], ["review", "assembly"]):
        assert packs.enable(arg)["note"] == EXPECTED_NOTE


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


def test_the_locked_refusal_names_an_administrator(monkeypatch):
    monkeypatch.setenv(packs.ENV_PACK_POLICY, "locked")
    with pytest.raises(WordMcpError) as exc:
        packs.enable(["references"])
    message = str(exc.value)
    assert "An administrator locked the tool packs for this install." in message
    assert getattr(exc.value, "code", None) == "CONFLICT"


# ------------------------------------------------------------- house rules


def test_no_em_dash_in_any_of_it():
    for text in (packs.LIST_CHANGED_NOTE, packs.WORKER_PACK_SENTENCE,
                 server.mcp.instructions or ""):
        assert "—" not in text
