"""The star nudge: once, quietly, and never in the way.

What these prove is mostly about what does NOT happen. A second start says
nothing. An unwritable state directory says nothing. Nothing reaches stdout.
The wording is pinned, because it is the one string a user reads.
"""

import io
import os

import pytest

from word_mcp.core.star_nudge import (
    MARKER_NAME,
    MESSAGE,
    announce_once,
    marker_path,
)


@pytest.fixture
def state(tmp_path, monkeypatch):
    """Point the product's state directory at a temporary one."""
    monkeypatch.setenv("KS4W_UPDATE_CACHE_DIR", str(tmp_path))
    monkeypatch.delenv("KS4W_STAR_NUDGE", raising=False)
    return tmp_path


def test_the_wording_is_exactly_what_was_approved():
    assert MESSAGE == (
        "KitchenSink4Word is community-supported. If it helps you, star "
        "github.com/KitchenSink4AI/KitchenSink4Word. Shown once."
    )


def test_it_is_shown_on_the_first_start(state):
    out = io.StringIO()
    assert announce_once(stream=out) is True
    assert out.getvalue() == MESSAGE + "\n"
    assert marker_path().name == MARKER_NAME
    assert marker_path().exists(), "the marker records that it was shown"


def test_it_is_not_shown_again(state):
    first, second = io.StringIO(), io.StringIO()
    announce_once(stream=first)
    assert announce_once(stream=second) is False
    assert second.getvalue() == ""


def test_the_toggle_silences_it(state, monkeypatch):
    monkeypatch.setenv("KS4W_STAR_NUDGE", "off")
    out = io.StringIO()
    assert announce_once(stream=out) is False
    assert out.getvalue() == ""
    assert not marker_path().exists(), "a silenced start does not burn the marker"


def test_an_unwritable_state_directory_is_silence_not_a_crash(state, monkeypatch):
    monkeypatch.setattr(
        "word_mcp.core.star_nudge.marker_path",
        lambda: (_ for _ in ()).throw(OSError("read-only")),
    )
    out = io.StringIO()
    assert announce_once(stream=out) is False
    assert out.getvalue() == ""


def test_it_never_writes_to_stdout(state, capsys):
    announce_once()
    captured = capsys.readouterr()
    assert captured.out == "", "stdout carries the protocol and stays clean"
    assert MESSAGE in captured.err
