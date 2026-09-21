"""The invisible-instance lifecycle in com/bridge._word().

THE DEFECT. Through 2.2.0 the finally block was:

    if app is not None:
        with contextlib.suppress(Exception):
            app.Quit(_WD_DO_NOT_SAVE)
    _INVISIBLE_PIDS.pop(tid, None)

Three failures stacked in four lines. Documents this instance still had
open were never closed, and Word refuses to Quit an instance holding one
it considers unsaved. The refusal was suppressed, so nothing recorded that
the instance had not gone. The PID record was then dropped without anyone
checking whether the process had actually exited, which threw away the only
handle that could ever kill it. That is how the 2026-09-21 field test found
two WINWORD.EXE with /Automation -Embedding still running long after the
calls that made them had returned.

These tests drive the finally block with a fake app object, because the
failure is in the bookkeeping and a fake can be made to refuse Quit on
demand, which a real Word cannot. One live test at the end does the
round trip against a real hidden Word where one is installed.

SAFETY. Every kill path in bridge.py takes PIDs recorded at DispatchEx
time and, before terminating one, checks that its command line carries
/Automation. The user may have Word open with a document that matters. The
refusal-to-kill path is tested here explicitly.
"""

from __future__ import annotations

import subprocess
import sys

import pytest

from word_mcp.com import bridge


# ------------------------------------------------------------- fake Word


class _FakeDocuments:
    def __init__(self, count: int, refuse: bool = False):
        self._open = list(range(count))
        self.refuse = refuse
        self.closed = 0

    @property
    def Count(self):  # noqa: N802 - COM casing
        return len(self._open)

    def __call__(self, index):  # Documents(1)
        doc = self

        class _Doc:
            def Close(self, mode):  # noqa: N802 - COM casing
                if doc.refuse:
                    raise RuntimeError("document is in use")
                doc._open.pop(0)
                doc.closed += 1

        return _Doc()


class _FakeApp:
    """Only what _word()'s finally block touches."""

    def __init__(self, docs: int = 0, quit_raises: bool = False,
                 refuse_close: bool = False):
        self.Documents = _FakeDocuments(docs, refuse=refuse_close)
        self.quit_raises = quit_raises
        self.quit_calls = 0
        self.Visible = True
        self.DisplayAlerts = None

    def Quit(self, mode):  # noqa: N802 - COM casing
        self.quit_calls += 1
        if self.quit_raises:
            raise RuntimeError("Word cannot quit: a document is open")


@pytest.fixture(autouse=True)
def _clean(monkeypatch):
    bridge._INVISIBLE_PIDS.clear()
    bridge._LIFECYCLE_NOTES.clear()
    yield
    bridge._INVISIBLE_PIDS.clear()


# ------------------------------------------- 1. documents close before Quit


def test_open_documents_are_closed_before_quit():
    app = _FakeApp(docs=3)
    bridge._close_open_documents(app)
    assert app.Documents.Count == 0
    assert app.Documents.closed == 3


def test_a_document_that_refuses_to_close_is_reported_not_retried():
    """One refusal ends the loop: spinning on the same document is how a
    cleanup path turns into the hang it was meant to prevent."""
    app = _FakeApp(docs=2, refuse_close=True)
    bridge._close_open_documents(app)
    assert app.Documents.closed == 0
    assert any("could not close a document" in n
               for n in bridge._LIFECYCLE_NOTES)


def test_a_dead_app_object_closes_nothing_and_raises_nothing():
    class _Dead:
        @property
        def Documents(self):
            raise RuntimeError("RPC server is unavailable")

    bridge._close_open_documents(_Dead())  # must not raise


# --------------------------------------------- 2. a failed Quit is not silent


def test_a_failed_quit_is_recorded(monkeypatch):
    """contextlib.suppress(Exception) around Quit is the line that let an
    orphan exist with nothing anywhere recording that it did."""
    app = _FakeApp(quit_raises=True)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    _drive_finally(monkeypatch, app, pids={4242})
    assert app.quit_calls == 1
    assert any("Word.Quit failed" in n for n in bridge._LIFECYCLE_NOTES), (
        f"nothing recorded the failed Quit: {list(bridge._LIFECYCLE_NOTES)}"
    )


def test_a_failed_quit_does_not_break_a_successful_operation(monkeypatch):
    """The operation already returned its result. A cleanup problem is
    announced, not raised over the top of a success."""
    app = _FakeApp(quit_raises=True)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    out = _drive_finally(monkeypatch, app, pids={4242})
    assert out == "body ran"


# ------------------------------------- 3. the PID is verified, not assumed


def test_a_surviving_pid_is_killed_before_the_record_is_dropped(monkeypatch):
    app = _FakeApp(quit_raises=True)
    killed: list = []
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set(pids))
    monkeypatch.setattr(
        bridge, "_kill_pids",
        lambda pids, why="": (killed.extend(sorted(pids)), set(pids))[1],
    )
    _drive_finally(monkeypatch, app, pids={4242})
    assert killed == [4242], "the surviving instance was never terminated"
    assert not bridge._INVISIBLE_PIDS, "a killed pid stayed on the books"


def test_a_pid_that_survives_the_kill_stays_on_the_books(monkeypatch):
    """The record is the only handle that can end it. Dropping it
    unverified is the 2.2.0 defect; dropping it after a FAILED kill would
    be the same defect one step later."""
    app = _FakeApp(quit_raises=True)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set(pids))
    monkeypatch.setattr(bridge, "_kill_pids", lambda pids, why="": set())
    _drive_finally(monkeypatch, app, pids={4242})
    remaining = {p for s in bridge._INVISIBLE_PIDS.values() for p in s}
    assert remaining == {4242}


def test_a_pid_that_exited_leaves_no_record(monkeypatch):
    app = _FakeApp()
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    _drive_finally(monkeypatch, app, pids={4242})
    assert not bridge._INVISIBLE_PIDS
    assert app.quit_calls == 1


def test_survivors_waits_rather_than_checking_once(monkeypatch):
    """Quit is asynchronous. A single check right after it would report a
    false orphan and kill a process that was already on its way out."""
    calls = {"n": 0}

    def _pids():
        calls["n"] += 1
        return {7} if calls["n"] < 3 else set()

    monkeypatch.setattr(bridge, "_winword_pids", _pids)
    monkeypatch.setattr(bridge.time, "sleep", lambda s: None)
    assert bridge._survivors({7}, timeout=5.0) == set()
    assert calls["n"] >= 3


def test_survivors_gives_up_at_the_deadline(monkeypatch):
    monkeypatch.setattr(bridge, "_winword_pids", lambda: {7})
    monkeypatch.setattr(bridge.time, "sleep", lambda s: None)
    assert bridge._survivors({7}, timeout=0.0) == {7}


# ------------------------------------------------- 4. never the user's Word


def test_a_non_automation_pid_is_never_killed(monkeypatch):
    """The paranoid gate. PID bookkeeping is a diff of two tasklist
    snapshots, so a Word the user started in the same instant could in
    principle land in it. An interactive WINWORD.EXE carries no
    /Automation flag and must survive whatever the record says."""
    ran: list = []
    monkeypatch.setattr(bridge, "_is_automation_instance", lambda pid: False)
    monkeypatch.setattr(
        bridge.subprocess, "run",
        lambda *a, **k: ran.append(a) or subprocess.CompletedProcess(a, 0),
    )
    assert bridge._kill_pids({1234}, why="test") == set()
    assert not ran, "taskkill was invoked on a non-automation instance"
    assert any("refusing to terminate" in n for n in bridge._LIFECYCLE_NOTES)


def test_an_unreadable_command_line_does_not_block_the_kill_switch(
        monkeypatch):
    """None means "could not read", not "not mine". The timeout kill
    switch exists to end a hang; a failed probe must not disarm it."""
    monkeypatch.setattr(bridge, "_is_automation_instance", lambda pid: None)
    monkeypatch.setattr(
        bridge.subprocess, "run",
        lambda *a, **k: subprocess.CompletedProcess(a, 0),
    )
    assert bridge._kill_pids({1234}, why="test") == {1234}


def test_kill_invisible_for_thread_only_touches_recorded_pids(monkeypatch):
    seen: list = []
    monkeypatch.setattr(
        bridge, "_kill_pids",
        lambda pids, why="": (seen.append(sorted(pids)), set(pids))[1],
    )
    bridge._INVISIBLE_PIDS[99] = {11, 12}
    assert bridge._kill_invisible_for_thread(99) is True
    assert seen == [[11, 12]]
    assert bridge._kill_invisible_for_thread(1234) is False


# ------------------------------------------------------ 5. the exit sweep


def test_an_atexit_sweep_is_registered():
    """The backstop for every path _word()'s finally block cannot reach.

    atexit keeps no readable registry, so the registration is asserted at
    the source and the callable is asserted to exist and be callable with
    no arguments, which is what atexit will do to it."""
    import inspect

    assert callable(bridge._sweep_invisible_instances)
    src = inspect.getsource(bridge)
    assert "@atexit.register\ndef _sweep_invisible_instances" in src, (
        "bridge no longer registers a process-exit sweep"
    )
    bridge._sweep_invisible_instances()  # empty books: a no-op, never raises


def test_the_sweep_kills_only_live_recorded_pids(monkeypatch):
    killed: list = []
    monkeypatch.setattr(bridge, "_winword_pids", lambda: {11, 77})
    monkeypatch.setattr(
        bridge, "_kill_pids",
        lambda pids, why="": (killed.extend(sorted(pids)), set(pids))[1],
    )
    bridge._INVISIBLE_PIDS[1] = {11, 12}   # 12 already gone
    bridge._INVISIBLE_PIDS[2] = {99}       # never existed
    bridge._sweep_invisible_instances()
    assert killed == [11]
    assert not bridge._INVISIBLE_PIDS


# --------------------------------------------------- 6. no desktop windows


def test_every_console_helper_hides_its_window():
    """Author directive 2026-09-06: nothing this server spawns may flash a
    console. tasklist, taskkill and the command-line probe are console
    programs."""
    import inspect

    src = inspect.getsource(bridge)
    assert src.count("creationflags=_NO_WINDOW") >= 3, (
        "a subprocess in bridge.py launches without CREATE_NO_WINDOW"
    )
    if sys.platform == "win32":
        assert bridge._NO_WINDOW == subprocess.CREATE_NO_WINDOW


# ------------------------------------------------------------ the harness


def _drive_finally(monkeypatch, app, pids):
    """Run _word()'s body and finally block against a fake app.

    DispatchEx and pythoncom are replaced, so no COM is touched and no
    process is created; everything else in _word() is the shipped code.
    """
    import types

    fake_com = types.ModuleType("pythoncom")
    fake_com.CoInitialize = lambda: None
    fake_com.CoUninitialize = lambda: None
    fake_client = types.ModuleType("win32com.client")
    fake_client.DispatchEx = lambda progid: app
    fake_parent = types.ModuleType("win32com")
    fake_parent.client = fake_client
    monkeypatch.setitem(sys.modules, "pythoncom", fake_com)
    monkeypatch.setitem(sys.modules, "win32com", fake_parent)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    seq = [set(), set(pids)]
    monkeypatch.setattr(bridge, "_winword_pids", lambda: seq.pop(0) if seq
                        else set(pids))
    with bridge._word() as got:
        assert got is app
        result = "body ran"
    return result


# ----------------------------------------------- 7. the real hidden Word


def _word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application").Close()
        return True
    except OSError:
        return False


@pytest.mark.live
@pytest.mark.skipif(not _word_available(),
                    reason="Word/pywin32 not available on this machine")
def test_a_real_hidden_instance_leaves_no_orphan(tmp_path):
    """The end-to-end shape of the field-test defect, against real Word.

    The instance is created invisible, given a document (the state that
    used to make Quit refuse), and released. Nothing may be left running,
    and nothing that existed before may be disturbed.
    """
    import pythoncom  # noqa: F401  (imported by _word itself)

    before = bridge._winword_pids()
    created: set = set()
    with bridge._word() as app:
        created = bridge._INVISIBLE_PIDS.get(
            __import__("threading").get_ident(), set()
        ) - before
        doc = app.Documents.Add()
        doc.Content.Text = "orphan check"
        assert int(app.Documents.Count) >= 1
    after = bridge._winword_pids()
    assert not (created & after), (
        f"invisible WINWORD.EXE {sorted(created & after)} outlived the "
        "context manager"
    )
    assert before - after == set(), (
        "a Word instance this test did not create disappeared"
    )
    assert not bridge._INVISIBLE_PIDS
