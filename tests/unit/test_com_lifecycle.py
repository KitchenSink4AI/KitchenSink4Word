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
checking whether the process had actually exited. That is how the
2026-09-21 field test found two WINWORD.EXE with /Automation -Embedding
still running long after the calls that made them had returned.

These tests drive the finally block with a fake app object, because the
failure is in the bookkeeping and a fake can be made to refuse Quit on
demand, which a real Word cannot. One live test at the end does the
round trip against a real hidden Word where one is installed.

SAFETY: THE KITCHENSINK4PPT 1.3.1 RULE. Nothing in this package
force-ends a process, on any path: not a survivor of Quit, not a
timed-out operation, not at interpreter exit. The instance a call starts
is released with an orderly Quit() on its own COM object. Ownership needs
positive evidence and a failed reading is a no: an unreadable process
table is unknown, never empty, and an unreadable command line owns
nothing. A source guard and the behavioural tests below pin all of it.
"""

from __future__ import annotations

import pathlib
import subprocess
import sys
import threading

import pytest

import word_mcp
from word_mcp.com import bridge
from word_mcp.core.errors import WordBlocked


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


@pytest.fixture
def no_process_is_ended(monkeypatch):
    """Fail the test the moment anything tries to end a process through
    the one route this module ever had. Reading the process table is
    fine; ending a process is what is gone."""
    real_run = subprocess.run

    def guarded(cmd, *args, **kwargs):
        if cmd and str(cmd[0]).lower() in ("taskkill", "taskkill.exe"):
            pytest.fail(f"a process-ending command was run: {cmd}")
        return real_run(cmd, *args, **kwargs)

    monkeypatch.setattr(bridge.subprocess, "run", guarded)


_LINGER = "the Word instance this call launched did not exit"


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


# ------------------------- 3. a survivor is verified, reported, never ended


def test_an_owned_survivor_is_reported_and_not_ended(
        monkeypatch, no_process_is_ended):
    """The instance outlived Quit and its command line READS as an
    automation instance: it is reported with the advice a human can act
    on, and nothing here ends it."""
    app = _FakeApp(quit_raises=True)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set(pids))
    monkeypatch.setattr(bridge, "_is_automation_instance", lambda pid: True)
    out = _drive_finally(monkeypatch, app, pids={4242})
    assert out == "body ran"
    notes = list(bridge._LIFECYCLE_NOTES)
    assert any(n.startswith(_LINGER) for n in notes), notes
    assert any("safe to end via Task Manager" in n for n in notes)
    assert not any("terminat" in n for n in notes), notes
    assert not bridge._INVISIBLE_PIDS, "a survivor stayed on the books"


@pytest.mark.parametrize("reading", [None, False])
def test_an_unreadable_or_foreign_command_line_owns_nothing(
        monkeypatch, no_process_is_ended, reading):
    """None is "could not read", and that owns nothing, exactly like a
    definite no. No claim is made about the process and no advice is
    given about ending it. (2.2.0 onward briefly treated None as licence
    to proceed with a kill; that rule is gone with the kill itself.)"""
    app = _FakeApp(quit_raises=True)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set(pids))
    monkeypatch.setattr(bridge, "_is_automation_instance",
                        lambda pid: reading)
    _drive_finally(monkeypatch, app, pids={4242})
    notes = list(bridge._LIFECYCLE_NOTES)
    assert not any(_LINGER in n for n in notes), notes
    assert not any("Task Manager" in n for n in notes), notes
    assert not bridge._INVISIBLE_PIDS


def test_a_pid_that_exited_leaves_no_record_and_no_claim(monkeypatch):
    app = _FakeApp()
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    monkeypatch.setattr(
        bridge, "_is_automation_instance",
        lambda pid: pytest.fail("nothing survived, so nothing is probed"),
    )
    _drive_finally(monkeypatch, app, pids={4242})
    assert not bridge._INVISIBLE_PIDS
    assert app.quit_calls == 1
    assert not bridge._LIFECYCLE_NOTES


def test_quit_is_still_issued_on_the_instance_the_call_started(monkeypatch):
    """The orderly release stays: Quit on the call's own COM object is not
    a force-end and is unchanged by the no-kill rule."""
    app = _FakeApp(docs=1)
    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    _drive_finally(monkeypatch, app, pids={4242})
    assert app.quit_calls == 1
    assert app.Documents.closed == 1


def test_survivors_waits_rather_than_checking_once(monkeypatch):
    """Quit is asynchronous. A single check right after it would report a
    false orphan for a process that was already on its way out."""
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


def test_an_unreadable_table_is_not_evidence_of_a_survivor(monkeypatch):
    """G2b: unknown mid-poll is not a survivor, and polling again cannot
    turn it into one."""
    monkeypatch.setattr(bridge, "_winword_pids", lambda: None)
    monkeypatch.setattr(bridge.time, "sleep", lambda s: None)
    assert bridge._survivors({7}, timeout=5.0) == set()


# ------------------------- 4. the process table: unknown is never empty


def _tasklist(monkeypatch, *, returncode=0, stdout="", raises=None):
    seen: list = []

    def fake_run(cmd, *args, **kwargs):
        seen.append((cmd, kwargs))
        if raises is not None:
            raise raises
        return subprocess.CompletedProcess(cmd, returncode, stdout, "")

    monkeypatch.setattr(bridge.subprocess, "run", fake_run)
    return seen


def test_a_readable_table_lists_the_word_pids(monkeypatch):
    seen = _tasklist(monkeypatch, stdout=(
        '"WINWORD.EXE","1234","Console","1","250,000 K"\n'
        '"WINWORD.EXE","5678","Console","1","90,000 K"\n'
    ))
    assert bridge._winword_pids() == {1234, 5678}
    assert seen[0][1].get("creationflags") == bridge._NO_WINDOW


def test_no_word_running_is_an_empty_set_not_unknown(monkeypatch):
    _tasklist(monkeypatch, stdout=(
        "INFO: No tasks are running which match the specified criteria.\n"
    ))
    assert bridge._winword_pids() == set()


@pytest.mark.parametrize("kwargs", [
    {"raises": subprocess.TimeoutExpired("tasklist", 30)},
    {"raises": FileNotFoundError("tasklist")},
    {"returncode": 1, "stdout": "ERROR: Access is denied.\n"},
    {"returncode": 0, "stdout": ""},
    {"returncode": 0, "stdout": '"WINWORD.EXE","not-a-pid","Console"\n'},
    {"returncode": 0, "stdout": "WINWORD.EXE with no csv columns\n"},
], ids=["timeout", "missing", "access-denied", "no-output", "bad-pid",
        "bad-row"])
def test_an_unreadable_table_is_unknown(monkeypatch, kwargs):
    """None, never set(): this used to read as "nothing is running", which
    made every Word already running look like one this call started."""
    _tasklist(monkeypatch, **kwargs)
    assert bridge._winword_pids() is None


@pytest.mark.parametrize("before, after", [
    (None, {1, 2}),
    ({1}, None),
    (None, None),
    ({1}, {1}),          # nothing appeared
    ({1}, {1, 2, 3}),    # two appeared at once: no telling which is ours
    (set(), {2, 3}),
])
def test_ownership_needs_exactly_one_new_pid_from_two_readings(
        before, after):
    assert bridge._acquisition_token(before, after) is None


def test_one_new_pid_across_two_readings_is_the_token():
    """The user's Word (pid 1) is in both readings and never in the
    difference; the one process that appeared is the call's."""
    assert bridge._acquisition_token({1}, {1, 2}) == {2}
    assert bridge._acquisition_token(set(), {2}) == {2}


def test_an_unreadable_table_at_start_records_nothing(monkeypatch):
    """The whole path, not just the helper: a failed tasklist before
    DispatchEx means this call names no process at all."""
    app = _FakeApp()
    recorded: list = []

    def body_sees(_app):
        recorded.append(
            dict(bridge._INVISIBLE_PIDS).get(threading.get_ident())
        )

    monkeypatch.setattr(
        bridge, "_survivors",
        lambda pids, **kw: pytest.fail("no record, so nothing to verify"),
    )
    _drive_finally(monkeypatch, app, pids={4242}, before=None,
                   during=body_sees)
    assert recorded == [None]
    assert app.quit_calls == 1


def test_a_readable_start_records_the_one_new_pid(monkeypatch):
    app = _FakeApp()
    recorded: list = []

    def body_sees(_app):
        recorded.append(
            dict(bridge._INVISIBLE_PIDS).get(threading.get_ident())
        )

    monkeypatch.setattr(bridge, "_survivors", lambda pids, **kw: set())
    _drive_finally(monkeypatch, app, pids={4242}, before={99},
                   during=body_sees, after={99, 4242})
    assert recorded == [{4242}]
    assert not bridge._INVISIBLE_PIDS


# ------------------------------ 5. a timed-out operation ends nothing


def test_nothing_in_the_package_can_end_a_process():
    """The KitchenSink4PPT 1.3.1 source guard, over the whole package.
    Force-ending is easy to reintroduce and expensive to notice, so its
    absence is asserted at the source. (os.kill(pid, 0) in xproc and
    safesave is a POSIX liveness probe that never runs on Windows, and is
    not one of the routes this guard claims.)"""
    root = pathlib.Path(word_mcp.__file__).parent
    offenders = []
    for path in sorted(root.rglob("*.py")):
        body = path.read_text(encoding="utf-8")
        for forbidden in ("taskkill", "TerminateProcess", ".Terminate("):
            if forbidden in body:
                offenders.append(
                    f"{path.relative_to(root)} contains {forbidden}"
                )
    assert not offenders, "; ".join(offenders)


def test_there_is_no_exit_sweep_left_to_end_anything():
    """The unmerged field-test round added an atexit sweep whose only job
    was to force-end recorded instances. With nothing ever ended, it had
    no job left and was removed rather than kept as a trap."""
    src = pathlib.Path(bridge.__file__).read_text(encoding="utf-8")
    assert "atexit" not in src
    assert not hasattr(bridge, "_sweep_invisible_instances")
    assert not hasattr(bridge, "_kill_pids")
    assert not hasattr(bridge, "_kill_invisible_for_thread")


def _stuck_run(monkeypatch, *, record=None, dialogs_up=None):
    """_run_bounded on an operation that never answers. Deterministic: the
    worker waits on an event the test owns, and releases it afterwards so
    the COM lock is returned. The grace wait is shortened."""
    from word_mcp.com import dialogs, serial

    monkeypatch.setattr(bridge, "TIMEOUT_GRACE_SECONDS", 0.2)
    monkeypatch.setattr(dialogs, "pending_dialogs",
                        lambda pids=None: list(dialogs_up or []))
    release = threading.Event()

    def stuck():
        if record:
            bridge._INVISIBLE_PIDS[threading.get_ident()] = set(record)
        release.wait(30)
        return {"late": True}

    try:
        with pytest.raises(WordBlocked) as exc_info:
            bridge._run_bounded("stuck-op", 0.5, stuck)
    finally:
        release.set()
        for t in threading.enumerate():
            if t.name == "ks4w-stuck-op":
                t.join(10)
    assert serial.lock_snapshot()["held"] is False
    return str(exc_info.value)


def test_a_timed_out_operation_names_the_pid_and_ends_nothing(
        monkeypatch, no_process_is_ended):
    """It used to force-end that process. A timeout cannot revalidate a
    hung apartment, so the refusal reports and nothing is ended."""
    message = _stuck_run(monkeypatch, record={5150})
    assert "5150" in message
    assert "was not running when this call began" in message
    assert "this call did not force-end it" in message
    assert "was NOT cancelled" in message
    assert "terminated" not in message
    assert "was aborted" not in message


def test_a_timeout_with_no_record_claims_no_process(
        monkeypatch, no_process_is_ended):
    message = _stuck_run(monkeypatch, record=None)
    assert (
        "no newly started Word process could be identified; no process "
        "was force-ended"
    ) in message


def test_a_timeout_names_the_dialog_that_is_up(
        monkeypatch, no_process_is_ended):
    message = _stuck_run(
        monkeypatch, record={5150},
        dialogs_up=[{"title": "Microsoft Word", "class": "#32770"}],
    )
    assert "Word has a dialog open: Microsoft Word" in message


def test_a_worker_that_finishes_in_the_grace_window_is_a_success(
        monkeypatch):
    """R6-1: the deadline cancels nothing, so an operation that finishes
    while the refusal is being considered has succeeded, and saying it
    timed out would be a false failure the caller acts on."""
    from word_mcp.com import dialogs

    monkeypatch.setattr(dialogs, "pending_dialogs", lambda pids=None: [])
    gate = threading.Event()
    finish = threading.Timer(1.5, gate.set)
    finish.start()
    try:
        out = bridge._run_bounded(
            "grace-op", 0.5, lambda: (gate.wait(30), {"ok": 1})[1]
        )
    finally:
        finish.cancel()
        gate.set()
    assert out == {"ok": 1}


def test_a_worker_that_fails_in_the_grace_window_raises_its_own_error(
        monkeypatch):
    from word_mcp.com import dialogs
    from word_mcp.core.errors import TargetNotFound

    monkeypatch.setattr(dialogs, "pending_dialogs", lambda pids=None: [])
    gate = threading.Event()
    finish = threading.Timer(1.5, gate.set)
    finish.start()

    def fails_late():
        gate.wait(30)
        raise TargetNotFound("the operation's own answer")

    try:
        with pytest.raises(TargetNotFound, match="own answer"):
            bridge._run_bounded("grace-err", 0.5, fails_late)
    finally:
        finish.cancel()
        gate.set()


# --------------------------------------------------- 6. no desktop windows


def test_every_console_helper_hides_its_window():
    """Author directive 2026-09-06: nothing this server spawns may flash a
    console. tasklist (the process table and zombie_check) and the
    command-line probe are console programs."""
    import inspect

    src = inspect.getsource(bridge)
    assert src.count("subprocess.run(") == src.count(
        "creationflags=_NO_WINDOW"
    ), "a subprocess in bridge.py launches without CREATE_NO_WINDOW"
    assert src.count("creationflags=_NO_WINDOW") >= 3
    if sys.platform == "win32":
        assert bridge._NO_WINDOW == subprocess.CREATE_NO_WINDOW


# ------------------------------------------------------------ the harness


def _drive_finally(monkeypatch, app, pids, *, before=frozenset(),
                   after=None, during=None):
    """Run _word()'s body and finally block against a fake app.

    DispatchEx and pythoncom are replaced, so no COM is touched and no
    process is created; everything else in _word() is the shipped code.
    The process table reads `before` ahead of DispatchEx and `after`
    (default: before plus pids) behind it; None stands for a table that
    could not be read.
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
    first = None if before is None else set(before)
    if after is None:
        after = (set(before) if before is not None else set()) | set(pids)
    seq = [first, set(after)]
    monkeypatch.setattr(bridge, "_winword_pids", lambda: seq.pop(0) if seq
                        else set(pids))
    with bridge._word() as got:
        assert got is app
        if during is not None:
            during(got)
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
    used to make Quit refuse), and released by an orderly Quit. Nothing
    may be left running, and nothing that existed before may be disturbed.
    """
    import pythoncom  # noqa: F401  (imported by _word itself)

    before = bridge._winword_pids()
    assert before is not None, "the process table could not be read"
    created: set = set()
    with bridge._word() as app:
        created = bridge._INVISIBLE_PIDS.get(
            threading.get_ident(), set()
        ) - before
        doc = app.Documents.Add()
        doc.Content.Text = "orphan check"
        assert int(app.Documents.Count) >= 1
    after = bridge._winword_pids()
    assert after is not None, "the process table could not be read"
    assert not (created & after), (
        f"invisible WINWORD.EXE {sorted(created & after)} outlived the "
        "context manager"
    )
    assert before - after == set(), (
        "a Word instance this test did not create disappeared"
    )
    assert not bridge._INVISIBLE_PIDS
