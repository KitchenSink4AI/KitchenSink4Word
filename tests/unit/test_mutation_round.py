"""One test per surviving mutant from the 2026-09-06 mutation round.

Spec: ``Draft/Working Files/Agent Results/20260906_wordppt_mutation_round.md``
(Word section: 319 mutants, 186 killed, 58% targeted kill rate; 55 class-(a)
survivors confirmed against the full 1390-test suite).

Every test here exists because a specific sabotage of the safety core passed
the whole suite. Each one was written against its mutant and hand-verified
RED with the mutant applied, GREEN with it reverted.

Test discipline carried over from the Excel sibling's round (whose method
notes recorded lock-timing tests false-redding under CPU load): NO test in
this file waits on a real elapsed interval to prove a policy. Ages come from
hand-written lockfiles carrying a chosen timestamp, idle gaps come from
``os.utime`` backdating, and the exact-value boundaries are measured against
a FROZEN wall clock injected into the module under test (monotonic and sleep
stay real), because a real elapsed interval can never land exactly ON the
boundary, which is the only place ``>`` and ``>=`` differ.
"""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path

import pytest
from docx import Document

from word_mcp import envelope
from word_mcp.com import xproc
from word_mcp.core import locate, safesave, sandbox
from word_mcp.core.errors import (
    AmbiguousTarget,
    DocumentNotFound,
    ValidationFailed,
    WordMcpError,
)
from word_mcp.core.package import DocxPackage
from word_mcp.core.safesave import MutationLockTimeout
from word_mcp.core.sandbox import SandboxViolation, check_path
from word_mcp.ops import backups as bk

WIN = sys.platform == "win32"
winonly = pytest.mark.skipif(not WIN, reason="Windows-only code path")


# ------------------------------------------------------------------ helpers


def _doc(where: Path, name: str = "doc.docx", text: str = "Anchor.") -> Path:
    f = where / name
    d = Document()
    d.add_paragraph(text)
    d.save(str(f))
    return f


class _FrozenClock:
    """A stand-in for the ``time`` module with a wall clock we choose.

    ``monotonic`` and ``sleep`` stay real so wait loops still terminate; only
    ``time.time()`` is frozen, which is what age arithmetic reads.
    """

    def __init__(self, now: float):
        self._now = now

    def time(self) -> float:
        return self._now

    def set(self, now: float) -> None:
        self._now = now

    def monotonic(self):
        return time.monotonic()

    def sleep(self, seconds):
        return time.sleep(seconds)


class _PollClock:
    """Real wall clock, CONTROLLED monotonic.

    ``sleep`` advances the monotonic reading instead of spending the time, so
    a wait loop's deadline arithmetic can be driven exactly and counted. This
    is what lets a test say "the refusal arrived only after the wait was
    spent" without ever asserting on elapsed real time.
    """

    def __init__(self, start: float = 1000.0):
        self._m = start
        self.sleeps = 0

    def time(self) -> float:
        return time.time()

    def monotonic(self) -> float:
        return self._m

    def sleep(self, seconds: float) -> None:
        self.sleeps += 1
        self._m += seconds


def _unused_drive():
    """A drive letter with nothing mounted on it, or None."""
    for letter in "QYXWVU":
        if not os.path.exists(letter + ":" + chr(92)):
            return letter
    return None


class _FakeKernel32:
    """Minimal kernel32 stand-in for the liveness-probe fallbacks."""

    def __init__(self, *, open_raises=False, open_returns=1, exitcode_ok=False):
        self._open_raises = open_raises
        self._open_returns = open_returns
        self._exitcode_ok = exitcode_ok
        self.closed = 0

    def OpenProcess(self, *args):
        if self._open_raises:
            raise OSError("simulated probe failure")
        return self._open_returns

    def GetExitCodeProcess(self, handle, out):
        return 1 if self._exitcode_ok else 0

    def GetLastError(self):
        return 0

    def CloseHandle(self, handle):
        self.closed += 1
        return 1

    def GetProcessTimes(self, *args):
        return 0


@pytest.fixture()
def sacrificial():
    """A live child process that reports its own PID on stdout.

    Its own report is used rather than ``Popen.pid`` because on Windows a
    venv launcher can mean the interpreter that ran is not the process that
    was spawned.
    """
    proc = subprocess.Popen(
        [sys.executable, "-c",
         "import os,sys,time;print(os.getpid(),flush=True);time.sleep(120)"],
        stdout=subprocess.PIPE, text=True,
    )
    try:
        pid = int(proc.stdout.readline().strip())
        yield proc, pid
    finally:
        proc.kill()
        proc.wait()


@pytest.fixture()
def dead_pid():
    """A PID whose process has certainly exited (and is still un-recycled:
    the Popen handle is held open until the fixture tears down)."""
    proc = subprocess.Popen(
        [sys.executable, "-c", "import os,sys;print(os.getpid(),flush=True)"],
        stdout=subprocess.PIPE, text=True,
    )
    pid = int(proc.stdout.readline().strip())
    proc.wait()
    yield pid
    proc.stdout.close()


def _mutate_in_place(doc: Path, text: str) -> None:
    """Replace a document's content the way the save path does: build the new
    bytes elsewhere and ``os.replace`` them onto the target.

    Writing through the document's own handle would be wrong here. Slots are
    HARDLINKS to the document's inode, so an in-place rewrite changes the
    backup too, and a test that did that would prove nothing about rotation.
    """
    tmp = doc.with_name(doc.name + ".newcontent")
    d = Document()
    d.add_paragraph(text)
    d.save(str(tmp))
    os.replace(tmp, doc)


def _lockfile(path: Path, **fields) -> Path:
    payload = {"pid": os.getpid(), "token": "foreign-token",
               "time": time.time(), "host": ""}
    payload.update(fields)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload), encoding="utf-8")
    return path


# ================================================== W-A1  the sandbox gate
# core/safesave.py:561  check_path(doc_path, "modify document") -> pass
#
# The mutation path's own gate. The sibling ppt repo carries this test; this
# repo, with the larger sandbox suite, did not.


class TestWriteLockSandboxGate:
    def test_write_lock_outside_the_roots_refuses(self, tmp_path, monkeypatch):
        inside = tmp_path / "inside"
        inside.mkdir()
        outside = tmp_path / "outside"
        outside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        with pytest.raises(SandboxViolation):
            with safesave.write_lock(outside / "escape.docx"):
                pass

    def test_write_lock_outside_creates_no_slot_folder(
        self, tmp_path, monkeypatch
    ):
        inside = tmp_path / "inside"
        inside.mkdir()
        outside = tmp_path / "outside"
        outside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        with pytest.raises(SandboxViolation):
            with safesave.write_lock(outside / "escape.docx"):
                pass
        assert not (outside / safesave.BACKUP_DIR_NAME).exists(), (
            "the refusal happened after the slot folder was created; the "
            "gate must fire before any filesystem work"
        )

    def test_write_lock_inside_the_roots_works(self, tmp_path, monkeypatch):
        inside = tmp_path / "inside"
        inside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        doc = _doc(inside)
        with safesave.write_lock(doc):
            pass
        assert (inside / safesave.BACKUP_DIR_NAME).exists()

    def test_write_lock_unrestricted_when_env_unset(self, tmp_path, monkeypatch):
        monkeypatch.delenv(sandbox.ENV_VAR, raising=False)
        doc = _doc(tmp_path)
        with safesave.write_lock(doc):
            pass


# ============================================ W-A2  staleness, leg by leg
# core/safesave.py:463  `or not _pid_alive(pid)` -> `and`
# core/safesave.py:465  `or age > LOCK_STALE_SECONDS` -> `and`
# core/safesave.py:466  `return True` -> `return False`
#
# The repo's existing stale-lock test is judged stale on a leg these mutants
# leave intact, so no individual leg had a pin of its own.


class TestStaleLockLegs:
    def test_dead_pid_with_a_fresh_timestamp_is_stale(self, dead_pid):
        """Kills line 463: a dead holder must be stale even when its
        lockfile was written a second ago."""
        info = {"pid": dead_pid, "token": "foreign", "time": time.time()}
        assert safesave._is_stale(info) is True

    def test_live_foreign_pid_with_an_ancient_timestamp_is_stale(
        self, sacrificial
    ):
        """Kills lines 465 and 466: age alone must break a lock whose holder
        is demonstrably still running."""
        _proc, pid = sacrificial
        info = {"pid": pid, "token": "foreign",
                "time": time.time() - (10 * 60) - 60}
        assert safesave._is_stale(info) is True

    def test_live_foreign_pid_with_a_fresh_timestamp_is_not_stale(
        self, sacrificial
    ):
        _proc, pid = sacrificial
        info = {"pid": pid, "token": "foreign", "time": time.time()}
        assert safesave._is_stale(info) is False

    def test_a_leaked_lockfile_from_a_dead_writer_is_broken(
        self, tmp_path, dead_pid
    ):
        """End to end: the permanent-lockout failure mode itself."""
        lock = _lockfile(tmp_path / safesave.LOCK_FILE_NAME, pid=dead_pid,
                         time=time.time())
        assert safesave._acquire_lockfile(lock, "leaked.docx") is True
        assert json.loads(lock.read_text())["token"] == safesave._OWNER_TOKEN

    def test_a_live_foreign_holder_still_refuses(self, tmp_path, sacrificial,
                                                 monkeypatch):
        _proc, pid = sacrificial
        lock = _lockfile(tmp_path / safesave.LOCK_FILE_NAME, pid=pid,
                         time=time.time())
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.3)
        with pytest.raises(MutationLockTimeout) as exc:
            safesave._acquire_lockfile(lock, "held.docx")
        assert str(pid) in str(exc.value)
        assert lock.exists(), "a live holder's lockfile was stolen"


# ================================== W-A3  _pid_alive must never kill anyone
# core/safesave.py:286  `if sys.platform == "win32":` negated
#
# Under the mutant a Windows liveness probe routes to os.kill(pid, 0), which
# on Windows TERMINATES the process being probed: a staleness check that
# kills the live lock holder.


class TestPidLivenessNeverKills:
    def test_probing_a_live_process_leaves_it_running(self, sacrificial):
        proc, pid = sacrificial
        assert safesave._pid_alive(pid) is True
        # poll() alone is not enough: process termination is asynchronous,
        # so a just-killed child can still report None. Waiting for it NOT to
        # exit is the assertion. (The module comment says os.kill(pid, 0)
        # terminates the target on Windows; that is CPython-version
        # dependent and does not reproduce on 3.14, so the dead-PID
        # misreading below is what actually pins the branch today.)
        with pytest.raises(subprocess.TimeoutExpired):
            proc.wait(timeout=1.5)

    def test_a_dead_pid_reads_as_dead(self, dead_pid):
        assert safesave._pid_alive(dead_pid) is False

    def test_nonpositive_pids_are_dead(self):
        assert safesave._pid_alive(0) is False
        assert safesave._pid_alive(-1) is False

    def test_xproc_probe_also_leaves_the_process_running(self, sacrificial):
        proc, pid = sacrificial
        assert xproc._pid_alive(pid) is True
        with pytest.raises(subprocess.TimeoutExpired):
            proc.wait(timeout=1.5)

    def test_xproc_reads_a_dead_pid_as_dead(self, dead_pid):
        """The other half of why the ctypes route exists on Windows.

        ``os.kill(pid, 0)`` cannot answer this question there: against a
        process that has exited but whose handle is still open it raises
        WinError 87, which the POSIX branch's ``except OSError`` reads as
        "alive". A live-session lock left by a dead server would then never
        be broken.
        """
        assert xproc._pid_alive(dead_pid) is False

    def test_xproc_nonpositive_pids_are_dead(self):
        assert xproc._pid_alive(0) is False
        assert xproc._pid_alive(-1) is False


# ============================== W-A4 / W-D2  the "assume alive" fallbacks
# core/safesave.py:304 and :308 (and com/xproc.py:181 and :185)
#   `return True  # assume alive` -> `return False`
#
# The comment states an invariant the mutated code no longer keeps: an
# unreadable process handle means "dead", so a live holder's lock is broken
# and two writers proceed.


@winonly
@pytest.mark.parametrize("module", [safesave, xproc], ids=["safesave", "xproc"])
class TestAssumeAliveFallbacks:
    def test_a_raising_probe_assumes_alive(self, module, monkeypatch):
        import ctypes

        ctypes.windll.kernel32  # force the loader to cache it
        monkeypatch.setattr(ctypes.windll, "kernel32",
                            _FakeKernel32(open_raises=True))
        assert module._pid_alive(os.getpid()) is True

    def test_an_unqueryable_handle_assumes_alive(self, module, monkeypatch):
        import ctypes

        fake = _FakeKernel32(open_returns=7, exitcode_ok=False)
        ctypes.windll.kernel32
        monkeypatch.setattr(ctypes.windll, "kernel32", fake)
        assert module._pid_alive(os.getpid()) is True
        assert fake.closed == 1, "the process handle leaked"

    def test_a_raising_probe_does_not_break_a_live_lock(self, module,
                                                        monkeypatch):
        import ctypes

        ctypes.windll.kernel32
        monkeypatch.setattr(ctypes.windll, "kernel32",
                            _FakeKernel32(open_raises=True))
        info = {"pid": os.getpid() + 1, "token": "foreign", "time": time.time()}
        assert module._is_stale(info) is False


# ================================ W-A5 / W-D4  PID-recycle detection is on
# core/safesave.py:325, :336, :350 (and com/xproc.py:121)
#
# With any of these the function returns None for every real PID, the
# `born != now_born` comparison is skipped, and staleness degrades to plain
# PID liveness -- the pre-fix behaviour. The existing recycle test feeds
# creation times in directly, so it never touches this function.


@pytest.mark.parametrize("module", [safesave, xproc], ids=["safesave", "xproc"])
class TestProcessCreateTime:
    @winonly
    def test_our_own_creation_time_is_readable(self, module):
        born = module._process_create_time(os.getpid())
        assert born is not None
        assert born > 0

    @winonly
    def test_two_reads_agree(self, module):
        assert (module._process_create_time(os.getpid())
                == module._process_create_time(os.getpid()))

    @winonly
    def test_a_live_child_has_a_readable_creation_time(self, module,
                                                       sacrificial):
        _proc, pid = sacrificial
        assert module._process_create_time(pid) is not None

    def test_nonpositive_pids_read_none(self, module):
        assert module._process_create_time(0) is None
        assert module._process_create_time(-1) is None

    @winonly
    def test_the_owner_creation_time_was_recorded_at_import(self, module):
        assert module._OWNER_CREATED is not None

    @winonly
    def test_a_recycled_pid_is_stale(self, module, sacrificial):
        """A live PID whose creation time does not match the payload is a
        recycled number, not the holder."""
        _proc, pid = sacrificial
        info = {"pid": pid, "token": "foreign", "time": time.time(),
                "pid_created": 1.0}
        assert module._is_stale(info) is True


# ============================== W-A6 / W-D3  the non-NTFS lockfile fallback
# core/safesave.py:433, :438 and com/xproc.py:254
#
# Every test runs on NTFS, where the hardlink path always wins, so the
# fallback that carries network shares and non-NTFS volumes had never been
# executed. The xproc mutant is the dangerous one: `except FileExistsError:
# return True` tells a process that LOST the race that it won.


class TestNonNtfsPublishFallback:
    @pytest.fixture()
    def no_hardlinks(self, monkeypatch):
        def boom(*args, **kwargs):
            raise OSError("simulated filesystem without hardlink support")

        monkeypatch.setattr(os, "link", boom)

    def test_safesave_publishes_a_complete_lockfile(self, tmp_path,
                                                    no_hardlinks):
        lock = tmp_path / safesave.LOCK_FILE_NAME
        assert safesave._publish_lockfile(lock) is True
        info = json.loads(lock.read_text(encoding="utf-8"))
        assert info["token"] == safesave._OWNER_TOKEN
        assert info["pid"] == os.getpid()

    def test_safesave_reports_an_existing_lock_rather_than_claiming_it(
        self, tmp_path, no_hardlinks
    ):
        lock = _lockfile(tmp_path / safesave.LOCK_FILE_NAME)
        before = lock.read_text(encoding="utf-8")
        assert safesave._publish_lockfile(lock) is False
        assert lock.read_text(encoding="utf-8") == before

    def test_safesave_full_cycle_without_hardlinks(self, tmp_path,
                                                   no_hardlinks):
        lock = tmp_path / safesave.LOCK_FILE_NAME
        assert safesave._acquire_lockfile(lock, "share.docx") is True
        safesave._release_lockfile(lock)
        assert not lock.exists()

    def test_safesave_leaves_no_temp_files_behind(self, tmp_path,
                                                  no_hardlinks):
        lock = tmp_path / safesave.LOCK_FILE_NAME
        safesave._publish_lockfile(lock)
        assert [p.name for p in tmp_path.glob(".lock-*.tmp")] == []

    def test_xproc_publishes_a_complete_lockfile(self, tmp_path, no_hardlinks):
        lock = tmp_path / "word-app.lock"
        assert xproc._publish_lockfile(lock, "unit-test") is True
        info = json.loads(lock.read_text(encoding="utf-8"))
        assert info["token"] == xproc._OWNER_TOKEN
        assert info["holder"] == "unit-test"

    def test_xproc_loser_of_the_race_is_told_it_lost(self, tmp_path,
                                                     no_hardlinks):
        """The mutant returns True here: a process that lost the race is
        told it won, and will delete the winner's lockfile on release."""
        lock = _lockfile(tmp_path / "word-app.lock", holder="the winner")
        assert xproc._publish_lockfile(lock, "the loser") is False
        assert json.loads(lock.read_text())["holder"] == "the winner"


# =============================== W-A7  the unreadable-lock timeout branch
# core/safesave.py:504  `elif time.monotonic() > deadline:` negated
#
# An unreadable lockfile inside its grace window must refuse once the wait
# is spent, not spin forever.


class TestUnreadableLockTimeout:
    def test_an_unreadable_lock_inside_its_grace_refuses_on_deadline(
        self, tmp_path, monkeypatch
    ):
        lock = tmp_path / safesave.LOCK_FILE_NAME
        lock.write_text("{torn", encoding="utf-8")
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.2)
        monkeypatch.setattr(safesave, "_UNREADABLE_GRACE_SECONDS", 300.0)
        with pytest.raises(MutationLockTimeout) as exc:
            safesave._acquire_lockfile(lock, "torn.docx")
        assert "cannot read" in str(exc.value)
        assert lock.exists(), "an in-grace unreadable lock was stolen"

    def test_the_refusal_arrives_only_after_the_wait_is_spent(
        self, tmp_path, monkeypatch
    ):
        """The deadline test is what makes this a WAIT rather than an instant
        refusal. Negated, the branch fires on the FIRST poll and the caller is
        told the lock is unreadable before anyone waited for it. Driven
        through a controlled monotonic clock, so no real time passes and the
        poll count is exact."""
        lock = tmp_path / safesave.LOCK_FILE_NAME
        lock.write_text("{torn", encoding="utf-8")
        clock = _PollClock()
        monkeypatch.setattr(safesave, "time", clock)
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 1.0)
        monkeypatch.setattr(safesave, "_UNREADABLE_GRACE_SECONDS", 300.0)

        with pytest.raises(MutationLockTimeout):
            safesave._acquire_lockfile(lock, "torn.docx")
        assert clock.sleeps >= 5, (
            f"refused after only {clock.sleeps} poll(s); the wait window was "
            "never spent"
        )

    def test_the_refusal_names_the_lockfile(self, tmp_path, monkeypatch):
        lock = tmp_path / safesave.LOCK_FILE_NAME
        lock.write_text("{torn", encoding="utf-8")
        monkeypatch.setattr(safesave, "LOCK_WAIT_SECONDS", 0.2)
        monkeypatch.setattr(safesave, "_UNREADABLE_GRACE_SECONDS", 300.0)
        with pytest.raises(MutationLockTimeout) as exc:
            safesave._acquire_lockfile(lock, "torn.docx")
        assert str(lock) in str(exc.value)


# =============================== W-A8  the antivirus-retry policy
# core/safesave.py:193 `if not _is_transient(exc):` negated
# core/safesave.py:184 `or` -> `and`
# core/safesave.py:96  the _TRANSIENT_WINERRORS membership itself
#
# Nothing in the suite forced a transient OSError through _with_retry, so
# the inversion (retry the fatal ones, raise the recoverable ones) passed.


class TestTransientRetryPolicy:
    def test_a_sharing_violation_is_retried_then_succeeds(self):
        calls = []

        def flaky(*args):
            calls.append(1)
            if len(calls) < 3:
                raise PermissionError(13, "sharing violation", None, 32)
            return "done"

        assert safesave._with_retry(flaky) == "done"
        assert len(calls) == 3

    def test_a_fatal_error_is_raised_on_the_first_attempt(self):
        calls = []

        def fatal(*args):
            calls.append(1)
            raise FileNotFoundError(2, "no such file")

        with pytest.raises(FileNotFoundError):
            safesave._with_retry(fatal)
        assert len(calls) == 1, (
            "a non-transient error was retried; the retry predicate is "
            "inverted"
        )

    def test_retries_are_bounded(self):
        calls = []

        def always(*args):
            calls.append(1)
            raise PermissionError(13, "sharing violation", None, 32)

        with pytest.raises(PermissionError):
            safesave._with_retry(always)
        assert len(calls) == len(safesave._RETRY_DELAYS) + 1

    @winonly
    @pytest.mark.parametrize("winerror", [5, 32, 33])
    def test_each_transient_winerror_is_recognized(self, winerror):
        """A plain OSError (not a PermissionError) carrying one of the
        sharing/access codes must still be transient: this is the leg the
        `or` -> `and` mutant removes."""
        exc = OSError(13, "busy", None, winerror)
        assert safesave._is_transient(exc) is True

    def test_a_permission_error_without_a_winerror_is_transient(self):
        assert safesave._is_transient(PermissionError(13, "denied")) is True

    def test_an_unrelated_oserror_is_not_transient(self):
        assert safesave._is_transient(OSError(2, "missing", None, 2)) is False

    def test_the_transient_code_set_is_the_measured_one(self):
        assert safesave._TRANSIENT_WINERRORS == {5, 32, 33}


# ================== W-A9  policy constants and the long-name boundary
# LOCK_STALE_SECONDS, ANCHOR_IDLE_SECONDS, _MAX_FOLDER_NAME, the `<=` at
# line 115 and the `- 9` truncation arithmetic at line 120.
#
# The existing long-name test uses a name far past the threshold, so the
# boundary itself was free. Constants are asserted as literals AND measured
# behaviourally, because a test that reads the constant moves with the
# mutant.


class TestMeasuredBoundaries:
    def test_lock_stale_seconds_is_ten_minutes(self):
        assert safesave.LOCK_STALE_SECONDS == 600

    def test_a_lock_exactly_at_the_stale_boundary_is_live(self, monkeypatch,
                                                          sacrificial):
        """`age > LOCK_STALE_SECONDS`: exactly 600s old is NOT stale. Only a
        frozen clock can land exactly on the boundary."""
        _proc, pid = sacrificial
        now = 1_700_000_000.0
        monkeypatch.setattr(safesave, "time", _FrozenClock(now))
        info = {"pid": pid, "token": "foreign", "time": now - 600.0}
        assert safesave._is_stale(info) is False

    def test_a_lock_just_past_the_stale_boundary_is_stale(self, monkeypatch,
                                                          sacrificial):
        _proc, pid = sacrificial
        now = 1_700_000_000.0
        monkeypatch.setattr(safesave, "time", _FrozenClock(now))
        info = {"pid": pid, "token": "foreign", "time": now - 600.5}
        assert safesave._is_stale(info) is True

    def test_anchor_idle_seconds_is_one_hour(self):
        assert safesave.ANCHOR_IDLE_SECONDS == 3600

    def test_the_anchor_rotates_exactly_at_the_idle_boundary(self, tmp_path,
                                                             monkeypatch):
        """`idle >= ANCHOR_IDLE_SECONDS`: exactly one hour idle DOES rotate.
        The idle gap is produced by backdating prev's mtime, never by
        sleeping."""
        doc = _doc(tmp_path, text="first")
        safesave.rotate_slots(doc)  # creates anchor + prev
        d = safesave.slot_dir(doc)
        anchor_before = (d / safesave.ANCHOR_SLOT).read_bytes()

        _mutate_in_place(doc, "second")
        now = 1_700_000_000.0
        prev = d / safesave.PREV_SLOT
        os.utime(prev, (now - 3600.0, now - 3600.0))
        monkeypatch.setattr(safesave, "time", _FrozenClock(now))

        rotated = safesave.rotate_slots(doc)
        assert rotated["anchor"] is True
        assert (d / safesave.ANCHOR_SLOT).read_bytes() != anchor_before

    def test_the_anchor_holds_still_just_inside_the_idle_boundary(
        self, tmp_path, monkeypatch
    ):
        doc = _doc(tmp_path, text="first")
        safesave.rotate_slots(doc)
        d = safesave.slot_dir(doc)
        anchor_before = (d / safesave.ANCHOR_SLOT).read_bytes()

        _mutate_in_place(doc, "second")
        now = 1_700_000_000.0
        prev = d / safesave.PREV_SLOT
        os.utime(prev, (now - 3599.0, now - 3599.0))
        monkeypatch.setattr(safesave, "time", _FrozenClock(now))

        rotated = safesave.rotate_slots(doc)
        assert rotated["anchor"] is False
        assert (d / safesave.ANCHOR_SLOT).read_bytes() == anchor_before

    def test_the_rotated_labels_are_not_swapped(self, tmp_path):
        doc = _doc(tmp_path)
        first = safesave.rotate_slots(doc)
        assert first == {"prev": True, "anchor": True}
        second = safesave.rotate_slots(doc)
        assert second == {"prev": True, "anchor": False}

    def test_the_folder_name_threshold_is_eighty(self):
        assert safesave._MAX_FOLDER_NAME == 80

    def test_a_name_exactly_at_the_threshold_keeps_its_own_folder(self,
                                                                  tmp_path):
        stem = "n" * (80 - len(".docx"))
        name = stem + ".docx"
        assert len(name) == 80
        doc = _doc(tmp_path, name)
        d = safesave.slot_dir(doc, create=True)
        assert d.name == name
        assert not (d / safesave._SOURCE_NAME_FILE).exists()
        assert safesave.source_doc_for(d) == doc

    def test_a_name_one_past_the_threshold_is_truncated_and_hashed(self,
                                                                   tmp_path):
        stem = "n" * (81 - len(".docx"))
        name = stem + ".docx"
        assert len(name) == 81
        doc = _doc(tmp_path, name)
        d = safesave.slot_dir(doc, create=True)
        assert d.name != name
        assert len(d.name) == 80, (
            "truncated folder names must land exactly on the budget; the "
            "`- 9` arithmetic reserves the dash plus 8 hex digits"
        )
        assert d.name[71] == "-"
        assert (d / safesave._SOURCE_NAME_FILE).read_text(encoding="utf-8") == name
        assert safesave.source_doc_for(d) == doc

    def test_long_unicode_names_round_trip(self, tmp_path):
        # 81 chars: one past _MAX_FOLDER_NAME, so the truncate+hash path
        # runs on a multibyte name -- while the FILENAME stays 248 bytes
        # in UTF-8, under the 255-byte per-name limit of Linux filesystems
        # (90 chars of Hangul = 270 bytes, which ext4 refuses to create).
        name = "한국어" * 27 + ".docx"
        doc = _doc(tmp_path, name)
        d = safesave.slot_dir(doc, create=True)
        assert safesave.source_doc_for(d) == doc


# ================================ W-B1  restore validates what it writes
# ops/backups.py:175  DocxPackage._validate_payload(payload) -> pass
#
# Restore is the one path where a corrupt payload lands on a document the
# user still wants.


class TestRestoreValidatesTheBackup:
    def test_a_garbage_prev_slot_refuses(self, tmp_path):
        doc = _doc(tmp_path, text="the good content")
        before = doc.read_bytes()
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(b"not a docx at all")

        with pytest.raises(ValidationFailed):
            bk.restore_backup(str(doc), "prev")
        assert doc.read_bytes() == before, (
            "a corrupt backup was written over the live document"
        )

    def test_a_truncated_prev_slot_refuses(self, tmp_path):
        doc = _doc(tmp_path)
        before = doc.read_bytes()
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(before[: len(before) // 2])

        with pytest.raises(ValidationFailed):
            bk.restore_backup(str(doc), "prev")
        assert doc.read_bytes() == before

    def test_a_zip_that_is_not_a_docx_refuses(self, tmp_path):
        doc = _doc(tmp_path)
        before = doc.read_bytes()
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("hello.txt", "not word content")
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(buf.getvalue())

        with pytest.raises(ValidationFailed):
            bk.restore_backup(str(doc), "prev")
        assert doc.read_bytes() == before

    def test_a_legacy_source_is_validated_too(self, tmp_path):
        doc = _doc(tmp_path)
        before = doc.read_bytes()
        legacy = tmp_path / "doc.bak-20200101.docx"
        legacy.write_bytes(b"garbage")
        with pytest.raises(ValidationFailed):
            bk.restore_backup(str(doc), str(legacy))
        assert doc.read_bytes() == before

    def test_a_good_backup_still_restores(self, tmp_path):
        doc = _doc(tmp_path, text="original")
        original = doc.read_bytes()
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(original)
        doc.write_bytes(_doc(tmp_path, "other.docx", text="changed").read_bytes())

        out = bk.restore_backup(str(doc), "prev")
        assert doc.read_bytes() == original
        assert out["prev_rotated"] is True
        assert "undo" in out


# ============ W-B2  the sandbox gate on the file_path branch of list/purge
# ops/backups.py:92 and :217  check_path(file_path, ...) -> pass
#
# The existing sandbox tests exercise the `directory=` argument, gated by a
# different call. The `file_path=` spelling was ungated as far as the suite
# could tell -- and purge with dry_run=False deletes.


class TestBackupSandboxOnFilePath:
    @pytest.fixture()
    def rooted(self, tmp_path, monkeypatch):
        inside = tmp_path / "inside"
        inside.mkdir()
        outside = tmp_path / "outside"
        outside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        return inside, outside

    def test_list_backups_file_path_outside_blocked(self, rooted):
        _inside, outside = rooted
        with pytest.raises(SandboxViolation):
            bk.list_backups(file_path=str(outside / "doc.docx"))

    def test_purge_slots_file_path_outside_blocked(self, rooted):
        _inside, outside = rooted
        with pytest.raises(SandboxViolation):
            bk.purge_backups("slots", file_path=str(outside / "doc.docx"),
                             dry_run=False)

    def test_purge_legacy_file_path_outside_blocked(self, rooted):
        _inside, outside = rooted
        with pytest.raises(SandboxViolation):
            bk.purge_backups("legacy", file_path=str(outside / "doc.docx"),
                             dry_run=False)

    def test_purge_outside_deletes_nothing(self, rooted):
        _inside, outside = rooted
        doc = _doc(outside)
        safesave.rotate_slots(doc)
        slot = safesave.slot_dir(doc) / safesave.PREV_SLOT
        assert slot.exists()
        with pytest.raises(SandboxViolation):
            bk.purge_backups("slots", file_path=str(doc), dry_run=False)
        assert slot.exists(), "a purge outside the roots deleted files"

    def test_list_backups_file_path_inside_allowed(self, rooted):
        inside, _outside = rooted
        doc = _doc(inside)
        out = bk.list_backups(file_path=str(doc))
        assert out["document_exists"] is True


# ============================ W-B3  purge scope-argument validation
# ops/backups.py:222 and :243  raise WordMcpError(...) -> pass
#
# Without the refusals these fall through to a None-path crash.


class TestPurgeArgumentValidation:
    def test_slots_without_file_path_refuses(self, tmp_path):
        with pytest.raises(WordMcpError) as exc:
            bk.purge_backups("slots", directory=str(tmp_path), dry_run=True)
        assert "file_path" in str(exc.value)

    def test_orphans_without_either_argument_refuses(self):
        with pytest.raises(WordMcpError) as exc:
            bk.purge_backups("orphans", dry_run=True)
        assert "orphans" in str(exc.value)

    def test_legacy_without_either_argument_refuses(self):
        with pytest.raises(WordMcpError) as exc:
            bk.purge_backups("legacy", dry_run=True)
        assert "needs file_path or directory" in str(exc.value), (
            "falling through to the unknown-scope refusal reports a real "
            "scope as unknown and never names the missing argument"
        )

    def test_an_unknown_scope_refuses(self, tmp_path):
        with pytest.raises(WordMcpError) as exc:
            bk.purge_backups("everything", directory=str(tmp_path))
        assert "unknown purge scope" in str(exc.value)

    def test_orphans_accepts_file_path_as_the_directory_hint(self, tmp_path):
        doc = _doc(tmp_path)
        out = bk.purge_backups("orphans", file_path=str(doc), dry_run=True)
        assert out["scope"] == "orphans"


# ============================ W-B4  list_backups reports what is really there
# ops/backups.py:119, :127, :202
#
# The tool's output SHAPE had assertions; its per-entry truth did not.


class TestListBackupsTruth:
    def test_only_the_slots_that_exist_are_listed(self, tmp_path):
        doc = _doc(tmp_path)
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(doc.read_bytes())

        out = bk.list_backups(file_path=str(doc))
        assert [s["slot"] for s in out["slots"]] == ["prev"]

    def test_both_slots_are_listed_when_both_exist(self, tmp_path):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        out = bk.list_backups(file_path=str(doc))
        assert sorted(s["slot"] for s in out["slots"]) == ["anchor", "prev"]

    def test_a_document_with_no_slots_lists_none(self, tmp_path):
        doc = _doc(tmp_path)
        out = bk.list_backups(file_path=str(doc))
        assert out["slots"] == []

    def test_the_directory_branch_finds_the_document(self, tmp_path):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        out = bk.list_backups(directory=str(tmp_path))
        assert [Path(e["document"]).name for e in out["documents"]] == ["doc.docx"]
        assert sorted(s["slot"] for s in out["documents"][0]["slots"]) == [
            "anchor", "prev"
        ]

    def test_a_directory_with_no_backup_root_reports_nothing(self, tmp_path):
        out = bk.list_backups(directory=str(tmp_path))
        assert out["documents"] == []
        assert out["orphaned_folders"] == []

    def test_an_orphan_folder_is_reported_and_not_listed_as_a_document(
        self, tmp_path
    ):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        doc.unlink()
        out = bk.list_backups(directory=str(tmp_path))
        assert out["documents"] == []
        assert len(out["orphaned_folders"]) == 1
        assert Path(out["orphaned_folders"][0]["missing_document"]).name == "doc.docx"

    def test_restoring_onto_a_missing_document_says_nothing_rotated(self,
                                                                    tmp_path):
        doc = _doc(tmp_path)
        d = safesave.slot_dir(doc, create=True)
        (d / safesave.PREV_SLOT).write_bytes(doc.read_bytes())
        doc.unlink()

        out = bk.restore_backup(str(doc), "prev")
        assert out["prev_rotated"] is False
        assert "note" in out and "undo" not in out

    def test_restoring_a_missing_slot_refuses(self, tmp_path):
        doc = _doc(tmp_path)
        with pytest.raises(DocumentNotFound):
            bk.restore_backup(str(doc), "anchor")


# ============================ W-B5  the post-purge folder cleanup
# ops/backups.py:284  `and` -> `or`


class TestPurgeCleanup:
    def test_purging_slots_drops_the_empty_folder(self, tmp_path):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        d = safesave.slot_dir(doc)
        bk.purge_backups("slots", file_path=str(doc), dry_run=False)
        assert not d.exists()

    def test_purging_legacy_leaves_the_slot_folder_alone(self, tmp_path):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        legacy = tmp_path / "doc.bak-20200101.docx"
        legacy.write_bytes(b"x")
        d = safesave.slot_dir(doc)
        bk.purge_backups("legacy", file_path=str(doc), dry_run=False)
        assert d.exists() and (d / safesave.PREV_SLOT).exists()
        assert not legacy.exists()

    def test_a_dry_run_deletes_nothing(self, tmp_path):
        doc = _doc(tmp_path)
        safesave.rotate_slots(doc)
        d = safesave.slot_dir(doc)
        out = bk.purge_backups("slots", file_path=str(doc), dry_run=True)
        assert out["count"] == 2
        assert (d / safesave.PREV_SLOT).exists()


# =============================== W-C1  the save gate's part vocabulary
# core/package.py:230  raise ValidationFailed("output lost [Content_Types].xml")
#   -> pass
#
# Half the validator's part-presence vocabulary had a test and the other
# half did not. A package missing [Content_Types].xml is exactly the file
# Word offers to repair.


class TestValidatePayloadVocabulary:
    @staticmethod
    def _zip(parts: dict[str, bytes]) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            for name, data in parts.items():
                zf.writestr(name, data)
        return buf.getvalue()

    def test_a_payload_without_content_types_refuses(self):
        payload = self._zip({"word/document.xml": b"<document/>"})
        with pytest.raises(ValidationFailed) as exc:
            DocxPackage._validate_payload(payload)
        assert "[Content_Types].xml" in str(exc.value)

    def test_a_payload_without_document_xml_refuses(self):
        payload = self._zip({"[Content_Types].xml": b"<Types/>"})
        with pytest.raises(ValidationFailed) as exc:
            DocxPackage._validate_payload(payload)
        assert "word/document.xml" in str(exc.value)

    def test_a_payload_with_both_parts_passes(self):
        payload = self._zip({
            "[Content_Types].xml": b"<Types/>",
            "word/document.xml": b"<document/>",
        })
        DocxPackage._validate_payload(payload)

    def test_malformed_xml_in_any_part_refuses(self):
        payload = self._zip({
            "[Content_Types].xml": b"<Types/>",
            "word/document.xml": b"<document/>",
            "word/_rels/document.xml.rels": b"<Relationships><broken>",
        })
        with pytest.raises(ValidationFailed) as exc:
            DocxPackage._validate_payload(payload)
        assert "well-formed" in str(exc.value)

    def test_a_non_zip_payload_refuses(self):
        with pytest.raises(ValidationFailed):
            DocxPackage._validate_payload(b"this is not a zip file")


# =============================== W-C2  the post-save baseline refresh
# core/package.py:213  `if in_place:` negated
#
# After an in-place save the _dirty set must clear and _raw must hold the
# written bytes; a save-as must NOT rebase.


class TestBaselineRefreshAfterSave:
    @staticmethod
    def _edit(pkg: DocxPackage, text: str) -> None:
        from word_mcp.core.package import qn

        body = pkg.body()
        para = body.findall(qn("w:p"))[0]
        for t in para.iter(qn("w:t")):
            t.text = text
        pkg.mark_dirty()

    def test_an_in_place_save_rebases_the_in_memory_bytes(self, tmp_path):
        doc = _doc(tmp_path, text="before")
        pkg = DocxPackage(doc)
        self._edit(pkg, "after")
        pkg.save()

        on_disk = zipfile.ZipFile(doc).read("word/document.xml")
        assert pkg.raw_part("word/document.xml") == on_disk, (
            "the in-memory package keeps serving pre-save bytes"
        )
        assert b"after" in pkg.raw_part("word/document.xml")
        assert pkg._dirty == set()

    def test_a_save_as_does_not_rebase(self, tmp_path):
        doc = _doc(tmp_path, text="before")
        pkg = DocxPackage(doc)
        original = pkg.raw_part("word/document.xml")
        self._edit(pkg, "after")
        pkg.save(dest=tmp_path / "copy.docx")

        assert pkg.raw_part("word/document.xml") == original
        assert pkg._dirty == {"word/document.xml"}

    def test_a_second_in_place_save_after_a_reread_sees_the_first(self,
                                                                  tmp_path):
        doc = _doc(tmp_path, text="before")
        pkg = DocxPackage(doc)
        self._edit(pkg, "after")
        pkg.save()
        assert b"after" in DocxPackage(doc).raw_part("word/document.xml")


# =============================== W-C3  the mark_dirty ordering invariant
# core/package.py:156  raise RuntimeError(...) -> pass


class TestMarkDirtyOrdering:
    def test_mark_dirty_before_tree_refuses(self, tmp_path):
        pkg = DocxPackage(_doc(tmp_path))
        with pytest.raises(RuntimeError):
            pkg.mark_dirty("word/settings.xml")

    def test_mark_dirty_after_tree_is_fine(self, tmp_path):
        pkg = DocxPackage(_doc(tmp_path))
        pkg.tree()
        pkg.mark_dirty()
        assert "word/document.xml" in pkg._dirty


# ================= W-D1 / W-D4 / W-D6  the live-session lock's staleness
# com/xproc.py:281 (age -> `and`), :293 (recycled-PID branch), :276 (the
# missing-pid sentinel)


class TestXprocStaleness:
    def test_an_abandoned_lock_ages_out(self, sacrificial):
        _proc, pid = sacrificial
        info = {"pid": pid, "token": "foreign",
                "time": time.time() - (10 * 60) - 60}
        assert xproc._is_stale(info) is True

    def test_a_working_holder_is_not_aged_out(self, sacrificial):
        _proc, pid = sacrificial
        info = {"pid": pid, "token": "foreign", "time": time.time() - 5}
        assert xproc._is_stale(info) is False

    def test_the_stale_window_is_ten_minutes(self):
        assert xproc.LOCK_STALE_SECONDS == 600

    def test_the_exact_stale_boundary_is_still_live(self, monkeypatch,
                                                    sacrificial):
        _proc, pid = sacrificial
        now = 1_700_000_000.0
        monkeypatch.setattr(xproc, "time", _FrozenClock(now))
        info = {"pid": pid, "token": "foreign", "time": now - 600.0}
        assert xproc._is_stale(info) is False

    def test_a_lock_without_a_pid_is_stale(self):
        assert xproc._is_stale({"token": "foreign", "time": time.time()}) is True

    @pytest.mark.parametrize("module", [safesave, xproc],
                             ids=["safesave", "xproc"])
    def test_the_missing_pid_sentinel_can_never_name_a_live_process(
        self, module, monkeypatch
    ):
        """The sentinel for a lockfile with no pid is NEGATIVE on purpose.

        Staleness for such a lock rides entirely on the liveness probe
        saying no, and the only reason it always says no is that the
        sentinel is not a number any process can have. A positive sentinel
        would make a pid-less lockfile look held by PID 1, which exists on
        every POSIX host. The probe is modelled here rather than trusted,
        because on Windows PID 1 happens not to exist and the bug would
        hide.
        """
        monkeypatch.setattr(module, "_pid_alive", lambda pid: pid > 0)
        info = {"token": "foreign", "time": time.time()}
        assert module._is_stale(info) is True

    def test_holder_info_ignores_a_stale_lock(self, tmp_path, monkeypatch,
                                              dead_pid):
        monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(tmp_path))
        _lockfile(tmp_path / "word-app.lock", pid=dead_pid, holder="gone")
        assert xproc.holder_info() is None

    def test_holder_info_reports_a_live_holder(self, tmp_path, monkeypatch,
                                               sacrificial):
        _proc, pid = sacrificial
        monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(tmp_path))
        _lockfile(tmp_path / "word-app.lock", pid=pid, holder="live_edit")
        info = xproc.holder_info()
        assert info is not None
        assert info["pid"] == pid
        assert info["holder"] == "live_edit"
        assert info["ours"] is False


# =============================== W-D5  the lock-directory fallback
# com/xproc.py:110  `os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()`
#   -> `and`
#
# No test ran with LOCALAPPDATA cleared, so the documented degradation path
# that lock_state() reports on had no coverage at all.


class TestLockDirFallback:
    def test_the_override_wins(self, tmp_path, monkeypatch):
        monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(tmp_path / "chosen"))
        assert xproc._lock_dir() == tmp_path / "chosen"

    def test_localappdata_is_used_when_set(self, tmp_path, monkeypatch):
        monkeypatch.delenv(xproc._LOCK_DIR_ENV, raising=False)
        monkeypatch.setenv("LOCALAPPDATA", str(tmp_path / "appdata"))
        assert xproc._lock_dir() == tmp_path / "appdata" / "word-mcp" / "live-locks"

    def test_the_temp_dir_carries_it_when_localappdata_is_gone(self, tmp_path,
                                                              monkeypatch):
        monkeypatch.delenv(xproc._LOCK_DIR_ENV, raising=False)
        monkeypatch.delenv("LOCALAPPDATA", raising=False)
        monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
        assert xproc._lock_dir() == tmp_path / "word-mcp" / "live-locks"

    def test_cross_process_coverage_survives_the_fallback(self, tmp_path,
                                                          monkeypatch):
        monkeypatch.delenv(xproc._LOCK_DIR_ENV, raising=False)
        monkeypatch.delenv("LOCALAPPDATA", raising=False)
        monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
        state = xproc.lock_state()
        assert state["cross_process"] is True
        assert str(tmp_path) in state["lock_dir"]

    def test_an_empty_localappdata_still_falls_back(self, tmp_path,
                                                    monkeypatch):
        monkeypatch.delenv(xproc._LOCK_DIR_ENV, raising=False)
        monkeypatch.setenv("LOCALAPPDATA", "")
        monkeypatch.setattr(tempfile, "gettempdir", lambda: str(tmp_path))
        assert xproc._lock_dir() == tmp_path / "word-mcp" / "live-locks"


# ======================= sandbox canonicalization internals (4 findings)
# core/sandbox.py:100, :109, :112, :125, :83
#
# The POLICY is well tested (junction, symlink, prefix collision, UNC, case
# all die). What survived is the canonicalizer's plumbing for paths that do
# not exist yet.


class TestCanonicalizationPlumbing:
    def test_a_target_several_levels_below_a_missing_directory_is_contained(
        self, tmp_path, monkeypatch
    ):
        inside = tmp_path / "inside"
        inside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        target = inside / "a" / "b" / "c" / "new.docx"
        out = check_path(target, "save document")
        assert Path(out) == Path(os.path.realpath(inside)) / "a" / "b" / "c" / "new.docx"

    def test_a_deep_missing_target_outside_is_still_blocked(self, tmp_path,
                                                            monkeypatch):
        inside = tmp_path / "inside"
        inside.mkdir()
        monkeypatch.setenv(sandbox.ENV_VAR, str(inside))
        with pytest.raises(SandboxViolation):
            check_path(tmp_path / "outside" / "a" / "b" / "new.docx", "save")

    def test_canonicalizing_an_existing_path_resolves_it(self, tmp_path):
        f = _doc(tmp_path)
        assert sandbox._canonicalize(str(f)) == os.path.realpath(str(f))

    @winonly
    def test_a_path_on_an_unmounted_drive_keeps_its_own_spelling(self):
        """The walk-up ends on a drive root that does not exist, and the
        fallback has to keep THAT as the head. Substituting the full target
        instead re-appends the tail onto itself, so Q:\\a\\b canonicalizes to
        Q:\\a\\b\\a\\b and every containment answer after it is computed
        about a path nobody named."""
        letter = _unused_drive()
        if letter is None:
            pytest.skip("every drive letter probed is mounted")
        target = letter + ":\\nowhere\\deep\\file.docx"
        assert sandbox._canonicalize(target).lower() == target.lower()

    @winonly
    def test_a_bare_drive_root_contains_paths_on_that_drive(self):
        assert sandbox._contained("C:", "c:\\") is True
        assert sandbox._contained("C:", "c:\\users\\x\\doc.docx") is True

    def test_a_bare_drive_root_excludes_other_drives(self):
        assert sandbox._contained("C:", "d:\\users\\x\\doc.docx") is False

    def test_the_forward_slash_extended_unc_spelling_is_recognized(self):
        assert sandbox._strip_extended_prefix("//?/UNC/srv/share/f.docx") == (
            "\\\\srv/share/f.docx"
        )
        assert sandbox._is_unc(
            sandbox._strip_extended_prefix("//?/UNC/srv/share/f.docx")
        ) is True

    def test_the_forward_slash_extended_unc_spelling_is_refused(self, tmp_path,
                                                                monkeypatch):
        monkeypatch.setenv(sandbox.ENV_VAR, str(tmp_path))
        with pytest.raises(SandboxViolation) as exc:
            check_path("//?/UNC/srv/share/secret.docx", "open document")
        assert "UNC" in str(exc.value)

    @winonly
    def test_the_backslash_extended_local_prefix_still_normalizes(self,
                                                                  tmp_path,
                                                                  monkeypatch):
        monkeypatch.setenv(sandbox.ENV_VAR, str(tmp_path))
        f = _doc(tmp_path)
        assert check_path("\\\\?\\" + str(f), "open document")


# ============================ locate.py selector validation (5 findings)
# core/locate.py:236, :363, :691, :718, :388, :440


class TestSelectorValidation:
    @pytest.fixture()
    def pkg(self, tmp_path):
        d = Document()
        d.add_heading("Chapter 3", level=1)
        d.add_paragraph("Body text under chapter three.")
        f = tmp_path / "doc.docx"
        d.save(str(f))
        return DocxPackage(f)

    def test_a_wrong_typed_required_key_refuses(self, pkg):
        with pytest.raises(WordMcpError) as exc:
            locate.resolve_location(pkg, {"after_heading": {"text": 3}})
        assert "must be a str" in str(exc.value)

    def test_a_boolean_is_not_accepted_as_a_string(self, pkg):
        with pytest.raises(WordMcpError):
            locate.resolve_location(pkg, {"after_heading": {"text": True}})

    def test_a_wrong_typed_optional_key_refuses(self, pkg):
        with pytest.raises(WordMcpError) as exc:
            locate.resolve_location(
                pkg, {"after_heading": {"text": "Chapter 3", "occurrence": "1"}}
            )
        assert "must be a int" in str(exc.value)

    def test_an_empty_heading_text_refuses(self, pkg):
        with pytest.raises(WordMcpError) as exc:
            locate.resolve_location(pkg, {"after_heading": {"text": "   "}})
        assert "non-empty" in str(exc.value)

    def test_a_negative_cursor_index_refuses(self, pkg):
        with pytest.raises(WordMcpError) as exc:
            locate.resolve_location(pkg, {"cursor": True},
                                    cursor_reader=lambda: -1)
        assert "0-based" in str(exc.value)

    def test_a_boolean_cursor_index_refuses(self, pkg):
        with pytest.raises(WordMcpError):
            locate.resolve_location(pkg, {"cursor": True},
                                    cursor_reader=lambda: True)

    def test_a_valid_cursor_index_resolves(self, pkg):
        out = locate.resolve_location(pkg, {"cursor": True},
                                      cursor_reader=lambda: 0)
        assert out.paragraph_index == 0


class TestRangeSpecRecognition:
    def test_a_start_only_object_is_a_range(self):
        assert locate.is_range_spec({"start": {"paragraph": 1}}) is True

    def test_an_end_only_object_is_a_range(self):
        assert locate.is_range_spec({"end": {"paragraph": 4}}) is True

    def test_both_endpoints_are_a_range(self):
        assert locate.is_range_spec(
            {"start": {"paragraph": 1}, "end": {"paragraph": 4}}
        ) is True

    def test_a_plain_selector_is_not_a_range(self):
        assert locate.is_range_spec({"paragraph": 1}) is False

    def test_an_empty_object_is_not_a_range(self):
        assert locate.is_range_spec({}) is False

    def test_a_non_dict_is_not_a_range(self):
        assert locate.is_range_spec("paragraph 1") is False


# ---------------------------------------------------------------------------


def _headings_doc(path: Path, count: int) -> Path:
    d = Document()
    for i in range(count):
        d.add_heading(f"Heading {i:03d}", level=1)
        d.add_paragraph(f"Body {i}")
    d.save(str(path))
    return path


class TestProseRecurrenceScan:
    """core/locate.py:388  `if pos >= 0:` -> `if pos > 0:`

    Heading text recurs in body prose on real documents, and the resolver
    refuses to resolve first-match past a recurrence. The boundary shift
    makes a recurrence at CHARACTER ZERO invisible, which is exactly where
    a paragraph that opens by naming the heading puts it.
    """

    def _pkg(self, tmp_path, prose: str):
        d = Document()
        d.add_heading("Chapter 3", level=1)
        d.add_paragraph("Ordinary body text.")
        d.add_paragraph(prose)
        f = tmp_path / "doc.docx"
        d.save(str(f))
        return DocxPackage(f)

    def test_a_recurrence_at_character_zero_is_seen(self, tmp_path):
        pkg = self._pkg(tmp_path, "Chapter 3 opened the argument.")
        with pytest.raises(AmbiguousTarget) as exc:
            locate.resolve_location(pkg, {"after_heading": {"text": "Chapter 3"}})
        assert any(m["paragraph"] == 2 for m in exc.value.matches)

    def test_a_recurrence_mid_paragraph_is_seen(self, tmp_path):
        pkg = self._pkg(tmp_path, "As Chapter 3 argued, the case holds.")
        with pytest.raises(AmbiguousTarget):
            locate.resolve_location(pkg, {"after_heading": {"text": "Chapter 3"}})

    def test_no_recurrence_resolves_cleanly(self, tmp_path):
        pkg = self._pkg(tmp_path, "Nothing here names the heading.")
        out = locate.resolve_location(
            pkg, {"after_heading": {"text": "Chapter 3"}}
        )
        assert out.selector == "after_heading"


class TestListedCandidateCap:
    """core/locate.py:440  `if len(entries) > _MAX_LISTED:` -> `>=`

    The cap on how many outline entries a refusal echoes. At exactly the
    cap the message must list them all and say nothing about "more".
    """

    def _doc_with(self, tmp_path, count: int):
        d = Document()
        for i in range(count):
            d.add_heading(f"Heading {i:03d}", level=1)
            d.add_paragraph(f"Body {i}")
        f = tmp_path / f"h{count}.docx"
        d.save(str(f))
        return DocxPackage(f)

    def test_the_cap_is_twenty_five(self):
        assert locate._MAX_LISTED == 25

    def test_exactly_the_cap_lists_everything(self, tmp_path):
        pkg = self._doc_with(tmp_path, 25)
        with pytest.raises(Exception) as exc:
            locate.resolve_location(pkg, {"outline": "99"})
        assert "more)" not in str(exc.value)

    def test_one_past_the_cap_says_how_many_more(self, tmp_path):
        pkg = self._doc_with(tmp_path, 26)
        with pytest.raises(Exception) as exc:
            locate.resolve_location(pkg, {"outline": "99"})
        assert "(and 1 more)" in str(exc.value)


# ================================ envelope.py — the closed code vocabulary
#
# This is the one finding in the Word section that is NOT a mutant: the
# report noted that CLOSED_CODES is referenced only by the tests.
# ``refusal()`` took ``getattr(exc, "code", None)`` verbatim, so the closed
# refusal vocabulary was a convention enforced by three test assertions
# rather than by the code, and any exception carrying a ``.code`` attribute
# reached the wire with it. The clamp (_declared_code) closes it; these
# tests pin both the clamp and the declared codes that must survive it.


class _Declared(WordMcpError):
    pass


class TestClosedCodeVocabulary:
    def test_a_made_up_code_does_not_reach_the_wire(self):
        exc = _Declared("something went sideways")
        exc.code = "MADE_UP_CODE"
        out = envelope.refusal(exc)
        assert out["error"]["code"] in envelope.CLOSED_CODES
        assert out["error"]["code"] != "MADE_UP_CODE"

    def test_a_declared_valid_code_is_honoured(self):
        exc = _Declared("stale")
        exc.code = "STALE_ANCHOR"
        assert envelope.refusal(exc)["error"]["code"] == "STALE_ANCHOR"

    def test_a_non_string_code_falls_back_to_classification(self):
        exc = _Declared("weird")
        exc.code = 42
        assert envelope.refusal(exc)["error"]["code"] == "BAD_PARAMS"

    def test_a_third_party_exception_carrying_a_code_is_clamped(self):
        class Vendor(Exception):
            code = "vendor.timeout.7"

        out = envelope.refusal(Vendor("upstream said no"))
        assert out["error"]["code"] in envelope.CLOSED_CODES

    def test_a_lowercase_spelling_of_a_real_code_is_clamped(self):
        exc = _Declared("nearly right")
        exc.code = "not_found"
        assert envelope.refusal(exc)["error"]["code"] in envelope.CLOSED_CODES

    def test_the_hint_matches_the_clamped_code(self):
        exc = _Declared("nope")
        exc.code = "MADE_UP_CODE"
        out = envelope.refusal(exc)
        assert out["error"]["hint"] == envelope.HINTS.get(
            out["error"]["code"], ""
        )

    def test_every_declared_code_in_the_source_is_closed(self):
        """The five .code assignments in src/ (ops/batch.py, packs.py) must
        all survive the clamp; a typo at a sixth is what this catches."""
        for code in ("STALE_ANCHOR", "NOT_FOUND", "UNSUPPORTED_CONTENT",
                     "CONFLICT"):
            exc = _Declared("x")
            exc.code = code
            assert envelope.refusal(exc)["error"]["code"] == code


class TestRefusalPayloadShape:
    """envelope.py:178  `isinstance(exc, LookupError) and len(message) < 40`
    -> `or`, plus the matches/detail legs."""

    def test_a_bare_key_error_gets_a_real_message(self):
        out = envelope.refusal(KeyError(0))
        assert "internal lookup failed" in out["error"]["message"]

    def test_a_long_lookup_message_is_left_alone(self):
        long = "x" * 60
        out = envelope.refusal(KeyError(long))
        assert "internal lookup failed" not in out["error"]["message"]

    def test_a_short_non_lookup_message_is_left_alone(self):
        out = envelope.refusal(WordMcpError("nope"))
        assert out["error"]["message"] == "nope"
        assert "internal lookup failed" not in out["error"]["message"]

    def test_ambiguity_matches_ride_out_on_the_payload(self):
        exc = AmbiguousTarget("two matched")
        exc.matches = [{"paragraph": 1}, {"paragraph": 7}]
        out = envelope.refusal(exc)
        assert out["error"]["matches"] == [{"paragraph": 1}, {"paragraph": 7}]

    def test_a_plain_refusal_carries_no_matches_key(self):
        out = envelope.refusal(WordMcpError("plain"))
        assert "matches" not in out["error"]

    def test_the_refusal_result_serializes_as_an_error(self):
        result = envelope.refuse(WordMcpError("plain"))
        assert result.is_error is True
        assert result["ok"] is False
