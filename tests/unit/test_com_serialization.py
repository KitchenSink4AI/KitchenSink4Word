"""v2.0.0 live-mode fix batch (2026-09-03 live COM stress report).

Non-live regression coverage for the five MUST-FIX items:
1. COM serialization: lock mechanics, no interleaving under threads, and
   an entry-point audit proving every COM path takes the lock.
2. Tracked-replace loop: an emulated-COM Word 2016 harness reproduces the
   report's failure mechanism (tracked assignment inserts BEFORE the
   deletion markup and the deleted copy stays findable) and proves the
   three-layer fix: revision skip fails closed, resume advances past own
   markup, and tracked mode caps at the pre-edit match count.
3. apply_edits atomicity: preflight conflict simulation and rollback
   note contract.
4. Dialog prevention: alerts-suppression context manager contract.
5. Bounded timeouts: fast path, error propagation, queue-vs-stuck
   distinction, and parameter validation.
6. Name-collision guard and OS-layer dialog detection (synthetic #32770).

The live halves (real Word) are in test_live_concurrency.py, marked live
for the next zero-foreign-WINWORD round.
"""

from __future__ import annotations

import contextlib
import inspect
import sys
import threading
import time

import pytest

from word_mcp.com import bridge, convert, dialogs
from word_mcp.com import serial as com_serial
from word_mcp.com import live_batch, live_ops
from word_mcp.core.errors import (
    TargetNotFound,
    WordBlocked,
    WordBusy,
    WordMcpError,
)

# ------------------------------------------------------- 1. lock mechanics


def test_com_operation_records_state_and_reenters():
    snap0 = com_serial.lock_snapshot()
    assert snap0["held"] is False
    with com_serial.com_operation("outer-op"):
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "outer-op"
        with com_serial.com_operation("nested-op"):  # RLock: no deadlock
            assert com_serial.lock_snapshot()["held"] is True
    snap2 = com_serial.lock_snapshot()
    assert snap2["held"] is False
    assert snap2["last_op"]["name"] == "outer-op"
    assert snap2["last_op"]["duration_ms"] >= 0


def test_threads_serialize_no_interleaving():
    """Two threads running COM-op bodies must never overlap in time."""
    spans = []

    def op(name):
        with com_serial.com_operation(name):
            t0 = time.monotonic()
            time.sleep(0.15)
            spans.append((t0, time.monotonic()))

    threads = [
        threading.Thread(target=op, args=(f"t{i}",)) for i in range(4)
    ]
    for t in threads:
        t.start()
    for t in threads:
        t.join(10)
    assert len(spans) == 4
    spans.sort()
    for (s1, e1), (s2, _e2) in zip(spans, spans[1:]):
        assert e1 <= s2 + 1e-4, "COM operations overlapped in time"


def test_bounded_acquire_reports_busy_from_other_thread():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("holding-op"):
            held.set()
            release.wait(5)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        assert com_serial.acquire(timeout=0.05) is False
        snap = com_serial.lock_snapshot()
        assert snap["held"] is True
        assert snap["current_op"]["name"] == "holding-op"
    finally:
        release.set()
        t.join(5)


# ---------------------------------------------- 1b. entry-point coverage


BRIDGE_EXEMPT = {
    # tasklist only, no COM:
    "zombie_check",
    # bounded try-acquire form, marked manually and asserted below:
    "word_status",
}


def _public_functions(module):
    return {
        name: fn
        for name, fn in vars(module).items()
        if callable(fn)
        and not name.startswith("_")
        and inspect.isfunction(inspect.unwrap(fn))
        and getattr(fn, "__module__", "") == module.__name__
    }


def test_every_bridge_entry_point_is_serialized():
    for name, fn in _public_functions(bridge).items():
        if name in BRIDGE_EXEMPT and name != "word_status":
            continue
        assert getattr(fn, "_com_serialized", None), (
            f"bridge.{name} does not take the COM serialization lock "
            "(missing @_serial.serialized / @_bounded_op)"
        )


def test_convert_entry_point_is_serialized():
    assert getattr(convert.import_pdf, "_com_serialized", None)


def test_live_session_acquires_the_lock(monkeypatch):
    """live_session (every live tool and dual-mode auto-route funnels
    through it) must hold the lock before touching COM."""
    from word_mcp.com import live

    if sys.platform != "win32":  # pragma: no cover
        pytest.skip("live layer is Windows-only")
    try:
        import pythoncom  # noqa: F401
    except ImportError:  # pragma: no cover
        pytest.skip("pywin32 not installed")

    seen = {}

    class Sentinel(Exception):
        pass

    def fake_attach(win32com, pythoncom):
        seen["snapshot"] = com_serial.lock_snapshot()
        raise Sentinel()

    monkeypatch.setattr(live, "_attach_app", fake_attach)
    with pytest.raises(Sentinel):
        live.run_live("C:/nonexistent.docx", "probe tool", lambda s: {})
    assert seen["snapshot"]["held"] is True
    assert seen["snapshot"]["current_op"]["name"] == "live:probe tool"


def test_interactive_status_reports_serving_when_lock_held():
    from word_mcp.com import live

    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        out = live.interactive_status()
        assert out["interactive_state"] == "serving"
        assert out["com_serialization"]["held"] is True
        assert out["com_serialization"]["current_op"]["name"] == "long-op"
    finally:
        release.set()
        t.join(5)


def test_word_status_skips_probe_when_lock_held():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("long-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        out = bridge.word_status()
        assert "note" in out and "not probed" in out["note"]
    finally:
        release.set()
        t.join(5)


# ------------------------------------------------- 5. bounded timeouts


def test_run_bounded_fast_path_and_error_propagation():
    assert bridge._run_bounded("fast", 10, lambda: {"ok": 1}) == {"ok": 1}
    with pytest.raises(TargetNotFound):
        bridge._run_bounded(
            "err", 10, lambda: (_ for _ in ()).throw(TargetNotFound("x"))
        )


def test_run_bounded_stuck_op_raises_word_blocked(monkeypatch):
    """A worker still inside its call after the deadline AND the grace
    wait is reported as blocked. Deterministic: the worker waits on an
    event this test owns, and the grace wait is shortened, so nothing here
    depends on how fast the machine is. Nothing is force-ended, so the
    worker is released at the end to give the COM lock back."""
    monkeypatch.setattr(bridge, "TIMEOUT_GRACE_SECONDS", 0.2)
    release = threading.Event()

    def stuck():
        release.wait(30)
        return {}

    t0 = time.monotonic()
    try:
        with pytest.raises(WordBlocked, match="did not answer within"):
            bridge._run_bounded("stuck-op", 0.3, stuck)
        assert time.monotonic() - t0 < 8
    finally:
        release.set()
        for t in threading.enumerate():
            if t.name == "ks4w-stuck-op":
                t.join(10)
    assert com_serial.lock_snapshot()["held"] is False


def test_run_bounded_queued_behind_lock_raises_word_busy():
    release = threading.Event()
    held = threading.Event()

    def holder():
        with com_serial.com_operation("blocking-op"):
            held.set()
            release.wait(10)

    t = threading.Thread(target=holder)
    t.start()
    try:
        assert held.wait(5)
        with pytest.raises(WordBusy, match="blocking-op"):
            bridge._run_bounded("queued-op", 0.3, lambda: {})
    finally:
        release.set()
        t.join(5)


def test_bounded_op_timeout_validation():
    @bridge._bounded_op("val-op", default=60.0)
    def sample():
        return {"ok": True}

    assert sample() == {"ok": True}
    assert sample._com_serialized == "val-op"
    with pytest.raises(WordMcpError, match="between 5 and 3600"):
        sample(timeout=1)
    with pytest.raises(WordMcpError, match="number of seconds"):
        sample(timeout="soon")


# ------------------------------------ 5b. stage control (PPT 1.3.1, R7-3)
#
# _COMPLETION_EVENT_FACTORY is the seam. The event below is real, and the
# worker sets it exactly as in production; what the test decides is WHICH
# WAIT observes it, by running an action at a chosen wait. _run_bounded
# waits on it exactly twice, the caller's deadline and then the grace, so
# the stage is named by index and nothing is left to timing.

_STAGE_JOIN = 30.0


def test_the_completion_seam_is_the_real_event_in_production():
    """The seam exists for tests only. Unpatched, it is threading.Event
    itself, so production waits are the real Event.wait."""
    assert bridge._COMPLETION_EVENT_FACTORY is threading.Event
    assert type(bridge._COMPLETION_EVENT_FACTORY()) is threading.Event


class _StagedCompletion:
    """_run_bounded's internal completion event, with the observing stage
    chosen by the test rather than by the scheduler."""

    INITIAL_WAIT = 0
    GRACE_WAIT = 1

    def __init__(self, release, actions=None, started=None):
        self._inner = threading.Event()
        self._release = release
        self._actions = dict(actions or {})
        self._started = started
        self.waits = []

    # the event interface _run_bounded uses
    def set(self):
        self._inner.set()

    def is_set(self):
        return self._inner.is_set()

    def wait(self, timeout=None):
        stage = len(self.waits)
        self.waits.append(timeout)
        action = self._actions.get(stage)
        if action is not None:
            action(self)
        elif stage == self.INITIAL_WAIT and self._started is not None:
            # The deadline "expires" only once fn has provably started, so
            # the worker has already won the start decision (R7-1) and the
            # stage under test is the grace path, not queue contention.
            # Without this the caller could reach its abandon decision
            # before the new thread reached its start decision.
            assert self._started.wait(_STAGE_JOIN), "fn never started"
        return self._inner.is_set()

    # what the test drives
    def finish_worker(self, *_):
        """Let the worker run to completion and block until its OWN
        completion signal is set, so no later step can outrun it."""
        self._release.set()
        assert self._inner.wait(_STAGE_JOIN), "the worker never signalled"


@contextlib.contextmanager
def _staged_run(monkeypatch, actions=None, during_diagnostics=None):
    """One _run_bounded call with the completion seam installed and the
    dialog probe counted. Yields (run, staged, probes)."""
    release = threading.Event()
    started = threading.Event()
    staged = _StagedCompletion(release, actions, started=started)
    probes = []

    def probe(*_a, **_k):
        probes.append(1)
        if during_diagnostics is not None:
            during_diagnostics(staged)
        return []

    monkeypatch.setattr(bridge, "_COMPLETION_EVENT_FACTORY", lambda: staged)
    monkeypatch.setattr(dialogs, "pending_dialogs", probe)

    def body():
        started.set()
        release.wait(_STAGE_JOIN)
        return {"paragraphs": 7}

    try:
        yield (lambda: bridge._run_bounded("staged-op", 30.0, body),
               staged, probes)
    finally:
        release.set()
        _join_worker("staged-op")


def _join_worker(name):
    for t in threading.enumerate():
        if t.name == f"ks4w-{name}":
            t.join(_STAGE_JOIN)


def test_completion_during_the_initial_wait_returns_without_any_grace(
    monkeypatch,
):
    """Stage ZERO: the ordinary success. The deadline wait itself observes
    completion, so no grace and no inspection happen at all."""
    with _staged_run(
        monkeypatch,
        actions={_StagedCompletion.INITIAL_WAIT:
                 _StagedCompletion.finish_worker},
    ) as (run, staged, probes):
        assert run() == {"paragraphs": 7}
    assert len(staged.waits) == 1, "the grace wait ran on a finished worker"
    assert probes == [], "a finished worker was inspected anyway"


def test_completion_during_the_grace_wait_skips_the_diagnostics(
    monkeypatch,
):
    """Stage ONE, the first post-grace check: a finished success is never
    delayed behind the dialog inspection."""
    with _staged_run(
        monkeypatch,
        actions={_StagedCompletion.GRACE_WAIT:
                 _StagedCompletion.finish_worker},
    ) as (run, staged, probes):
        assert run() == {"paragraphs": 7}
    assert len(staged.waits) == 2, "the grace wait was not reached"
    assert probes == [], (
        "a completed result waited behind the dialog inspection"
    )


def test_completion_during_the_diagnostics_returns_on_the_second_check(
    monkeypatch,
):
    """Stage TWO, the second check: the inspection can itself take long
    enough for the worker to finish inside it, and the answer it gives
    then is still the worker's own."""
    with _staged_run(
        monkeypatch,
        during_diagnostics=_StagedCompletion.finish_worker,
    ) as (run, staged, probes):
        assert run() == {"paragraphs": 7}
    assert len(staged.waits) == 2
    assert probes == [1], "the second check ran without any inspection"


def test_a_worker_that_never_completes_gets_the_refusal(monkeypatch):
    """Stage THREE: neither check sees completion, the inspection did run,
    and the refusal is what comes back."""
    with _staged_run(monkeypatch) as (run, staged, probes):
        with pytest.raises(WordBlocked) as exc_info:
            run()
        assert len(staged.waits) == 2
        assert probes == [1]
    assert "The operation was NOT cancelled" in str(exc_info.value)


# ------------------------- 5c. queued ghost write (PPT 1.3.1, R7-1)
#
# A caller that gives up while its worker is still QUEUED on the COM
# serialization lock used to get WordBusy and an invitation to retry, while
# the worker stayed in the queue and ran fn() anyway once the holder
# released. A retried call therefore ran twice. One atomic decision per
# call now settles it, and these tests prove both sides of it.

_ABANDONED_SENTENCE = (
    " The queued call was abandoned before its operation ran; it made no "
    "document changes. It is safe to retry after com_word_status reports "
    "no COM operation in progress."
)


@contextlib.contextmanager
def _lock_holder(name="existing-write"):
    """Hold the process-wide COM serialization lock until released."""
    held = threading.Event()
    release = threading.Event()

    def hold():
        with com_serial.com_operation(name):
            held.set()
            release.wait(_STAGE_JOIN)

    t = threading.Thread(target=hold, daemon=True, name="ks4w-test-holder")
    t.start()
    assert held.wait(_STAGE_JOIN), "the holder never took the lock"
    try:
        yield release
    finally:
        release.set()
        t.join(_STAGE_JOIN)


def _queued_worker(name):
    for t in threading.enumerate():
        if t.name == f"ks4w-{name}":
            return t
    return None


def test_a_call_abandoned_while_queued_never_runs_its_operation():
    """R7-1(a). The holder keeps the lock past the caller's wait, so the
    caller is told the call was abandoned. The proof is what happens
    AFTER the holder releases: the queued worker reaches the front of the
    queue, loses the decision, and exits without ever entering fn."""
    calls = []
    with _lock_holder() as release_holder:
        with pytest.raises(WordBusy) as exc_info:
            bridge._run_bounded(
                "queued-write", 0.05, lambda: calls.append("ran")
            )
        assert calls == [], "the operation ran before the caller gave up"
        worker = _queued_worker("queued-write")
        assert worker is not None, "the queued worker was not found"
        release_holder.set()
        worker.join(_STAGE_JOIN)
        assert not worker.is_alive(), "the queued worker never finished"
    assert calls == [], (
        "the abandoned worker executed the operation after the holder "
        "released the lock, so a caller told nothing had happened had it "
        "happen behind its back"
    )
    message = str(exc_info.value)
    assert message.endswith(_ABANDONED_SENTENCE), message
    assert message == (
        "queued-write waited 0s for the COM serialization lock "
        "(existing-write is still running); retry when it finishes. "
        "com_word_status reports the running operation."
        + _ABANDONED_SENTENCE
    )


def test_an_abandoned_worker_takes_the_lock_and_still_never_enters_fn():
    """The abandoned worker is not merely slow: it DOES reach the front of
    the queue and take the lock after the holder releases, and even then it
    never enters fn. Repeated, with the holder released the instant the
    caller gives up, so the tightest ordering is exercised every time.
    last_op proves the worker really held the lock, which keeps the
    assertion from passing vacuously on a worker that never ran at all."""
    entered = []

    def fn():
        entered.append(threading.get_ident())
        return {"changed": True}

    for i in range(25):
        name = f"abandon-{i}"
        with _lock_holder() as release_holder:
            with pytest.raises(WordBusy, match="was abandoned"):
                bridge._run_bounded(name, 0.02, fn)
            worker = _queued_worker(name)
            release_holder.set()
        if worker is not None:
            worker.join(_STAGE_JOIN)
            assert not worker.is_alive()
        snap = com_serial.lock_snapshot()
        assert snap["held"] is False
        assert snap["last_op"]["name"] == name, (
            "the abandoned worker never took the lock, so the test proved "
            "nothing"
        )
    assert entered == [], (
        f"an abandoned worker entered fn {len(entered)} time(s) after the "
        "holder released"
    )


def test_a_retry_after_an_abandoned_call_runs_the_operation_once():
    """R7-1(c). The refusal invites a retry, so the retry must be the only
    execution there ever is."""
    calls = []

    def refresh_fields():
        calls.append("ran")
        return {"fields_refreshed": True}

    with _lock_holder() as release_holder:
        with pytest.raises(WordBusy):
            bridge._run_bounded("queued-write", 0.05, refresh_fields)
        worker = _queued_worker("queued-write")
        release_holder.set()
        if worker is not None:
            worker.join(_STAGE_JOIN)
    assert bridge._run_bounded(
        "queued-write", 30.0, refresh_fields
    ) == {"fields_refreshed": True}
    assert calls == ["ran"], (
        "the operation ran twice, once from the abandoned queued worker and "
        "once from the retry"
    )


def test_a_worker_that_wins_the_start_decision_is_never_told_to_retry(
    monkeypatch,
):
    """R7-1(b). The other side of the same decision: the worker takes the
    lock and starts fn in the instant before the caller would have
    abandoned it. The caller must then follow the grace path and get the
    worker's own answer, never a queue-contention retry.

    The completion seam makes the order explicit rather than hoped for:
    the lock is handed over during the caller's deadline wait, and that
    wait does not return until fn has provably been entered."""
    calls = []
    entered = threading.Event()
    finish = threading.Event()

    def op():
        calls.append("ran")
        entered.set()
        finish.wait(_STAGE_JOIN)
        return {"paragraphs": 4}

    with _lock_holder() as release_holder:

        def start_the_worker(_staged):
            release_holder.set()
            assert entered.wait(_STAGE_JOIN), "the worker never started fn"

        staged = _StagedCompletion(
            finish,
            {
                _StagedCompletion.INITIAL_WAIT: start_the_worker,
                _StagedCompletion.GRACE_WAIT: _StagedCompletion.finish_worker,
            },
        )
        monkeypatch.setattr(
            bridge, "_COMPLETION_EVENT_FACTORY", lambda: staged
        )
        monkeypatch.setattr(dialogs, "pending_dialogs", lambda *a, **k: [])
        try:
            assert bridge._run_bounded("racing-write", 30.0, op) == {
                "paragraphs": 4
            }
        except WordBusy as exc:  # pragma: no cover - the defect
            pytest.fail(f"a started operation was reported as queued: {exc}")
    assert calls == ["ran"], "the operation did not run exactly once"


# ---------------------------------------- 2. tracked-replace loop (fake COM)


class _FakeRev:
    def __init__(self, start, end):
        self.Type = 2  # wdRevisionDelete

        class _R:
            pass

        self.Range = _R()
        self.Range.Start = start
        self.Range.End = end


class _FakeFind:
    def __init__(self, rng):
        self._rng = rng
        self.Text = ""
        self.Forward = True
        self.Wrap = 0
        self.MatchWildcards = False
        self.MatchCase = True

    def ClearFormatting(self):
        pass

    def Execute(self):
        story = self._rng._story
        i = story.text.find(self.Text, self._rng.Start)
        if i < 0:
            return False
        self._rng.SetRange(i, i + len(self.Text))
        return True


class _FakeRange:
    def __init__(self, story, start, end):
        self._story = story
        self.Start = start
        self.End = end

    def SetRange(self, start, end):
        self.Start, self.End = start, end

    @property
    def Duplicate(self):
        return _FakeRange(self._story, self.Start, self.End)

    @property
    def Find(self):
        return _FakeFind(self)

    @property
    def Text(self):
        return self._story.text[self.Start:self.End]

    @Text.setter
    def Text(self, value):
        self._story.assign(self, value)

    @property
    def Revisions(self):
        return self._story.revisions_for(self.Start, self.End)

    @property
    def ParentContentControl(self):
        return None


class FakeStory(_FakeRange):
    """Emulates the Word 2016 behavior the stress report hit: a tracked
    Range.Text assignment INSERTS the replacement at the range start and
    keeps the old text findable in the story as a tracked deletion AFTER
    the insertion (so a naive resume point sits right before the deleted
    copy of the find text)."""

    def __init__(self, text, *, tracked=False, revisions_visible=True,
                 revisions_raise=False, deletions=None):
        self._story = self
        self.text = text
        self.tracked = tracked
        self.revisions_visible = revisions_visible
        self.revisions_raise = revisions_raise
        self.deletions = list(deletions or [])
        self.assignments = 0

    @property
    def Start(self):
        return 0

    @Start.setter
    def Start(self, v):  # story range is the whole story
        pass

    @property
    def End(self):
        return len(self.text)

    @End.setter
    def End(self, v):
        pass

    @property
    def Fields(self):
        return []

    def assign(self, rng, value):
        self.assignments += 1
        if not self.tracked:
            old = self.text
            self.text = old[:rng.Start] + value + old[rng.End:]
            rng.End = rng.Start + len(value)
            return
        old_len = rng.End - rng.Start
        self.text = (
            self.text[:rng.Start] + value + self.text[rng.Start:]
        )
        shift = len(value)
        self.deletions = [
            (s + shift if s >= rng.Start else s,
             e + shift if e > rng.Start else e)
            for s, e in self.deletions
        ]
        self.deletions.append(
            (rng.Start + shift, rng.Start + shift + old_len)
        )
        rng.End = rng.Start + shift  # range covers the insertion only

    def revisions_for(self, start, end):
        if self.revisions_raise:
            raise RuntimeError("COM call rejected (emulated contention)")
        if not self.revisions_visible:
            return []
        return [
            _FakeRev(s, e) for s, e in self.deletions
            if s < end and e > start
        ]


ERP_TEXT = (
    "The enterprise resource planning rollout continues. "
    "Other text follows the single occurrence."
)


def test_tracked_replace_does_not_rematch_own_markup():
    """The report's bug 1 mechanism: 1 occurrence, tracked replace,
    insertion lands before the still-findable deleted copy. Must replace
    exactly once (the pre-fix loop produced ERP x51)."""
    story = FakeStory(ERP_TEXT, tracked=True)
    done, skipped, skipped_del = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert story.text.count("ERP") == 1
    assert "ERPERP" not in story.text


def test_tracked_replace_two_occurrences_both_replaced_once():
    """Guarantee 2 in action: after replacing occurrence 1, the resume
    point advances past its own deletion markup, so occurrence 2 (the
    real one) is the next match, not the deleted copy of occurrence 1."""
    text = ERP_TEXT + " Later the enterprise resource planning gets cut."
    story = FakeStory(text, tracked=True)
    done, _, _ = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 2
    assert story.text.count("ERP") == 2
    assert "ERPERP" not in story.text


def test_tracked_replace_capped_even_if_revisions_invisible():
    """Defense in depth: if COM misreports the deletion markup as
    revision-free (the concurrency failure mode), the pre-edit match
    count caps replacements at 1 — garbage growth is impossible."""
    story = FakeStory(ERP_TEXT, tracked=True, revisions_visible=False)
    done, _, _ = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert story.text.count("ERP") == 1
    assert "ERPERP" not in story.text


def test_tracked_replace_fails_closed_on_revision_read_error():
    story = FakeStory(ERP_TEXT, tracked=True, revisions_raise=True)
    with pytest.raises(WordBusy, match="fail-closed"):
        live_ops._replace_literal(
            story, "enterprise resource planning", "ERP", tracked=True
        )
    assert story.assignments == 0  # nothing replaced unverified


def test_tracked_replace_skips_preexisting_deletion():
    """A document that ALREADY contains tracked changes (the report's
    single-agent hypothesis): a match inside an existing tracked deletion
    is skipped, the live one is replaced."""
    text = "old enterprise resource planning gone. " + ERP_TEXT
    start = text.find("enterprise")
    end = start + len("enterprise resource planning")
    story = FakeStory(text, tracked=True, deletions=[(start, end)])
    done, _, skipped_del = live_ops._replace_literal(
        story, "enterprise resource planning", "ERP", tracked=True
    )
    assert done == 1
    assert skipped_del == 1
    assert story.text.count("ERP") == 1


def test_untracked_self_referencing_replacement_still_terminates():
    story = FakeStory("alliance one alliance two", tracked=False)
    done, _, _ = live_ops._replace_literal(
        story, "alliance", "alliance-x", tracked=False
    )
    assert done == 2
    assert story.text == "alliance-x one alliance-x two"


# ------------------------------------ 3. apply_edits atomicity (unit level)


def _spec(op, **kw):
    out = {"op": op, "index": kw.pop("index", None)}
    out.update(kw)
    return out


def test_preflight_conflicts_catches_deleted_target():
    specs = [
        _spec("delete", indices=[2]),
        _spec("set_text", index=2, text="x"),
    ]
    with pytest.raises(WordMcpError, match="Nothing was applied"):
        live_batch._preflight_conflicts(specs, n_paras=5)


def test_preflight_conflicts_accepts_valid_sequence():
    specs = [
        _spec("insert", index=1, items=[{"text": "a"}], mode="after"),
        _spec("set_text", index=3, text="x"),
        _spec("delete", indices=[0, 4]),
    ]
    live_batch._preflight_conflicts(specs, n_paras=5)  # no raise


class _FakeUndo:
    def __init__(self, recording=True):
        self.IsRecordingCustomRecord = recording
        self.ended = 0

    def EndCustomRecord(self):
        self.ended += 1
        self.IsRecordingCustomRecord = False


class _FakeSession:
    def __init__(self, grouped=True, undo_raises=False):
        self.undo_grouped = grouped

        class _App:
            pass

        class _Doc:
            def __init__(self):
                self.undone = 0
                self._raises = undo_raises

            def Undo(self):
                if self._raises:
                    raise RuntimeError("no undo")
                self.undone += 1

        self.app = _App()
        self.app.UndoRecord = _FakeUndo()
        self.doc = _Doc()


def test_rollback_via_undo_group():
    s = _FakeSession(grouped=True)
    note = live_batch._rollback(s)
    assert "ROLLED BACK" in note
    assert s.app.UndoRecord.ended == 1
    assert s.doc.undone == 1


def test_rollback_honest_when_ungrouped_or_failing():
    assert "PARTIALLY APPLIED" in live_batch._rollback(
        _FakeSession(grouped=False)
    )
    assert "PARTIALLY APPLIED" in live_batch._rollback(
        _FakeSession(grouped=True, undo_raises=True)
    )


# ------------------------------------------- 4. alerts suppression contract


def test_alerts_suppressed_sets_and_restores():
    class _App:
        DisplayAlerts = -1

    app = _App()
    with bridge._alerts_suppressed(app):
        assert app.DisplayAlerts == 0
    assert app.DisplayAlerts == -1


def test_alerts_suppressed_restores_on_error():
    class _App:
        DisplayAlerts = -1

    app = _App()
    with pytest.raises(RuntimeError):
        with bridge._alerts_suppressed(app):
            assert app.DisplayAlerts == 0
            raise RuntimeError("boom")
    assert app.DisplayAlerts == -1


def test_retry_word_call_backs_off_and_reports():
    calls = []

    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise RuntimeError("transient")
        return "ok"

    value, retries = bridge._retry_word_call(flaky, first_delay=0.01)
    assert value == "ok" and retries == 2

    def always_fails():
        raise RuntimeError("permission error")

    with pytest.raises(WordBusy, match="no dialog is pending"):
        bridge._retry_word_call(
            always_fails, attempts=2, first_delay=0.01
        )


# --------------------------------------- 6. dialog detection (OS layer)


def test_pending_dialogs_empty_without_word():
    assert dialogs.pending_dialogs(pids=set()) == []
    assert dialogs.pending_dialogs(pids={999999999}) == []


@pytest.mark.skipif(sys.platform != "win32", reason="Win32 only")
def test_pending_dialogs_sees_synthetic_dialog_window():
    """Create a hidden-offscreen but WS_VISIBLE #32770 window in this
    process and detect it via the same enumeration com_word_status uses."""
    import ctypes
    import os

    user32 = ctypes.windll.user32
    user32.CreateWindowExW.restype = ctypes.c_void_p
    WS_POPUP = 0x80000000
    WS_VISIBLE = 0x10000000
    title = "ks4w dialog probe: file permission error"
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, WS_POPUP | WS_VISIBLE,
        -32000, -32000, 1, 1, None, None, None, None,
    )
    assert hwnd, "could not create the synthetic dialog window"
    try:
        found = dialogs.pending_dialogs(pids={os.getpid()})
        assert any(
            d["title"] == title and d["class"] == "#32770" for d in found
        ), f"synthetic dialog not detected: {found}"
    finally:
        user32.DestroyWindow(hwnd)
    assert not any(
        d["title"] == title
        for d in dialogs.pending_dialogs(pids={os.getpid()})
    )


# ----------------------------------------- name-collision guard (unit)


def test_proofing_refuses_open_document(monkeypatch, tmp_path):
    doc = tmp_path / "open_doc.docx"
    doc.write_bytes(b"PK\x03\x04stub")
    monkeypatch.setattr(bridge, "_open_in_running_word", lambda p: True)
    with pytest.raises(WordBusy, match="same-name dialog"):
        bridge.proofing_errors(str(doc))
    with pytest.raises(WordBusy, match="same-name dialog"):
        bridge.readability_statistics(str(doc))
