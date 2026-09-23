"""Process-wide COM serialization (live COM stress report, 2026-09-03).

Word's COM interface is single-threaded (STA). The stress test proved that
concurrent tool calls reaching the Word proxy interleave at the character
level (garbled Range.Text writes), collide on the save temp-file swap
(modal dialogs), and starve invisible instances (30-minute hangs). The fix
is architectural and small: exactly ONE tool call reaches any Word COM
proxy at a time, enforced by this module's process-wide lock.

Coverage contract: every COM entry point acquires the lock —
- the live layer (live.live_session wraps every live tool and every
  dual-mode auto-route, live_repair, interactive_status),
- the bridge layer (every public invisible-instance and message-the-
  visible-instance function; enforced by test_com_serialization's audit
  of the _com_serialized marker),
- the convert layer (import_pdf).

The lock is an RLock: nested acquisitions on one thread are legal (a
bridge helper called under an already-held lock must not deadlock).

Scope honesty: this lock serializes THIS server process only. That was
the whole scope until the 2026-09-05 concurrency matrix measured what it
left open — two server processes against one Word instance overlapped in
49 of 50 operation pairs, double-applied a find-then-assign replace, and
leaked DisplayAlerts to wdAlertsNone. Cross-PROCESS serialization is a
separate, file-based lock in com/xproc.py, which every live session now
holds on top of this one. Keep the two distinct: this one is cheap,
in-memory, and covers threads; that one is a lockfile and covers the
machine.

Individual operations are fast; serialization latency is negligible
(report finding). Status reporting (lock_snapshot) lets com_word_status
tell callers honestly when a call would queue.

NO LATE ENTRY (M1, 2.2.1 release review X-20260924-011). A caller that
waits on this lock without a bound can outlive its client: behind a hung
invisible Word, which nothing force-ends any more, a live edit or a save
sat queued after the client had given up, the client retried, and both
bodies ran against the user's open document once the holder let go.
Callers that change the user's open document therefore opt into
``refuse_if_busy``: they take the lock only if it is free (or already
theirs, since it is re-entrant) and otherwise raise ``CallNotStarted``
BEFORE their body, leaving nothing queued. A finite wait would not do: the
server cannot know how long its client will wait.
"""

from __future__ import annotations

import contextlib
import threading
import time

from ..core.errors import CallNotStarted

COM_LOCK = threading.RLock()

_state_lock = threading.Lock()
_current: dict | None = None      # {name, thread, started_wall, started_mono}
_last: dict | None = None         # {name, duration_ms, waited_ms, finished_wall}
_depth = 0                        # re-entrant depth on the owning thread
_serialized_total = 0             # ops that ran under the lock
_waited_total_ms = 0.0            # cumulative wait time across ops

#: The refusal every no-late-entry route raises, pinned by
#: test_m1_no_late_entry. {holder} names what is using Word. COPY SLOT
#: P1-P-3 (copy packet 2, CM-20260924-012): the sentence is Codex's to
#: write; replace this placeholder and its pin in the same commit.
BUSY_NOT_STARTED = (
    "[[COPY: P1-P-3, see COPY_PACKET_2_FACT_SHEET §6 (P-3). The "
    "APP_BUSY refusal for a live call, or a com_save_document save or "
    "close, refused before it started because Word is already in use by "
    "{holder}. Facts: this call did not start and made no document "
    "changes; nothing was queued, so nothing from this call will run "
    "later; wait until com_word_status reports no COM operation in "
    "progress, then retry. No existing approved string says this "
    "truthfully: the R7-1 sentence speaks of a queued call that was "
    "abandoned.]]"
)


def busy_refusal(holder: str) -> CallNotStarted:
    """The no-late-entry refusal, naming what holds Word."""
    return CallNotStarted(BUSY_NOT_STARTED.format(holder=holder))


def _current_holder() -> str:
    with _state_lock:
        if _current is not None:
            return _current["name"]
    return "another COM operation"


@contextlib.contextmanager
def com_operation(name: str, *, refuse_if_busy: bool = False):
    """Hold the process-wide COM lock for the duration of one COM-touching
    operation. Records timing so com_word_status can report contention.

    refuse_if_busy: take the lock only if it is free or already held by
    this thread; otherwise raise CallNotStarted before the caller's body
    runs, with nothing left waiting (see NO LATE ENTRY above)."""
    global _current, _last, _depth, _serialized_total, _waited_total_ms
    t0 = time.monotonic()
    if refuse_if_busy:
        if not COM_LOCK.acquire(blocking=False):
            raise busy_refusal(_current_holder())
    else:
        COM_LOCK.acquire()
    waited_ms = (time.monotonic() - t0) * 1000.0
    with _state_lock:
        _depth += 1
        outermost = _depth == 1
        if outermost:
            _current = {
                "name": name,
                "thread": threading.get_ident(),
                "started_wall": time.time(),
                "started_mono": time.monotonic(),
            }
            _serialized_total += 1
            _waited_total_ms += waited_ms
    started = time.monotonic()
    try:
        yield
    finally:
        with _state_lock:
            _depth -= 1
            if outermost:
                _last = {
                    "name": name,
                    "duration_ms": round(
                        (time.monotonic() - started) * 1000.0, 1
                    ),
                    "waited_ms": round(waited_ms, 1),
                    "finished_wall": time.time(),
                }
                _current = None
        COM_LOCK.release()


def serialized(name: str, *, refuse_if_busy: bool = False):
    """Decorator form of com_operation for whole-function COM operations.
    Marks the function so the coverage audit test can verify every COM
    entry point takes the lock (and, with refuse_if_busy, that it refuses
    rather than queues)."""

    def deco(fn):
        import functools

        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            with com_operation(name, refuse_if_busy=refuse_if_busy):
                return fn(*args, **kwargs)

        wrapper._com_serialized = name
        wrapper._refuses_if_busy = refuse_if_busy
        return wrapper

    return deco


def lock_snapshot() -> dict:
    """Contention report for com_word_status: is a COM operation running
    right now, what is it, and how long has it held the lock. last_op
    carries the previous operation's duration and queue wait."""
    with _state_lock:
        out: dict = {
            "held": _current is not None,
            "ops_serialized": _serialized_total,
        }
        if _current is not None:
            out["current_op"] = {
                "name": _current["name"],
                "running_ms": round(
                    (time.monotonic() - _current["started_mono"]) * 1000.0, 1
                ),
            }
        if _last is not None:
            out["last_op"] = {
                "name": _last["name"],
                "duration_ms": _last["duration_ms"],
                "waited_ms": _last["waited_ms"],
            }
        return out


def acquire(timeout: float) -> bool:
    """Bounded acquisition for callers that must stay responsive
    (com_word_status). Pair with release() only when this returns True."""
    return COM_LOCK.acquire(timeout=timeout)


def release() -> None:
    COM_LOCK.release()
