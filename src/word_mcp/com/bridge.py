"""COM bridge: the operations only a running Word can perform.

Every function opens its own dedicated invisible Word instance (DispatchEx —
never attaches to the user's visible Word), disables alerts, and guarantees
Quit in a finally block, EXCEPT the message-the-visible-instance trio
(word_status, save_open_document, close_open_document), which talk to the
user's Word deliberately and suppress alerts around every call.

Serialization (2026-09-03 stress report): every public function here runs
under the process-wide COM lock (com/serial.py) — via @_serial.serialized,
via _run_bounded's worker, or via a bounded try-acquire for status — so
no bridge call can interleave with a live-layer edit or another bridge
call. Long invisible-instance operations additionally run under a bounded
timeout, which turns the report's 30-minute silent hang into a structured
WordBlocked. The deadline bounds how long the CALLER waits; it cancels
nothing.

NOTHING IN THIS PACKAGE FORCE-ENDS A WORD PROCESS (the KitchenSink4PPT
1.3.1 rule, adopted here for the 2026-09-22 field-test fixes). The
invisible instance an operation starts is released with an orderly
Quit() on its own COM object, and that is all. A timed-out operation is
reported, not killed: a timeout cannot revalidate a hung apartment, so a
PID recorded minutes earlier is not proof of what that process is now.
Ownership needs POSITIVE evidence, and a reading that fails is a no: an
unreadable process table is unknown, never empty, and an unreadable
command line owns nothing.
"""

from __future__ import annotations

import collections
import contextlib
import functools
import subprocess
import sys
import threading
import time
from pathlib import Path

from ..core.errors import (
    DocumentNotFound,
    WordBlocked,
    WordBusy,
    WordMcpError,
)
from ..core.sandbox import check_path
from . import callargs as _args
from . import rot as _rot
from . import serial as _serial

_WD_ALERTS_NONE = 0
_WD_FORMAT_PDF = 17
_WD_FORMAT_DOCX_DEFAULT = 16
#: An encrypted Office document is an OLE compound file; a plain
#: one is a zip starting PK. Four bytes settle it, with no
#: password prompt and no second Word instance.
_OLE_MAGIC = bytes.fromhex("d0cf11e0")
_WD_DO_NOT_SAVE = 0
_WD_SAVE = -1

# Invisible WINWORD.EXE PIDs started by _word(), keyed by spawning thread.
# EVIDENCE, not a target list: nothing in this package force-ends a
# process. An entry exists only when _word() READ the process table before
# and after DispatchEx and found exactly one new WINWORD.EXE
# (_acquisition_token). It records what was OBSERVED at start-up, that the
# pid was not running when the call began; it is not proof of what that
# process is now. A timed-out operation only NAMES these in its refusal.
_INVISIBLE_PIDS: dict[int, set] = {}

#: Nothing this module spawns may flash a console on the user's desktop
#: (author directive 2026-09-06). tasklist and the command-line probe are
#: console programs, so every launch carries the flag on Windows and 0
#: elsewhere.
_NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)

#: How long to wait for a PID to actually leave the process table after
#: Quit. Word exits in well under a second when Quit is honored; the
#: budget is for the case where it is not.
_EXIT_WAIT_SECONDS = 6.0

#: How much longer a timed-out operation is given before the refusal is
#: built. The deadline bounds the CALLER's wait and cancels nothing, so a
#: worker that finishes in this window has finished and its result is
#: returned instead of a false timeout (KitchenSink4PPT 1.3.1, R6-1).
TIMEOUT_GRACE_SECONDS = 10.0

#: TEST SEAM (KitchenSink4PPT 1.3.1, R7-3). _run_bounded builds its internal
#: completion signal through this factory, so a test can substitute an event
#: whose waits it drives and pin WHICH stage observed completion without
#: assuming anything about thread scheduling. Production behaviour is
#: unchanged: this is threading.Event, with its real wait(), and nothing in
#: the package reassigns it.
_COMPLETION_EVENT_FACTORY = threading.Event

#: The last few lifecycle diagnostics, for tests and for com_word_status.
#: A failed Quit used to be swallowed by contextlib.suppress, which is how
#: an orphan could exist with nothing anywhere recording that it did.
_LIFECYCLE_NOTES: collections.deque = collections.deque(maxlen=20)


def _note(message: str) -> None:
    """Record and announce one lifecycle event. stderr, never stdout:
    stdout carries the MCP protocol."""
    _LIFECYCLE_NOTES.append(message)
    with contextlib.suppress(Exception):
        sys.stderr.write(f"[kitchensink4word] {message}\n")


def _winword_pids() -> set | None:
    """WINWORD.EXE process ids via the process table (never COM), or None
    when the table could not be READ.

    NONE MEANS UNKNOWN AND NEVER MEANS "NOTHING WAS RUNNING" (the
    KitchenSink4PPT 1.3.1 rule, G2b and R4-1). This used to return an
    empty set for both, so a failed tasklist at the start of a call read as
    an empty machine, and every Word already running was then counted as
    one this call had started. Nothing that decides ownership may act on
    None.

    A genuine no-match is NOT a failure: tasklist prints its INFO line,
    exits 0, and the answer is an empty set. Unknown is the subprocess
    raising or timing out, ANY non-zero exit (tasklist prints "ERROR:
    Access is denied." to stdout and exits 1), no output at all, or a row
    naming WINWORD.EXE whose pid column will not parse.
    """
    try:
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV",
             "/NH"],
            capture_output=True,
            text=True,
            timeout=30,
            creationflags=_NO_WINDOW,
        )
    except Exception:
        return None
    if result.returncode != 0:
        return None
    stdout = result.stdout or ""
    if not stdout.strip():
        return None
    pids = set()
    for ln in stdout.splitlines():
        if "WINWORD.EXE" not in ln.upper():
            continue
        parts = ln.split('","')
        if len(parts) < 2:
            return None
        try:
            pids.add(int(parts[1].strip('"')))
        except ValueError:
            return None
    return pids


def _acquisition_token(before, after) -> set | None:
    """The pids this call may record as the instance it started, or None.

    ALL of these must hold, and a read that failed is a no:

    - the process table was READ before DispatchEx (before is not None);
    - it was READ again after DispatchEx (after is not None);
    - exactly ONE WINWORD.EXE appeared in between. Word starts one private
      process per DispatchEx, so one is the only count that names it; two
      means someone else started Word in the same instant, and there is
      then no telling which one is ours.

    Word already running is normal here, unlike PowerPoint's singleton: the
    user's Word is simply in both readings and never in the difference.
    """
    if before is None or after is None:
        return None
    created = set(after) - set(before)
    if len(created) != 1:
        return None
    return created


def _is_automation_instance(pid: int) -> bool | None:
    """Is this PID a Word started for automation, rather than the user's?

    True / False when the command line settles it, None when it cannot be
    read. An interactive WINWORD.EXE never carries /Automation, and the
    one started by DispatchEx always does. The policy at every call site
    is: this server claims a process only on a definite True. None means
    the command line could not be read, and an unreadable command line
    OWNS NOTHING, exactly like False: no claim is made about it and no
    advice is given about ending it.
    """
    if sys.platform != "win32":
        return None
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command",
             "(Get-CimInstance Win32_Process -Filter "
             f"\"ProcessId={int(pid)}\").CommandLine"],
            capture_output=True, text=True, timeout=20,
            creationflags=_NO_WINDOW,
        )
    except Exception:
        return None
    if out.returncode != 0:
        return None
    line = (out.stdout or "").strip()
    if not line:
        return None  # already gone, or unreadable: not a definite "no"
    return "/automation" in line.lower()


def _claims(pid: int) -> bool:
    """Whether this server may say a recorded pid is the instance it
    launched. Only a command line READ as carrying /Automation says yes;
    False and None (unreadable) both own nothing."""
    return _is_automation_instance(pid) is True


def _survivors(pids, timeout: float = _EXIT_WAIT_SECONDS) -> set:
    """Which of these PIDs are still in the process table after a bounded
    wait. Quit is asynchronous: Word acknowledges it and then takes a
    moment, so checking once would report a false orphan.

    If the table goes unreadable mid-poll the answer is the empty set:
    unknown is not evidence that anything survived, and polling it again
    cannot become evidence either (G2b)."""
    remaining = {int(p) for p in pids}
    if not remaining:
        return remaining
    deadline = time.monotonic() + timeout
    while True:
        now = _winword_pids()
        if now is None:
            return set()
        remaining &= now
        if not remaining or time.monotonic() >= deadline:
            return remaining
        time.sleep(0.25)


def _close_open_documents(app) -> None:
    """Close whatever this instance still has open, before Quit.

    This is where the orphan came from. Individual operations close their
    own document, but any path that raised before its close left one open,
    and Word will not Quit an instance holding a document it considers
    unsaved. The refusal was swallowed and the PID record was dropped on
    the very next line, unchecked: exactly the two /Automation orphans the
    2026-09-21 field test found still running.

    wdDoNotSaveChanges, always: this instance is private to one operation
    and everything it was asked to persist was persisted by that operation.
    """
    try:
        count = int(app.Documents.Count)
    except Exception:  # noqa: BLE001 - a dead app has nothing to close
        return
    for _ in range(count):
        try:
            app.Documents(1).Close(_WD_DO_NOT_SAVE)
        except Exception as exc:  # noqa: BLE001
            _note(f"could not close a document before Quit: {exc}")
            return  # one refusal is enough; do not spin on the same doc


@contextlib.contextmanager
def _word():
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:  # pragma: no cover
        raise WordMcpError(
            "pywin32 is not available; COM operations need it"
        ) from exc
    pythoncom.CoInitialize()
    app = None
    tid = threading.get_ident()
    # OWNERSHIP NEEDS POSITIVE EVIDENCE. before is None when tasklist could
    # not be read, and then nothing is recorded for this call at all.
    before = _winword_pids()
    try:
        app = win32com.client.DispatchEx("Word.Application")
        created = _acquisition_token(before, _winword_pids())
        if created:
            _INVISIBLE_PIDS[tid] = created
        app.Visible = False
        app.DisplayAlerts = _WD_ALERTS_NONE
        yield app
    finally:
        # Take the PID record BEFORE anything can fail. It is evidence for
        # a timeout refusal while the call runs, and nothing after this
        # call acts on it: nothing in this package force-ends a process.
        pids = _INVISIBLE_PIDS.pop(tid, None) or set()
        if app is not None:
            _close_open_documents(app)
            try:
                app.Quit(_WD_DO_NOT_SAVE)
            except Exception as exc:  # noqa: BLE001
                _note(
                    f"Word.Quit failed on the invisible instance {sorted(pids) or '(pid unknown)'}: "
                    f"{exc}"
                )
        if pids:
            # Verified, not assumed: Quit is asynchronous, so the record is
            # checked against the process table before it is dropped. A
            # survivor is REPORTED, never ended, and only when its command
            # line is read as an automation instance: an unreadable one owns
            # nothing, so no claim is made and no advice is given about it.
            # A cleanup problem is announced, not raised over the top of an
            # operation that already succeeded.
            alive = _survivors(pids)
            if any(_claims(pid) for pid in sorted(alive)):
                _note(
                    "the Word instance this call launched did not exit "
                    f"within {_EXIT_WAIT_SECONDS:.0f}s of Quit(); a "
                    "WINWORD.EXE process may be lingering (zombie_check() "
                    "to confirm). It was launched by this tool and is safe "
                    "to end via Task Manager."
                )
        pythoncom.CoUninitialize()


def _run_bounded(name: str, timeout: float, fn):
    """Run fn on a worker thread under the COM lock, bounding how long the
    CALLER waits.

    On expiry: if the worker never got the lock, that is queue contention
    and WordBusy names the operation actually running. If it got the lock
    and then stalled, NOTHING IS CANCELLED AND NOTHING IS FORCE-ENDED: the
    deadline bounds how long the caller waits, not how long the work runs.
    The worker is still inside the COM call, can finish later, and can
    still save its output.

    It used to force-end the invisible instance the worker had started.
    That is gone (the KitchenSink4PPT 1.3.1 rule): a timeout cannot
    revalidate a hung apartment, so the PID recorded at start-up is not
    proof that the process is still only this call's.

    The completion check is TWO-STAGE, as in KitchenSink4PPT 1.3.1: once
    the instant the grace wait ends, before any dialog inspection, and once
    more after it. Either way the worker's own answer wins, its result
    returned or its exception raised, because reporting a timeout for an
    operation that succeeded is a false failure the caller acts on (R6-1).
    Only if it is still not done is the refusal built, and everything it
    says about a process is state-neutral: what was observed, and what
    this path did not do (R6-2).

    QUEUE ABANDONMENT (KitchenSink4PPT 1.3.1, R7-1). A caller that gave up
    while its worker was still QUEUED on the lock used to get WordBusy and
    an invitation to retry, while the worker stayed in the queue and ran fn
    anyway once the holder released, so a retried call ran twice. There is
    now ONE decision per call, taken atomically, with two outcomes: the
    caller marks the call ABANDONED when its wait expires, or the worker
    marks it STARTED after it acquires the lock and BEFORE it calls fn.
    Whichever side reaches the decision first wins and the other stands
    down, so an abandoned call never runs and a started one is never
    reported as queued."""
    result: dict = {}
    done = _COMPLETION_EVENT_FACTORY()
    worker_tid: list = []

    # ONE decision per call, taken under this lock, with exactly two
    # outcomes: the caller ABANDONED the queued call, or the worker STARTED
    # it. The first to arrive wins; the other is told it lost (R7-1).
    decision_lock = threading.Lock()
    decision: list = []

    def _elect(outcome: str) -> bool:
        """Take this call's one decision, or lose it to the other side."""
        with decision_lock:
            if decision:
                return False
            decision.append(outcome)
            return True

    def worker():
        worker_tid.append(threading.get_ident())
        try:
            with _serial.com_operation(name):
                # Winning the serialization lock is not permission to run.
                # The caller may have given up while this thread sat in the
                # queue, and it was told the operation had not happened; an
                # abandoned call releases the lock and exits WITHOUT
                # touching Word.
                if not _elect("started"):
                    return
                result["value"] = fn()
        except BaseException as exc:  # noqa: BLE001 — re-raised in caller
            result["error"] = exc
        finally:
            done.set()

    t = threading.Thread(target=worker, daemon=True, name=f"ks4w-{name}")
    t.start()
    if done.wait(timeout):
        if "error" in result:
            raise result["error"]
        return result["value"]
    if _elect("abandoned"):
        # The worker had NOT started fn when this decision was taken, so it
        # never will: it is still queued on the lock, and when it reaches
        # the front it loses this same decision, releases the lock and
        # exits. Nothing ran and nothing was changed.
        if "error" in result:
            # It failed before it could reach the decision at all, so its
            # own error is the honest answer, not a queue-contention one.
            raise result["error"]
        snap = _serial.lock_snapshot()
        running = (snap.get("current_op") or {}).get(
            "name", "another COM operation"
        )
        raise WordBusy(
            f"{name} waited {timeout:.0f}s for the COM serialization lock "
            f"({running} is still running); retry when it finishes. "
            "com_word_status reports the running operation."
            " The queued call was abandoned before its operation ran; it "
            "made no document changes. It is safe to retry after "
            "com_word_status reports no COM operation in progress."
        )
    # The worker won the decision, so fn IS running, or has already
    # finished. This is no longer queue contention and the caller must not
    # be told to simply retry. The deadline has passed but the work has NOT
    # been cancelled, so the grace wait is a real second chance rather than
    # a formality.
    done.wait(TIMEOUT_GRACE_SECONDS)

    def _completed():
        """The worker's own answer, when it has one."""
        if not done.is_set():
            return False
        if "error" in result:
            raise result["error"]
        return "value" in result

    if _completed():
        return result["value"]

    dialogs_seen = []
    with contextlib.suppress(Exception):
        from . import dialogs as _dialogs

        dialogs_seen = _dialogs.pending_dialogs()

    if _completed():
        return result["value"]
    # Evidence is read HERE, at the moment the refusal is built: a worker
    # that finished and released its instance meanwhile has cleared it.
    ours = set(
        _INVISIBLE_PIDS.get(worker_tid[0] if worker_tid else None) or set()
    )
    # STATE-NEUTRAL: an empty record is not evidence that no Word was
    # started. The table may have been unreadable, two processes may have
    # appeared at once, or the worker may never have reached DispatchEx.
    if ours:
        pids = ", ".join(str(p) for p in sorted(ours))
        detail = (
            f" (Word process pid {pids} was not running when this "
            "call began; this call did not force-end it, and its current "
            "state was not re-checked)"
        )
    else:
        detail = (
            " (no newly started Word process could be identified; no "
            "process was force-ended)"
        )
    if dialogs_seen:
        titles = ", ".join(
            d.get("title") or d.get("class", "?") for d in dialogs_seen[:3]
        )
        detail += f". Word has a dialog open: {titles}"
    # NOTHING WAS CANCELLED. The old sentence said the operation "was
    # aborted", which stopped being true the moment nothing was ended: a
    # caller told it was aborted retries on top of a live operation (R5-1).
    raise WordBlocked(
        f"Word did not answer within {timeout:.0f} s. The operation "
        "was NOT cancelled: it may still be running and may still finish "
        "and save its output. Do not retry yet: call com_word_status and "
        "wait until it reports no COM operation in progress, then check the "
        "output file before repeating the call."
        + detail
    )


def _bounded_op(name: str, default: float):
    """Public-function wrapper: adds a tunable timeout parameter and runs
    the body via _run_bounded (which serializes). Marks the function for
    the entry-point coverage audit."""

    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*args, timeout: float = default, **kwargs):
            try:
                timeout = float(timeout)
            except (TypeError, ValueError):
                raise WordMcpError(
                    f"timeout must be a number of seconds, got {timeout!r}"
                ) from None
            if not 5 <= timeout <= 3600:
                raise WordMcpError(
                    "timeout must be between 5 and 3600 seconds"
                )
            return _run_bounded(name, timeout, lambda: fn(*args, **kwargs))

        wrapper._com_serialized = name
        wrapper._com_bounded = default
        return wrapper

    return deco


def _open_in_running_word(path) -> bool:
    """Is this file open in ANY running Word instance? (primary instance
    scan + ROT file-moniker scan; read-only, best-effort). Guards the
    invisible-open name-collision cascade from the stress report: opening
    a second copy of an open document triggers Word's same-name dialog."""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return False
    pythoncom.CoInitialize()
    target = str(Path(path).resolve()).lower()
    try:
        with contextlib.suppress(Exception):
            app = win32com.client.GetActiveObject("Word.Application")
            for doc in app.Documents:
                if doc.FullName.lower() == target:
                    return True
        with contextlib.suppress(Exception):
            for _, _moniker, name in _rot.iter_entries(pythoncom):
                if name == target:
                    return True
        return False
    finally:
        pythoncom.CoUninitialize()


def _open(app, path: Path, *, read_only: bool):
    """Open a document in this invisible Word.

    Routed through callargs because a keyword argument here does not
    survive the trip: pywin32 drops it when it is bound late, and this
    call is exactly where that mattered, since read_only silently became
    false on every machine with no makepy cache.
    """
    return _args.call(
        app.Documents.Open, "Documents.Open",
        FileName=str(path.resolve()),
        ConfirmConversions=False,
        ReadOnly=read_only,
        AddToRecentFiles=False,
        OpenAndRepair=False,
    )


def word_status() -> dict:
    """Is Word installed / running, and which documents are open in the USER's
    Word instance (best effort via the running-object table)? Uses a bounded
    lock acquire so a long-running COM operation cannot hang the status
    probe; when the lock is busy the probe is skipped and says so."""
    out = {"word_running": False, "open_documents": []}
    if not _serial.acquire(timeout=2.0):
        out["note"] = (
            "COM serialization lock held by a running operation; process "
            "state not probed"
        )
        return out
    try:
        try:
            import pythoncom
            import win32com.client

            pythoncom.CoInitialize()
            try:
                app = win32com.client.GetActiveObject("Word.Application")
                out["word_running"] = True
                for doc in app.Documents:
                    out["open_documents"].append(doc.FullName)
            except Exception:
                pass
            finally:
                pythoncom.CoUninitialize()
        except ImportError:
            out["error"] = "pywin32 not installed"
        return out
    finally:
        _serial.release()


word_status._com_serialized = "com_word_status"  # bounded try-acquire form


# wdStoryType -> the name reported in fields_by_story.
_STORY_NAMES = {
    1: "body",
    2: "footnotes",
    3: "endnotes",
    4: "comments",
    5: "text_frames",
    6: "even_pages_header",
    7: "primary_header",
    8: "even_pages_footer",
    9: "primary_footer",
    10: "first_page_header",
    11: "first_page_footer",
}


def _merge_counts(a: dict, b: dict) -> dict:
    """Per-story field counts across two update passes over the SAME fields:
    the larger count per story, never the sum."""
    return {k: max(a.get(k, 0), b.get(k, 0)) for k in set(a) | set(b)}


def _update_every_story(doc) -> dict:
    """Update fields in EVERY story, walking NextStoryRange so each
    section's header and footer is reached and not just the first story of
    each type (punchlist #862). Returns per-story field counts."""
    counts: dict[str, int] = {}
    for story in doc.StoryRanges:
        rng = story
        hops = 0
        while rng is not None and hops < 500:
            hops += 1
            try:
                story_type = int(rng.StoryType)
            except Exception:
                story_type = 0
            name = _STORY_NAMES.get(story_type, f"story_{story_type}")
            try:
                n = int(rng.Fields.Count)
                if n:
                    rng.Fields.Update()
                    counts[name] = counts.get(name, 0) + n
            except Exception:
                pass  # a story that refuses enumeration blocks nothing
            try:
                rng = rng.NextStoryRange
            except Exception:
                rng = None
    return counts


@_bounded_op("com_refresh_fields", default=300.0)
def refresh_fields(path: str) -> dict:
    """Open invisibly, update every field in every story (body, headers,
    footers, notes, text frames), update TOC-family tables explicitly,
    save, close. This is the immediate alternative to the update-on-open
    flag."""
    path = check_path(path, "refresh fields")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    # Word silently refuses to update fields in a protected document.
    from ..core.package import DocxPackage
    from ..ops.protection import get_protection

    state = get_protection(DocxPackage(p))
    if state.get("protected"):
        raise WordMcpError(
            f"document is protected (edit={state.get('edit')}); Word will not "
            "update fields under protection; remove_document_protection "
            "first, refresh, then re-protect (build → refresh → protect)"
        )
    with _word() as app:
        doc = _open(app, p, read_only=False)
        try:
            by_story = _update_every_story(doc)
            for i in range(1, doc.TablesOfContents.Count + 1):
                doc.TablesOfContents(i).Update()
            for i in range(1, doc.TablesOfFigures.Count + 1):
                doc.TablesOfFigures(i).Update()
            toc_count = doc.TablesOfContents.Count
            tof_count = doc.TablesOfFigures.Count
            # Page-dependent fields (PAGE, NUMPAGES, PAGEREF) resolve off the
            # layout, so repaginate and update the page furniture again.
            with contextlib.suppress(Exception):
                doc.Repaginate()
            by_story = _merge_counts(by_story, _update_every_story(doc))
            doc.Save()
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    return {
        "fields_refreshed": True,
        "tocs_updated": toc_count,
        "tables_of_figures_updated": tof_count,
        "fields_by_story": by_story,
    }


@_bounded_op("com_export_pdf", default=300.0)
def export_pdf(path: str, pdf_path: str | None = None) -> dict:
    """Export to PDF via Word (highest fidelity available on this machine)."""
    path = check_path(path, "PDF export source")
    if pdf_path:
        pdf_path = check_path(pdf_path, "PDF export output")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    out = Path(pdf_path) if pdf_path else p.with_suffix(".pdf")
    with _word() as app:
        doc = _open(app, p, read_only=True)
        try:
            _args.call(
                doc.SaveAs2, "Document.SaveAs2",
                FileName=str(out.resolve()), FileFormat=_WD_FORMAT_PDF,
            )
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but no PDF was produced")
    return {"pdf": str(out), "bytes": out.stat().st_size}


@_bounded_op("com_multi_document", default=300.0)
def compare_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
    *,
    author: str = "word-mcp compare",
) -> dict:
    """Word-native compare: produces a NEW document where every difference
    between original and revised appears as a tracked change. Neither input is
    modified. Perfect for diffing two DTG versions of a draft."""
    original_path = check_path(original_path, "compare original")
    revised_path = check_path(revised_path, "compare revised")
    if output_path:
        output_path = check_path(output_path, "compare output")
    orig = Path(original_path)
    rev = Path(revised_path)
    for p in (orig, rev):
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = (
        Path(output_path)
        if output_path
        else rev.with_name(f"{rev.stem}_COMPARE{rev.suffix}")
    )
    with _word() as app:
        doc_a = _open(app, orig, read_only=True)
        doc_b = _open(app, rev, read_only=True)
        try:
            result = _args.call(
                app.CompareDocuments, "Application.CompareDocuments",
                OriginalDocument=doc_a,
                RevisedDocument=doc_b,
                Destination=2,  # wdCompareDestinationNew
                Granularity=1,  # wdGranularityWordLevel
                CompareFormatting=True,
                CompareTables=True,
                CompareFootnotes=True,
                CompareHeaders=True,
                CompareFields=False,
                RevisedAuthor=author,
            )
            result.SaveAs2(str(out.resolve()))
            result.Close(_WD_DO_NOT_SAVE)
        finally:
            doc_a.Close(_WD_DO_NOT_SAVE)
            doc_b.Close(_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but produced no output")
    # Summarize what changed, using our own revision reader.
    from ..core.package import DocxPackage
    from ..ops.read import revision_summary

    summary = revision_summary(DocxPackage(out))
    return {"comparison": str(out), "revisions": summary}


@_bounded_op("com_validate_opens_clean", default=60.0)
def validate_opens_clean(path: str) -> dict:
    """Open in invisible Word and confirm no repair/recovery path triggers.

    Name-collision guard (stress report bug 6): when the file is ALREADY
    open in a running Word, no second invisible copy is opened — the check
    routes to the open document (which by definition opened clean) and
    says so, instead of triggering Word's same-name dialog cascade."""
    path = check_path(path, "validate opens clean")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        pythoncom, app, doc = _find_open_document(str(p))
        try:
            paragraphs = doc.Paragraphs.Count
            words = int(doc.ComputeStatistics(0))  # wdStatisticWords
        finally:
            pythoncom.CoUninitialize()
        return {
            "opens_clean": True,
            "paragraphs": paragraphs,
            "words": words,
            "note": (
                "document is OPEN in Word: the verdict reflects the open "
                "copy (unsaved changes included); no second invisible copy "
                "was opened, avoiding Word's same-name dialog"
            ),
        }
    with _word() as app:
        try:
            doc = _open(app, p, read_only=True)
        except Exception as exc:
            return {"opens_clean": False, "error": str(exc)}
        try:
            paragraphs = doc.Paragraphs.Count
            # NEVER doc.Words.Count: it counts punctuation runs and paragraph
            # marks as "words" (~18% high on real prose — L1, 2026-08-28).
            # ComputeStatistics(wdStatisticWords) is the number Word's own
            # status bar shows.
            words = int(doc.ComputeStatistics(0))  # wdStatisticWords
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    return {"opens_clean": True, "paragraphs": paragraphs, "words": words}


@_bounded_op("com_multi_document", default=300.0)
def merge_documents(
    paths: list[str],
    output_path: str,
    *,
    section_break_between: bool = True,
) -> dict:
    """Merge documents in order into one file via Word's InsertFile (full
    fidelity: styles, footnotes, numbering survive). Chapters get a next-page
    section break between them so per-chapter headers/numbering stay possible."""
    if len(paths) < 2:
        raise WordMcpError("give at least two documents to merge")
    paths = [check_path(p, "merge input") for p in paths]
    output_path = check_path(output_path, "merge output")
    srcs = [Path(p) for p in paths]
    for p in srcs:
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = Path(output_path)
    with _word() as app:
        doc = app.Documents.Add()
        try:
            rng = doc.Range(0, 0)
            for i, src in enumerate(srcs):
                rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                if i > 0:
                    if section_break_between:
                        rng.InsertBreak(2)  # wdSectionBreakNextPage
                    else:
                        rng.InsertBreak(7)  # wdPageBreak
                    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                rng.InsertFile(str(src.resolve()))
            doc.SaveAs2(str(out.resolve()))
            paragraphs = doc.Paragraphs.Count
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word reported success but produced no output")
    # Word's InsertFile does not carry the customXml bibliography store;
    # union the inputs' sources into the merged output ourselves.
    sources_merged = _merge_bibliography_stores(out, srcs)
    result = {
        "merged": [str(p) for p in srcs],
        "output": str(out),
        "paragraphs": paragraphs,
    }
    if sources_merged:
        result["bibliography_sources_carried"] = sources_merged
    return result


def _merge_bibliography_stores(out_path, src_paths) -> int:
    import copy as _copy

    from ..core.package import DocxPackage
    from ..ops import bibliography as _bib

    collected = []
    seen_tags = set()
    for sp in src_paths:
        try:
            spkg = DocxPackage(sp)
        except Exception:
            continue
        store = _bib._find_store(spkg)
        if not store:
            continue
        for s in spkg.root(store).findall(_bib._bq("Source")):
            tag = s.findtext(_bib._bq("Tag")) or ""
            if tag and tag not in seen_tags:
                seen_tags.add(tag)
                collected.append(_copy.deepcopy(s))
    if not collected:
        return 0
    opkg = DocxPackage(out_path)
    part = _bib._ensure_store(opkg)
    root = opkg.root(part)
    existing = {
        s.findtext(_bib._bq("Tag")) for s in root.findall(_bib._bq("Source"))
    }
    added = 0
    for s in collected:
        if s.findtext(_bib._bq("Tag")) not in existing:
            root.append(s)
            added += 1
    if added:
        opkg.mark_dirty(part)
        opkg.save(do_backup=False)
    return added


@_bounded_op("com_multi_document", default=300.0)
def combine_documents(
    original_path: str,
    revised_path: str,
    output_path: str | None = None,
) -> dict:
    """Combine two documents' TRACKED CHANGES into one (Word's Combine — for
    merging two reviewers' edits of the same draft; both authors' revisions
    survive as separate attributions). Distinct from compare, which diffs
    content."""
    original_path = check_path(original_path, "combine original")
    revised_path = check_path(revised_path, "combine revised")
    if output_path:
        output_path = check_path(output_path, "combine output")
    orig = Path(original_path)
    rev = Path(revised_path)
    for p in (orig, rev):
        if not p.exists():
            raise DocumentNotFound(f"no file at {p}")
    out = (
        Path(output_path)
        if output_path
        else rev.with_name(f"{rev.stem}_COMBINED{rev.suffix}")
    )
    with _word() as app:
        doc_a = _open(app, orig, read_only=True)
        doc_b = _open(app, rev, read_only=True)
        try:
            result = _args.call(
                app.MergeDocuments, "Application.MergeDocuments",
                OriginalDocument=doc_a,
                RevisedDocument=doc_b,
                Destination=2,  # wdCompareDestinationNew
            )
            result.SaveAs2(str(out.resolve()))
            result.Close(_WD_DO_NOT_SAVE)
        finally:
            doc_a.Close(_WD_DO_NOT_SAVE)
            doc_b.Close(_WD_DO_NOT_SAVE)
    from ..core.package import DocxPackage
    from ..ops.read import revision_summary

    summary = revision_summary(DocxPackage(out))
    return {"combined": str(out), "revisions": summary}


def _find_open_document(path: str):
    """(pythoncom, app, doc) for a document open in ANY running interactive
    Word instance, else raise.

    Multi-instance aware (WS-L, 2026-08-28): GetActiveObject returns only
    whichever instance registered first — and that instance can be busy
    (modal dialog) or simply not the one holding the document. Open
    documents register their full path as a ROT file moniker, so the
    fallback binds the document directly, exactly like the live layer's
    _find_doc_via_rot."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    path = check_path(path, "find open document")
    target = str(Path(path).resolve()).lower()
    word_seen = False
    try:
        app = win32com.client.GetActiveObject("Word.Application")
        word_seen = True
        for doc in app.Documents:
            if doc.FullName.lower() == target:
                return pythoncom, app, doc
    except Exception:
        # not running, busy, or dying — the ROT scan below still works for
        # documents held by OTHER instances
        pass
    with contextlib.suppress(Exception):
        for rot, moniker, name in _rot.iter_entries(pythoncom):
            if name != target:
                continue
            with contextlib.suppress(Exception):
                doc = _rot.bind_document(
                    pythoncom, win32com.client, rot, moniker, name
                )
                if doc is not None:
                    return pythoncom, doc.Application, doc
            word_seen = True  # moniker exists but binding failed
    pythoncom.CoUninitialize()
    if not word_seen:
        raise WordMcpError("Word is not running")
    raise DocumentNotFound(
        f"{Path(path).name} is not open in any running Word instance"
    )


@contextlib.contextmanager
def _alerts_suppressed(app):
    """DisplayAlerts off around a call into the USER's Word, restored
    after — contention then raises a catchable com_error instead of
    freezing Word behind a modal dialog (stress report bug 3)."""
    prev = None
    with contextlib.suppress(Exception):
        prev = app.DisplayAlerts
        app.DisplayAlerts = _WD_ALERTS_NONE
    try:
        yield
    finally:
        if prev is not None:
            with contextlib.suppress(Exception):
                app.DisplayAlerts = prev


def _retry_word_call(fn, *, attempts: int = 5, first_delay: float = 0.5):
    """Bounded retry with exponential backoff for save/close contention
    (Word's write-to-temp-then-rename save can transiently collide).
    Total worst-case wait: 0.5+1+2+4 = 7.5s."""
    delay = first_delay
    last_exc = None
    for attempt in range(attempts):
        try:
            return fn(), attempt
        except Exception as exc:
            last_exc = exc
            if attempt < attempts - 1:
                time.sleep(delay)
                delay *= 2
    raise WordBusy(
        f"Word refused the operation {attempts} times (file contention or "
        f"a permission error; last error: {last_exc}). Alerts were "
        "suppressed, so no dialog is pending; wait a moment and retry, "
        "or save manually in Word."
    ) from last_exc


@_serial.serialized("com_save_document")
def save_open_document(path: str) -> dict:
    """Tell the USER's running Word to save a document it has open, so
    file-based tools can read the current state. Runs under the COM
    serialization lock with alerts suppressed and a bounded retry —
    concurrent saves were the report's 11/11 dialog-blocked failure."""
    pythoncom, app, doc = _find_open_document(path)
    try:
        with _alerts_suppressed(app):
            _, retries = _retry_word_call(doc.Save)
        out = {"saved": doc.FullName}
        if retries:
            out["retries"] = retries
        return out
    finally:
        pythoncom.CoUninitialize()


@_serial.serialized("com_save_document")
def close_open_document(path: str, *, save: bool = True) -> dict:
    """Tell the USER's running Word to close a document (saving by default),
    releasing the lock so file-based tools can edit it. Serialized, alerts
    suppressed, bounded retry (same contention path as save)."""
    pythoncom, app, doc = _find_open_document(path)
    try:
        with _alerts_suppressed(app):
            _, retries = _retry_word_call(
                lambda: doc.Close(
                    _WD_SAVE if save else _WD_DO_NOT_SAVE
                ),
                attempts=3,
            )
        out = {"closed": str(Path(path).resolve()), "saved": save}
        if retries:
            out["retries"] = retries
        return out
    finally:
        pythoncom.CoUninitialize()


@_bounded_op("com_proofing_errors", default=60.0)
def proofing_errors(path: str, *, limit: int = 100) -> dict:
    """Word's own spelling and grammar error lists (with context). Slower on
    long documents; capped by limit per category.

    Name-collision guard (stress report bug 6): proofing needs a WRITE
    open in the invisible instance (the proofed flags are reset), and
    opening a second writable copy of an open document triggers Word's
    same-name dialog cascade — refused up front instead."""
    path = check_path(path, "proofing errors")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        raise WordBusy(
            f"{p.name} is open in Word right now; proofing opens a second "
            "writable copy invisibly, which triggers Word's same-name "
            "dialog cascade. Save and close it first (com_save_document "
            "with close=true), or proof a copy (copy_document)."
        )
    with _word() as app:
        doc = _open(app, p, read_only=False)
        try:
            # Word marks documents as already-proofed; without resetting these
            # flags SpellingErrors is silently empty (verified 2026-08-28).
            doc.SpellingChecked = False
            doc.GrammarChecked = False
            spelling = []
            for err in doc.SpellingErrors:
                if len(spelling) >= limit:
                    break
                spelling.append(err.Text)
            grammar = []
            for err in doc.GrammaticalErrors:
                if len(grammar) >= limit:
                    break
                text = err.Text
                grammar.append(text[:120])
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    return {
        "spelling_errors": spelling,
        "spelling_truncated": len(spelling) >= limit,
        "grammar_flagged_ranges": grammar,
        "grammar_truncated": len(grammar) >= limit,
        "note": (
            "Word's proofing engine; includes proper nouns and citations it "
            "does not recognize; review, do not auto-fix"
        ),
    }


@_bounded_op("com_readability_statistics", default=60.0)
def readability_statistics(path: str) -> dict:
    """Word's readability statistics (Flesch Reading Ease, grade level, word
    and sentence counts...). Refuses when the file is open in Word (the
    invisible read-only copy would compute stats for the stale saved state
    while risking instance contention); save and close first, or run on a
    copy."""
    path = check_path(path, "readability statistics")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if _open_in_running_word(p):
        raise WordBusy(
            f"{p.name} is open in Word right now; close it first "
            "(com_save_document with close=true) or run the statistics on "
            "a copy (copy_document); a second invisible copy of an open "
            "document risks Word's same-name dialog cascade."
        )
    with _word() as app:
        doc = _open(app, p, read_only=True)
        try:
            stats = {}
            for st in doc.Content.ReadabilityStatistics:
                stats[st.Name] = st.Value
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    return {"readability": stats}


@_bounded_op("com_save_document", default=300.0)
def save_with_password(
    path: str, output_path: str | None = None, *, password: str
) -> dict:
    """Save a copy encrypted with an open-password (real encryption, unlike
    document protection). Word will demand the password to open the copy."""
    path = check_path(path, "encrypt source")
    if output_path:
        output_path = check_path(output_path, "encrypt output")
    p = Path(path)
    if not p.exists():
        raise DocumentNotFound(f"no file at {path}")
    if not password:
        raise WordMcpError("password must be non-empty")
    out = (
        Path(output_path)
        if output_path
        else p.with_name(f"{p.stem}_PROTECTED{p.suffix}")
    )
    with _word() as app:
        # Read-write on purpose: the Password PROPERTY is what actually
        # encrypts, and Word refuses to set it on a read-only document.
        # The source is never saved, only SaveAs2'd to the output path.
        doc = _open(app, p, read_only=False)
        try:
            doc.Password = password
            _args.call(
                doc.SaveAs2, "Document.SaveAs2",
                FileName=str(out.resolve()),
                FileFormat=_WD_FORMAT_DOCX_DEFAULT,
            )
        finally:
            doc.Close(_WD_DO_NOT_SAVE)
    if not out.exists():
        raise WordMcpError("Word produced no output")
    if not _is_encrypted(out):
        # Never report an encryption that did not happen. An encrypted
        # OOXML file is an OLE compound file; a plain one is a ZIP, so the
        # first four bytes settle it without a password prompt.
        out.unlink(missing_ok=True)
        raise WordMcpError(
            f"Word saved {out.name} without applying the password, so the "
            f"copy was NOT encrypted and has been removed rather than left "
            f"looking protected. Encrypt it in Word (File, Info, Protect "
            f"Document, Encrypt with Password); set_document_protection is "
            f"a different thing and does not encrypt."
        )
    return {
        "encrypted_copy": str(out),
        "verified": "the copy is an encrypted OLE container, not a plain zip",
        "note": "keep the password safe; there is no recovery",
    }


def _is_encrypted(path: Path) -> bool:
    """Is this file an encrypted Office document?

    Encryption rewraps the .docx zip inside an OLE compound file, so the
    magic bytes answer it: D0CF11E0 for encrypted, PK for a plain package.
    Cheap, offline, and no password prompt.
    """
    try:
        with open(path, "rb") as fh:
            return fh.read(4) == _OLE_MAGIC
    except OSError:
        return False


def zombie_check() -> dict:
    """Count WINWORD.EXE processes (diagnostic for leak detection)."""
    import subprocess

    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
        timeout=30,
        creationflags=_NO_WINDOW,
    )
    lines = [
        ln for ln in result.stdout.splitlines() if "WINWORD.EXE" in ln.upper()
    ]
    return {"winword_processes": len(lines)}
