"""M1: a live call or a passwordless save/close that the server did not start
must never run later.

THE DEFECT (Codex release review of 2.2.1 candidate bf54af4, X-20260924-011).
``live.live_session`` took the process-wide COM lock with a bare
``COM_LOCK.acquire()`` and then the cross-process live lock, whose local
mutex had no bound either and whose lockfile wait was 120 s.
``bridge.save_open_document`` and ``close_open_document`` took the same
unbounded lock through ``@_serial.serialized``. Behind a hung invisible Word,
which 2.2.1 no longer force-ends, such a call sat in an invisible queue. A
client whose own timeout expired retried, the retry queued behind it, and
when the holder let go BOTH ran against the user's open document. Codex's
exact-wheel probe got ``['original', 'retry']``.

THE CONTRACT. These routes now acquire every lock layer they need without
waiting: when the process COM lock, the local cross-process mutex or a live
foreign lockfile is already held, the call is refused with
``CallNotStarted`` (APP_BUSY) BEFORE the body runs, and nothing is left
queued that could run it later. A finite wait is not enough: the server
cannot know how long the client will wait, so any queue can outlive it.

Every case below is fail-first against bf54af4: there, each route either
blocked behind the holder and then ran the body once it released (process
lock, local mutex, foreign lockfile on the live route), or ran at once with
no cross-process check at all (save and close against the cross-process
layers). Each case holds one layer, fires an ORIGINAL call and a RETRY the
way a client that gave up would, releases the holder, and proves zero
mutations. A call made after the release then runs exactly once, so the
refusal wedges nothing.

Pure Python: the live route runs through a faked COM layer and the bridge
routes through a faked document lookup (the shape of Codex's probe), so
these run on CI's Linux and Windows runners alike.
"""

from __future__ import annotations

import contextlib
import json
import os
import subprocess
import sys
import textwrap
import threading
import time
from pathlib import Path

import pytest

import word_mcp.server as srv
from word_mcp import envelope
from word_mcp.com import bridge, live, xproc
from word_mcp.com import serial as com_serial
from word_mcp.core.errors import CallNotStarted, WordBusy

#: The refusal text as it must read, pinned literally: copy slot P1-P-3,
#: landed from Codex X-20260924-013. The constant in com/serial.py and
#: this pin change together.
PINNED_BUSY_TEXT = (
    "Word is already in use by {holder}. This call did not start or "
    "change the document, and nothing from it is queued to run later. "
    "Wait until com_word_status reports no COM operation in progress, "
    "then retry."
)

#: How long the simulated client waits before it gives up on a call.
CLIENT_PATIENCE = 0.3


# ------------------------------------------------------------ fixtures


@pytest.fixture()
def lock_dir(tmp_path, monkeypatch):
    d = tmp_path / "locks"
    monkeypatch.setenv(xproc._LOCK_DIR_ENV, str(d))
    return d


@pytest.fixture()
def mutations():
    return []


class _PW:
    class com_error(Exception):
        pass


class _PC:
    @staticmethod
    def CoInitialize():
        return None

    @staticmethod
    def CoUninitialize():
        return None


class _FakeDoc:
    Saved = True
    AutoSaveOn = False

    def __init__(self, label: str, mutations: list):
        self.FullName = label
        self._mutations = mutations

    def Save(self):
        self._mutations.append(self.FullName)

    def Close(self, _how):
        self._mutations.append(self.FullName)


@pytest.fixture()
def fake_word(monkeypatch, mutations):
    """Word discovery replaced, every lock left exactly as shipped."""

    @contextlib.contextmanager
    def _no_undo_group(app, name, doc=None):
        yield False

    monkeypatch.setattr(live, "_com_modules", lambda: (_PC, _PW, object()))
    monkeypatch.setattr(live, "_ensure_com", lambda pythoncom: None)
    monkeypatch.setattr(live, "_attach_app", lambda w32, pc: object())
    monkeypatch.setattr(
        live, "_resolve_document",
        lambda pc, pw, w32, app, path: (app, _FakeDoc(path, mutations)),
    )
    monkeypatch.setattr(live, "probe_ready", lambda *a, **k: None)
    monkeypatch.setattr(live, "_suppress_alerts_owned", lambda g, app: False)
    monkeypatch.setattr(live, "undo_group", _no_undo_group)
    monkeypatch.setattr(live, "probe_read_only", lambda session, doc: None)
    monkeypatch.setattr(live, "_check_protection", lambda doc, p, s: None)
    monkeypatch.setattr(
        bridge, "_find_open_document",
        lambda path: (_PC, None, _FakeDoc(path, mutations)),
    )
    return mutations


# -------------------------------------------------------------- routes


def _live_edit(label: str) -> dict:
    """A live mutation: the body is the edit, and it records itself."""
    def body(session):
        session.doc._mutations.append(label)
        return {"edited": label}

    return live.run_live(label, "m1 probe edit", body)


def _save(label: str) -> dict:
    return bridge.save_open_document(label)


def _close(label: str) -> dict:
    return bridge.close_open_document(label)


ROUTES = {"live": _live_edit, "save": _save, "close": _close}


# ------------------------------------------------------------- holders


class _ProcessLockHolder:
    """The hung invisible Word: another thread inside a COM operation."""

    def __init__(self):
        self._held = threading.Event()
        self._release = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def _run(self):
        with com_serial.com_operation("com_export_pdf"):
            self._held.set()
            self._release.wait(30)

    def start(self):
        self._thread.start()
        assert self._held.wait(5), "holder never took the COM lock"

    def release(self):
        self._release.set()
        self._thread.join(5)
        assert not self._thread.is_alive()


class _LocalMutexHolder:
    """Another thread in this server holding ONLY the cross-process local
    mutex (the process COM lock stays free), so this layer is proven on its
    own rather than behind the process lock."""

    def __init__(self):
        self._held = threading.Event()
        self._release = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def _run(self):
        mutex = xproc._mutex_for(xproc.APP_SCOPE)
        mutex.acquire()
        try:
            self._held.set()
            self._release.wait(30)
        finally:
            mutex.release()

    def start(self):
        self._thread.start()
        assert self._held.wait(5), "holder never took the local mutex"
        assert com_serial.lock_snapshot()["held"] is False

    def release(self):
        self._release.set()
        self._thread.join(5)
        assert not self._thread.is_alive()


_FOREIGN = textwrap.dedent(
    """
    import os, sys
    os.environ["KS4W_LIVE_LOCK_DIR"] = sys.argv[1]
    sys.path.insert(0, sys.argv[2])
    from word_mcp.com import xproc
    with xproc.cross_process_lock("live:foreign edit", wait=5.0):
        print("HELD", flush=True)
        sys.stdin.readline()
    print("RELEASED", flush=True)
    """
)


class _ForeignLockfileHolder:
    """A second kitchensink4word server PROCESS holding the lockfile."""

    def __init__(self, lock_dir: Path):
        self._lock_dir = lock_dir
        self._proc = None

    def start(self):
        src = str(Path(srv.__file__).resolve().parents[1])
        self._proc = subprocess.Popen(
            [sys.executable, "-c", _FOREIGN, str(self._lock_dir), src],
            stdin=subprocess.PIPE, stdout=subprocess.PIPE,
            stderr=subprocess.PIPE, text=True,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        line = self._proc.stdout.readline()
        if line.strip() != "HELD":
            self._proc.kill()
            raise AssertionError(
                f"foreign holder failed: {self._proc.stderr.read()[:2000]}"
            )
        assert com_serial.lock_snapshot()["held"] is False
        held = xproc.holder_info()
        assert held is not None and held["ours"] is False

    def release(self):
        try:
            self._proc.stdin.write("go\n")
            self._proc.stdin.flush()
            out, _err = self._proc.communicate(timeout=30)
            assert "RELEASED" in out
        finally:
            if self._proc.poll() is None:
                self._proc.kill()
                self._proc.wait(10)
        assert xproc.holder_info() is None


def _holder(layer: str, lock_dir: Path):
    if layer == "process_lock":
        return _ProcessLockHolder()
    if layer == "local_xproc_mutex":
        return _LocalMutexHolder()
    if layer == "foreign_lockfile":
        return _ForeignLockfileHolder(lock_dir)
    raise AssertionError(layer)


# -------------------------------------------------------------- driver


def _fire(route, label: str, outcomes: dict) -> threading.Thread:
    def run():
        try:
            outcomes[label] = ("returned", route(label))
        except BaseException as exc:  # noqa: BLE001 - recorded for asserts
            outcomes[label] = ("raised", exc)

    t = threading.Thread(target=run, daemon=True, name=f"m1-{label}")
    t.start()
    return t


# --------------------------------------------------------------- tests


@pytest.mark.parametrize("layer", [
    "process_lock", "local_xproc_mutex", "foreign_lockfile",
])
@pytest.mark.parametrize("route_name", ["live", "save", "close"])
def test_a_refused_call_never_runs_after_the_holder_releases(
    route_name, layer, fake_word, lock_dir
):
    """Original plus retry against a held layer: both refused at once,
    nothing left behind, zero mutations after the release, and a call made
    after the release runs exactly once."""
    mutations = fake_word
    route = ROUTES[route_name]
    holder = _holder(layer, lock_dir)
    holder.start()
    outcomes: dict = {}
    threads = []
    try:
        for label in ("original", "retry"):
            t = _fire(route, label, outcomes)
            threads.append(t)
            # the client stops waiting after CLIENT_PATIENCE
            t.join(CLIENT_PATIENCE)
        # a call still waiting here is a queued body that can run after its
        # client has given up
        still_waiting = [t.name for t in threads if t.is_alive()]
        ran_while_held = list(mutations)
    finally:
        holder.release()
        for t in threads:
            t.join(10)
    # give anything that was secretly queued every chance to run
    time.sleep(0.3)
    assert mutations == [], (
        f"{route_name} behind {layer}: {mutations} ran although the client "
        f"had stopped waiting (still queued at give-up: {still_waiting}; "
        f"ran while held: {ran_while_held})"
    )
    assert not still_waiting, (
        f"{route_name} behind {layer}: {still_waiting} still queued after "
        "the client gave up"
    )
    for label in ("original", "retry"):
        kind, value = outcomes[label]
        assert kind == "raised", f"{label} returned {value!r}"
        assert isinstance(value, CallNotStarted), repr(value)
        assert isinstance(value, WordBusy)
        assert envelope.classify(value) == "APP_BUSY"
    # nothing of ours is left holding anything
    assert com_serial.lock_snapshot()["held"] is False
    assert xproc.holder_info() is None
    # and the route is not wedged: one fresh call runs exactly once
    route("after")
    assert mutations == ["after"]


def test_the_refusal_text_is_pinned(fake_word, lock_dir):
    holder = _ProcessLockHolder()
    holder.start()
    try:
        with pytest.raises(CallNotStarted) as excinfo:
            _live_edit("pinned")
    finally:
        holder.release()
    assert str(excinfo.value) == PINNED_BUSY_TEXT.format(
        holder="com_export_pdf"
    )
    assert com_serial.BUSY_NOT_STARTED == PINNED_BUSY_TEXT
    assert fake_word == []


def test_foreign_lockfile_refusal_names_the_other_process(fake_word, lock_dir):
    holder = _ForeignLockfileHolder(lock_dir)
    holder.start()
    held = xproc.holder_info()
    try:
        with pytest.raises(CallNotStarted) as excinfo:
            _save("named")
    finally:
        holder.release()
    assert f"PID {held['pid']}, running 'live:foreign edit'" in str(
        excinfo.value
    )
    assert fake_word == []


def test_same_thread_reentry_is_not_refused(fake_word, lock_dir):
    """The no-wait acquisition must not refuse its own thread: the COM lock
    and the local mutex are re-entrant, and the lockfile is ours."""
    with com_serial.com_operation("outer-op", refuse_if_busy=True):
        with xproc.cross_process_lock("outer", refuse_if_held=True) as owns:
            assert owns is True
            assert _live_edit("nested")["edited"] == "nested"
            assert _save("nested-save")["saved"] == "nested-save"
    assert fake_word == ["nested", "nested-save"]


def test_a_stale_lockfile_is_reclaimed_not_refused(fake_word, lock_dir):
    """Refusing without waiting applies to a LIVE holder only. A lockfile
    whose writer is gone is broken exactly as before."""
    lock_dir.mkdir(parents=True, exist_ok=True)
    (lock_dir / f"{xproc.APP_SCOPE}.lock").write_text(json.dumps({
        "pid": 999_999_999, "token": "gone", "time": time.time(),
        "holder": "live:crashed", "host": xproc._local_host(),
    }), encoding="utf-8")
    assert _save("reclaimed")["saved"] == "reclaimed"
    assert fake_word == ["reclaimed"]
    assert not (lock_dir / f"{xproc.APP_SCOPE}.lock").exists()


def test_the_blocking_default_is_unchanged_for_other_callers(lock_dir):
    """cross_process_lock and com_operation keep their waiting behavior for
    callers that did not opt in (live_repair, the screen-repair daemon)."""
    holder = _ProcessLockHolder()
    holder.start()
    entered = threading.Event()

    def waiter():
        with com_serial.com_operation("waits"):
            entered.set()

    t = threading.Thread(target=waiter, daemon=True)
    t.start()
    try:
        assert not entered.wait(0.3)
    finally:
        holder.release()
    t.join(5)
    assert entered.is_set()


def test_the_passwordless_save_and_close_are_marked_no_wait():
    """Structural pin beside the behavioural ones: both user-document
    bridge routes carry the no-wait form of the serialization decorator,
    so the entry-point audit in test_com_serialization sees it too."""
    for fn in (bridge.save_open_document, bridge.close_open_document):
        assert fn._com_serialized == "com_save_document"
        assert fn._refuses_if_busy is True
