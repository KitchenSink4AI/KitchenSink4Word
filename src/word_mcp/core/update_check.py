"""Is a newer KitchenSink4Word published? A check-on-demand version report.

The family spec (identical in KS4W, KS4P, and KS4XL, modulo the package and
variable names):

- CHECK ON DEMAND ONLY. Nothing runs at import, nothing runs at server
  startup, no thread, no scheduler. The one call that can reach the network
  is status(), and exactly one tool calls it: the server's info/diagnostic
  surface. Every other tool path reads the cache or nothing at all.
- It asks PyPI's public JSON endpoint for the package it was installed
  from, at most once every seven days. The answer is cached in a tiny JSON
  file under the server's own state directory, never beside a user's
  documents.
- Two seconds or nothing. A timeout, a network error, or a malformed
  payload never raises and never blocks; it is RECORDED and REPORTED, with
  the reason and the age of the last answer that did arrive. The report is
  honest about not knowing rather than silent.
- The PyPI payload is untrusted input. Comparison uses packaging.version,
  is prerelease-aware (a prerelease is never offered to a stable build),
  and anything unparseable yields "unknown", never a fabricated version.
- OFF SWITCH: KS4W_UPDATE_CHECK=off disables the whole thing (no network
  call, no cache read, no cache write); the surface then says so. The
  older KS4W_NO_UPDATE_CHECK=1 still disables it too.
- PRIVACY: the check is one plain HTTPS GET to pypi.org. It sends nothing
  but the request itself: no document, no path, no identifier, no
  telemetry. The payload discloses this the first time it runs.
- The server never downloads, installs, or executes anything. It reports
  version numbers and facts.
"""

from __future__ import annotations

import json
import os
import tempfile
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from packaging.version import InvalidVersion, Version

#: PyPI distribution name (what `pip install -U` takes).
PACKAGE = "kitchensink4word"

#: Product name as a human reads it, used in the one surfaced line.
PRODUCT = "KitchenSink4Word"

#: The public JSON endpoint. Read-only, unauthenticated, no payload.
PYPI_URL = f"https://pypi.org/pypi/{PACKAGE}/json"

#: The off switch. The value "off" (case-insensitive) turns everything off.
OFF_ENV = "KS4W_UPDATE_CHECK"

#: The value that means off.
OFF_VALUE = "off"

#: The original opt-out, still honored: "1"/"true"/"yes"/"on" turns it off.
LEGACY_OFF_ENV = "KS4W_NO_UPDATE_CHECK"

#: Back-compat alias for the legacy name.
OPT_OUT_ENV = LEGACY_OFF_ENV

#: Test/ops seam: point the cache somewhere else. Undocumented on purpose.
CACHE_DIR_ENV = "KS4W_UPDATE_CACHE_DIR"

#: How long any answer, good or bad, is trusted before asking again. One
#: real network call per week, whatever happened last time. A week, not a
#: day: these products ship small releases often, and a user does not want
#: a notice for each of them. The report always dates its own answer, so a
#: week-old "latest version" is still an honest one.
CHECK_INTERVAL = timedelta(days=7)

#: Kept as a separate name for readability; same horizon (see above).
RETRY_INTERVAL = CHECK_INTERVAL

#: Hard cap on the request. Two seconds or nothing.
TIMEOUT_SECONDS = 2.0

_CACHE_NAME = "update-check.json"
_STATE_DIR_NAME = "word-mcp"

# --------------------------------------------------------------- placeholders
# PLACEHOLDER STRINGS. The wording is ratified by the main thread; the facts
# each one must convey are listed in the wave report
# (Agent Results/20260908_update_notice.md). Machine-readable values (state
# names, version numbers, ISO timestamps) are final; these sentences are not.

NOTE_UPDATE_AVAILABLE = "[PLACEHOLDER ks4-update-notice/note_update_available]"
NOTE_CURRENT = "[PLACEHOLDER ks4-update-notice/note_current]"
NOTE_UNKNOWN = "[PLACEHOLDER ks4-update-notice/note_unknown]"
NOTE_NOT_REACHED = "[PLACEHOLDER ks4-update-notice/note_not_reached]"
NOTE_DISABLED = "[PLACEHOLDER ks4-update-notice/note_disabled]"
INSTALL_NOTE = "[PLACEHOLDER ks4-update-notice/install_note]"
PRIVACY_NOTE = "[PLACEHOLDER ks4-update-notice/privacy_note]"


# ---------------------------------------------------------------- off switch


def disabled() -> bool:
    """True when the operator has turned the update check off.

    Checked BEFORE any file or network I/O in every entry point here, so a
    machine with the check off performs neither.
    """
    if (os.environ.get(OFF_ENV) or "").strip().lower() == OFF_VALUE:
        return True
    return (os.environ.get(LEGACY_OFF_ENV) or "").strip().lower() in {
        "1", "true", "yes", "on",
    }


# ------------------------------------------------------------- the version


def current_version() -> str | None:
    """The running build's version, or None when it cannot be read.

    Deliberately NOT importlib.metadata: an editable install's recorded
    metadata goes stale the moment the version constant is bumped, and a
    wrong "you are running X" line is worse than no line at all.
    """
    try:
        from .. import __version__

        return str(__version__)
    except Exception:
        return None


# --------------------------------------------------------------- the cache


def cache_path() -> Path:
    """The cache file: a couple of hundred bytes in the server's own state
    directory, alongside the live-lock directory the COM tier already uses.
    Never beside a user's documents."""
    override = os.environ.get(CACHE_DIR_ENV)
    if override:
        return Path(override) / _CACHE_NAME
    base = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("XDG_STATE_HOME")
        or tempfile.gettempdir()
    )
    return Path(base) / _STATE_DIR_NAME / _CACHE_NAME


def read_cache(path: Path | None = None) -> dict | None:
    """The cached answer, or None when there is not a readable one."""
    try:
        raw = (path or cache_path()).read_text(encoding="utf-8")
        data = json.loads(raw)
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def write_cache(data: dict, path: Path | None = None) -> None:
    """Replace the cache file. Silent on any failure (a read-only state
    directory must not break the server)."""
    target = path or cache_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_suffix(".tmp")
        tmp.write_text(json.dumps(data), encoding="utf-8")
        os.replace(tmp, target)
    except Exception:
        pass


def _parse_stamp(value) -> datetime | None:
    try:
        stamp = datetime.fromisoformat(str(value))
    except Exception:
        return None
    if stamp.tzinfo is None:
        stamp = stamp.replace(tzinfo=timezone.utc)
    return stamp


def is_due(cache: dict | None, now: datetime | None = None) -> bool:
    """Has the cached answer aged out?

    One horizon, seven days, whether the last attempt succeeded or failed:
    that is what caps the server at one real network call a week. No cache,
    or an unreadable timestamp, means due.
    """
    if not cache:
        return True
    stamp = _parse_stamp(cache.get("last_check"))
    if stamp is None:
        return True
    horizon = CHECK_INTERVAL if cache.get("ok") else RETRY_INTERVAL
    return (now or datetime.now(timezone.utc)) - stamp >= horizon


# ------------------------------------------------------------ the question


def _latest_stable(payload: dict) -> str | None:
    """The newest NON-prerelease release in a PyPI JSON payload.

    The payload is untrusted input: every field is type-checked, every
    version string goes through packaging.version, and anything that does
    not parse is skipped rather than passed along. Prefers the full
    releases map (so a prerelease published after the last stable cannot
    win), and falls back to info.version when the map is absent. Returns
    None for anything malformed.
    """
    if not isinstance(payload, dict):
        return None
    best: Version | None = None
    releases = payload.get("releases")
    if isinstance(releases, dict):
        for name, files in releases.items():
            try:
                candidate = Version(str(name))
            except InvalidVersion:
                continue
            if candidate.is_prerelease or candidate.is_devrelease:
                continue
            if isinstance(files, list):
                if not files:
                    continue  # no artifacts: not installable
                if all(
                    isinstance(f, dict) and f.get("yanked") for f in files
                ):
                    continue  # every artifact yanked
            if best is None or candidate > best:
                best = candidate
    if best is not None:
        return str(best)
    info = payload.get("info")
    if isinstance(info, dict):
        try:
            candidate = Version(str(info.get("version")))
        except InvalidVersion:
            return None
        if not (candidate.is_prerelease or candidate.is_devrelease):
            return str(candidate)
    return None


def _fetch(url: str = PYPI_URL) -> dict | None:
    """One GET, 2-second cap, no payload beyond a standard request. Returns
    the decoded JSON or None."""
    request = urllib.request.Request(
        url,
        headers={"Accept": "application/json",
                 "User-Agent": f"{PACKAGE}/{current_version() or 'unknown'}"},
    )
    with urllib.request.urlopen(request, timeout=TIMEOUT_SECONDS) as resp:
        body = resp.read(1_000_000)
    data = json.loads(body.decode("utf-8"))
    return data if isinstance(data, dict) else None


def _reason(exc: BaseException) -> str:
    """A short, mechanical description of why the call did not land. Carries
    the exception class and its own message, truncated; no user data can
    reach it because nothing user-supplied goes into the request."""
    text = f"{type(exc).__name__}: {exc}".strip()
    return text[:200]


def run_check(force: bool = False, path: Path | None = None) -> str | None:
    """Perform the check if it is due, refresh the cache, and return the
    latest stable version when one is known.

    Never raises. Never prints. Called from status() and from nowhere else,
    so no other tool path can reach the network.
    """
    if disabled():
        return None
    try:
        return _run_check(force, path)
    except Exception:
        return None


def _run_check(force: bool, path: Path | None) -> str | None:
    target = path or cache_path()
    cache = read_cache(target)
    if not force and not is_due(cache):
        return (cache or {}).get("latest_version")
    known = (cache or {}).get("latest_version")
    last_success = (cache or {}).get("last_success")
    disclosed = bool((cache or {}).get("disclosed"))
    stamp = datetime.now(timezone.utc).isoformat()
    error: str | None = None
    try:
        latest = _latest_stable(_fetch())
        if latest is None:
            error = "unusable response: no parseable stable release"
    except Exception as exc:
        latest = None
        error = _reason(exc)
    if latest is None:
        # The attempt is recorded (that is what caps the network at one call
        # a week), the last answer that DID arrive is kept, and the reason is
        # kept with it so the surface can say what went wrong.
        write_cache({"last_check": stamp, "latest_version": known,
                     "ok": False, "last_success": last_success,
                     "error": error, "disclosed": disclosed}, target)
        return None
    write_cache({"last_check": stamp, "latest_version": latest, "ok": True,
                 "last_success": stamp, "disclosed": disclosed}, target)
    return latest


# ------------------------------------------------------------- the surface


def status(force: bool = False) -> dict:
    """The honest version report: what is running, what is published, when
    that was last confirmed, and what went wrong when it could not be.

    THIS is the only function that can reach the network, and it does so at
    most once every seven days. It never raises, and a failed or disabled
    check produces a report saying so rather than an absent one.
    """
    running = current_version()
    if disabled():
        return {
            "state": "disabled",
            "current_version": running,
            "disabled_by": OFF_ENV,
            "off_value": OFF_VALUE,
            "note": NOTE_DISABLED,
        }

    try:
        before = read_cache() or {}
    except Exception:
        before = {}
    first_time = not before.get("disclosed")

    try:
        run_check(force=force)
    except Exception:
        pass

    try:
        cache = read_cache() or {}
    except Exception:
        cache = {}

    out: dict = {"state": "unknown", "current_version": running}
    latest = cache.get("latest_version")
    reachable = bool(cache.get("ok"))
    out["reachable"] = reachable
    last_success = cache.get("last_success")
    out["last_successful_check"] = str(last_success) if last_success else None
    if cache.get("last_check"):
        out["last_attempt"] = str(cache.get("last_check"))
    if not reachable:
        out["error"] = str(cache.get("error") or "unknown")
        out["note"] = NOTE_NOT_REACHED

    if latest:
        out["latest_version"] = str(latest)
    try:
        newer = bool(
            latest and running and Version(str(latest)) > Version(str(running))
        )
        known = bool(latest and running)
    except Exception:
        newer, known = False, False
    if not known:
        out["state"] = "unknown"
        out.setdefault("note", NOTE_UNKNOWN)
    elif newer:
        out["state"] = "update_available"
        out["install_note"] = INSTALL_NOTE
        if reachable:
            out["note"] = NOTE_UPDATE_AVAILABLE
    else:
        out["state"] = "current"
        if reachable:
            out["note"] = NOTE_CURRENT

    if first_time:
        out["privacy"] = PRIVACY_NOTE
        out["privacy_endpoint"] = PYPI_URL
        if cache:
            write_cache({**cache, "disclosed": True})

    return out


def update_notice() -> str | None:
    """One calm line when a newer stable release exists, else None.

    Reads the cache only: no network, no blocking, safe to call from any
    tool. Returns None when the check is off, when nothing is cached, when
    the versions cannot be parsed, or when the running build is current.
    """
    if disabled():
        return None
    try:
        cache = read_cache()
        if not cache:
            return None
        latest_raw = cache.get("latest_version")
        current_raw = current_version()
        if not latest_raw or not current_raw:
            return None
        latest = Version(str(latest_raw))
        running = Version(str(current_raw))
        if latest.is_prerelease or latest.is_devrelease:
            return None
        if latest <= running:
            return None
        return (
            f"{PRODUCT} {running} is running; {latest} is available. "
            f"Update: pip install -U {PACKAGE} (or download the new "
            f"installer from the releases page)."
        )
    except Exception:
        return None
