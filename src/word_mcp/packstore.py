"""Saved pack choices: a pack switched on or off stays that way until a
person or the AI changes it again (owner ruling, 2026-09-26).

Product-neutral: this file is identical in KitchenSink4Word,
KitchenSink4PPT and KitchenSink4XL, and each repo's
test_pack_persistence.py pins its hash, so a change here is made in all
three or in none.

What is saved. Every enable_tools or disable_tools call records, per pack
it names, the last choice made: on or off. The file lives in the server's
own state directory (the one the update check already uses), never beside
a user's documents. Nothing in it expires.

Keyed by the launch settings. Choices are kept under the set of packs the
launch settings turn on at startup (the MODE setting and the launch
toggles). A process started with the same launch settings starts from
those settings plus the saved choices; a process started with different
launch settings uses its own entry. So a person who changes the launch
settings has made a newer choice, and it is the one that applies. A choice
that matches what the launch settings already give is not stored.

Precedence. A launch-time lock (the pack policy "locked", or the lock
toggle) always wins: the product's packs.py does not apply saved choices
while the surface is locked, and enable_tools and disable_tools refuse
there, so nothing is saved either. The file is left as it is.

Failure is quiet and honest. An unreadable, damaged or foreign file reads
as "no saved choices"; a state directory that cannot be written leaves the
change in force for the running process and the call reports it was not
saved. Neither stops the server.

Several processes. Each save re-reads the file and changes only the packs
the call named, then replaces the file in one step, so two servers of the
same product that run at once do not undo each other's other choices.
"""

from __future__ import annotations

import json
import os
import tempfile
import threading
from pathlib import Path
from typing import Iterable, Mapping

#: The file name inside the product's state directory.
FILE_NAME = "tool-packs.json"

#: The format of the file. A file with any other value is ignored.
FORMAT = 1

#: Serializes this process's read-modify-write of the file.
_LOCK = threading.Lock()


def store_path(dir_env: str, state_dir_name: str) -> Path:
    """Where the choices live: the directory named by `dir_env` when that
    is set, else `state_dir_name` under %LOCALAPPDATA% (XDG_STATE_HOME, then
    the temp directory, off Windows)."""
    override = os.environ.get(dir_env, "").strip()
    if override:
        return Path(override) / FILE_NAME
    base = (
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("XDG_STATE_HOME")
        or tempfile.gettempdir()
    )
    return Path(base) / state_dir_name / FILE_NAME


def launch_key(startup_packs: Iterable[str]) -> str:
    """The entry a launch reads and writes: its startup packs, sorted and
    comma-joined, or "lite" when the launch turns no pack on."""
    packs = sorted({str(p) for p in startup_packs})
    return ",".join(packs) if packs else "lite"


def _read(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001 - missing, unreadable or damaged
        return {"format": FORMAT, "launches": {}}
    if (
        not isinstance(data, dict)
        or data.get("format") != FORMAT
        or not isinstance(data.get("launches"), dict)
    ):
        return {"format": FORMAT, "launches": {}}
    return data


def load(path: Path, key: str, known_packs: Iterable[str]) -> dict[str, bool]:
    """The saved choices for one launch: {pack: on?}. Packs this build does
    not know and values that are not true/false are skipped."""
    known = set(known_packs)
    entry = _read(path)["launches"].get(key)
    if not isinstance(entry, dict):
        return {}
    return {
        pack: value
        for pack, value in entry.items()
        if pack in known and isinstance(value, bool)
    }


def save(
    path: Path,
    key: str,
    changes: Mapping[str, bool],
    startup_packs: Iterable[str],
) -> bool:
    """Record `changes` ({pack: on?}) under `key`. A pack whose choice
    equals its launch-settings state is dropped from the entry. Returns
    True when the file now holds the choices, False when it could not be
    written."""
    startup = set(startup_packs)
    with _LOCK:
        data = _read(path)
        launches = data["launches"]
        entry = launches.get(key)
        entry = dict(entry) if isinstance(entry, dict) else {}
        for pack, on in changes.items():
            if bool(on) == (pack in startup):
                entry.pop(pack, None)
            else:
                entry[pack] = bool(on)
        if entry:
            launches[key] = entry
        else:
            launches.pop(key, None)
        data["format"] = FORMAT
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            fd, tmp = tempfile.mkstemp(
                prefix=".tool-packs.", suffix=".tmp", dir=str(path.parent))
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    json.dump(data, fh, indent=1, sort_keys=True)
                os.replace(tmp, path)
            except BaseException:
                try:
                    os.unlink(tmp)
                except OSError:
                    pass
                raise
        except Exception:  # noqa: BLE001 - read-only or vanished directory
            return False
        return True
