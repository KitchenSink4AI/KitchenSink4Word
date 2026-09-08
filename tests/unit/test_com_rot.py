"""The Running Object Table scan, and the guard that must cover all of it.

Binding a STARTUP add-in template (Zotero.dotm) and touching an attribute
is a hard ACCESS VIOLATION: a native crash, not a catchable exception. The
guard was found once and lived in one of four identical ROT loops. These
tests pin the consolidation, so the guard cannot diverge again.

No COM here. com/rot.py takes pythoncom and win32com as parameters
precisely so it can be exercised with fakes on any machine.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from word_mcp.com import rot

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src" / "word_mcp"


# ------------------------------------------------------------------ fakes


class _Moniker:
    def __init__(self, name, readable=True):
        self.name = name
        self._readable = readable

    def GetDisplayName(self, ctx, other):
        if not self._readable:
            raise RuntimeError("moniker is dying")
        return self.name


class _Doc:
    def __init__(self, name):
        self.FullName = name
        self.touched = False


class _Rot:
    """Bind a startup template and the fake process dies, the way the
    real one does."""

    def __init__(self, monikers, bindable=None, crash_on=()):
        self._monikers = monikers
        self._bindable = bindable if bindable is not None else {}
        self._crash_on = tuple(crash_on)
        self.bound = []

    def EnumRunning(self):
        return list(self._monikers)

    def GetObject(self, moniker):
        self.bound.append(moniker.name)
        if moniker.name in self._crash_on:
            raise SystemExit("ACCESS VIOLATION (stand-in for a native crash)")
        if moniker.name not in self._bindable:
            raise RuntimeError("no longer registered")
        return _Obj(self._bindable[moniker.name])


class _Obj:
    def __init__(self, doc):
        self._doc = doc

    def QueryInterface(self, iid):
        return self._doc


class _Pythoncom:
    IID_IDispatch = object()

    def __init__(self, rot_obj):
        self._rot = rot_obj

    def GetRunningObjectTable(self):
        return self._rot

    def CreateBindCtx(self, flags):
        return object()


class _Win32Com:
    @staticmethod
    def Dispatch(obj):
        return obj


# ------------------------------------------------------------ enumeration


def test_iter_entries_lowercases_and_skips_unreadable_monikers():
    r = _Rot([
        _Moniker(r"C:\Docs\Chapter4.DOCX"),
        _Moniker("broken", readable=False),
        _Moniker(r"C:\Docs\Ch5.docx"),
    ])
    got = [name for _rot, _mk, name in rot.iter_entries(_Pythoncom(r))]
    assert got == [r"c:\docs\chapter4.docx", r"c:\docs\ch5.docx"]


def test_iter_entries_is_lazy():
    """A caller looking for one path stops at the match instead of naming
    every entry in the table."""
    monikers = [_Moniker(f"c:\\docs\\{i}.docx") for i in range(50)]
    r = _Rot(monikers)
    seen = 0
    for _rot, _mk, name in rot.iter_entries(_Pythoncom(r)):
        seen += 1
        if name == r"c:\docs\2.docx":
            break
    assert seen == 3


# ------------------------------------------------------------- the guard


def test_startup_template_is_never_bound():
    startup = (
        r"c:\users\a\appdata\roaming\microsoft\word\startup\zotero.dotm"
    )
    r = _Rot([_Moniker(startup)], crash_on=(startup,))
    pyc = _Pythoncom(r)
    entry = next(rot.iter_entries(pyc))
    assert rot.bind_document(pyc, _Win32Com, entry[0], entry[1], entry[2]) is None
    assert r.bound == [], "the guard let a startup template reach GetObject"


def test_is_bindable_matches_the_startup_folder_not_the_words():
    """iter_entries lowercases, so is_bindable works on a lowered name.
    It must catch the folder and leave an innocent filename alone."""
    assert not rot.is_bindable(
        r"c:\users\a\appdata\roaming\microsoft\word\startup\zotero.dotm"
    )
    assert not rot.is_bindable(
        r"d:\office\microsoft\word\startup\grammarly.dotm"
    )
    assert rot.is_bindable(r"c:\docs\microsoft word startup notes.docx")
    assert rot.is_bindable(r"c:\docs\chapter4.docx")


def test_bind_document_returns_none_when_binding_fails():
    r = _Rot([_Moniker(r"c:\docs\gone.docx")])
    pyc = _Pythoncom(r)
    entry = next(rot.iter_entries(pyc))
    assert rot.bind_document(pyc, _Win32Com, *entry) is None


def test_bind_document_returns_the_document():
    doc = _Doc(r"C:\Docs\Ch4.docx")
    r = _Rot(
        [_Moniker(r"C:\Docs\Ch4.docx")],
        bindable={r"C:\Docs\Ch4.docx": doc},
    )
    pyc = _Pythoncom(r)
    entry = next(rot.iter_entries(pyc))
    assert rot.bind_document(pyc, _Win32Com, *entry) is doc


# --------------------------------------------- no fifth copy, ever again


def test_only_rot_module_enumerates_or_binds_the_running_object_table():
    """The consolidation itself. Four copies is how a hard-crash guard
    ends up in one of them; this is what stops the fifth."""
    offenders = []
    for py in sorted(SRC.rglob("*.py")):
        if py.name == "rot.py":
            continue
        text = py.read_text(encoding="utf-8")
        for i, line in enumerate(text.splitlines(), 1):
            if re.search(r"GetRunningObjectTable|\.EnumRunning\(", line):
                offenders.append(f"{py.relative_to(ROOT)}:{i} {line.strip()}")
    assert not offenders, (
        "ROT enumeration outside com/rot.py:\n" + "\n".join(offenders)
    )


def test_the_guard_has_exactly_one_home():
    """The startup-template marker appears in one file. A second copy is
    the divergence starting over."""
    carriers = sorted(
        py.name for py in SRC.rglob("*.py")
        if re.search(r"word\\+startup", py.read_text(encoding="utf-8").lower())
    )
    assert carriers == ["rot.py"], carriers


@pytest.mark.parametrize(
    "module,func",
    [("bridge", "_open_in_running_word"), ("live", "_find_doc_via_rot")],
)
def test_rot_callers_go_through_the_shared_module(module, func):
    src = (SRC / "com" / f"{module}.py").read_text(encoding="utf-8")
    assert "from . import rot as _rot" in src
    assert func in src
