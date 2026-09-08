"""Running Object Table enumeration, in one place.

An open Word document registers its full path as a file moniker in the COM
Running Object Table, and that is how this server reaches a document held
by a Word instance it cannot attach to directly. Four copies of the
enumerate-the-ROT-and-read-the-display-name loop had grown across
bridge.py and live.py, and three of them bound monikers without the guard
the fourth carried.

THE GUARD: binding a STARTUP add-in template (Zotero.dotm is the one that
found it) and touching any attribute is a hard ACCESS VIOLATION, a native
crash rather than a catchable exception, and a loaded add-in template is
not an open document anyway. The guard was discovered once, in the
enumerate-open-documents path, and never reached the three loops that also
bind. bind_document() is now the only place in the server that binds a ROT
moniker, so every caller inherits it.

Reading a display name is safe on any moniker; only binding is dangerous.
That is why iter_entries yields every entry and the filtering lives in
bind_document, where the crash actually happens.

The COM modules arrive as parameters rather than imports so this module
stays importable (and testable) on a machine with no pywin32.
"""

from __future__ import annotations

import contextlib

# Word loads every template in this folder at startup; those monikers are
# add-ins, not documents, and binding one crashes the process.
_STARTUP_MARKER = "\\microsoft\\word\\startup\\"


def iter_entries(pythoncom):
    """Yield (rot, moniker, display_name_lower) for each running object.

    Lazy, so a caller looking for one path stops at the match instead of
    naming every entry. A moniker whose display name cannot be read is
    skipped, which is what all four original loops did.
    """
    rot = pythoncom.GetRunningObjectTable()
    for moniker in rot.EnumRunning():
        ctx = pythoncom.CreateBindCtx(0)
        try:
            name = moniker.GetDisplayName(ctx, None)
        except Exception:
            continue
        yield rot, moniker, name.lower()


def is_bindable(name_lower: str) -> bool:
    """False for a moniker that must never be bound. See the module
    docstring: a STARTUP template takes the process down natively."""
    return _STARTUP_MARKER not in name_lower


def bind_document(pythoncom, win32com, rot, moniker, name_lower):
    """Bind a ROT moniker to a Word Document, or return None.

    None means the guard refused the moniker, or the bind failed (a
    moniker can outlive the process that registered it). Callers treat
    both the same way: skip it.
    """
    if not is_bindable(name_lower):
        return None
    with contextlib.suppress(Exception):
        obj = rot.GetObject(moniker)
        return win32com.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
    return None
