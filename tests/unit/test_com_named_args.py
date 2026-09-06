"""No COM method may be called with a named argument.

pywin32 drops keyword arguments on the floor when it talks to Word
late-bound, and late-bound is what a machine gets whenever its makepy
type-library cache is absent. That cache is generated on demand into the
user's temp directory, so it is missing on a fresh machine and can vanish
at any time. The call still succeeds, the return value still looks right,
and the argument simply never arrives: a read-only open is not read-only,
and an export that asked for PDF gets a .docx wearing a .pdf name.

com/callargs.py is the sanctioned route. It keeps the names at the call
site, maps them onto Word's parameter ORDER, and passes everything
positionally, so both binding modes behave identically. This file fails
the build if anyone writes the raw form again.

The scan works on TEXT rather than the syntax tree, because several live
tests embed a COM script inside a string and run it in a subprocess.
Those are real calls to Word at run time, and an AST walk of the outer
file cannot see a single one of them.
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from word_mcp.com import callargs
from word_mcp.core.errors import WordMcpError

ROOT = Path(__file__).resolve().parents[2]

#: Word object-model methods whose arguments reach COM. A named argument
#: on any of these is the bug this file exists to prevent.
COM_METHODS = {
    "Open", "SaveAs", "SaveAs2", "Close", "Quit", "CompareDocuments",
    "MergeDocuments", "ExportAsFixedFormat", "PrintOut", "Protect",
    "Unprotect", "CheckSpelling", "InsertFile", "GoTo",
}

#: Files allowed to name COM parameters, because naming them is what they
#: are for: callargs.py holds the order tables, and this file describes
#: the bug it is guarding against.
EXEMPT = {"callargs.py", "test_com_named_args.py"}

#: A capitalized keyword argument inside a COM call. Every Word parameter
#: name is capitalized, so this is what the bug looks like in text.
_NAMED = re.compile(r"\b[A-Z][A-Za-z0-9_]*\s*=(?!=)")

_CALL = re.compile(r"\.(" + "|".join(sorted(COM_METHODS)) + r")\s*\(")


def _call_text(source: str, start: int) -> str:
    """One call's text, from its opening paren to the matching one."""
    depth = 0
    for i in range(start, len(source)):
        if source[i] == "(":
            depth += 1
        elif source[i] == ")":
            depth -= 1
            if depth == 0:
                return source[start:i + 1]
    return source[start:start + 400]


def _offenders() -> list[str]:
    out = []
    for base in (ROOT / "src", ROOT / "tests"):
        for path in sorted(base.rglob("*.py")):
            if "__pycache__" in str(path) or path.name in EXEMPT:
                continue
            source = path.read_text(encoding="utf-8")
            for match in _CALL.finditer(source):
                body = _call_text(source, match.end() - 1)
                named = sorted(set(_NAMED.findall(body)))
                if not named:
                    continue
                line = source.count("\n", 0, match.start()) + 1
                out.append(
                    f"{path.relative_to(ROOT).as_posix()}:{line} "
                    f"{match.group(1)}({named})"
                )
    return out


def test_no_com_call_passes_a_named_argument():
    offenders = _offenders()
    assert not offenders, (
        "these COM calls name their arguments, which pywin32 drops "
        "silently when it is bound late (the state of any machine with no "
        "makepy cache). Pass them positionally, or route them through "
        "com/callargs.py:\n" + "\n".join(offenders)
    )


def test_the_guard_would_notice():
    """A scanner that finds nothing proves nothing, so give it something,
    including the embedded-script form no AST walk can see."""
    embedded = 'script = """\napp.Documents.Open(p, ReadOnly=True)\n"""'
    for text in ("app.Documents.Open(p, ReadOnly=True)", embedded):
        match = _CALL.search(text)
        assert match is not None
        body = _call_text(text, match.end() - 1)
        assert _NAMED.findall(body) == ["ReadOnly="]


def test_the_guard_does_not_flag_a_positional_call():
    text = "app.Documents.Open(p, False, True, False)"
    match = _CALL.search(text)
    assert not _NAMED.findall(_call_text(text, match.end() - 1))


def test_the_guard_ignores_a_lowercase_python_keyword():
    """Some of these names are also Python method names. `close(force=True)`
    on a plain object is not a COM call and must not be flagged."""
    text = "session.Close(0)  # cleanup(save_changes=False)"
    match = _CALL.search(text)
    assert not _NAMED.findall(_call_text(text, match.end() - 1))


# ---------------------------------------------------------- the order tables


def test_every_signature_has_unique_parameter_names():
    for name, params in callargs.SIGNATURES.items():
        assert len(set(params)) == len(params), f"{name} repeats a parameter"


def test_positional_fills_the_gaps_and_stops_at_the_last_named_argument():
    import pythoncom

    args = callargs.positional(
        "Documents.Open", FileName="x.docx", ReadOnly=True,
    )
    assert args[0] == "x.docx"
    assert args[1] is pythoncom.Missing, (
        "ConfirmConversions was not asked for, so it must be omitted "
        "rather than guessed at"
    )
    assert args[2] is True
    assert len(args) == 3, "nothing past the last named argument is passed"


def test_the_read_only_flag_lands_in_the_right_slot():
    """The exact bug: ReadOnly is parameter THREE, and the keyword form put
    its value nowhere at all."""
    args = callargs.positional("Documents.Open", FileName="x", ReadOnly=True)
    assert args.index(True) == 2


def test_open_and_repair_lands_past_the_gap():
    args = callargs.positional(
        "Documents.Open", FileName="x", ConfirmConversions=False,
        ReadOnly=True, AddToRecentFiles=False, OpenAndRepair=False,
    )
    assert len(args) == 13, "OpenAndRepair is Word's thirteenth parameter"
    assert args[12] is False


def test_the_password_lands_in_the_right_slot():
    args = callargs.positional(
        "Document.SaveAs2", FileName="x", FileFormat=16, Password="pw",
    )
    assert args[1] == 16
    assert args[3] == "pw", "Password is parameter four, not two"


def test_an_unknown_parameter_is_a_loud_error():
    with pytest.raises(WordMcpError, match="no parameter"):
        callargs.positional("Documents.Open", ReadOnley=True)


def test_an_unrecorded_method_is_a_loud_error():
    with pytest.raises(WordMcpError, match="no recorded parameter order"):
        callargs.positional("Document.Whatever", Thing=1)


def test_no_arguments_means_no_arguments():
    assert callargs.positional("Document.Close") == []
