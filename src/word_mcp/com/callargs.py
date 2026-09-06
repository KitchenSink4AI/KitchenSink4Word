"""Calling Word's object model without trusting named arguments.

pywin32 talks to Word LATE-BOUND whenever its makepy type-library cache is
absent, and absent is the state of a fresh machine: nothing generates that
cache until something asks for it, and it lives in the user's temp
directory, so it can vanish at any time. On the late-bound path a keyword
argument to a Document or Documents method is SILENTLY DROPPED. The call
succeeds, the return value looks right, and the argument never reached
Word.

Measured directly, with the makepy cache removed:

    Documents.Open(path, ReadOnly=True)   -> doc.ReadOnly is False
    Documents.Open(path, False, True)     -> doc.ReadOnly is True
    Document.SaveAs2(path, FileFormat=17) -> a .docx wearing a .pdf name
    Document.SaveAs2(path, 17)            -> a real PDF

Application-level methods behave differently, because pywin32 resolves
type information for Application and not for Document or Documents. That
inconsistency is the reason this module does not try to reason about which
call is safe on which machine: it makes EVERY call positional, so both
binding modes, every Word version, and a machine with or without the cache
all behave the same.

Call sites keep naming their arguments. call() maps the names onto Word's
parameter ORDER and fills the gaps with pythoncom.Missing, so the wire
form is positional while the source stays readable, and a name Word does
not have becomes a loud error here instead of a silent no-op in Word.
"""

from __future__ import annotations

from ..core.errors import WordMcpError

#: Word's parameter ORDER for the methods this server calls, from the Word
#: object-model reference. Order is the whole contract: everything below
#: is passed positionally, so a wrong entry here is a wrong argument in
#: Word. Only the leading parameters each call site needs are listed, since
#: call() never passes further than the last argument given.
SIGNATURES: dict[str, tuple[str, ...]] = {
    "Documents.Open": (
        "FileName", "ConfirmConversions", "ReadOnly", "AddToRecentFiles",
        "PasswordDocument", "PasswordTemplate", "Revert",
        "WritePasswordDocument", "WritePasswordTemplate", "Format",
        "Encoding", "Visible", "OpenAndRepair", "DocumentDirection",
        "NoEncodingDialog",
    ),
    "Document.SaveAs2": (
        "FileName", "FileFormat", "LockComments", "Password",
        "AddToRecentFiles", "WritePassword", "ReadOnlyRecommended",
        "EmbedTrueTypeFonts", "SaveNativePictureFormat", "SaveFormsData",
        "SaveAsAOCELetter", "Encoding", "InsertLineBreaks",
        "AllowSubstitutions", "LineEnding", "AddBiDiMarks",
        "CompatibilityMode",
    ),
    "Document.Close": ("SaveChanges", "OriginalFormat", "RouteDocument"),
    "Application.Quit": ("SaveChanges", "OriginalFormat", "RouteDocument"),
    "Application.CompareDocuments": (
        "OriginalDocument", "RevisedDocument", "Destination", "Granularity",
        "CompareFormatting", "CompareCaseChanges", "CompareWhitespace",
        "CompareTables", "CompareHeaders", "CompareFootnotes",
        "CompareTextboxes", "CompareFields", "CompareComments",
        "CompareMoves", "RevisedAuthor", "IgnoreAllComparisonWarnings",
    ),
    "Application.MergeDocuments": (
        "OriginalDocument", "RevisedDocument", "Destination", "Granularity",
        "CompareFormatting", "CompareCaseChanges", "CompareWhitespace",
        "CompareTables", "CompareHeaders", "CompareFootnotes",
        "CompareTextboxes", "CompareFields", "CompareComments",
        "CompareMoves", "RevisedAuthor", "FormatFrom",
    ),
}


def positional(signature: str, **kwargs) -> list:
    """The argument LIST for one COM call, in Word's own parameter order.

    Gaps between named arguments are filled with pythoncom.Missing, which
    is how an omitted optional argument is spelled on the wire, and
    nothing past the last named argument is passed at all.
    """
    import pythoncom

    try:
        names = SIGNATURES[signature]
    except KeyError:
        raise WordMcpError(
            f"no recorded parameter order for {signature}; add it to "
            f"com/callargs.py from the Word object-model reference before "
            f"calling it, because this module passes every argument "
            f"positionally"
        ) from None
    unknown = [key for key in kwargs if key not in names]
    if unknown:
        raise WordMcpError(
            f"{signature} has no parameter(s) {sorted(unknown)}; Word takes "
            f"{list(names)}"
        )
    if not kwargs:
        return []
    last = max(names.index(key) for key in kwargs)
    return [
        kwargs[names[i]] if names[i] in kwargs else pythoncom.Missing
        for i in range(last + 1)
    ]


def call(method, signature: str, **kwargs):
    """Invoke a COM method with named intent and a positional wire form."""
    return method(*positional(signature, **kwargs))
