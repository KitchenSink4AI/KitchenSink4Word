"""The test suite runs against a corpus of documents in tests/corpus/.

Missing corpus files are generated as structural stand-ins (see
tests/make_corpus.py), so the full suite runs anywhere. Real documents
placed in tests/corpus/ under the expected names below take precedence;
documents with genuine footnotes, tracked changes, and comments give the
most meaningful coverage.
"""

import os
from pathlib import Path

import pytest

CORPUS = Path(__file__).parent / "corpus"

# name -> features the tests expect it to contain
EXPECTED = {
    "report_tables.docx": "ordinary prose document with a couple of tables",
    "report_plain.docx": "ordinary prose document",
    "long_report.docx": "long document with headings, sections, and a real TOC",
    "many_tables.docx": "document with many tables",
    "with_images.docx": "document containing images",
    "report_second.docx": "any second prose document",
    "many_footnotes.docx": "document with many real footnotes (tests expect 171)",
    "tracked_changes.docx": "document with tracked changes (tests expect 126) and comments",
    "threaded_comments.docx": "document with threaded comments",
}


def pytest_configure(config):
    """Missing corpus files are GENERATED as structural stand-ins (see
    tests/make_corpus.py), so the full suite runs anywhere — CI included.
    Real local documents, when present, always take precedence."""
    # NO NETWORK IN TESTS. get_server_info runs the on-demand update check,
    # so the suite turns it off for every test by default; the update tests
    # switch it back on for themselves and mock the fetch. setdefault, so a
    # developer who exports the variable still wins.
    os.environ.setdefault("KS4W_UPDATE_CHECK", "off")
    present = {p.name for p in CORPUS.glob("*.docx")}
    if set(EXPECTED) - present:
        import make_corpus  # noqa: F401  (lives beside this file)

        made = make_corpus.generate_missing(verbose=True)
        if made:
            print(f"conftest: generated {len(made)} synthetic corpus file(s)")


def pytest_collection_modifyitems(config, items):
    present = {p.name for p in CORPUS.glob("*.docx")}
    missing = set(EXPECTED) - present
    if not missing:
        return
    skip = pytest.mark.skip(
        reason=(
            f"test corpus incomplete (missing: {sorted(missing)}) and "
            "generation failed — run python tests/make_corpus.py for details."
        )
    )
    uses_corpus: dict[str, bool] = {}
    for item in items:
        fname = str(item.fspath)
        if fname not in uses_corpus:
            uses_corpus[fname] = "corpus" in Path(fname).read_text(
                encoding="utf-8", errors="ignore"
            )
        # Only corpus-dependent tests are skipped; corpus-free tests still run.
        if uses_corpus[fname]:
            item.add_marker(skip)
