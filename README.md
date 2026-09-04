<!-- mcp-name: io.github.KitchenSink4AI/kitchensink4word -->

# 🚰 KitchenSink4Word

[![Tests](https://github.com/KitchenSink4AI/KitchenSink4Word/actions/workflows/tests.yml/badge.svg)](https://github.com/KitchenSink4AI/KitchenSink4Word/actions/workflows/tests.yml)
[![PyPI](https://img.shields.io/pypi/v/kitchensink4word)](https://pypi.org/project/kitchensink4word/)
[![License: AGPL-3.0](https://img.shields.io/badge/license-AGPL--3.0-blue)](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/LICENSE)

**Edit Word documents with your AI assistant: tracked changes, citations, templates, and live edits in an open document.**

KitchenSink4Word is an MCP server that lets Claude Code, the Codex CLI, Copilot CLI, Claude Desktop and other local MCP clients read and edit .docx files. It can also work on a document while it is open in Microsoft Word, so you watch the changes arrive. This is the free Community edition, licensed under the AGPL-3.0.

[Product page](https://kitchensink4.ai/KitchenSink4Word/) · [llms.txt](https://kitchensink4.ai/KitchenSink4Word/llms.txt) · [Changelog](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/CHANGELOG.md)

**Platform:** Windows. Editing files works without Word installed. Live editing and the Word-powered features (PDF import and export, document compare, proofing, field refresh) need Microsoft Word.

## Install

The package needs Python 3.12 or newer. Run the commands below in PowerShell or a terminal, not in an AI chat.

**Claude Code or Codex CLI.** Install [uv](https://docs.astral.sh/uv/getting-started/installation/), run the line for your app, and start a new session:

```sh
claude mcp add word -s user -- uvx kitchensink4word
```

```sh
codex mcp add word -- uvx kitchensink4word
```

**Claude Desktop.** Install uv and restart Claude Desktop. Download the `.mcpb` file from the [latest release](https://github.com/KitchenSink4AI/KitchenSink4Word/releases/latest), then open Settings > Extensions > Advanced settings > Install extension and choose the file. The first launch downloads the Python package, so it needs a network connection.

**Any other local MCP client.** Use `uvx` as the command and `kitchensink4word` as its argument, or install the package and run `kitchensink4word` as the server command:

```sh
pip install kitchensink4word
```

Installing the package by itself does not connect it to an AI app; add it as a local MCP server in your client's settings.

## What it can do

- Revise a report while it stays open in Word, with each edit visible as it lands.
- Read, accept and reject tracked changes, write edits as tracked changes, and work through reviewer comments.
- Fill templates, assemble chapter files into one document, and run mail merges.
- Build and format tables, lists, headings, sections and page layouts.
- Insert citations, bibliographies, footnotes, endnotes, captions and cross-references, including Zotero library citations.
- Check formatting, accessibility and document structure before you hand a file over.
- Compare two versions of a document, redact selected text, and import or export PDF through Word.

In total the server offers 110 document tools covering 222 operations, plus two tools that switch capability packs on and off.

## Highlights

- **Table columns through merged cells.** Inserting or deleting a column works correctly when the table has horizontally or vertically merged cells.
- **Bulk edits in one call.** Twenty cell edits are one `set_cells` call, not twenty round trips.
- **Real tracked changes.** Edit tools accept `track=True` and an author name and produce Word revisions the recipient can accept or reject.
- **Word-native compare.** `com_multi_document(action="compare")` produces a redline between two versions.
- **Built not to corrupt.** Saves are atomic and validated, the previous version is backed up before every change, and anything the call does not touch (equations, text boxes, content controls) passes through byte for byte. A file that is open in Word is refused with a clear message instead of a hang.
- **Find and replace across split runs.** Matches work even when Word has broken the text into many runs, and per-character formatting is preserved. User regular expressions run with a timeout.

## Tool packs and context cost

A session starts with a small core of everyday tools and loads larger capability packs only when a task needs them, which keeps the tool list your AI carries short. All figures count the full `tools/list` entry the client receives and come from `scripts/measure_surface.py`.

| Pack | Tools | Approx. tokens | What it covers |
|---|---:|---:|---|
| **lite** (loaded at start) | 29 | ~10.2k | Text, paragraphs, tables, cells, lists, find and replace, outline, document view, backups, workflow guides, server info, pack controls |
| references | 8 | ~2.9k | Word-native citations and bibliography, Zotero search and cite, citation checks, style conversion |
| review | 9 | ~2.4k | Tracked changes, threaded comments, structured diff, anonymize and restore |
| academic | 24 | ~7.6k | Footnotes and endnotes, TOC, index, captions, cross-references, front matter, sections, styles, numbering, word counts, validation, accessibility |
| assembly | 7 | ~2.1k | Insert and split documents, move sections, copy tables between files, templates, mail merge |
| media-forms | 16 | ~5.6k | Images, charts, equations, text boxes, hyperlinks, table structure and styling, form fields, content controls, field codes |
| com-live | 13 | ~3.1k | Drives a local Word: PDF import and export, compare and combine, proofing, readability, field refresh, live editing |
| protection-io | 6 | ~1.6k | Document protection, watermarks, verified redaction, table data import and export |
| **Everything** | **112** | **~35.5k** | All packs plus `enable_tools` and `disable_tools` |

Your AI turns a pack on by name when it needs one:

```
enable_tools(["references"])
enable_tools(["academic", "review"])
```

A lite tool that cannot do something names the pack and the exact `enable_tools` call to use, and the `get_workflows` recipes list the packs each workflow needs. Clients that load tool schemas on first use, such as Claude Code, pay close to nothing until a tool is called.

To load everything from the start, set `KS4W_MODE=full`. In the Claude Desktop extension, the settings **Load every tool at startup** and **Lock the tool set at startup** do the same through `KS4W_ALL_TOOLS` and `KS4W_LOCK_TOOLS` (`true` or `false`; any other value stops startup with a clear message). `KS4W_MODE` and `KS4W_PACK_POLICY` take precedence when set, and the server prints one line at startup saying what decided the tool set.

**Fewer permission prompts:** in Claude Desktop's tool permissions for this server, set the read-only tools to Always Allow. They cannot change anything. Leave the editing tools on Ask.

**If a re-enabled tool reports "no such tool":** some clients drop a tool's schema when its pack is turned off and do not reload it when the pack comes back, even though the server announces the change. Refresh the client's tool list (in Claude Code, a `ToolSearch` call for the tool reloads it).

## Safety model

- Before every change to a closed file, the current version is copied into a hidden `.ks4w-backups/` folder next to the document. Each document keeps two copies: `prev.docx` (the state before the most recent change) and `anchor.docx` (the state at the start of the session, refreshed after an hour of inactivity). Storage stays at about two copies per document no matter how many edits you make. `manage_backups` lists, restores and purges them, and a restore can itself be undone.
- Edits to a document that is open in Word are not part of this backup system. Word's own AutoRecover and undo cover the open document.
- Saves are atomic and validated. If an operation fails, the original file is left exactly as it was.
- Changes to the same file are serialized, within one server and across server processes, so parallel calls cannot overwrite each other.
- Deleting content that holds footnote references removes the orphaned notes, and `validate` checks integrity in both directions. Paragraph deletion refuses ranges that would cut a field such as a TOC in half or remove a section break by accident.
- Exclude `.ks4w-backups/` from cloud sync (OneDrive, Dropbox, Google Drive). The backup files change with every edit, and sync clients can hold locks that slow saves down.

## Privacy

The server runs on your computer. KitchenSink4AI never receives your files or usage data. Your AI app may send prompts, file contents and tool results to its own provider under that app's settings and terms. The optional update check described below contacts PyPI and sends nothing beyond the request itself. The [Privacy Policy](https://kitchensink4.ai/privacy/) covers the product, purchases and support.

## Configuration

**Update check.** The server checks PyPI for a newer release only when you call `get_server_info`: never at startup or on a timer, at most once every seven days, with a two-second limit. It never downloads or installs anything. Set `KS4W_UPDATE_CHECK=off` to turn it off completely (the older `KS4W_NO_UPDATE_CHECK=1` also works).

**Sandboxing (optional).** Set `KS4W_ALLOWED_ROOTS` to a list of folders separated by `;` on Windows (`:` elsewhere), for example `%USERPROFILE%\Documents;D:\Work`. Every path the server reads or writes must then resolve inside one of them. The check runs on canonical paths, so `..\` traversal, symbolic links and junctions, 8.3 short names, extended-length prefixes, case tricks and lookalike folder names (`Documents2` next to an allowed `Documents`) are all caught. Network (UNC) paths are refused unless an allowed root is itself a UNC path that contains them. A blocked call returns a typed error naming the path and the allowed roots before any file is opened. We recommend it whenever the server handles requests you do not fully trust.

## Live editing in Word

Which tools work on a document that is OPEN in Word. Dual-mode tools auto-route (`live='auto'`); everything else refuses with `DOCUMENT_LOCKED` until the file is closed. Word COM calls within one server process are serialized. Live calls and passwordless save or close also use a machine-local cross-process lock; a concurrent caller receives `APP_BUSY` before touching Word and nothing is queued. Live edits stay unsaved until `com_save_document`; `get_workflows(task='live-editing')` explains the model.

| Dual-mode tool (file + live) | Live-mode notes |
|---|---|
| `apply_edits` | one COM undo group; markdown lists/pipe tables in insert ops refuse up front with a close-the-file hint; replace verifies matched text in the live range; all other index-addressed ops run the snapshot staleness guard against the last saved state |
| `delete_paragraphs` | section-break/field-crossing ranges refuse; range={} location endpoints run the snapshot staleness guard |
| `find_text` | plain queries, 500-match cap; formatting mode and include_textboxes file-mode only |
| `format_text` | formatting mode routes live; case mode file-mode only (live='force' refuses); range={} locations run the snapshot staleness guard |
| `get_comments` | same shape |
| `get_document_info` | same key names; adds words/track_revisions, omits part list |
| `get_outline` | same flat-list shape |
| `get_text` | body shape identical; include_textboxes/textbox modes file-mode only |
| `insert_paragraphs` | heading_level maps to built-in Heading styles by numeric constant (outline-based docs get direct outline levels); inherit_format/copy_format_from file-mode only; text-selector locations run the snapshot staleness guard |
| `search_and_replace` | plain items route live (255-char finds handled); preview and find_formatting modes are file-mode only (find_formatting with live='force' refuses) |
| `set_cells` | plain text cells; vertical merges refuse live |
| `set_paragraph_format` | shading/borders/tab_stops keys refuse live (XML-level); raw indices keep the v1 index-trust contract |
| `set_paragraph_text` | expect guard honored; tracked-revision paragraphs refuse; text-selector locations run the snapshot staleness guard |
| `word_count` | Word's own statistics engine; exclusions mode file-mode only |

**Live-only tools** (need the document open): `live_insert_at_cursor`, `live_repair`, `live_scroll_to`, `live_set_track_changes`.  
**COM tier** (drive the Word application, open or closed): `com_export_pdf`, `com_import_pdf`, `com_multi_document`, `com_proofing_errors`, `com_readability_statistics`, `com_refresh_fields`, `com_save_document`, `com_validate_opens_clean`, `com_word_status`.  
**Open-document reader**: `get_document_view` (reads the last SAVED state and says so).

The remaining 63 writers and 14 readers are file-only: they refuse while the document is open (close it, or `com_save_document` with `close=true`).

## Known limits

- Live regular-expression replacements skip matches that sit after complex fields in the same story, because Word's character offsets drift there. The number skipped is reported, and a plain-text find still works.
- Live `set_cells` refuses tables with vertically merged cells. The file-based tool handles them; close the document first.
- Live tracked-change attribution follows Word: if Word is signed in to an Office account, revisions are attributed to that account. Results report the author Word actually used.
- Table of contents and caption-list page numbers need a field update. Word does it the next time the file opens, or call `com_refresh_fields`.
- Some tools find styles by their displayed name, so Word installs in other languages may need the localized style names. English and Korean are handled.

**Beta tools.** `convert_citation_style`, `anonymize_for_review`, `validate(checks=["defined_terms"])` and the text-reference scan in `validate(checks=["cross_references"])` are heuristic. Each returns a list of what it could not handle with confidence, so review that list rather than trusting the result blindly.

## Upgrading from 1.x

Every 1.x tool name changed in 2.0, when 189 tools were consolidated into 110 that cover every earlier capability. Read the [migration guide](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/docs/MIGRATION_V2.md) first; it maps each old name to its new home, and `get_workflows("migrate-from-v1")` returns the same map inside a session.

## Testing

The suite has more than 2,000 tests. About 95 of them drive a real Word instance and run only on Windows with Word installed. Tests that need sample documents generate structurally equivalent synthetic files with `tests/make_corpus.py`, so the full suite runs on any machine and on every pull request. `tests/word_validator.py` opens each output in an invisible Word and fails on any repair prompt.

To run the suite without Word:

```sh
pytest tests -m "not live"
```

## Reporting problems

Please [open an issue](https://github.com/KitchenSink4AI/KitchenSink4Word/issues) with the tool name, the call you made, what happened and what you expected. Ask your AI to run `get_server_info` and paste the output; it is designed to be safe to share. Please do not attach documents that contain anything private. For security issues, follow [SECURITY.md](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/SECURITY.md) instead.

## Business edition

The Business edition pairs these tools with a Windows installer, a signed update channel and support under the Business terms. Update checks tell you when a covered release is available, and nothing installs on its own. See [pricing](https://kitchensink4.ai/pricing/). The Community edition stays free under the AGPL, including business use that meets its terms.

## Contributing

Bug reports, fixes, tests and documentation are welcome. See [CONTRIBUTING.md](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/CONTRIBUTING.md); first-time contributors agree to the [Contributor License Agreement](https://github.com/KitchenSink4AI/KitchenSink4Word/blob/main/CLA.md).

## License

KitchenSink4Word is dual-licensed.

- **AGPL-3.0.** Free for anyone, including businesses, for any use that complies with the AGPL. Its terms include sharing your source code, with your modifications, when you distribute the software or make it available over a network.
- **Commercial license.** For organizations that want to build KitchenSink4Word into their own products or services without the AGPL's source-sharing obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line of Alvut Consulting, LLC. KitchenSink4AI tools are built with the help of AI coding assistants, and every release goes through the full test suite before it ships.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation. Microsoft and Word are trademarks of the Microsoft group of companies.
