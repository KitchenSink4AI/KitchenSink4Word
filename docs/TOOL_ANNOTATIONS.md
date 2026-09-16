# Tool annotations

Every tool this server registers ships four MCP annotations: `title`
(see [TOOL_TITLES.md](TOOL_TITLES.md)), `readOnlyHint`, `destructiveHint`
and `openWorldHint`, plus `idempotentHint` where it is obviously true.
The tables are GENERATED from `src/word_mcp/core/tool_annotations.py` and a test in
`tests/unit/test_tool_annotations.py` fails if they disagree, so a row
here is what goes on the wire.

## The rule for `destructiveHint`

ONE rule decides every row, and it is stated here so a reviewer can
dispute any single one of them.

**`false`** only when every code path either

* **(a)** adds new content or a new file without replacing anything that
  was already there, with creation refusing an existing target, or
* **(b)** changes no user data at all: this session's tool surface, the
  viewport, a read performed through a hidden or already-running Office
  instance, or a read taken through a temporary copy.

**`true`** otherwise. That covers everything that deletes, replaces,
overwrites, clears, reorders, moves, applies a batch of edits, saves over
the document the user has open, writes an output file it may silently
overwrite, or acts on a live page. It also covers **every tool that writes
into an existing document**, because the pre-write backup these servers
take is defeatable by the tool's own `backup=False` argument and therefore
cannot be claimed as guaranteed reversibility. And it covers everything
whose reversibility could not be PROVEN by reading the code: an unproven
claim of safety is the one thing this annotation must not make.

Read-only tools carry no `destructiveHint`. The field is meaningful only
when `readOnlyHint` is false, and a value there would be noise.

`idempotentHint` is set true only where repeating the identical call
obviously lands the same state: the pack switches, whose own docstrings
say idempotent, and the whole-value setters that take an address and a
value and carry no action selector. Everything else is left unset rather
than guessed; an absent hint means undeclared, never false.

`openWorldHint` is `false` on every tool. Nothing here reaches a remote service; the
optional update check is not a tool.

## Tools that may perform destructive updates (52)

`destructiveHint: true`.

| Tool | Why |
|---|---|
| `anonymize_for_review` | rewrites author names and self-citations through the document |
| `apply_edits` | applies a batch of addressed edits whose ops include replace and delete |
| `apply_manuscript_format` | writes over the styling or content already in place |
| `apply_style` | writes over the styling or content already in place |
| `apply_template` | writes over the styling or content already in place |
| `assemble_front_matter` | spec.force replaces front matter that is already there |
| `change_heading_level` | reorders or moves existing content |
| `com_export_pdf` | SaveAs2 writes the output path with no existing-file check |
| `com_multi_document` | writes a comparison or merge output with no proven collision check |
| `com_refresh_fields` | rewrites every field result and saves the source file in place |
| `com_save_document` | saves over the document the user has open |
| `convert_citation_style` | rewrites existing content in place |
| `convert_notes` | rewrites existing content in place |
| `deanonymize_document` | writes the mapping back over the anonymized text |
| `define_style` | creates OR REPLACES a style definition of the same name |
| `delete_element` | deletes existing content |
| `delete_paragraphs` | deletes existing content |
| `delete_table` | deletes existing content |
| `fill_template` | fills the template IN PLACE, replacing every placeholder run |
| `fix_accessibility` | rewrites existing content in place |
| `format_cells` | writes a new value over the one already stored |
| `format_text` | writes a new value over the one already stored |
| `import_table` | addressing an existing table refills it in place, overwriting its cells |
| `manage_backups` | action='restore' overwrites the document and action='purge' deletes backups |
| `manage_comment` | carries delete or remove actions alongside its read actions |
| `manage_note` | carries delete or remove actions alongside its read actions |
| `manage_source` | carries delete or remove actions alongside its read actions |
| `modify_table_structure` | reorders or moves existing content |
| `move_section` | reorders or moves existing content |
| `prepare_for_submission` | accepts every tracked change and deletes every comment |
| `redact_text` | rewrites existing content in place |
| `resolve_revisions` | rewrites existing content in place |
| `search_and_replace` | replaces matched text throughout the document |
| `set_bibliography_style` | writes a new value over the one already stored |
| `set_cells` | writes a new value over the one already stored |
| `set_chart_data` | writes a new value over the one already stored |
| `set_content_control` | writes a new value over the one already stored |
| `set_document_properties` | writes a new value over the one already stored |
| `set_document_protection` | writes a new value over the one already stored |
| `set_form_fields` | writes a new value over the one already stored |
| `set_header_footer` | writes a new value over the one already stored |
| `set_image` | writes a new value over the one already stored |
| `set_list_numbering` | writes a new value over the one already stored |
| `set_page_numbers` | writes a new value over the one already stored |
| `set_paragraph_format` | writes a new value over the one already stored |
| `set_paragraph_text` | writes a new value over the one already stored |
| `set_section_properties` | writes a new value over the one already stored |
| `set_table_properties` | writes a new value over the one already stored |
| `set_textbox_text` | writes a new value over the one already stored |
| `set_watermark` | writes a new value over the one already stored |
| `setup_chapter_headers` | replaces whatever other content the running header held |
| `sort_table` | reorders or moves existing content |
## Tools that may not (37)

`destructiveHint: false`.

| Tool | Why |
|---|---|
| `com_import_pdf` | creates a new object or file and refuses an existing target |
| `com_proofing_errors` | drives a hidden or already-running Office instance to read; the file is not modified |
| `com_readability_statistics` | drives a hidden or already-running Office instance to read; the file is not modified |
| `com_validate_opens_clean` | drives a hidden or already-running Office instance to read; the file is not modified |
| `copy_document` | writes to a separate destination; an overwrite rotates the destination into its backup slot first |
| `copy_table` | writes to a separate destination; an overwrite rotates the destination into its backup slot first |
| `create_document` | creates a new object or file and refuses an existing target |
| `create_table` | adds new content at a position; nothing existing is replaced or removed |
| `disable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
| `enable_tools` | changes only this session's tool surface, which the counterpart tool reverses |
| `export_images` | writes a new output file and refuses a collision; the source is never modified |
| `export_table` | writes a new output file and refuses a collision; the source is never modified |
| `get_document_view` | adds w14:paraId attributes only; content is never rewritten or removed |
| `insert_bookmark` | adds new content at a position; nothing existing is replaced or removed |
| `insert_break` | adds new content at a position; nothing existing is replaced or removed |
| `insert_caption` | adds new content at a position; nothing existing is replaced or removed |
| `insert_chart` | adds new content at a position; nothing existing is replaced or removed |
| `insert_citation` | adds new content at a position; nothing existing is replaced or removed |
| `insert_content_control` | adds new content at a position; nothing existing is replaced or removed |
| `insert_cross_reference` | adds new content at a position; nothing existing is replaced or removed |
| `insert_document` | adds new content at a position; nothing existing is replaced or removed |
| `insert_equation` | adds new content at a position; nothing existing is replaced or removed |
| `insert_field` | adds new content at a position; nothing existing is replaced or removed |
| `insert_hyperlink` | adds new content at a position; nothing existing is replaced or removed |
| `insert_image` | adds new content at a position; nothing existing is replaced or removed |
| `insert_list` | adds new content at a position; nothing existing is replaced or removed |
| `insert_paragraphs` | adds new content at a position; nothing existing is replaced or removed |
| `insert_reference_list` | adds new content at a position; nothing existing is replaced or removed |
| `insert_zotero_citation` | adds new content at a position; nothing existing is replaced or removed |
| `live_insert_at_cursor` | adds new content at a position; nothing existing is replaced or removed |
| `live_repair` | resets the live layer's own state; touches no document content and saves nothing |
| `live_scroll_to` | moves what the user is looking at; no document byte changes |
| `live_set_track_changes` | flips a document mode flag and returns the previous value; no content is touched |
| `mail_merge` | writes one new file per row and refuses before writing on any collision |
| `mark_index_entry` | adds new content at a position; nothing existing is replaced or removed |
| `search_zotero_library` | reads through a temporary copy; the real database is never touched |
| `split_document` | writes new files and refuses a collision; the source is never modified |
## Read-only tools (23)

`readOnlyHint: true`, and no `destructiveHint`: the field carries no
meaning for a tool that changes nothing. The classification itself lives
in `core/readonly.py` and is guarded by
`tests/unit/test_readonly_annotations.py`.

`com_word_status`, `comment_report`, `detect_citation_system`, `diagnose_document`, `find_text`, `get_comments`, `get_document_info`, `get_headers_footers`, `get_outline`, `get_paragraph_format`, `get_protection`, `get_revision_report`, `get_server_info`, `get_styles`, `get_table`, `get_text`, `get_tracked_changes`, `get_workflows`, `list_elements`, `parse_references`, `structured_diff`, `validate`, `word_count`

## Idempotent tools (19)

`idempotentHint: true`.

`disable_tools`, `enable_tools`, `set_bibliography_style`, `set_cells`, `set_chart_data`, `set_content_control`, `set_document_properties`, `set_document_protection`, `set_form_fields`, `set_header_footer`, `set_image`, `set_list_numbering`, `set_page_numbers`, `set_paragraph_format`, `set_paragraph_text`, `set_section_properties`, `set_table_properties`, `set_textbox_text`, `set_watermark`
