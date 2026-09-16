# Tool titles

Every tool this server registers ships a `title` annotation, which is what
the Anthropic Connectors Directory requires and what a client shows in
place of the raw tool name. This table is GENERATED from
`src/word_mcp/core/tool_annotations.py` and a test in `tests/unit/test_tool_annotations.py`
fails if the two ever disagree, so the English here is the English on the
wire.

## How a title is derived

The rule is mechanical, so that a new tool gets a title without anyone
inventing one:

1. split the tool name on underscores;
2. replace a leading `com_` with the host application's name and collapse
   an immediately repeated word (`com_word_status` -> `Word Status`);
3. upper-case known acronyms (`pdf` -> `PDF`, `svg` -> `SVG`), Title Case
   the rest, and keep short function words lowercase inside the phrase
   (`add_equation_to_shape` -> `Add Equation to Shape`);
4. a name too short or too generic to read as a title takes a phrase from
   the first clause of its own description instead. Those are the only
   hand-written strings here and the Source column marks them.

Titles are unique within this server and none exceeds 40 characters. No
title carries product or marketing language.

112 tools.

| Tool | Title | Source |
|---|---|---|
| `anonymize_for_review` | Anonymize for Review | mechanical |
| `apply_edits` | Apply Edits | mechanical |
| `apply_manuscript_format` | Apply Manuscript Format | mechanical |
| `apply_style` | Apply Style | mechanical |
| `apply_template` | Apply Template | mechanical |
| `assemble_front_matter` | Assemble Front Matter | mechanical |
| `change_heading_level` | Change Heading Level | mechanical |
| `com_export_pdf` | Word Export PDF | mechanical |
| `com_import_pdf` | Word Import PDF | mechanical |
| `com_multi_document` | Word Multi-Document Ops | first clause of its description |
| `com_proofing_errors` | Word Proofing Errors | mechanical |
| `com_readability_statistics` | Word Readability Statistics | mechanical |
| `com_refresh_fields` | Word Refresh Fields | mechanical |
| `com_save_document` | Word Save Document | mechanical |
| `com_validate_opens_clean` | Word Validate Opens Clean | mechanical |
| `com_word_status` | Word Status | mechanical |
| `comment_report` | Comment Report | mechanical |
| `convert_citation_style` | Convert Citation Style | mechanical |
| `convert_notes` | Convert Notes | mechanical |
| `copy_document` | Copy Document | mechanical |
| `copy_table` | Copy Table | mechanical |
| `create_document` | Create Document | mechanical |
| `create_table` | Create Table | mechanical |
| `deanonymize_document` | Deanonymize Document | mechanical |
| `define_style` | Define Style | mechanical |
| `delete_element` | Delete Element | mechanical |
| `delete_paragraphs` | Delete Paragraphs | mechanical |
| `delete_table` | Delete Table | mechanical |
| `detect_citation_system` | Detect Citation System | mechanical |
| `diagnose_document` | Diagnose Document | mechanical |
| `disable_tools` | Disable Tools | mechanical |
| `enable_tools` | Enable Tools | mechanical |
| `export_images` | Export Images | mechanical |
| `export_table` | Export Table | mechanical |
| `fill_template` | Fill Template | mechanical |
| `find_text` | Find Text | mechanical |
| `fix_accessibility` | Fix Accessibility | mechanical |
| `format_cells` | Format Cells | mechanical |
| `format_text` | Format Text | mechanical |
| `get_comments` | Get Comments | mechanical |
| `get_document_info` | Get Document Info | mechanical |
| `get_document_view` | Get Document View | mechanical |
| `get_headers_footers` | Get Headers and Footers | first clause of its description |
| `get_outline` | Get Outline | mechanical |
| `get_paragraph_format` | Get Paragraph Format | mechanical |
| `get_protection` | Get Protection | mechanical |
| `get_revision_report` | Get Revision Report | mechanical |
| `get_server_info` | Get Server Info | mechanical |
| `get_styles` | Get Styles | mechanical |
| `get_table` | Get Table | mechanical |
| `get_text` | Get Text | mechanical |
| `get_tracked_changes` | Get Tracked Changes | mechanical |
| `get_workflows` | Get Workflows | mechanical |
| `import_table` | Import Table | mechanical |
| `insert_bookmark` | Insert Bookmark | mechanical |
| `insert_break` | Insert Break | mechanical |
| `insert_caption` | Insert Caption | mechanical |
| `insert_chart` | Insert Chart | mechanical |
| `insert_citation` | Insert Citation | mechanical |
| `insert_content_control` | Insert Content Control | mechanical |
| `insert_cross_reference` | Insert Cross Reference | mechanical |
| `insert_document` | Insert Document | mechanical |
| `insert_equation` | Insert Equation | mechanical |
| `insert_field` | Insert Field | mechanical |
| `insert_hyperlink` | Insert Hyperlink | mechanical |
| `insert_image` | Insert Image | mechanical |
| `insert_list` | Insert List | mechanical |
| `insert_paragraphs` | Insert Paragraphs | mechanical |
| `insert_reference_list` | Insert Reference List | mechanical |
| `insert_zotero_citation` | Insert Zotero Citation | mechanical |
| `list_elements` | List Elements | mechanical |
| `live_insert_at_cursor` | Live Insert at Cursor | mechanical |
| `live_repair` | Live Repair | mechanical |
| `live_scroll_to` | Live Scroll to Location | first clause of its description |
| `live_set_track_changes` | Live Set Track Changes | mechanical |
| `mail_merge` | Mail Merge | mechanical |
| `manage_backups` | Manage Backups | mechanical |
| `manage_comment` | Manage Comment | mechanical |
| `manage_note` | Manage Note | mechanical |
| `manage_source` | Manage Source | mechanical |
| `mark_index_entry` | Mark Index Entry | mechanical |
| `modify_table_structure` | Modify Table Structure | mechanical |
| `move_section` | Move Section | mechanical |
| `parse_references` | Parse References | mechanical |
| `prepare_for_submission` | Prepare for Submission | mechanical |
| `redact_text` | Redact Text | mechanical |
| `resolve_revisions` | Resolve Revisions | mechanical |
| `search_and_replace` | Search and Replace | mechanical |
| `search_zotero_library` | Search Zotero Library | mechanical |
| `set_bibliography_style` | Set Bibliography Style | mechanical |
| `set_cells` | Set Cells | mechanical |
| `set_chart_data` | Set Chart Data | mechanical |
| `set_content_control` | Set Content Control | mechanical |
| `set_document_properties` | Set Document Properties | mechanical |
| `set_document_protection` | Set Document Protection | mechanical |
| `set_form_fields` | Set Form Fields | mechanical |
| `set_header_footer` | Set Header Footer | mechanical |
| `set_image` | Set Image | mechanical |
| `set_list_numbering` | Set List Numbering | mechanical |
| `set_page_numbers` | Set Page Numbers | mechanical |
| `set_paragraph_format` | Set Paragraph Format | mechanical |
| `set_paragraph_text` | Set Paragraph Text | mechanical |
| `set_section_properties` | Set Section Properties | mechanical |
| `set_table_properties` | Set Table Properties | mechanical |
| `set_textbox_text` | Set Textbox Text | mechanical |
| `set_watermark` | Set Watermark | mechanical |
| `setup_chapter_headers` | Setup Chapter Headers | mechanical |
| `sort_table` | Sort Table | mechanical |
| `split_document` | Split Document | mechanical |
| `structured_diff` | Structured Diff | mechanical |
| `validate` | Validate Document | first clause of its description |
| `word_count` | Word Count | mechanical |
