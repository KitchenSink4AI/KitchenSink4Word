"""Tool titles and the write-shape annotations, declared by name.

Every registered tool ships four MCP annotations. `readOnlyHint` lives in
core/readonly.py and is not repeated here. The other three live here:

* **title** -- the human-readable name a client shows instead of the raw
  tool name. Derived mechanically: the tool name, split on underscores,
  Title Cased, with known acronyms upper-cased, a leading `com_` replaced
  by the host application's name, and short function words kept lowercase
  inside the phrase. A handful of names too short or too generic to read
  as a title take a phrase from the first clause of their own description
  instead; those are the only hand-written strings in the table and
  docs/TOOL_TITLES.md marks each one. Titles are unique within this
  server and none exceeds 40 characters.

* **destructiveHint** -- whether the tool may perform a destructive
  update. ONE rule decides it, and docs/TOOL_ANNOTATIONS.md states the
  rule alongside a per-tool reason so a reviewer can dispute any single
  row:

      false only when every code path either (a) adds new content or a new
      file without replacing anything that was already there, creation
      refusing an existing target, or (b) changes no user data at all
      (this session's tool surface, the viewport, a read performed through
      a hidden or already-running Office instance, a read taken through a
      temporary copy).

      true otherwise. That covers everything that deletes, replaces,
      overwrites, clears, reorders, moves, applies a batch of edits, saves
      over the document the user has open, or writes an output file it may
      silently overwrite. It also covers every tool that writes into an
      existing document, because the pre-write backup these servers take
      is defeatable by the tool's own `backup=False` argument and so
      cannot be claimed as guaranteed reversibility, and everything whose
      reversibility could not be PROVEN by reading the code.

  Read-only tools carry no destructiveHint: the field is meaningful only
  when readOnlyHint is false, and a value there would be noise.

* **idempotentHint** -- true only where repeating the identical call
  obviously lands the same state: the pack switches, whose own docstrings
  say idempotent, and the whole-value setters that take an address and a
  value and carry no action selector. Everything else is left unset rather
  than guessed.

* **openWorldHint** -- false on every tool. Nothing here reaches a remote
  service; the optional update check is not a tool.

An unclassified name RAISES at registration, the same contract
core/readonly.py already holds this surface to.
"""

from __future__ import annotations

#: tool name -> the title a client shows. Unique, at most 40
#: characters, and mirrored into docs/TOOL_TITLES.md, which a test
#: checks against this table so the published English cannot drift
#: from what goes on the wire.
TITLES: dict[str, str] = {
    "anonymize_for_review":       "Anonymize for Review",
    "apply_edits":                "Apply Edits",
    "apply_manuscript_format":    "Apply Manuscript Format",
    "apply_style":                "Apply Style",
    "apply_template":             "Apply Template",
    "assemble_front_matter":      "Assemble Front Matter",
    "change_heading_level":       "Change Heading Level",
    "com_export_pdf":             "Word Export PDF",
    "com_import_pdf":             "Word Import PDF",
    "com_multi_document":         "Word Multi-Document Ops",
    "com_proofing_errors":        "Word Proofing Errors",
    "com_readability_statistics": "Word Readability Statistics",
    "com_refresh_fields":         "Word Refresh Fields",
    "com_save_document":          "Word Save Document",
    "com_validate_opens_clean":   "Word Validate Opens Clean",
    "com_word_status":            "Word Status",
    "comment_report":             "Comment Report",
    "convert_citation_style":     "Convert Citation Style",
    "convert_notes":              "Convert Notes",
    "copy_document":              "Copy Document",
    "copy_table":                 "Copy Table",
    "create_document":            "Create Document",
    "create_table":               "Create Table",
    "deanonymize_document":       "Deanonymize Document",
    "define_style":               "Define Style",
    "delete_element":             "Delete Element",
    "delete_paragraphs":          "Delete Paragraphs",
    "delete_table":               "Delete Table",
    "detect_citation_system":     "Detect Citation System",
    "diagnose_document":          "Diagnose Document",
    "disable_tools":              "Disable Tools",
    "enable_tools":               "Enable Tools",
    "export_images":              "Export Images",
    "export_table":               "Export Table",
    "fill_template":              "Fill Template",
    "find_text":                  "Find Text",
    "fix_accessibility":          "Fix Accessibility",
    "format_cells":               "Format Cells",
    "format_text":                "Format Text",
    "get_comments":               "Get Comments",
    "get_document_info":          "Get Document Info",
    "get_document_view":          "Get Document View",
    "get_headers_footers":        "Get Headers and Footers",
    "get_outline":                "Get Outline",
    "get_paragraph_format":       "Get Paragraph Format",
    "get_protection":             "Get Protection",
    "get_revision_report":        "Get Revision Report",
    "get_server_info":            "Get Server Info",
    "get_styles":                 "Get Styles",
    "get_table":                  "Get Table",
    "get_text":                   "Get Text",
    "get_tracked_changes":        "Get Tracked Changes",
    "get_workflows":              "Get Workflows",
    "import_table":               "Import Table",
    "insert_bookmark":            "Insert Bookmark",
    "insert_break":               "Insert Break",
    "insert_caption":             "Insert Caption",
    "insert_chart":               "Insert Chart",
    "insert_citation":            "Insert Citation",
    "insert_content_control":     "Insert Content Control",
    "insert_cross_reference":     "Insert Cross Reference",
    "insert_document":            "Insert Document",
    "insert_equation":            "Insert Equation",
    "insert_field":               "Insert Field",
    "insert_hyperlink":           "Insert Hyperlink",
    "insert_image":               "Insert Image",
    "insert_list":                "Insert List",
    "insert_paragraphs":          "Insert Paragraphs",
    "insert_reference_list":      "Insert Reference List",
    "insert_zotero_citation":     "Insert Zotero Citation",
    "list_elements":              "List Elements",
    "live_insert_at_cursor":      "Live Insert at Cursor",
    "live_repair":                "Live Repair",
    "live_scroll_to":             "Live Scroll to Location",
    "live_set_track_changes":     "Live Set Track Changes",
    "mail_merge":                 "Mail Merge",
    "manage_backups":             "Manage Backups",
    "manage_comment":             "Manage Comment",
    "manage_note":                "Manage Note",
    "manage_source":              "Manage Source",
    "mark_index_entry":           "Mark Index Entry",
    "modify_table_structure":     "Modify Table Structure",
    "move_section":               "Move Section",
    "parse_references":           "Parse References",
    "prepare_for_submission":     "Prepare for Submission",
    "redact_text":                "Redact Text",
    "resolve_revisions":          "Resolve Revisions",
    "search_and_replace":         "Search and Replace",
    "search_zotero_library":      "Search Zotero Library",
    "set_bibliography_style":     "Set Bibliography Style",
    "set_cells":                  "Set Cells",
    "set_chart_data":             "Set Chart Data",
    "set_content_control":        "Set Content Control",
    "set_document_properties":    "Set Document Properties",
    "set_document_protection":    "Set Document Protection",
    "set_form_fields":            "Set Form Fields",
    "set_header_footer":          "Set Header Footer",
    "set_image":                  "Set Image",
    "set_list_numbering":         "Set List Numbering",
    "set_page_numbers":           "Set Page Numbers",
    "set_paragraph_format":       "Set Paragraph Format",
    "set_paragraph_text":         "Set Paragraph Text",
    "set_section_properties":     "Set Section Properties",
    "set_table_properties":       "Set Table Properties",
    "set_textbox_text":           "Set Textbox Text",
    "set_watermark":              "Set Watermark",
    "setup_chapter_headers":      "Setup Chapter Headers",
    "sort_table":                 "Sort Table",
    "split_document":             "Split Document",
    "structured_diff":            "Structured Diff",
    "validate":                   "Validate Document",
    "word_count":                 "Word Count",
}

#: destructiveHint: true. Deletes, replaces, overwrites, clears,
#: reorders, moves, batch-edits, saves over the open document, writes
#: an output file it may overwrite, acts on a live page, or could not
#: be proven reversible. docs/TOOL_ANNOTATIONS.md carries the reason
#: for every row.
DESTRUCTIVE: frozenset[str] = frozenset({
    "anonymize_for_review", "apply_edits", "apply_manuscript_format",
    "apply_style", "apply_template", "assemble_front_matter",
    "change_heading_level", "com_export_pdf", "com_multi_document",
    "com_refresh_fields", "com_save_document",
    "convert_citation_style", "convert_notes", "deanonymize_document",
    "define_style", "delete_element", "delete_paragraphs",
    "delete_table", "fill_template", "fix_accessibility",
    "format_cells", "format_text", "import_table", "manage_backups",
    "manage_comment", "manage_note", "manage_source",
    "modify_table_structure", "move_section", "prepare_for_submission",
    "redact_text", "resolve_revisions", "search_and_replace",
    "set_bibliography_style", "set_cells", "set_chart_data",
    "set_content_control", "set_document_properties",
    "set_document_protection", "set_form_fields", "set_header_footer",
    "set_image", "set_list_numbering", "set_page_numbers",
    "set_paragraph_format", "set_paragraph_text",
    "set_section_properties", "set_table_properties",
    "set_textbox_text", "set_watermark", "setup_chapter_headers",
    "sort_table"
})

#: destructiveHint: false. Every code path either only ADDS, or
#: changes no user data at all. Each row's reason is in
#: docs/TOOL_ANNOTATIONS.md.
NON_DESTRUCTIVE: frozenset[str] = frozenset({
    "com_import_pdf", "com_proofing_errors",
    "com_readability_statistics", "com_validate_opens_clean",
    "copy_document", "copy_table", "create_document", "create_table",
    "disable_tools", "enable_tools", "export_images", "export_table",
    "get_document_view", "insert_bookmark", "insert_break",
    "insert_caption", "insert_chart", "insert_citation",
    "insert_content_control", "insert_cross_reference",
    "insert_document", "insert_equation", "insert_field",
    "insert_hyperlink", "insert_image", "insert_list",
    "insert_paragraphs", "insert_reference_list",
    "insert_zotero_citation", "live_insert_at_cursor", "live_repair",
    "live_scroll_to", "live_set_track_changes", "mail_merge",
    "mark_index_entry", "search_zotero_library", "split_document"
})

#: idempotentHint: true. Repeating the identical call lands the same
#: state. Anything absent here is left UNSET rather than guessed.
IDEMPOTENT: frozenset[str] = frozenset({
    "disable_tools", "enable_tools", "set_bibliography_style",
    "set_cells", "set_chart_data", "set_content_control",
    "set_document_properties", "set_document_protection",
    "set_form_fields", "set_header_footer", "set_image",
    "set_list_numbering", "set_page_numbers", "set_paragraph_format",
    "set_paragraph_text", "set_section_properties",
    "set_table_properties", "set_textbox_text", "set_watermark"
})


def title(name: str) -> str:
    """The MCP title for one tool. Unknown names RAISE, because a tool that
    reaches tools/list without a title fails the directory's annotation
    requirement and a name-shaped fallback would hide that from the test."""
    try:
        return TITLES[name]
    except KeyError:
        raise RuntimeError(
            f"tool {name!r} has no title in this module. Every registered "
            f"tool needs one: the Anthropic directory requires a title on "
            f"every tool, and a generated fallback would pass the check "
            f"while shipping 'Com Export Pdf' to a user."
        ) from None


def destructive_hint(name: str) -> bool | None:
    """The MCP destructiveHint, or None for a read-only tool, where the
    field carries no meaning. An unclassified mutating name returns None
    and `annotations` raises on it."""
    if name in NON_DESTRUCTIVE:
        return False
    if name in DESTRUCTIVE:
        return True
    return None


def idempotent_hint(name: str) -> bool | None:
    """True where repeating the call lands the same state, else None. Never
    false: an unlisted tool is unclassified, not proven non-idempotent."""
    return True if name in IDEMPOTENT else None


def open_world_hint(name: str) -> bool:
    """False for every tool: this server talks to local files and to a local
    Office installation, never to a remote service."""
    return False


def annotations(name: str, read_only: bool) -> dict:
    """The full annotation dict for one tool, ready for registration.
    destructiveHint is omitted on read-only tools and idempotentHint is
    omitted where it was not classified, so an absent field means
    "undeclared" rather than "false"."""
    ann: dict = {
        "title": title(name),
        "readOnlyHint": read_only,
        "openWorldHint": open_world_hint(name),
    }
    if not read_only:
        destructive = destructive_hint(name)
        if destructive is None:
            raise RuntimeError(
                f"tool {name!r} is not classified DESTRUCTIVE or "
                f"NON_DESTRUCTIVE in this module. Every tool that can "
                f"change something must declare whether the change may be "
                f"destructive before it can be registered."
            )
        ann["destructiveHint"] = destructive
    idempotent = idempotent_hint(name)
    if idempotent is not None:
        ann["idempotentHint"] = idempotent
    return ann
