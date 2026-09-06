"""Which tools can change something, and which cannot.

Every registered tool declares a readOnlyHint annotation, and this module
is where that decision lives. Claude Desktop groups tools by the hint:
the read-only ones land in a "Read-only tools" group the user can approve
once, and everything else keeps asking before it runs. That makes the
classification a safety boundary rather than documentation, so it is
declared here by name and never inferred from what a tool is called.

The bar for true is that the tool cannot change ANYTHING: not the
document, not any other file (temporary files included), not a process,
not what this server exposes. Three exclusions are worth naming, because
all three read as reads. get_document_view takes stamp_anchors, and
stamping writes paragraph identifiers into the file. search_zotero_library
copies the Zotero database to a temporary directory before opening the
copy. The com_ tools each start their own hidden Word and open the
document in it. The live route the dual-mode readers take is different:
it attaches to a Word that is already running and reads, so those tools
stay read-only.

The sibling web server (kitchensink4web) set the pattern and the
discipline: an honest hint buys client-side permission lenience, and an
optimistic one would be a false safety claim in metadata.
"""

from __future__ import annotations


#: Tools that cannot change anything: no document, no file, no process,
#: no server state. These carry readOnlyHint: true.
READ_ONLY: frozenset[str] = frozenset({
    "com_word_status", "comment_report", "detect_citation_system",
    "diagnose_document", "find_text", "get_comments",
    "get_document_info", "get_headers_footers", "get_outline",
    "get_paragraph_format", "get_protection", "get_revision_report",
    "get_styles", "get_table", "get_text", "get_tracked_changes",
    "get_workflows", "list_elements", "parse_references",
    "structured_diff", "validate", "word_count",
})

#: Everything else, declared explicitly rather than inferred. A tool is
#: here because it writes a document, writes a file, starts or drives an
#: Office process, or changes what this server exposes.
MUTATING: frozenset[str] = frozenset({
    "anonymize_for_review", "apply_edits", "apply_manuscript_format",
    "apply_style", "apply_template", "assemble_front_matter",
    "change_heading_level", "com_export_pdf", "com_import_pdf",
    "com_multi_document", "com_proofing_errors",
    "com_readability_statistics", "com_refresh_fields",
    "com_save_document", "com_validate_opens_clean",
    "convert_citation_style", "convert_notes", "copy_document",
    "copy_table", "create_document", "create_table",
    "deanonymize_document", "define_style", "delete_element",
    "delete_paragraphs", "delete_table", "disable_tools",
    "enable_tools", "export_images", "export_table", "fill_template",
    "fix_accessibility", "format_cells", "format_text",
    "get_document_view", "import_table", "insert_bookmark",
    "insert_break", "insert_caption", "insert_chart", "insert_citation",
    "insert_content_control", "insert_cross_reference",
    "insert_document", "insert_equation", "insert_field",
    "insert_hyperlink", "insert_image", "insert_list",
    "insert_paragraphs", "insert_reference_list",
    "insert_zotero_citation", "live_insert_at_cursor", "live_repair",
    "live_scroll_to", "live_set_track_changes", "mail_merge",
    "manage_backups", "manage_comment", "manage_note", "manage_source",
    "mark_index_entry", "modify_table_structure", "move_section",
    "prepare_for_submission", "redact_text", "resolve_revisions",
    "search_and_replace", "search_zotero_library",
    "set_bibliography_style", "set_cells", "set_chart_data",
    "set_content_control", "set_document_properties",
    "set_document_protection", "set_form_fields", "set_header_footer",
    "set_image", "set_page_numbers", "set_paragraph_format",
    "set_paragraph_text", "set_section_properties",
    "set_table_properties", "set_textbox_text", "set_watermark",
    "setup_chapter_headers", "sort_table", "split_document",
})


def read_only_hint(name: str) -> bool:
    """The MCP readOnlyHint for one tool.

    Unknown names RAISE at registration time rather than defaulting.
    Defaulting to false would quietly drop a read tool out of the
    client's read-only group; defaulting to true would put a tool that
    can change a file into a group the user bulk-approves. Neither is a
    decision this module is willing to make on someone's behalf."""
    if name in READ_ONLY:
        return True
    if name in MUTATING:
        return False
    raise RuntimeError(
        f"tool {name!r} is not classified in core/readonly.py. Every "
        f"tool must be declared READ_ONLY or MUTATING before it can be "
        f"registered, because the annotation drives a bulk-approval "
        f"group in the client and an unclassified tool would land in it "
        f"by accident."
    )
