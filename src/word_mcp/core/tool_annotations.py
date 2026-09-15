"""Tool titles, declared by name.

Every registered tool ships a `title` annotation: the Anthropic Connectors
Directory requires one, and a client shows it to a human instead of the
raw tool name. `readOnlyHint` lives in core/readonly.py and is not repeated here.

The title is derived mechanically, so that a new tool gets one without
anyone inventing it: the tool name, split on underscores, Title Cased,
with known acronyms upper-cased, a leading `com_` replaced by the host
application's name, and short function words kept lowercase inside the
phrase. A handful of names too short or too generic to read as a title
take a phrase from the first clause of their own description instead;
those are the only hand-written strings in the table, and
docs/TOOL_TITLES.md marks each one.

Titles are unique within this server and none exceeds 40 characters. An
unknown name RAISES at registration, the same contract core/readonly.py
already holds this surface to.
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


def annotations(name: str, read_only: bool) -> dict:
    """The annotation dict for one tool, ready for registration."""
    return {
        "title": title(name),
        "readOnlyHint": read_only,
    }
