## KitchenSink4Word Quickstart

**Install.** Claude Desktop: the KitchenSink4Word extension; two checkboxes,
both fine off. Anywhere else: `uvx kitchensink4word`.

**First edit.** Ask Claude: "Open report.docx and fix the citations in
section 2." It reads the outline, makes the edit after an automatic backup,
and verifies the save. If the document is open in your Word, it can edit it
live while you watch; if a change cannot be done safely, it says what stood
in the way instead of guessing.

**The habits that matter.** Point Claude at files by full path the first
time. Backups land in the document's folder before every change, so undo is
always a file away. `get_server_info` reports the server's health, version,
and whether an update exists (checked at most weekly, off with
KS4W_UPDATE_CHECK=off). If Claude says a tool is missing, the full set loads
with the "Load every tool" checkbox or mid-session on request.
