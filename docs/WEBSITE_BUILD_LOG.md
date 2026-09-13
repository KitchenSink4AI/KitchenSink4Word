# Website build log

## 2026-09-13 23:55:21 KST — Product-page polish

Responsive mastheads/cards/footers, corporate/contact/privacy links, localized service routing and shortened footer links. Word/Web teaser lines and redundant release headings removed. Major-feature panels dated2026-09-09 from release records. License/trademark notices retained. PPT alternate theme changed from olive to indigo/slate.

Validation:420 layout states across five pages,7locales,6widths320–1440px and2themes; no page overflow or script errors. Theme controls exercised. Independent markup/i18n checks pass; desktop/mobile screenshots visually inspected. No core/Worker changes. External badges blocked during local render; live validation follows publication.


## 2026-09-14 00:23:41 KST - MCP client compatibility
Broadened installer introduction to Claude Desktop, Claude Code, Codex and local-MCP-capable apps including open-source clients. Added optional Codex commands and generic command/argument guidance in all seven locales. Distinguished ChatGPT desktop local MCP support from separate web remote setup. Preserved Claude-specific bundle/command instructions. Official reference: https://learn.chatgpt.com/docs/extend/mcp and https://developers.openai.com/api/docs/guides/developer-mode ; local codex mcp add --help confirms syntax. No new installation or client settings changes. Validation and copy review pending.

Validation: shared420-state layout matrix passes with setup details expanded, zero page errors/overflow;10 focused desktop/mobile renders match viewport. Codex command syntax checked against installed CLI help. Copy review requested from Claude before publication.


## 2026-09-14 00:43:41 KST - Installation typography and beginner setup
Nyk requested visual consistency, removal of Codex dropdown and combined Claude Code/Codex card. Matched compatibility lead to existing Fraunces typography, secondary copy to13px body. Moved labeled commands into existing install cards/order form. Added shared one-time Windows uv instructions naming Start, PowerShell, paste, Enter and restart; linked official alternatives for Mac/help. Replaced Desktop PATH jargon and unreliable double-click/drag-drop promises with settings-based installation. Desktop card also explains local command/argument setup for compatible apps. Removed Word permission-tip tile per Nyk. Seven locales. Official uv WinGet command verified at https://docs.astral.sh/uv/getting-started/installation/#winget . No installers executed. Tests/review pending.

Validation: final420locale/viewport/theme states pass with zero overflow/page errors. Full installation sections visually reviewed, commands wrap on narrow screens. PPT latest release independently confirmed to contain kitchensink4ppt.mcpb. C101 approves eleven setup strings. Storefront product-settings pointer submitted with frozen candidates. No fresh app installation executed.


## 2026-09-14 01:05:40 KST - Platform labels and beginner desktop routes
Nyk authorizes push/main/live for revised pages. All permission-tip tiles removed. Desktop app routes are styled expandable sections inside existing card, with manual Claude extension import and numbered settings steps for compatible other desktop apps; raw Command/Arguments tile removed. Product hero labels: Word Windows-only, XL/PPT/Web Windows/macOS/Linux. Word directs Mac users to siblings; Web optional OCR Windows-only. File-vs-Office-control limits in installation note. Cleaned retired i18n keys and changed German Enter wording to avoid naive beta-substring guard. Exact platform evidence and review in private report. Checks pending final candidates.

Validation:420final expanded-card locale/viewport/theme states pass;10desktop/mobile full-section renders inspected. Existing guards: PPT site copy OK; Excel22passed; Web26passed2skipped; Word10passed. C104/C105 requested German Enter grammar correction and Word tools wording both applied. Earlier unused-key and beta-substring CI failures repaired without changing test rules.


## 2026-09-14 01:38:14 KST - Desktop-first installation follow-up
C108 usability follow-up within existing website assignment: desktop card first, launcher prerequisite second, CLI route below. Launcher explicitly does not install editors; desktop and manual-client numbered steps require launcher setup and return before proceeding. Direct release asset links remove GitHub Assets navigation. Links pinned to verified current release tags: XL/Web filenames contain versions, so latest/download would break on later release. Future release checklist must refresh these links or adopt stable asset naming. All seven locales updated. No MCP implementation, bundle manifests, installer behavior, corporate or enterprise page changes. Copy review and visual verification pending.

Validation 2026-09-14 01:46:59 KST: C110 approves anchored launcher structure and final copy with product-framed skip applied. Final420layout states pass; desktop/mobile render review completed. Four direct asset URLs return200. Existing guards: PPT OK, Word10 passed, XL22 passed, Web26 passed/2 skipped. CLI anchor direction corrected to above. No end-to-end desktop extension installation performed.


## 2026-09-14 02:01:03 KST - Correct installation intent mismatch
Nyk flagged the previously published structure did not match intended UX. C115 confirms C110 accepted an unintended shared-launcher alternative; C109 self-contained design governs. Removed standalone launcher tile and cross-card navigation. Three complete app disclosures: Claude Desktop open, other apps and CLI closed. Launcher command and guidance embedded inside each path; Claude sequence download, extension import, launcher, restart. Seven locales updated; platform caveats/direct verified downloads retained. Storefront unchanged per spec scope. This corrects user intent, not cache propagation. Checks pending.

Direct Nyk correction during implementation: Claude Desktop must start collapsed like ChatGPT Desktop, with clean equal styling. All three app paths now start collapsed; this supersedes C109/C115 default-open detail.

Validation 2026-09-14 02:07:18 KST: 420layout states pass, explicit three-path/all-collapsed/no-launcher-anchor assertions pass in7languages, desktop/mobile screenshots inspected. Word10/XL22/Web26+2skip copy tests pass; PPTcheck_site passes after removing unused retired inst.form translation key. No runtime tests weakened. Final structure follows direct Nyk collapse correction over C115 open-default.
