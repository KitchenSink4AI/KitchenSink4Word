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
