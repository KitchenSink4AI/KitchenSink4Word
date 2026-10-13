# KitchenSink4Word : License Notice

KitchenSink4Word is Copyright (c) 2026 Alvut Consulting, LLC.

## License

KitchenSink4Word is dual-licensed:

**AGPL-3.0 (open source).** Free for anyone (individuals, academics, and
businesses) for any use that complies with the AGPL's terms. Those terms
include sharing source, including your modifications, when you distribute the
software or make it available over a network.

**Commercial license.** For organizations that want to build KitchenSink4Word
into their own products or services without the AGPL's source-sharing
obligations. Contact licensing@kitchensink4.ai.

Copyright (c) 2026 Alvut Consulting, LLC. KitchenSink4AI is a product line of
Alvut Consulting, LLC.

Not affiliated with, endorsed by, or sponsored by Microsoft Corporation.
Microsoft and Word are trademarks of the Microsoft group of companies.

Required Notice: Copyright (c) 2026 Alvut Consulting, LLC
(https://github.com/KitchenSink4AI/KitchenSink4Word)

## Third-party software

KitchenSink4Word does not bundle or redistribute these packages. `pip`
installs each one from its own publisher under its own license.

Required dependencies:

| Package | License | Use |
|---|---|---|
| `fastmcp` | Apache-2.0 | MCP server framework |
| `python-docx` | MIT | .docx object model |
| `lxml` | BSD-3-Clause | XML engine |
| `latex2mathml` | MIT | LaTeX to MathML for equations |
| `mathml2omml` | MIT | MathML to OMML, the format Word stores |
| `regex` | Apache-2.0 AND CNRI-Python | Pattern matching with timeouts |
| `packaging` | Apache-2.0 OR BSD-2-Clause | Version comparison |
| `pywin32` | PSF | Word COM integration (Windows only) |

Development dependencies (not installed for normal use): `pytest` (MIT) and
`pytest-timeout` (MIT).

## Fonts

`docs/fonts/` includes Fraunces and IBM Plex Mono as `.woff2` files for the
documentation page. Both are licensed under the SIL Open Font License 1.1.
The license texts are `docs/fonts/OFL-Fraunces.txt` and
`docs/fonts/OFL-IBMPlexMono.txt`.
