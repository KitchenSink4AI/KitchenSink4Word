# Dependency license ledger

Every declared dependency, its license, and why it is here. Enforced by
`tests/unit/test_dependency_ledger.py`, which fails the build if
`pyproject.toml` grows a dependency that is not listed here. Ported from
KitchenSink4Web, which carried the only ledger in the family until the
2026-09-15 licence audit (finding D-02).

Licenses below were read from the installed package metadata in this repo's
virtual environment (`importlib.metadata`), not from a search result.

## The rule

**No copyleft and no source-available dependency anywhere in the REQUIRED
install.** Anything questionable lives behind an optional extra, so a
dependency's license never becomes the server's problem. The ship license is
`AGPL-3.0-only` and every permissive license below is one-way compatible into
it, which is the direction that matters: the obligations run to whoever
redistributes those packages, and `pip` resolves each one from its own
publisher with its own license files. KitchenSink4Word redistributes none of
them, so no attribution obligation attaches to the wheel.

## Required install

The package name is the FIRST column and the license the SECOND, in every
table in this file. The enforcing test reads them positionally.

| Package | License | Direction | Note |
|---|---|---|---|
| `fastmcp` | Apache-2.0 | permissive, one-way into anything | The MCP server framework. Pinned `>=3.4,<4` because a minor bump moved the visibility API mid-build. |
| `python-docx` | MIT | permissive | The .docx object model the file tier is built on. |
| `lxml` | BSD-3-Clause | permissive | The XML engine, reached directly in 44 source files. The BSD-3 no-endorsement clause is the only obligation and it binds redistribution, which does not happen here. |
| `latex2mathml` | MIT | permissive | First half of the equation path: LaTeX in, MathML out. |
| `mathml2omml` | MIT | permissive | Second half: MathML to the OMML that Word actually stores. The wheel's `License` metadata field reads `UNKNOWN`, which is a packaging defect rather than an unlicensed release: the classifier says `License :: OSI Approved :: MIT License` and the wheel ships an MIT license text naming amedama, matching the upstream repository. Confirmed 2026-09-15 against the package's own license file, closing the caveat the licence audit raised. |
| `regex` | Apache-2.0 AND CNRI-Python | both permissive | Backs the caller-pattern guard in `ops/_regex.py`. Needed rather than convenient: stdlib `re` has no match timeout, the server is single-threaded stdio, and one pathological caller pattern would deny service to the whole session. |
| `packaging` | Apache-2.0 OR BSD-2-Clause | either, permissive | Version comparison in `core/update_check.py`. |
| `pywin32` | PSF | permissive | The COM tier, which drives a real Word. Declared under a `sys_platform == 'win32'` marker, so it is never installed anywhere it cannot work. |

No copyleft. No obligation triggered by the current distribution model.

## Optional extras

| Package | License | Extra | Why it is optional |
|---|---|---|---|
| `pytest` | MIT | `dev` | Test-time only, never distributed. |
| `pytest-timeout` | MIT | `dev` | Test-time only. A hung COM call is this family's most common CI failure and an unbounded one wedges the runner. |

## Transitive obligations

Transitive dependencies arrive through the packages above and are resolved by
`pip` from their own publishers. None is redistributed by this project, so
none creates an attribution obligation here. A full transitive audit has not
been run; it is the right companion to the one the KitchenSink4Web ledger
also defers.

## Fonts, which are the one real attribution obligation

`docs/fonts/` ships Fraunces and IBM Plex Mono as `.woff2` files on the
documentation page. Both are under the SIL Open Font License 1.1 (OFL),
which requires the copyright notice and license to travel with the font
software. `docs/fonts/OFL-Fraunces.txt` and `docs/fonts/OFL-IBMPlexMono.txt`
are those notices. Unlike every package
above, these files ARE redistributed, which is why the obligation is real
here and nowhere else in this ledger.
