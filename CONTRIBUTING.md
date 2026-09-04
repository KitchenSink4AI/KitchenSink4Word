# Contributing to KitchenSink4Word

Contributions are welcome: bug reports, fixes, tests, documentation and feature proposals.

## Before your first pull request

KitchenSink4Word is dual-licensed (AGPL-3.0, with commercial licenses available), and a Contributor License Agreement keeps that possible. Read [CLA.md](CLA.md) and include this line in the description of your first pull request:

    I have read the CLA.md of this repository and agree to its terms.

We cannot merge a pull request without it.

## Reporting bugs

Include the tool name, the exact call with its parameters, what happened and what you expected. A small file that reproduces the problem helps a great deal, but please rebuild it with placeholder content rather than sharing anything private. Security problems go to the address in [SECURITY.md](SECURITY.md), not to a public issue.

## Making changes

- Every change in behavior needs a test, and the suite must stay green:

  ```sh
  pytest tests/unit -m "not live"
  ```

  The excluded tests run separately (live tests need Microsoft Word installed and closed).
- Match the style of the surrounding code. Tool descriptions are kept short on purpose, and a test checks their size, so run it after editing one.
- Counts in the README and documentation come from the scripts in `scripts/`. Re-run them rather than editing a number by hand.

## What gets accepted

Small, well-tested fixes are merged fastest. For a larger feature, please open an issue describing the use case before writing code. The tool set is deliberately compact, so a new tool needs a strong reason to exist.
