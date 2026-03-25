# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**PL-TEKST** is a VBA add-in for Microsoft Excel that adds the `=PL_TEKST()` worksheet function, converting numeric amounts to Polish text with correct grammatical inflection. Intended for financial documents (invoices, contracts, checks).

Supported range: `0` to `999 999 999,99` PLN.

## Repository Structure

```
src/PL_TEKST.bas     — Single VBA module (the entire source code)
releases/PL_TEKST.xlam — Compiled Excel add-in (binary, pre-built)
docs/instalacja.md   — Detailed installation instructions (Polish)
```

There is no build system, no automated tests, and no CI/CD pipeline.

## Development Workflow

Since this is a VBA project, there is no CLI build or test command. Development cycle:

1. **Edit source**: Modify `src/PL_TEKST.bas` (plain text VBA module).
2. **Test in Excel**: Open Excel, press `ALT+F11`, import or paste the module, run manually.
3. **Rebuild the add-in**: Open Excel, load the module, save as `.xlam` to `releases/PL_TEKST.xlam`.

The `.xlam` binary in `releases/` must be manually rebuilt after source changes — it is not auto-generated.

## Code Architecture

All logic lives in a single VBA module (`src/PL_TEKST.bas`):

- **`PL_TEKST(kwota)`** — Public worksheet function. Validates input, splits the number into millions/thousands/remainder/grosze, assembles Polish text, capitalises the first letter.
- **`TrojkaSlownie(n, zenski)`** — Converts a 3-digit number to Polish words; `zenski=True` applies feminine forms (needed for thousands: *dwie* tysiące, not *dwa* tysiące).
- **`Odmiana(n, f0, f1, f2)`** — Picks the correct grammatical form based on Polish declension rules (singular / 2–4 / 5+ pattern with special case for teens ending in 10–19).
- **`Jednosci` / `Nastki` / `Dziesiatki` / `Setki`** — Return Polish words for units, teens, tens, and hundreds respectively.

### Polish Character Encoding

All Polish diacritics are embedded via `ChrW()` (Unicode code points) to avoid `.bas` file encoding issues:

| Character | ChrW code |
|-----------|-----------|
| ą | 261 |
| ć | 263 |
| ę | 281 |
| ł | 322 |
| ó | 243 |
| ś | 347 |

When adding new Polish words, always use `ChrW()` for diacritics — never paste raw UTF-8 characters into the `.bas` file.

## Roadmap Context

Planned future work (from CHANGELOG.md):
- v1.1 — multi-currency support (EUR, USD, GBP)
- v1.2 — legal invoice format
- v2.0 — web API (Python/FastAPI)
- v2.1 — web application
