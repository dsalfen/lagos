# Tie-Out

A Microsoft Word add-in for tying selections in a Word document to specific
locations in PDF evidence. Built for SOX/audit workpaper review.

## Status

Phase 1 — PDF only, no hover preview, click-to-jump navigation.

## Setup

See [SETUP.md](./SETUP.md) for step-by-step instructions to host this on
GitHub Pages and sideload into Word.

## Files

- `taskpane.html` — the entire add-in UI and logic (one file, no build).
- `manifest.xml` — Office add-in manifest. Edit before use.
- `assets/` — ribbon icons.
- `SETUP.md` — installation instructions.

## How it works

Tie-out records are stored as a Custom XML Part inside the `.docx` itself,
so the links travel with the file (email, SharePoint, etc.). Each tied
selection is wrapped in a Word content control tagged `tieout:<id>`; the
matching XML record stores the source PDF, page number, and a normalized
(0–1) coordinate within the page.
