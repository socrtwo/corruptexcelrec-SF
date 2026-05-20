# CLAUDE.md

S2 Recovery Tools for Microsoft Excel — recover and repair corrupt `.xls`
and `.xlsx` files. Two implementations live here: a **cross-platform PWA**
under `web/` (the canonical user-facing app) and a **legacy VB.NET WinForms
tool** under `Excel Recovery/` that keeps the Excel COM Interop recovery
routines for the most stubborn corruption cases. Assume edits target `web/`
unless told otherwise.

## Repo map

- `web/` — the PWA (HTML / JS / manifest / service worker). Powers the
  live app, the Web release, and every cross-platform release archive.
- `Excel Recovery/`, `Excel Recovery.sln` — legacy VB.NET WinForms app
  (Visual Studio solution). Windows-only; bundles Excel COM Interop and
  CLI utilities for deep recovery passes.
- `releases/` — pre-packaged release archives committed to the repo.
- `scripts/` — release packaging helpers.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy `web/` to
  Pages on push to `main`), `release.yml` (build per-platform zips on
  `v*` tag).

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. The Pages deploy and Release pipelines fire from `main`, so nothing
   ships until the PR lands.

## Releasing

- Push a `v*` tag to `main` (or use Actions → Release → workflow_dispatch)
  to produce platform zips. The Windows release bundles both the PWA and
  the legacy WinForms `.exe`; other platforms ship the PWA only.

## Verifying changes

- PWA: serve `web/` locally (`python3 -m http.server` from inside `web/`)
  and exercise the six recovery strategies (Auto, Strict Read, Lenient
  XML Repair, ZIP Recovery, Salvage Cells, Convert to CSV) against a
  corrupt-xlsx fixture.
- VB.NET app: open `Excel Recovery.sln` in Visual Studio and build the
  WinForms project. CI on `build.yml` validates this build.
- Test save formats: `.xlsx`, `.xls`, `.ods`, `.csv` — a regression in
  any one of them tends to be subtle.

## Gotchas

- 100% client-side is a hard requirement: the PWA must never upload files
  anywhere. Don't introduce `fetch()` of user data.
- The legacy VB.NET app uses Excel COM Interop — it requires a real
  Excel installation on the target machine. Don't assume Interop calls
  are available everywhere; they aren't even on the build CI runners
  without a step that installs Office.
- Service worker caches the entire app shell for offline use. Bump the
  cache version in `web/service-worker.js` whenever you change cached
  assets, or returning users serve stale code.
- Six recovery strategies share UI but not logic. A bug in one strategy
  rarely indicates a bug in the others — fix narrowly.
