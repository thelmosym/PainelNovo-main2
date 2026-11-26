---
type: project
created: 2026-05-25
updated: 2026-11-25
---

# Project Conventions

## Workbook Backup & Versioning (Painel & Monitoramento)
- Whenever modifying any main workbook (`Painel de controle MemoriaPetrobras *.xlsm` or `Monitoramento/Monitoramento individual_*.xlsm`):
  1. First save an untouched backup copy of the current file inside `Backups/`.
  2. Save the modified workbook as a **new file incrementing the version number at the end of the filename** (e.g., `V.8` → `V.9`, `V.8.6` → `V.8.7`).
  3. Archive/remove the superseded version from `Monitoramento/` (keeping it in `Backups/`) so Power Query `Folder.Files` never ingests duplicate versions, and verify `Shape.OnAction` macro bindings match the new filename.

## Git Workflow
- Always create a new dedicated branch for major code changes.
- Branch name format should follow: `feature/[task-slug]` or `fix/[bug-slug]`.

## Supported AI platforms (AG Kit)
- AG Kit **only supports Gemini CLI and Google Antigravity**.
- Do not claim compatibility with Claude Code, Cursor, Copilot, Windsurf, or other assistants unless the user explicitly expands scope.
- Copy on the website, docs, FAQ, README, and marketing should describe AG Kit as a toolkit for Gemini CLI / Antigravity-style agent setups.
