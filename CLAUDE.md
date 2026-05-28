# Thais.2 — FROZEN REFERENCE (retired 2026-05-06)

⚠️  This directory is a frozen reference. The live codebase is at:
    ../thais2-ceo/

## What's here

- `tools/app.py.retired` — original Streamlit UI (retired; do not modify)
- `tools/` — Python engine files (expense_engine, pdf_report, etc.)
  These are the SOURCE OF TRUTH copies. The live copies that Vercel uses
  are at ../thais2-ceo/tools/ — keep them in sync if you edit here.
- `data/presets.backup.2026-05-06.json` — disaster-recovery backup of
  classification rules. Do not delete. Do not read programmatically.
  Manual restore only: copy contents into Supabase presets table.

## Active architecture

See ../thais2-ceo/CLAUDE.md for the current architecture notes.
