# Workflow: Fixed Expenses Executive Report (TAAG)

## Objective
Generate a CEO-ready executive summary (PDF) and a detailed Excel breakdown of **fixed expenses** for TAAG, split across the 5 companies, for a user-chosen date range.

## Companies (from `Projeto` column)
1. Rio de Janeiro
2. Alameda 470
3. Arthur
4. Mazzini
5. Alameda 334

## Inputs
- **Source xlsx** uploaded by user via web UI. Single sheet, header on row 10. Key columns: `Pagto` (date), `Banco`, `Cliente/Fornecedor/Favorecido`, `Descrição`, `R$` (negative = debit), `Cód. Despesa`, `Despesas`, `Conta Sintética`, `Projeto`.
- **Plano de Contas** reference: [data/plano_de_contas.xlsx](../data/plano_de_contas.xlsx)
- **Date range**: start + end (DD/MM/AAAA), inclusive, applied to `Pagto`.
- **User preset**: [data/presets.json](../data/presets.json) — last-saved category whitelist & company name normalization.

## Filters / Rules
1. Drop rows where `Pagto` is empty or outside selected range.
2. Drop rows where `R$` >= 0 (we only want debits / fixed expenses).
3. **Exclude all rows tied to "Cesar Valverde"** — match anywhere in `Cliente/Fornecedor/Favorecido`, `Descrição`, or `Conta Sintética` (case-insensitive). These are personal expenses of the owner.
4. Normalize `Projeto` to one of the 5 companies (fuzzy: "alameda 470" / "470" → Alameda 470, "rio" → Rio de Janeiro, etc.). Unrecognized → bucket "Outros" (shown but flagged).
5. **Auto-classify as fixed** using two signals (OR):
   - `Cód. Despesa` is in the fixed-code whitelist (default: 301, 302, 303, 304, 305, 306, 309, 310 — see plano_de_contas).
   - `Descrição` or `Conta Sintética` matches a fixed-keyword (Aluguel, IPTU, Enel, Sabesp, Claro, Net, Vivo, Telefone, Hagana, Limpa vidros, Segurança, Grupo Gabriel, Controle de Pragas, Supricorp, Gimba, Pão de queijo, Água personalizada, Impressora, Cartucho, Seguro Incêndio, Auto de licença, Extintor, Laudo bombeiro, Galão, Garagem, Box, Motoboy, Faxineira, Lalamove, Regus, Helena, Uniart, Ralph, RMF).
6. Vendor ≠ category: e.g., a Regus row whose description says "internet" is still fixed (utility), not rent. Classification reads description, not vendor.

## User Review Step (critical)
After auto-classification, the UI shows a table where the user can:
- Tick / untick individual rows (toggle "is fixed?")
- Tick / untick whole categories
- Add custom keywords that auto-flag future rows
- Save current selection as the new preset

No PDF/Excel is generated until the user clicks **Confirm & Generate**.

## Outputs
1. **PDF** — `Resumo_Executivo_DDMMAAAA_DDMMAAAA.pdf`
   - Cover: TAAG logo, title "Resumo Executivo — Despesas Fixas", period, generation date.
   - Page 2: Total summary (grand total, by company), donut chart by company.
   - Pages 3–7: One page per company — total, top categories, bar chart by category, top 10 vendors.
   - Style: Montserrat font, black + `#34b3d3` cyan accent, white background, generous whitespace.
2. **Excel** — `Despesas_Fixas_DDMMAAAA_DDMMAAAA.xlsx`
   - Sheet 1: "Resumo" — pivot of company × category totals.
   - Sheets 2–6: One per company, listing every row included with full detail.
   - Sheet 7: "Excluídos" — rows the user unticked (audit trail).

## Tools
- [tools/expense_engine.py](../tools/expense_engine.py) — load + clean + classify (pure functions, no UI).
- [tools/pdf_report.py](../tools/pdf_report.py) — PDF generation with ReportLab + matplotlib charts.
- [tools/excel_report.py](../tools/excel_report.py) — Excel generation with openpyxl.
- [tools/app.py](../tools/app.py) — Streamlit UI orchestrating all of the above.

## Deployment
Streamlit app deployed to Streamlit Community Cloud (free) or Render. No login (small audience). The user uploads the latest xlsx each session — file is processed in-memory, never persisted to the server.

## Edge cases & lessons (append as discovered)
- (none yet)
