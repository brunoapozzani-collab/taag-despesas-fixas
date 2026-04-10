"""Excel detailed report generator."""
from __future__ import annotations

import io
from datetime import date

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from expense_engine import COMPANIES

CYAN = "FF34B3D3"
BLACK = "FF000000"
LIGHT = "FFF3F3F3"
WHITE = "FFFFFFFF"

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=WHITE)
TITLE_FONT = Font(name="Calibri", size=14, bold=True, color=BLACK)
BODY_FONT = Font(name="Calibri", size=10, color=BLACK)
HEADER_FILL = PatternFill("solid", fgColor=CYAN)
BLACK_FILL = PatternFill("solid", fgColor=BLACK)
LIGHT_FILL = PatternFill("solid", fgColor=LIGHT)
THIN = Side(style="thin", color="FFCCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DETAIL_COLS = [
    ("Pagto", "Data"),
    ("Banco", "Banco"),
    ("Favorecido", "Fornecedor"),
    ("Descricao", "Descrição"),
    ("CodDespesa", "Cód."),
    ("Despesas", "Categoria"),
    ("ContaSintetica", "Conta Sintética"),
    ("Valor", "Valor (R$)"),
]


def _style_header(ws, row: int, n_cols: int):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def _autosize(ws):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 10
        for cell in col:
            if cell.value is not None:
                max_len = max(max_len, min(60, len(str(cell.value)) + 2))
        ws.column_dimensions[col_letter].width = max_len


def _write_detail(ws, df: pd.DataFrame, title: str, period: str):
    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.cell(row=2, column=1, value=period).font = BODY_FONT
    headers = [h for _, h in DETAIL_COLS]
    for i, h in enumerate(headers, start=1):
        ws.cell(row=4, column=i, value=h)
    _style_header(ws, 4, len(headers))

    r = 5
    for _, row in df.iterrows():
        for i, (key, _) in enumerate(DETAIL_COLS, start=1):
            val = row.get(key)
            if key == "Pagto" and pd.notna(val):
                val = pd.to_datetime(val).strftime("%d/%m/%Y")
            elif key == "Valor" and pd.notna(val):
                val = float(abs(val))
            if val is pd.NA or (isinstance(val, float) and pd.isna(val)) or (hasattr(pd, "isna") and pd.isna(val) is True):
                val = None
            cell = ws.cell(row=r, column=i, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            if key == "Valor":
                cell.number_format = 'R$ #,##0.00'
                cell.alignment = Alignment(horizontal="right")
        r += 1

    # Total row
    total = float(df["Valor"].abs().sum()) if not df.empty else 0.0
    ws.cell(row=r + 1, column=1, value="TOTAL").font = Font(bold=True)
    last = ws.cell(row=r + 1, column=len(headers), value=total)
    last.font = Font(bold=True)
    last.number_format = 'R$ #,##0.00'
    last.fill = LIGHT_FILL

    ws.freeze_panes = "A5"
    _autosize(ws)


def build_excel(df_fixed: pd.DataFrame, df_excluded: pd.DataFrame, start: date, end: date) -> bytes:
    wb = Workbook()
    period = f"Período: {start.strftime('%d/%m/%Y')} a {end.strftime('%d/%m/%Y')}"

    # ---- Resumo sheet ----
    ws = wb.active
    ws.title = "Resumo"
    ws.cell(row=1, column=1, value="TAAG — Despesas Fixas").font = TITLE_FONT
    ws.cell(row=2, column=1, value=period).font = BODY_FONT

    if df_fixed.empty:
        ws.cell(row=4, column=1, value="Sem dados no período selecionado.").font = BODY_FONT
    else:
        pivot = (
            df_fixed.assign(V=df_fixed["Valor"].abs())
            .pivot_table(index="Despesas", columns="Empresa", values="V", aggfunc="sum", fill_value=0)
        )
        # ensure column order
        for c in COMPANIES + ["Outros"]:
            if c not in pivot.columns:
                pivot[c] = 0
        cols = [c for c in COMPANIES + ["Outros"] if c in pivot.columns]
        pivot = pivot[cols]
        pivot["Total"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("Total", ascending=False)

        headers = ["Categoria"] + cols + ["Total"]
        for i, h in enumerate(headers, start=1):
            ws.cell(row=4, column=i, value=h)
        _style_header(ws, 4, len(headers))

        r = 5
        for cat, row in pivot.iterrows():
            ws.cell(row=r, column=1, value=cat).font = BODY_FONT
            for i, c in enumerate(cols, start=2):
                cell = ws.cell(row=r, column=i, value=float(row[c]))
                cell.number_format = 'R$ #,##0.00'
                cell.font = BODY_FONT
                cell.border = BORDER
            tcell = ws.cell(row=r, column=len(headers), value=float(row["Total"]))
            tcell.font = Font(bold=True)
            tcell.number_format = 'R$ #,##0.00'
            tcell.fill = LIGHT_FILL
            r += 1

        # Grand total row
        ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
        for i, c in enumerate(cols, start=2):
            cell = ws.cell(row=r, column=i, value=float(pivot[c].sum()))
            cell.font = Font(bold=True)
            cell.number_format = 'R$ #,##0.00'
            cell.fill = LIGHT_FILL
        gt = ws.cell(row=r, column=len(headers), value=float(pivot["Total"].sum()))
        gt.font = Font(bold=True, color=WHITE)
        gt.fill = PatternFill("solid", fgColor=BLACK[2:])  # openpyxl wants RGB
        gt.fill = BLACK_FILL
        gt.number_format = 'R$ #,##0.00'

        ws.freeze_panes = "A5"
    _autosize(ws)

    # ---- Per-company sheets ----
    for empresa in COMPANIES + ["Outros"]:
        sub = df_fixed[df_fixed["Empresa"] == empresa]
        if sub.empty:
            continue
        sheet_name = empresa[:31]
        wsc = wb.create_sheet(sheet_name)
        _write_detail(wsc, sub, f"{empresa} — Despesas Fixas", period)

    # ---- Excluded sheet ----
    if not df_excluded.empty:
        wsx = wb.create_sheet("Excluídos")
        _write_detail(wsx, df_excluded, "Linhas Excluídas pelo Usuário", period)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
