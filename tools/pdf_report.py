"""PDF executive summary generator for TAAG fixed expenses.

Structure:
  1. Cover (logo + title + period)
  2. Visão Geral — KPIs, donut by company, narrative
  3. Evolução Mensal (Total) — line chart + narrative
  4. Evolução por Empresa — stacked bar + narrative
  5..N. One page per company — monthly progression line, top categories,
        top vendors, narrative

Outros is excluded from the PDF (per requirement). Excel still keeps it.
"""
from __future__ import annotations

import io
from datetime import date
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from expense_engine import (
    COMPANIES, assign_ceo_category, monthly_by_company, monthly_total,
    summarize_by_company, summarize_by_company_category, top_vendors,
)
from narrative import write_narrative, write_strategic_insights

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
LOGO_PATH = DATA_DIR / "logo.png"

CYAN = colors.HexColor("#34b3d3")
BLACK = colors.HexColor("#000000")
GREY = colors.HexColor("#666666")
LIGHT = colors.HexColor("#f3f3f3")

PALETTE = ["#34b3d3", "#000000", "#666666", "#a8e1ee", "#cccccc", "#1d6e85"]

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Helvetica", "Arial", "DejaVu Sans"],
    "axes.edgecolor": "#000000",
    "axes.labelcolor": "#000000",
    "xtick.color": "#000000",
    "ytick.color": "#000000",
})


def brl(x: float) -> str:
    s = f"R$ {x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _styles():
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("title", parent=base["Title"], fontName="Helvetica-Bold",
                                fontSize=28, textColor=BLACK, leading=34, alignment=1),
        "subtitle": ParagraphStyle("subtitle", parent=base["Normal"], fontName="Helvetica",
                                   fontSize=14, textColor=GREY, leading=18, alignment=1),
        "h1": ParagraphStyle("h1", parent=base["Heading1"], fontName="Helvetica-Bold",
                             fontSize=20, textColor=BLACK, leading=24, spaceAfter=10),
        "h2": ParagraphStyle("h2", parent=base["Heading2"], fontName="Helvetica-Bold",
                             fontSize=13, textColor=CYAN, leading=16, spaceAfter=6, spaceBefore=10),
        "body": ParagraphStyle("body", parent=base["Normal"], fontName="Helvetica",
                               fontSize=10, textColor=BLACK, leading=14),
        "narrative": ParagraphStyle("narrative", parent=base["Normal"], fontName="Helvetica",
                                    fontSize=10.5, textColor=BLACK, leading=15,
                                    leftIndent=12, rightIndent=12,
                                    spaceBefore=8, spaceAfter=14,
                                    borderPadding=0),
        "kpi": ParagraphStyle("kpi", parent=base["Normal"], fontName="Helvetica-Bold",
                              fontSize=20, textColor=CYAN, leading=24, alignment=1),
        "kpi_label": ParagraphStyle("kpi_label", parent=base["Normal"], fontName="Helvetica",
                                    fontSize=9, textColor=GREY, leading=11, alignment=1),
        "footer": ParagraphStyle("footer", parent=base["Normal"], fontName="Helvetica",
                                 fontSize=8, textColor=GREY, alignment=1),
    }


def _img(buf: io.BytesIO, w: float, h: float) -> Image:
    return Image(buf, width=w, height=h)


def _narrative_box(text: str, body_style) -> Table:
    """Wrap narrative text in a bordered box that won't collide with headings."""
    p = Paragraph(text, body_style)
    t = Table([[p]], colWidths=[A4[0] - 4 * cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
        ("LINEBEFORE", (0, 0), (0, -1), 3, CYAN),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return t


def _donut(by_company: pd.DataFrame) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(5.4, 4.0), dpi=180)
    data = by_company[by_company["Total"] > 0]
    if data.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center"); ax.axis("off")
    else:
        wedges, _ = ax.pie(
            data["Total"], labels=None, startangle=90,
            colors=PALETTE[: len(data)],
            wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2),
        )
        total = data["Total"].sum()
        ax.text(0, 0, brl(total), ha="center", va="center", fontsize=11, fontweight="bold")
        ax.legend(
            wedges,
            [f"{e}  •  {brl(t)}" for e, t in zip(data["Empresa"], data["Total"])],
            loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False, fontsize=9,
        )
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
    buf.seek(0); return buf


def _line_progression(monthly: pd.DataFrame, title: str) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(7.5, 3.6), dpi=180)
    if monthly.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center"); ax.axis("off")
    else:
        ax.plot(monthly["Mes"], monthly["Total"], marker="o", linewidth=2.5,
                color="#34b3d3", markerfacecolor="#000000", markersize=7)
        ax.fill_between(range(len(monthly)), monthly["Total"], alpha=0.12, color="#34b3d3")
        for i, (m, t) in enumerate(zip(monthly["Mes"], monthly["Total"])):
            ax.annotate(brl(t), (i, t), textcoords="offset points", xytext=(0, 10),
                        ha="center", fontsize=8, fontweight="bold")
        ax.set_title(title, fontsize=11, fontweight="bold", color="black", pad=12)
        ax.set_ylabel("R$")
        ax.tick_params(axis="x", labelrotation=30, labelsize=9)
        ax.tick_params(axis="y", labelsize=8)
        ax.set_ylim(0, monthly["Total"].max() * 1.25 if monthly["Total"].max() > 0 else 1)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
    buf.seek(0); return buf


def _stacked_bar(monthly_co: pd.DataFrame, title: str) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(7.5, 3.8), dpi=180)
    if monthly_co.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center"); ax.axis("off")
    else:
        bottom = [0] * len(monthly_co)
        for i, col in enumerate(monthly_co.columns):
            ax.bar(monthly_co.index, monthly_co[col], bottom=bottom,
                   color=PALETTE[i % len(PALETTE)], label=col, edgecolor="white", linewidth=0.5)
            bottom = [b + v for b, v in zip(bottom, monthly_co[col])]
        ax.set_title(title, fontsize=11, fontweight="bold", color="black", pad=12)
        ax.set_ylabel("R$")
        ax.tick_params(axis="x", labelrotation=30, labelsize=9)
        ax.tick_params(axis="y", labelsize=8)
        ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1.0), frameon=False, fontsize=8)
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
    buf.seek(0); return buf


def _grouped_bar_categories(df_fixed: pd.DataFrame) -> io.BytesIO:
    """Grouped horizontal bar chart: top categories × companies side-by-side."""
    pivot = (
        df_fixed.assign(V=df_fixed["Valor"].abs())
        .pivot_table(index="Despesas", columns="Empresa", values="V", aggfunc="sum", fill_value=0)
    )
    if pivot.empty:
        fig, ax = plt.subplots(figsize=(8, 4), dpi=160)
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center"); ax.axis("off")
        buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
        buf.seek(0); return buf

    pivot["_total"] = pivot.sum(axis=1)
    pivot = pivot.sort_values("_total", ascending=False).head(12).drop(columns="_total")
    pivot = pivot.iloc[::-1]  # reverse so top category is at top
    companies = list(pivot.columns)
    n_cats = len(pivot)
    n_cos = len(companies)

    bar_h = 0.7 / max(n_cos, 1)
    fig, ax = plt.subplots(figsize=(8.5, max(4.5, n_cats * 0.65 + 1.5)), dpi=160)

    for i, col in enumerate(companies):
        offset = (i - n_cos / 2 + 0.5) * bar_h
        ax.barh(
            [p + offset for p in range(n_cats)],
            pivot[col],
            height=bar_h * 0.88,
            color=PALETTE[i % len(PALETTE)],
            label=col,
        )

    ax.set_yticks(list(range(n_cats)))
    ax.set_yticklabels([str(c)[:32] for c in pivot.index], fontsize=8)
    ax.set_xlabel("R$")
    ax.xaxis.set_major_formatter(
        plt.FuncFormatter(lambda x, _: f"R${x/1000:.0f}k" if x >= 1000 else f"R${x:.0f}")
    )
    ax.legend(bbox_to_anchor=(1.02, 1.0), loc="upper left", frameon=False, fontsize=8)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)
    ax.tick_params(axis="both", labelsize=8)
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
    buf.seek(0); return buf


def _bar_categories(df_cat: pd.DataFrame, title: str) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(7.0, 3.4), dpi=180)
    if df_cat.empty:
        ax.text(0.5, 0.5, "Sem dados", ha="center", va="center"); ax.axis("off")
    else:
        d = df_cat.sort_values("Total", ascending=True).tail(8)
        ax.barh(d["Despesas"].astype(str), d["Total"], color="#34b3d3", edgecolor="black", linewidth=0.5)
        ax.set_xlabel("Valor (R$)")
        ax.set_title(title, fontsize=10, fontweight="bold", color="black")
        for spine in ("top", "right"):
            ax.spines[spine].set_visible(False)
        ax.tick_params(axis="both", labelsize=8)
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", bbox_inches="tight", facecolor="white"); plt.close(fig)
    buf.seek(0); return buf


def _draw_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(GREY)
    canvas.drawCentredString(
        A4[0] / 2, 1.2 * cm,
        f"TAAG Brasil  •  Resumo Executivo Confidencial  •  Página {doc.page}",
    )
    canvas.setStrokeColor(CYAN)
    canvas.setLineWidth(1.5)
    canvas.line(2 * cm, 1.6 * cm, A4[0] - 2 * cm, 1.6 * cm)
    canvas.restoreState()


def build_pdf(df_fixed_all: pd.DataFrame, start: date, end: date) -> bytes:
    """Build the executive summary PDF and return its bytes."""
    # Exclude Outros from PDF entirely (per requirement)
    df_fixed = df_fixed_all[df_fixed_all["Empresa"] != "Outros"].copy()
    df_fixed["CeoCategoria"] = df_fixed.apply(assign_ceo_category, axis=1)

    s = _styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=2 * cm, bottomMargin=2.2 * cm,
        title="Resumo Executivo - Despesas Fixas TAAG",
    )
    story = []
    period = f"{start.strftime('%d/%m/%Y')}  —  {end.strftime('%d/%m/%Y')}"

    # ---- Cover ----
    if LOGO_PATH.exists():
        story.append(Spacer(1, 4 * cm))
        story.append(Image(str(LOGO_PATH), width=7 * cm, height=3.07 * cm))
    story.append(Spacer(1, 1.5 * cm))
    story.append(Paragraph("Resumo Executivo", s["title"]))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Paragraph("Despesas Fixas", s["title"]))
    story.append(Spacer(1, 1.2 * cm))
    story.append(Paragraph(period, s["subtitle"]))
    story.append(Spacer(1, 6 * cm))
    story.append(Paragraph(
        f"Gerado em {date.today().strftime('%d/%m/%Y')}  •  Confidencial", s["footer"]))
    story.append(PageBreak())

    if df_fixed.empty:
        story.append(Paragraph("Nenhuma despesa fixa atribuída a uma empresa no período.", s["body"]))
        doc.build(story, onLaterPages=_draw_footer)
        return buf.getvalue()

    by_co = summarize_by_company(df_fixed)
    grand_total = float(df_fixed["Valor"].abs().sum())
    n_rows = len(df_fixed)
    n_companies = int((by_co["Total"] > 0).sum())

    # ---- Page: Visão Geral ----
    story.append(Paragraph("Visão Geral", s["h1"]))
    story.append(Paragraph(period, s["body"]))
    story.append(Spacer(1, 0.4 * cm))

    kpi_table = Table(
        [[Paragraph(brl(grand_total), s["kpi"]),
          Paragraph(str(n_rows), s["kpi"]),
          Paragraph(str(n_companies), s["kpi"])],
         [Paragraph("Total Despesas Fixas", s["kpi_label"]),
          Paragraph("Lançamentos", s["kpi_label"]),
          Paragraph("Empresas", s["kpi_label"])]],
        colWidths=[(A4[0] - 4 * cm) / 3] * 3,
    )
    kpi_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 0.6, CYAN),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 0.5 * cm))

    story.append(Paragraph("Distribuição por Empresa", s["h2"]))
    story.append(_img(_donut(by_co), 15 * cm, 8.5 * cm))

    # Narrative for the whole period
    monthly_all = monthly_total(df_fixed)
    by_cat_all = (
        df_fixed.assign(V=df_fixed["Valor"].abs())
        .groupby("Despesas", as_index=False)["V"].sum()
        .rename(columns={"V": "Total"}).sort_values("Total", ascending=False)
    )
    by_vendor_all = (
        df_fixed.assign(V=df_fixed["Valor"].abs())
        .groupby("Favorecido", as_index=False)["V"].sum()
        .rename(columns={"V": "Total"}).sort_values("Total", ascending=False)
    )
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Análise do Período", s["h2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_narrative_box(
        write_narrative(monthly_all, by_cat_all, by_vendor_all,
                        scope="o consolidado de todas as empresas TAAG no período"),
        s["body"],
    ))
    story.append(PageBreak())

    # ---- Page: Evolução Mensal Total ----
    story.append(Paragraph("Evolução Mensal — Consolidado", s["h1"]))
    story.append(Paragraph(period, s["body"]))
    story.append(Spacer(1, 0.4 * cm))
    story.append(_img(_line_progression(monthly_all, "Despesas Fixas Mês a Mês"),
                      17 * cm, 8 * cm))
    story.append(Spacer(1, 0.3 * cm))
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Leitura do Gráfico", s["h2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_narrative_box(
        write_narrative(monthly_all, by_cat_all, by_vendor_all,
                        scope="a evolução mensal das despesas fixas consolidadas"),
        s["body"],
    ))
    # Monthly table
    if not monthly_all.empty:
        rows = [["Mês", "Total", "Variação"]]
        prev = None
        for _, r in monthly_all.iterrows():
            if prev is None:
                var = "—"
            else:
                pct = ((r["Total"] - prev) / prev * 100) if prev else 0
                arrow = "▲" if pct > 0 else ("▼" if pct < 0 else "→")
                var = f"{arrow} {pct:+.1f}%"
            rows.append([r["Mes"], brl(r["Total"]), var])
            prev = r["Total"]
        t = Table(rows, colWidths=[5 * cm, 5 * cm, 4 * cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), CYAN),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
        ]))
        story.append(Spacer(1, 0.3 * cm))
        story.append(t)
    story.append(PageBreak())

    # ---- Page: Evolução por Empresa ----
    story.append(Paragraph("Evolução Mensal por Empresa", s["h1"]))
    story.append(Paragraph(period, s["body"]))
    story.append(Spacer(1, 0.4 * cm))
    monthly_co = monthly_by_company(df_fixed)
    story.append(_img(_stacked_bar(monthly_co, "Composição mês a mês por empresa"),
                      17 * cm, 9 * cm))
    story.append(Spacer(1, 0.4 * cm))
    story.append(Paragraph("Leitura do Gráfico", s["h2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_narrative_box(
        write_narrative(monthly_all, by_cat_all, by_vendor_all,
                        scope="a composição mensal por empresa, comparando crescimento e participação"),
        s["body"],
    ))
    story.append(PageBreak())

    # ---- Page: Comparativo por Categoria × Empresa ----
    story.append(Paragraph("Comparativo por Categoria", s["h1"]))
    story.append(Paragraph(period, s["body"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(Paragraph(
        "Quanto cada endereço gasta em cada categoria de despesa fixa. "
        "As barras permitem comparar diretamente qual local tem o maior custo por item.",
        s["body"],
    ))
    story.append(Spacer(1, 0.4 * cm))
    story.append(_img(_grouped_bar_categories(df_fixed), 17 * cm, 12 * cm))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("Tabela Resumo — Categoria × Empresa (R$)", s["h2"]))

    # Build pivot table for PDF
    _pv = (
        df_fixed.assign(V=df_fixed["Valor"].abs())
        .pivot_table(index="Despesas", columns="Empresa", values="V", aggfunc="sum", fill_value=0)
    )
    _pv["_total"] = _pv.sum(axis=1)
    _pv = _pv.sort_values("_total", ascending=False).drop(columns="_total")
    _co_cols = [c for c in COMPANIES if c in _pv.columns]
    _pv = _pv[_co_cols] if _co_cols else _pv

    if not _pv.empty:
        _col_w = (A4[0] - 4 * cm) / (len(_co_cols) + 2)
        _tbl_rows = [["Categoria"] + _co_cols + ["Total"]]
        for cat, row in _pv.iterrows():
            _row_vals = [str(cat)[:30]]
            for c in _co_cols:
                v = float(row.get(c, 0))
                _row_vals.append(brl(v) if v > 0 else "—")
            _row_vals.append(brl(float(row[_co_cols].sum())))
            _tbl_rows.append(_row_vals)
        # Grand total row
        _gt_row = ["TOTAL"]
        for c in _co_cols:
            _gt_row.append(brl(float(_pv[c].sum())))
        _gt_row.append(brl(float(_pv[_co_cols].values.sum())))
        _tbl_rows.append(_gt_row)

        _t = Table(_tbl_rows, colWidths=[_col_w * 2] + [_col_w] * len(_co_cols) + [_col_w])
        _style = [
            ("BACKGROUND", (0, 0), (-1, 0), CYAN),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), BLACK),
            ("TEXTCOLOR", (0, -1), (-1, -1), colors.white),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, LIGHT]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]
        # Bold the max-value cell in each data row
        for r_idx, (_, row) in enumerate(_pv.iterrows(), start=1):
            row_vals = [float(row.get(c, 0)) for c in _co_cols]
            if any(v > 0 for v in row_vals):
                max_c_idx = row_vals.index(max(row_vals)) + 1  # +1 for Categoria col
                _style.append(("FONTNAME", (max_c_idx, r_idx), (max_c_idx, r_idx), "Helvetica-Bold"))
                _style.append(("TEXTCOLOR", (max_c_idx, r_idx), (max_c_idx, r_idx), CYAN))
        _t.setStyle(TableStyle(_style))
        story.append(_t)
    story.append(PageBreak())

    # ---- Per-company pages ----
    for empresa in COMPANIES:
        sub = df_fixed[df_fixed["Empresa"] == empresa]
        if sub.empty:
            continue
        total = float(sub["Valor"].abs().sum())
        story.append(Paragraph(empresa, s["h1"]))
        story.append(Paragraph(period, s["body"]))
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph(
            f"<b>Total no período:</b> <font color='#34b3d3'>{brl(total)}</font>  "
            f"&nbsp;&nbsp;&nbsp;<b>Lançamentos:</b> {len(sub)}", s["body"]))
        story.append(Spacer(1, 0.3 * cm))

        # Monthly progression
        monthly_e = monthly_total(df_fixed, empresa=empresa)
        story.append(Paragraph("Evolução Mensal", s["h2"]))
        story.append(_img(_line_progression(monthly_e, ""), 16 * cm, 6.5 * cm))

        # CEO categories with % breakdown
        cat = (
            sub.assign(V=sub["Valor"].abs())
            .groupby("CeoCategoria", as_index=False)["V"].sum()
            .rename(columns={"V": "Total", "CeoCategoria": "Despesas"})
            .sort_values("Total", ascending=False)
        )
        story.append(Paragraph("Categorias de Despesa", s["h2"]))
        story.append(_img(_bar_categories(cat, ""), 16 * cm, 6.5 * cm))

        # Category table with % of location total
        if not cat.empty:
            cat_rows = [["Categoria", "Total (R$)", "% do Local"]]
            for _, cr in cat.iterrows():
                pct = float(cr["Total"]) / total * 100 if total else 0
                cat_rows.append([str(cr["Despesas"]), brl(float(cr["Total"])), f"{pct:.1f}%"])
            ct = Table(cat_rows, colWidths=[9 * cm, 4 * cm, 3 * cm])
            ct.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), CYAN),
                ("TEXTCOLOR",  (0, 0), (-1, 0), colors.white),
                ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN",      (1, 0), (-1, -1), "RIGHT"),
                ("FONTSIZE",   (0, 0), (-1, -1), 9),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING",    (0, 0), (-1, -1), 4),
            ]))
            story.append(ct)
            story.append(Spacer(1, 0.3 * cm))

        # Top vendors
        tv = top_vendors(df_fixed, empresa, n=8)
        story.append(Paragraph("Top Fornecedores", s["h2"]))
        if tv.empty:
            story.append(Paragraph("Sem dados.", s["body"]))
        else:
            rows = [["Fornecedor", "Total"]]
            for _, r in tv.iterrows():
                name = (str(r["Favorecido"])[:55] + "…") if len(str(r["Favorecido"])) > 56 else str(r["Favorecido"])
                rows.append([name, brl(r["Total"])])
            t = Table(rows, colWidths=[12 * cm, 4 * cm])
            t.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), BLACK),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT]),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
            ]))
            story.append(t)

        # Narrative for this company
        story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph("Análise — " + empresa, s["h2"]))
        story.append(Spacer(1, 0.2 * cm))
        by_vendor_e = (
            sub.assign(V=sub["Valor"].abs())
            .groupby("Favorecido", as_index=False)["V"].sum()
            .rename(columns={"V": "Total"})
            .sort_values("Total", ascending=False)
        )
        story.append(_narrative_box(
            write_narrative(monthly_e, cat, by_vendor_e,
                            scope=f"as despesas fixas da empresa {empresa}"),
            s["body"],
        ))
        story.append(PageBreak())

    # ---- Page: Strategic Insights for CEO ----
    story.append(Paragraph("Análise Estratégica para o CEO", s["h1"]))
    story.append(Paragraph(period, s["body"]))
    story.append(Spacer(1, 0.4 * cm))

    # Key metrics table
    by_co_s = summarize_by_company(df_fixed)
    by_ceo_s = (
        df_fixed.assign(V=df_fixed["Valor"].abs())
        .groupby("CeoCategoria", as_index=False)["V"].sum()
        .rename(columns={"V": "Total"})
        .sort_values("Total", ascending=False)
    )
    monthly_s = monthly_total(df_fixed)
    n_outros_s = int((df_fixed_all["Empresa"] == "Outros").sum())

    story.append(Paragraph("Onde a empresa está gastando mais", s["h2"]))
    if not by_ceo_s.empty:
        insight_rows = [["Categoria", "Total", "% do Total"]]
        for _, r in by_ceo_s.head(8).iterrows():
            pct = float(r["Total"]) / grand_total * 100 if grand_total else 0
            insight_rows.append([str(r["CeoCategoria"]), brl(float(r["Total"])), f"{pct:.1f}%"])
        it = Table(insight_rows, colWidths=[9 * cm, 4 * cm, 3 * cm])
        it.setStyle(TableStyle([
            ("BACKGROUND",    (0, 0), (-1, 0), BLACK),
            ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
            ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
            ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
            ("FONTSIZE",      (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS",(0, 1), (-1, -1), [colors.white, LIGHT]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ]))
        story.append(it)

    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("Recomendações Estratégicas", s["h2"]))
    story.append(Spacer(1, 0.2 * cm))
    story.append(_narrative_box(
        write_strategic_insights(
            grand_total=grand_total,
            by_co=by_co_s,
            by_ceo=by_ceo_s,
            monthly=monthly_s,
            n_outros=n_outros_s,
        ),
        s["body"],
    ))

    if n_outros_s > 0:
        story.append(Spacer(1, 0.3 * cm))
        story.append(Paragraph("Atenção Operacional", s["h2"]))
        story.append(_narrative_box(
            f"{n_outros_s} lançamento(s) não têm localidade definida (classificados como 'Outros'). "
            "Estes valores não aparecem nos totais por empresa e podem distorcer a análise. "
            "Recomenda-se revisar e atribuir cada lançamento à sua unidade correspondente.",
            s["body"],
        ))

    doc.build(story, onLaterPages=_draw_footer)
    return buf.getvalue()
