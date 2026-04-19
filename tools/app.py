"""TAAG Despesas Fixas — Streamlit UI.

Run locally:    streamlit run tools/app.py
"""
from __future__ import annotations

import json
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

# Make sibling tool modules importable when launched from the project root
sys.path.insert(0, str(Path(__file__).resolve().parent))

from expense_engine import (
    COMPANIES, DEFAULT_FIXED_KEYWORDS, Preset, apply_vendor_map, auto_classify_fixed,
    exclude_personal, filter_by_date, load_workbook_dataframe, only_debits,
    summarize_by_company, summarize_by_company_category,
)
from excel_report import build_excel
from pdf_report import brl, build_pdf

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
LOGO = DATA_DIR / "logo.png"

st.set_page_config(
    page_title="TAAG • Despesas Fixas",
    page_icon=str(DATA_DIR / "symbol.png") if (DATA_DIR / "symbol.png").exists() else "💼",
    layout="wide",
)

# ---- Custom CSS ----
st.markdown(
    """
    <style>
    :root { --taag-cyan: #34b3d3; }
    .stApp { background: #ffffff; }
    h1, h2, h3 { color: #000000; font-family: 'Montserrat', sans-serif; }
    .stButton>button {
        background: #34b3d3; color: white; border: none; border-radius: 6px;
        padding: 0.55rem 1.2rem; font-weight: 600;
    }
    .stButton>button:hover { background: #1d6e85; color: white; }
    .stDownloadButton>button {
        background: #000000; color: white; border: none; border-radius: 6px;
        padding: 0.55rem 1.2rem; font-weight: 600;
    }
    .kpi {
        background: #f8fbfc; border-left: 4px solid #34b3d3;
        padding: 14px 18px; border-radius: 6px; margin-bottom: 8px;
    }
    .kpi .v { font-size: 28px; font-weight: 700; color: #000; }
    .kpi .l { font-size: 12px; color: #666; text-transform: uppercase; letter-spacing: 0.5px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---- Header ----
col_logo, col_title = st.columns([1, 5])
with col_logo:
    if LOGO.exists():
        st.image(str(LOGO), width=140)
with col_title:
    st.markdown("## Despesas Fixas — Resumo Executivo")
    st.caption("Faça o upload da planilha, escolha o período, revise as despesas fixas e gere o relatório.")

st.divider()

# ---- Session state ----
if "preset" not in st.session_state:
    st.session_state.preset = Preset.load()
if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
if "df_review" not in st.session_state:
    st.session_state.df_review = None

# ---- Sidebar ----
with st.sidebar:
    st.header("1. Planilha")
    upload = st.file_uploader("Arraste a planilha (.xlsx)", type=["xlsx"])
    if upload is not None:
        try:
            with st.spinner("Carregando planilha…"):
                st.session_state.df_raw = load_workbook_dataframe(upload)
            st.success(f"{len(st.session_state.df_raw):,} linhas carregadas.")
        except Exception as e:
            st.error(f"Erro ao ler a planilha: {e}")

    st.header("2. Datas")
    MESES_PT = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro",
    ]
    YEARS = list(range(2020, 2036))
    today = date.today()
    default_start = date(today.year, today.month, 1) - timedelta(days=1)
    default_start = date(default_start.year, default_start.month, 1)

    def _date_picker(label: str, default: date, key_prefix: str) -> date:
        st.markdown(f"**{label}**")
        c1, c2, c3 = st.columns([1, 2, 1])
        with c1:
            day = st.selectbox(
                "Dia", list(range(1, 32)),
                index=default.day - 1, key=f"{key_prefix}_d", label_visibility="collapsed",
            )
        with c2:
            month = st.selectbox(
                "Mês", MESES_PT,
                index=default.month - 1, key=f"{key_prefix}_m", label_visibility="collapsed",
            )
        with c3:
            year = st.selectbox(
                "Ano", YEARS,
                index=YEARS.index(default.year) if default.year in YEARS else len(YEARS) - 1,
                key=f"{key_prefix}_y", label_visibility="collapsed",
            )
        m_num = MESES_PT.index(month) + 1
        # Clamp day if invalid (e.g., 31 fev -> 28/29)
        from calendar import monthrange
        max_d = monthrange(year, m_num)[1]
        d = min(day, max_d)
        return date(year, m_num, d)

    start = _date_picker("Data inicial", default_start, "start")
    end = _date_picker("Data final", today, "end")

    st.caption("_Suas preferências (palavras-chave, regras de fornecedor, ajustes manuais) são salvas automaticamente._")

    preset = st.session_state.preset
    with st.expander("⚙️ Configurações avançadas", expanded=False):
        st.markdown("**Palavras-chave que identificam despesas fixas**")
        edited_kws = st.text_area(
            "Uma por linha", value="\n".join(preset.fixed_keywords), height=180,
            label_visibility="collapsed",
        )
        if st.button("Atualizar palavras-chave", use_container_width=True):
            preset.fixed_keywords = [k.strip().lower() for k in edited_kws.splitlines() if k.strip()]
            preset.save()
            st.session_state.df_review = None
            st.success("Lista atualizada.")
            st.rerun()

        st.divider()
        if st.button("↺ Restaurar padrões de fábrica", use_container_width=True):
            st.session_state.preset = Preset()
            st.session_state.preset.save()
            st.session_state.df_review = None
            st.success("Configurações restauradas.")
            st.rerun()

# ---- Main ----
if st.session_state.df_raw is None:
    st.info("⬅️ Comece fazendo o upload da planilha na barra lateral.")
    st.stop()

if start > end:
    st.error("Data inicial maior que a final.")
    st.stop()

# Pipeline
df = st.session_state.df_raw
df = filter_by_date(df, start, end)
df = exclude_personal(df)
df = only_debits(df)

if df.empty:
    st.warning("Nenhuma linha encontrada no período selecionado (após exclusões).")
    st.stop()

df = apply_vendor_map(df, st.session_state.preset)
df_classified = auto_classify_fixed(df, st.session_state.preset)
if st.session_state.df_review is None:
    st.session_state.df_review = df_classified.copy()
else:
    # Re-merge keeping any user toggles via row_id
    prev = st.session_state.df_review.set_index("row_id")["is_fixed"]
    df_classified["is_fixed"] = df_classified.apply(
        lambda r: prev.get(r["row_id"], r["is_fixed"]), axis=1
    )
    st.session_state.df_review = df_classified.copy()

tab_review, tab_summary, tab_audit, tab_generate = st.tabs(
    ["📋 Revisar Despesas", "📊 Resumo", "🔎 Auditoria", "📥 Gerar Relatórios"]
)

# ---- Review tab ----
with tab_review:
    st.markdown("### Revise as despesas fixas")
    st.caption(
        "Marque/desmarque a coluna **Fixa?** para incluir ou excluir uma linha do relatório. "
        "Use os filtros acima da tabela para focar em uma empresa ou categoria."
    )

    # ---- Apply pending checkbox edits from previous interaction ----
    # st.session_state.editor holds edits keyed by row index in the PREVIOUS
    # render's displayed DataFrame. _editor_row_ids maps those indices to row_ids.
    _editor_state = st.session_state.get("editor")
    _prev_row_ids = st.session_state.get("_editor_row_ids", [])
    if isinstance(_editor_state, dict) and _prev_row_ids:
        _edited_rows = _editor_state.get("edited_rows", {})
        _changed = False
        for _row_idx_str, _changes in _edited_rows.items():
            if "is_fixed" in _changes:
                _row_idx = int(_row_idx_str)
                if _row_idx < len(_prev_row_ids):
                    _rid = _prev_row_ids[_row_idx]
                    _new_val = bool(_changes["is_fixed"])
                    st.session_state.preset.manual_overrides[_rid] = _new_val
                    _mask = st.session_state.df_review["row_id"] == _rid
                    st.session_state.df_review.loc[_mask, "is_fixed"] = _new_val
                    _changed = True
        if _changed:
            st.session_state.preset.save()

    c1, c2, c3 = st.columns(3)
    with c1:
        f_emp = st.multiselect("Empresa", sorted(st.session_state.df_review["Empresa"].unique()))
    with c2:
        cats_avail = sorted(st.session_state.df_review["Despesas"].dropna().astype(str).unique())
        f_cat = st.multiselect("Categoria", cats_avail)
    with c3:
        f_show = st.radio("Mostrar", ["Todas", "Apenas fixas", "Apenas variáveis"], horizontal=True)

    view = st.session_state.df_review.copy()
    if f_emp:
        view = view[view["Empresa"].isin(f_emp)]
    if f_cat:
        view = view[view["Despesas"].astype(str).isin(f_cat)]
    if f_show == "Apenas fixas":
        view = view[view["is_fixed"]]
    elif f_show == "Apenas variáveis":
        view = view[~view["is_fixed"]]

    display_cols = ["is_fixed", "Pagto", "Empresa", "Favorecido", "Descricao",
                    "CodDespesa", "Despesas", "Valor", "row_id"]
    show_df = view[display_cols].copy()
    show_df["Pagto"] = pd.to_datetime(show_df["Pagto"]).dt.strftime("%d/%m/%Y")
    show_df["Valor"] = show_df["Valor"].abs()

    st.caption(
        "💡 Clique na coluna **Empresa** para mudar a empresa de uma linha. "
        "Você pode digitar uma única empresa (ex: `Alameda 470`) ou várias separadas "
        "por vírgula (ex: `Alameda 470, Mazzini, Artur de Azevedo`) — neste caso o "
        "valor será dividido igualmente entre elas."
    )
    edited = st.data_editor(
        show_df,
        column_config={
            "is_fixed": st.column_config.CheckboxColumn("Fixa?", default=False),
            "Pagto": st.column_config.TextColumn("Data", disabled=True),
            "Empresa": st.column_config.TextColumn("Empresa", help="Digite uma ou várias empresas separadas por vírgula."),
            "Favorecido": st.column_config.TextColumn("Fornecedor", disabled=True),
            "Descricao": st.column_config.TextColumn("Descrição", disabled=True),
            "CodDespesa": st.column_config.NumberColumn("Cód.", format="%d", disabled=True),
            "Despesas": st.column_config.TextColumn("Categoria", disabled=True),
            "Valor": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f", disabled=True),
            "row_id": None,  # hidden
        },
        hide_index=True, use_container_width=True, height=520,
        key="editor",
    )

    # Store row_id mapping for next rerun's checkbox edit detection
    st.session_state["_editor_row_ids"] = show_df["row_id"].tolist()

    # ---- Match free-text empresa input to canonical company names ----
    def _parse_companies(text: str) -> list[str]:
        if not text or pd.isna(text):
            return []
        from expense_engine import _strip
        # Split on comma, semicolon, slash, " e "
        import re as _re
        parts = _re.split(r"[,;/]| e ", str(text))
        out: list[str] = []
        for p in parts:
            ps = _strip(p)
            if not ps:
                continue
            matched = None
            for c in COMPANIES:
                if _strip(c) in ps or ps in _strip(c):
                    matched = c
                    break
            if matched is None:
                # try alias keywords
                from expense_engine import PROJETO_ALIASES
                for k, v in PROJETO_ALIASES.items():
                    if k in ps:
                        matched = v
                        break
            if matched and matched not in out:
                out.append(matched)
        return out

    edits_map_emp = {rid: str(v) for rid, v in zip(edited["row_id"], edited["Empresa"])}

    # Detect company changes that need confirmation
    if "pending_company_change" not in st.session_state:
        st.session_state.pending_company_change = None

    for rid, new_emp_text in edits_map_emp.items():
        orig = st.session_state.df_review[st.session_state.df_review["row_id"] == rid]
        if orig.empty:
            continue
        old_emp = str(orig.iloc[0]["Empresa"])
        if new_emp_text and new_emp_text != old_emp and new_emp_text.strip():
            companies = _parse_companies(new_emp_text)
            if companies and companies != [old_emp]:
                st.session_state.pending_company_change = {
                    "row_id": rid,
                    "favorecido": str(orig.iloc[0]["Favorecido"]),
                    "old": old_emp,
                    "new": companies,
                    "raw": new_emp_text,
                }
                break

    # Show confirmation dialog
    pcc = st.session_state.pending_company_change
    if pcc:
        with st.container(border=True):
            st.markdown(f"**Mudar empresa para esta linha**")
            st.write(f"Fornecedor: **{pcc['favorecido']}**")
            st.write(f"De: `{pcc['old']}`  →  Para: `{', '.join(pcc['new'])}`"
                     + (f"  *(valor dividido por {len(pcc['new'])})*" if len(pcc['new']) > 1 else ""))
            st.caption("Você quer aplicar essa mudança a **todas as linhas deste fornecedor** (regra permanente) ou só a esta linha?")
            cb1, cb2, cb3 = st.columns(3)
            with cb1:
                if st.button("✅ Aplicar a TODAS as linhas deste fornecedor (salvar regra)", use_container_width=True):
                    key = pcc["favorecido"].lower().strip()
                    st.session_state.preset.vendor_company_map[key] = pcc["new"]
                    st.session_state.preset.save()
                    st.session_state.df_review = None
                    st.session_state.pending_company_change = None
                    st.success(f"Regra salva: {pcc['favorecido']} → {', '.join(pcc['new'])}")
                    st.rerun()
            with cb2:
                if st.button("📍 Aplicar SÓ a esta linha", use_container_width=True):
                    # Row-level override: split this single row in place
                    df = st.session_state.df_review
                    row = df[df["row_id"] == pcc["row_id"]].iloc[0].copy()
                    df = df[df["row_id"] != pcc["row_id"]].copy()
                    share = len(pcc["new"])
                    new_rows = []
                    for co in pcc["new"]:
                        nr = row.copy()
                        nr["Empresa"] = co
                        nr["Valor"] = float(row["Valor"]) / share
                        if share > 1:
                            nr["Descricao"] = f"{row['Descricao']} [rateado {share}x]"
                        nr["row_id"] = f"{pcc['row_id']}::{co}"
                        new_rows.append(nr)
                    st.session_state.df_review = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)
                    st.session_state.pending_company_change = None
                    st.success("Linha atualizada.")
                    st.rerun()
            with cb3:
                if st.button("❌ Cancelar", use_container_width=True):
                    st.session_state.pending_company_change = None
                    st.rerun()

    n_fixed = int(st.session_state.df_review["is_fixed"].sum())
    n_total = len(st.session_state.df_review)
    st.info(f"**{n_fixed}** linhas marcadas como fixas (de {n_total} candidatas).")

# ---- Summary tab ----
with tab_summary:
    df_fixed = st.session_state.df_review[st.session_state.df_review["is_fixed"]]
    if df_fixed.empty:
        st.warning("Nenhuma despesa marcada como fixa.")
    else:
        total = float(df_fixed["Valor"].abs().sum())
        n_co = df_fixed["Empresa"].nunique()
        c1, c2, c3 = st.columns(3)
        c1.markdown(f'<div class="kpi"><div class="v">{brl(total)}</div><div class="l">Total despesas fixas</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="kpi"><div class="v">{len(df_fixed)}</div><div class="l">Lançamentos</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="kpi"><div class="v">{n_co}</div><div class="l">Empresas</div></div>', unsafe_allow_html=True)

        st.markdown("#### Por Empresa")
        by_co = summarize_by_company(df_fixed)
        st.bar_chart(by_co.set_index("Empresa")["Total"], color="#34b3d3")
        st.dataframe(
            by_co.assign(Total=by_co["Total"].map(brl)),
            hide_index=True, use_container_width=True,
        )

        st.markdown("#### Por Empresa × Categoria")
        by_cc = summarize_by_company_category(df_fixed)
        st.dataframe(
            by_cc.assign(Total=by_cc["Total"].map(brl)),
            hide_index=True, use_container_width=True,
        )

# ---- Audit tab ----
with tab_audit:
    st.markdown("### Auditoria — Despesas fixas sem empresa definida")
    st.caption(
        "Linhas classificadas como **fixas** cuja empresa ficou como **'Outros'** "
        "(Projeto estava em branco / 'x' e nenhuma palavra-chave de endereço foi encontrada). "
        "Atribua fornecedores a uma ou mais empresas — o valor será dividido igualmente "
        "entre as empresas escolhidas, e a regra é salva no preset para a próxima vez."
    )

    df_audit = st.session_state.df_review[
        (st.session_state.df_review["is_fixed"])
        & (st.session_state.df_review["Empresa"] == "Outros")
    ].copy()

    if df_audit.empty:
        st.success("Nenhuma despesa fixa sem empresa no período. 🎉")
    else:
        # Group by vendor
        grouped = (
            df_audit.assign(V=df_audit["Valor"].abs())
            .groupby("Favorecido", as_index=False)
            .agg(Linhas=("V", "size"), Total=("V", "sum"))
            .sort_values("Total", ascending=False)
        )
        st.info(f"{len(df_audit)} linhas fixas sem empresa, agrupadas em {len(grouped)} fornecedores.")

        st.markdown("#### Atribuir fornecedor → empresa(s)")
        vendor_choice = st.selectbox(
            "Escolha um fornecedor",
            options=grouped["Favorecido"].tolist(),
            format_func=lambda v: f"{v}  •  {int(grouped.loc[grouped['Favorecido']==v, 'Linhas'].iloc[0])} linhas  •  {brl(float(grouped.loc[grouped['Favorecido']==v, 'Total'].iloc[0]))}",
            key="audit_vendor",
        )
        companies_sel = st.multiselect(
            "Empresas que compartilham esta despesa (o valor será dividido igualmente)",
            options=COMPANIES,
            key="audit_companies",
        )

        col_a, col_b = st.columns([1, 1])
        with col_a:
            if st.button("💾 Salvar regra e aplicar", use_container_width=True):
                if not companies_sel:
                    st.error("Escolha pelo menos uma empresa.")
                else:
                    key = str(vendor_choice).lower().strip()
                    st.session_state.preset.vendor_company_map[key] = companies_sel
                    st.session_state.preset.save()
                    st.session_state.df_review = None  # force reclassification
                    st.success(f"Regra salva: {vendor_choice} → {', '.join(companies_sel)}")
                    st.rerun()
        with col_b:
            if st.button("📋 Mostrar linhas deste fornecedor", use_container_width=True):
                rows = df_audit[df_audit["Favorecido"] == vendor_choice][
                    ["Pagto", "Descricao", "Despesas", "Valor"]
                ].copy()
                rows["Pagto"] = pd.to_datetime(rows["Pagto"]).dt.strftime("%d/%m/%Y")
                rows["Valor"] = rows["Valor"].abs()
                st.dataframe(rows, hide_index=True, use_container_width=True)

        st.markdown("#### Fornecedores na fila")
        display = grouped.copy()
        display["Total"] = display["Total"].map(brl)
        st.dataframe(display, hide_index=True, use_container_width=True, height=350)

        with st.expander("Regras já salvas (fornecedor → empresas)", expanded=False):
            if st.session_state.preset.vendor_company_map:
                rules = pd.DataFrame(
                    [(k, ", ".join(v)) for k, v in st.session_state.preset.vendor_company_map.items()],
                    columns=["Fornecedor (normalizado)", "Empresas"],
                )
                st.dataframe(rules, hide_index=True, use_container_width=True)
                rm = st.selectbox("Remover regra", ["—"] + list(st.session_state.preset.vendor_company_map.keys()))
                if st.button("🗑 Remover selecionada") and rm != "—":
                    del st.session_state.preset.vendor_company_map[rm]
                    st.session_state.preset.save()
                    st.session_state.df_review = None
                    st.success(f"Regra removida: {rm}")
                    st.rerun()
            else:
                st.caption("Nenhuma regra salva ainda.")


# ---- Generate tab ----
with tab_generate:
    df_fixed = st.session_state.df_review[st.session_state.df_review["is_fixed"]].copy()
    df_excluded = st.session_state.df_review[~st.session_state.df_review["is_fixed"]].copy()

    st.markdown("### Confirmar e gerar relatórios")
    st.caption("O preset (palavras-chave + ajustes manuais) será salvo automaticamente ao gerar.")

    formato = st.radio(
        "Qual formato você quer gerar?",
        ["PDF + Excel", "Apenas PDF", "Apenas Excel"],
        horizontal=True,
        key="formato_output",
    )

    if st.button("✨ Confirmar e Gerar", type="primary", use_container_width=True):
        if df_fixed.empty:
            st.error("Marque ao menos uma linha como fixa antes de gerar.")
        else:
            try:
                st.session_state.preset.save()
            except Exception as e:
                st.warning(f"Não foi possível salvar o preset: {e}")
            ds = start.strftime("%d%m%Y")
            de = end.strftime("%d%m%Y")
            want_pdf = formato in ("PDF + Excel", "Apenas PDF")
            want_xls = formato in ("PDF + Excel", "Apenas Excel")
            if want_pdf:
                with st.spinner("Gerando PDF…"):
                    st.session_state.pdf_bytes = build_pdf(df_fixed, start, end)
            else:
                st.session_state.pdf_bytes = None
            if want_xls:
                with st.spinner("Gerando Excel…"):
                    st.session_state.xlsx_bytes = build_excel(df_fixed, df_excluded, start, end)
            else:
                st.session_state.xlsx_bytes = None
            st.session_state.output_period = (ds, de)
            st.success("Relatórios gerados com sucesso!")

    # Persistent download buttons (survive reruns caused by other widgets)
    if st.session_state.get("pdf_bytes") or st.session_state.get("xlsx_bytes"):
        ds, de = st.session_state.get("output_period", ("", ""))
        if st.session_state.get("pdf_bytes"):
            st.download_button(
                "⬇ Baixar PDF Executivo",
                data=st.session_state.pdf_bytes,
                file_name=f"Resumo_Executivo_{ds}_{de}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        if st.session_state.get("xlsx_bytes"):
            st.download_button(
                "⬇ Baixar Excel Detalhado",
                data=st.session_state.xlsx_bytes,
                file_name=f"Despesas_Fixas_{ds}_{de}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
