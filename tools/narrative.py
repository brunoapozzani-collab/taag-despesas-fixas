"""Generate executive narrative paragraphs via Claude API.

Falls back to a deterministic template if the API key is missing or the call fails.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

MODEL = "claude-sonnet-4-6"


def _brl(x: float) -> str:
    s = f"R$ {x:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _client():
    try:
        import anthropic
        key = os.getenv("ANTHROPIC_API_KEY")
        if not key:
            try:
                import streamlit as st
                key = st.secrets.get("ANTHROPIC_API_KEY")
            except Exception:
                pass
        if not key:
            return None
        return anthropic.Anthropic(api_key=key)
    except Exception:
        return None


def _call(client, system: str, user: str) -> str | None:
    try:
        resp = client.messages.create(
            model=MODEL,
            max_tokens=600,
            system=system,
            messages=[{"role": "user", "content": user}],
        )
        return resp.content[0].text.strip()
    except Exception as e:
        print(f"[narrative] Claude call failed: {e}")
        return None


SYSTEM_PROMPT = (
    "Você é um analista financeiro sênior escrevendo um resumo executivo "
    "para o CEO da TAAG Brasil. Use português brasileiro, tom profissional "
    "e direto. Cite valores em R$ no formato brasileiro (R$ 1.234,56). "
    "Nunca invente números — use apenas os fornecidos. Seja conciso: "
    "máximo 4 frases curtas. Destaque tendências (alta/baixa), maior "
    "categoria/fornecedor e qualquer alerta relevante."
)


def _fact_pack(monthly: pd.DataFrame, by_category: pd.DataFrame, by_vendor: pd.DataFrame, scope: str) -> dict:
    pack = {"escopo": scope}
    if not monthly.empty:
        first = float(monthly["Total"].iloc[0])
        last = float(monthly["Total"].iloc[-1])
        delta = last - first
        pct = (delta / first * 100) if first else 0
        pack["meses"] = [
            {"mes": str(m), "total": float(t)}
            for m, t in zip(monthly["Mes"], monthly["Total"])
        ]
        pack["primeiro_mes"] = _brl(first)
        pack["ultimo_mes"] = _brl(last)
        pack["variacao_pct"] = round(pct, 1)
        pack["variacao_abs"] = _brl(abs(delta))
        pack["tendencia"] = "alta" if delta > 0 else ("baixa" if delta < 0 else "estável")
    if not by_category.empty:
        top = by_category.head(3)
        pack["top_categorias"] = [
            {"categoria": c, "total": _brl(float(t))}
            for c, t in zip(top["Despesas"].astype(str), top["Total"])
        ]
    if not by_vendor.empty:
        top = by_vendor.head(5)
        pack["top_fornecedores"] = [
            {"fornecedor": str(v)[:50], "total": _brl(float(t))}
            for v, t in zip(top["Favorecido"], top["Total"])
        ]
    return pack


def _fallback(facts: dict) -> str:
    parts = []
    if "tendencia" in facts:
        parts.append(
            f"No período analisado, as despesas fixas apresentaram {facts['tendencia']} "
            f"({facts['variacao_pct']:+.1f}%), passando de {facts['primeiro_mes']} "
            f"para {facts['ultimo_mes']}."
        )
    if facts.get("top_categorias"):
        c = facts["top_categorias"][0]
        parts.append(f"A principal categoria foi {c['categoria']} ({c['total']}).")
    if facts.get("top_fornecedores"):
        v = facts["top_fornecedores"][0]
        parts.append(f"O maior fornecedor foi {v['fornecedor']} ({v['total']}).")
    return " ".join(parts) or "Sem dados suficientes para análise."


def write_strategic_insights(
    grand_total: float,
    by_co: pd.DataFrame,
    by_ceo: pd.DataFrame,
    monthly: pd.DataFrame,
    n_outros: int,
) -> str:
    """CEO strategic recommendations based on the full expense dataset."""
    facts: dict = {"total_geral": _brl(grand_total)}
    if not by_co.empty:
        facts["localidades"] = [
            {"empresa": str(r["Empresa"]), "total": _brl(float(r["Total"])),
             "pct": f"{float(r['Total']) / grand_total * 100:.1f}%" if grand_total else "0%"}
            for _, r in by_co.iterrows()
        ]
    if not by_ceo.empty:
        facts["top_categorias"] = [
            {"categoria": str(r["CeoCategoria"]), "total": _brl(float(r["Total"])),
             "pct": f"{float(r['Total']) / grand_total * 100:.1f}%" if grand_total else "0%"}
            for _, r in by_ceo.head(8).iterrows()
        ]
    if len(monthly) >= 2:
        last = float(monthly["Total"].iloc[-1])
        first = float(monthly["Total"].iloc[0])
        delta_pct = (last - first) / first * 100 if first else 0
        facts["tendencia"] = "alta" if delta_pct > 3 else ("baixa" if delta_pct < -3 else "estável")
        facts["variacao_periodo"] = f"{delta_pct:+.1f}%"
    facts["lancamentos_sem_localidade"] = n_outros

    client = _client()
    if client:
        system = (
            "Você é um consultor estratégico sênior preparando um briefing confidencial "
            "para o CEO da TAAG Brasil. Use português brasileiro, tom executivo e direto. "
            "Com base nos dados, produza 4 a 6 recomendações estratégicas concretas e acionáveis "
            "sobre como reduzir, renegociar ou otimizar as despesas fixas operacionais. "
            "Mencione localidades e categorias específicas. Máximo 200 palavras."
        )
        user = (
            f"Dados de despesas fixas operacionais da TAAG Brasil:\n\n"
            f"{json.dumps(facts, ensure_ascii=False, indent=2)}\n\n"
            "Escreva as recomendações estratégicas para o CEO."
        )
        text = _call(client, system, user)
        if text:
            return text

    # Programmatic fallback
    parts = []
    if facts.get("localidades"):
        top = facts["localidades"][0]
        parts.append(f"A localidade {top['empresa']} concentra {top['pct']} das despesas fixas — avaliar se o investimento é proporcional à receita da unidade.")
    if facts.get("top_categorias"):
        top_cat = facts["top_categorias"][0]
        if "Aluguel" in str(top_cat.get("categoria", "")):
            parts.append(f"Aluguel é o maior custo fixo ({top_cat['pct']}). Recomendar revisão dos contratos de locação e análise de consolidação de espaços.")
        else:
            parts.append(f"A categoria {top_cat['categoria']} lidera as despesas ({top_cat['pct']}). Avaliar oportunidades de otimização.")
    if n_outros > 0:
        parts.append(f"Existem {n_outros} lançamentos sem localidade definida. Atribuir corretamente para melhorar a gestão por unidade.")
    parts.append("Recomendar análise semestral dos fornecedores de serviços recorrentes (telecom, segurança, limpeza) para renegociação de contratos.")
    return " ".join(parts)


def write_narrative(monthly: pd.DataFrame, by_category: pd.DataFrame,
                    by_vendor: pd.DataFrame, scope: str) -> str:
    """Return a Portuguese narrative paragraph for the given data slice."""
    facts = _fact_pack(monthly, by_category, by_vendor, scope)
    client = _client()
    if client:
        user = (
            f"Escreva uma análise executiva (máximo 4 frases) sobre {scope}, "
            f"baseada nestes dados:\n\n{json.dumps(facts, ensure_ascii=False, indent=2)}\n\n"
            "Foque em: tendência mês-a-mês, maior categoria, maior fornecedor, "
            "e qualquer variação relevante. Não cite os números brutos do JSON, "
            "use prosa fluida."
        )
        text = _call(client, SYSTEM_PROMPT, user)
        if text:
            return text
    return _fallback(facts)
