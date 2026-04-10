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
