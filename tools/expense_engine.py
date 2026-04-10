"""Core expense loading, cleaning, and fixed-expense classification.

Pure functions; no Streamlit, no file output. UI imports from here.
"""
from __future__ import annotations

import json
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
PRESETS_PATH = DATA_DIR / "presets.json"

COMPANIES = ["Rio de Janeiro", "Alameda 470", "Artur de Azevedo", "Mazzini", "Alameda 334"]

DEFAULT_FIXED_CODES = [301, 302, 303, 304, 305, 306, 309, 310]

DEFAULT_FIXED_KEYWORDS = [
    "aluguel", "iptu", "enel", "sabesp", "claro", "net ", "vivo", "telefone",
    "hagana", "limpa vidros", "seguranca", "grupo gabriel", "controle de pragas",
    "supricorp", "gimba", "pao de queijo", "agua personalizada",
    "impressora", "cartucho", "plotter", "alcatoner", "seguro incendio", "auto de licenca",
    "extintor", "laudo bombeiro", "galao", "garagem", "box", "motoboy",
    "faxineira", "lalamove", "regus", "helena", "uniart", "ralph", "rmf",
    "internet", "energia", "condominio",
]

EXCLUDE_KEYWORDS = ["cesar valverde"]

# Keys are normalized (lowercased, accent-stripped). Order matters: more specific first.
PROJETO_ALIASES = {
    # Rio de Janeiro
    "rio de janeiro": "Rio de Janeiro",
    "regus": "Rio de Janeiro",
    # Alameda 470
    "alameda gabriel 470": "Alameda 470",
    "alameda gabriel, 470": "Alameda 470",
    "alameda 470": "Alameda 470",
    "gabriel 470": "Alameda 470",
    "helena cabral magano": "Alameda 470",
    # Artur de Azevedo
    "artur de azevedo": "Artur de Azevedo",
    "rmf": "Artur de Azevedo",
    # Mazzini
    "mazzini": "Mazzini",
    "uniart": "Mazzini",
    # Alameda 334
    "alameda 334": "Alameda 334",
    "alameda gabriel 334": "Alameda 334",
    "ralph": "Alameda 334",
    "focal": "Alameda 334",
}


# ---------- helpers ----------

def _strip(s) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return s.lower().strip()


def _match_company_in_text(text: str) -> str | None:
    """Scan a normalized text blob for any company-identifying keyword."""
    if not text:
        return None
    for key, canon in PROJETO_ALIASES.items():
        if key in text:
            return canon
    return None


def _match_all_companies_in_text(text: str) -> list[str]:
    """Return all distinct companies mentioned in a text blob, preserving order."""
    if not text:
        return []
    found: list[str] = []
    for key, canon in PROJETO_ALIASES.items():
        if key in text and canon not in found:
            found.append(canon)
    return found


def normalize_projeto(value) -> str:
    """Map a Projeto cell to a canonical company. 'x'/empty/unknown → 'Outros'."""
    s = _strip(value)
    if not s or s == "x":
        return "Outros"
    hit = _match_company_in_text(s)
    if hit:
        return hit
    for c in COMPANIES:
        if _strip(c) in s:
            return c
    return "Outros"


def assign_company_row(row) -> str:
    """Resolve company for a row.

    1. If Projeto is a real value, use normalize_projeto.
    2. Otherwise (Projeto = 'x' or empty), scan Descricao + Favorecido + ContaSintetica
       for a company keyword.
    3. Fallback: 'Outros'.
    """
    primary = normalize_projeto(row.get("Projeto"))
    if primary != "Outros":
        return primary
    blob = " ".join([
        _strip(row.get("Descricao")),
        _strip(row.get("Favorecido")),
        _strip(row.get("ContaSintetica")),
        _strip(row.get("CentroDeCusto")),
    ])
    hit = _match_company_in_text(blob)
    return hit or "Outros"


# ---------- loading ----------

def load_workbook_dataframe(file) -> pd.DataFrame:
    """Load the messy TAAG xlsx into a clean DataFrame.

    Header is on row 10 (1-indexed). Reads from openpyxl-compatible file path or buffer.
    """
    df = pd.read_excel(file, sheet_name=0, header=9, engine="openpyxl")
    # Drop fully empty rows
    df = df.dropna(how="all")
    # Standardize column names
    df.columns = [str(c).strip() for c in df.columns]
    # Required columns
    rename_map = {}
    for col in df.columns:
        cl = col.lower()
        if cl.startswith("pagto"): rename_map[col] = "Pagto"
        elif cl == "r$" or cl.startswith("r$"): rename_map[col] = "Valor"
        elif cl.startswith("descri"): rename_map[col] = "Descricao"
        elif cl.startswith("cliente"): rename_map[col] = "Favorecido"
        elif cl.startswith("cod") and "despesa" in cl: rename_map[col] = "CodDespesa"
        elif cl == "despesas": rename_map[col] = "Despesas"
        elif cl.startswith("conta"): rename_map[col] = "ContaSintetica"
        elif cl == "projeto": rename_map[col] = "Projeto"
        elif cl == "banco": rename_map[col] = "Banco"
        elif cl.startswith("centro"): rename_map[col] = "CentroDeCusto"
    df = df.rename(columns=rename_map)

    needed = ["Pagto", "Valor", "Descricao", "Favorecido", "CodDespesa", "Despesas", "Projeto"]
    for c in needed:
        if c not in df.columns:
            df[c] = None

    df["Pagto"] = pd.to_datetime(df["Pagto"], errors="coerce")
    df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce")
    df["CodDespesa"] = pd.to_numeric(df["CodDespesa"], errors="coerce").astype("Int64")
    df = df.dropna(subset=["Pagto", "Valor"])
    df["Empresa"] = df.apply(assign_company_row, axis=1)
    df = _expand_shared_rows(df)
    return df.reset_index(drop=True)


def _expand_shared_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Split rows whose Projeto/Descrição reference multiple companies.

    Example: Projeto = 'Alameda gabriel 470 e 334' with R$ -500 becomes two rows
    of R$ -250 each (one per company). The original row is replaced.
    """
    rows_out = []
    for _, row in df.iterrows():
        proj_text = _strip(row.get("Projeto"))
        desc_text = _strip(row.get("Descricao"))
        # Only look for multi-company in Projeto first, then description
        hits = _match_all_companies_in_text(proj_text)
        if len(hits) < 2:
            hits_desc = _match_all_companies_in_text(desc_text)
            if len(hits_desc) >= 2:
                hits = hits_desc
        if len(hits) < 2:
            rows_out.append(row)
            continue
        share = len(hits)
        val = row.get("Valor")
        for co in hits:
            new_row = row.copy()
            new_row["Empresa"] = co
            if pd.notna(val):
                new_row["Valor"] = float(val) / share
            new_row["Descricao"] = f"{row.get('Descricao')} [rateado {share}x]"
            rows_out.append(new_row)
    return pd.DataFrame(rows_out)


# ---------- filtering ----------

def filter_by_date(df: pd.DataFrame, start: date, end: date) -> pd.DataFrame:
    mask = (df["Pagto"].dt.date >= start) & (df["Pagto"].dt.date <= end)
    return df.loc[mask].copy()


def exclude_personal(df: pd.DataFrame) -> pd.DataFrame:
    def is_personal(row) -> bool:
        blob = " ".join([_strip(row.get(c)) for c in ("Favorecido", "Descricao", "ContaSintetica")])
        return any(kw in blob for kw in EXCLUDE_KEYWORDS)
    return df.loc[~df.apply(is_personal, axis=1)].copy()


def only_debits(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[df["Valor"] < 0].copy()


# ---------- classification ----------

@dataclass
class Preset:
    fixed_codes: list[int] = field(default_factory=lambda: list(DEFAULT_FIXED_CODES))
    fixed_keywords: list[str] = field(default_factory=lambda: list(DEFAULT_FIXED_KEYWORDS))
    manual_overrides: dict[str, bool] = field(default_factory=dict)  # row hash -> is_fixed
    # Vendor name (normalized) -> list of company names to split the value across.
    # Example: {"hagana": ["Alameda 470","Artur de Azevedo"]} splits each Hagana row 50/50.
    vendor_company_map: dict[str, list[str]] = field(default_factory=dict)

    def to_json(self) -> str:
        return json.dumps({
            "fixed_codes": self.fixed_codes,
            "fixed_keywords": self.fixed_keywords,
            "manual_overrides": self.manual_overrides,
            "vendor_company_map": self.vendor_company_map,
        }, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls) -> "Preset":
        if PRESETS_PATH.exists():
            data = json.loads(PRESETS_PATH.read_text())
            return cls(
                fixed_codes=data.get("fixed_codes", list(DEFAULT_FIXED_CODES)),
                fixed_keywords=data.get("fixed_keywords", list(DEFAULT_FIXED_KEYWORDS)),
                manual_overrides=data.get("manual_overrides", {}),
                vendor_company_map=data.get("vendor_company_map", {}),
            )
        return cls()

    def save(self) -> None:
        PRESETS_PATH.parent.mkdir(parents=True, exist_ok=True)
        PRESETS_PATH.write_text(self.to_json())


def row_hash(row) -> str:
    """Stable id for a row so manual overrides survive between runs."""
    parts = [
        str(row.get("Pagto")),
        f"{row.get('Valor'):.2f}" if pd.notna(row.get("Valor")) else "",
        _strip(row.get("Favorecido")),
        _strip(row.get("Descricao")),
        str(row.get("CodDespesa")),
    ]
    return "|".join(parts)


def apply_vendor_map(df: pd.DataFrame, preset: Preset) -> pd.DataFrame:
    """Apply user-defined vendor→companies map.

    Rows whose (normalized) Favorecido is in the map and whose current Empresa is
    "Outros" are replaced by one row per target company, with Valor divided evenly.
    """
    if not preset.vendor_company_map:
        return df
    vmap = {k.lower(): v for k, v in preset.vendor_company_map.items()}
    out = []
    for _, row in df.iterrows():
        fav = _strip(row.get("Favorecido"))
        if row.get("Empresa") != "Outros" or not fav:
            out.append(row); continue
        target = None
        for vendor_key, companies in vmap.items():
            if vendor_key and vendor_key in fav:
                target = companies
                break
        if not target:
            out.append(row); continue
        share = len(target)
        val = row.get("Valor")
        for co in target:
            new_row = row.copy()
            new_row["Empresa"] = co
            if pd.notna(val) and share > 0:
                new_row["Valor"] = float(val) / share
            if share > 1:
                new_row["Descricao"] = f"{row.get('Descricao')} [rateado {share}x]"
            out.append(new_row)
    return pd.DataFrame(out).reset_index(drop=True)


def auto_classify_fixed(df: pd.DataFrame, preset: Preset) -> pd.DataFrame:
    df = df.copy()
    kws = [_strip(k) for k in preset.fixed_keywords]
    codes = set(preset.fixed_codes)

    def classify(row) -> bool:
        blob = " ".join([
            _strip(row.get("Descricao")),
            _strip(row.get("ContaSintetica")),
            _strip(row.get("Despesas")),
            _strip(row.get("Favorecido")),
        ])
        code = row.get("CodDespesa")
        if pd.notna(code) and int(code) in codes:
            return True
        return any(kw and kw in blob for kw in kws)

    df["is_fixed_auto"] = df.apply(classify, axis=1)
    df["row_id"] = df.apply(row_hash, axis=1)
    df["is_fixed"] = df.apply(
        lambda r: preset.manual_overrides.get(r["row_id"], r["is_fixed_auto"]),
        axis=1,
    )
    return df


# ---------- aggregation ----------

def summarize_by_company(df_fixed: pd.DataFrame) -> pd.DataFrame:
    out = (
        df_fixed.assign(Valor_abs=df_fixed["Valor"].abs())
        .groupby("Empresa", as_index=False)["Valor_abs"]
        .sum()
        .rename(columns={"Valor_abs": "Total"})
        .sort_values("Total", ascending=False)
    )
    return out


def summarize_by_company_category(df_fixed: pd.DataFrame) -> pd.DataFrame:
    out = (
        df_fixed.assign(Valor_abs=df_fixed["Valor"].abs())
        .groupby(["Empresa", "Despesas"], as_index=False)["Valor_abs"]
        .sum()
        .rename(columns={"Valor_abs": "Total"})
        .sort_values(["Empresa", "Total"], ascending=[True, False])
    )
    return out


MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def monthly_total(df_fixed: pd.DataFrame, empresa: str | None = None) -> pd.DataFrame:
    """Return one row per (year, month) with the total fixed expense."""
    df = df_fixed if empresa is None else df_fixed[df_fixed["Empresa"] == empresa]
    if df.empty:
        return pd.DataFrame(columns=["Ano", "MesNum", "Mes", "Total"])
    df = df.assign(
        Ano=df["Pagto"].dt.year,
        MesNum=df["Pagto"].dt.month,
        Valor_abs=df["Valor"].abs(),
    )
    out = (
        df.groupby(["Ano", "MesNum"], as_index=False)["Valor_abs"]
        .sum()
        .rename(columns={"Valor_abs": "Total"})
        .sort_values(["Ano", "MesNum"])
    )
    out["Mes"] = out.apply(lambda r: f"{MESES_PT[int(r['MesNum'])-1]}/{int(r['Ano'])}", axis=1)
    return out.reset_index(drop=True)


def monthly_by_company(df_fixed: pd.DataFrame) -> pd.DataFrame:
    """Wide table: index = Mes label, columns = Empresa, values = Total."""
    if df_fixed.empty:
        return pd.DataFrame()
    df = df_fixed.assign(
        Ano=df_fixed["Pagto"].dt.year,
        MesNum=df_fixed["Pagto"].dt.month,
        Valor_abs=df_fixed["Valor"].abs(),
    )
    pivot = df.pivot_table(
        index=["Ano", "MesNum"], columns="Empresa", values="Valor_abs",
        aggfunc="sum", fill_value=0,
    )
    pivot = pivot.sort_index()
    pivot.index = [f"{MESES_PT[m-1]}/{y}" for y, m in pivot.index]
    return pivot


def top_vendors(df_fixed: pd.DataFrame, empresa: str, n: int = 10) -> pd.DataFrame:
    sub = df_fixed[df_fixed["Empresa"] == empresa]
    return (
        sub.assign(Valor_abs=sub["Valor"].abs())
        .groupby("Favorecido", as_index=False)["Valor_abs"]
        .sum()
        .rename(columns={"Valor_abs": "Total"})
        .sort_values("Total", ascending=False)
        .head(n)
    )
