"""Upload pipeline and CEO category sync for the Taag Despesas app.

Responsibilities:
- sync_ceo_categories(): UPSERT CEO_CATEGORY_MAP to Supabase on app start
- build_and_push_upload(): Storage upload + transaction insert + rollup write
  after operator confirms a ledger

All functions fail silently when offline. The Supabase presets table is the
unconditional source of truth for classification rules.
"""
from __future__ import annotations

import uuid
from datetime import date

import pandas as pd


# ---------------------------------------------------------------------------
# CEO category sync
# ---------------------------------------------------------------------------

def sync_ceo_categories() -> None:
    """UPSERT CEO_CATEGORY_MAP → ceo_categoria_lookup on app start.

    Python is canonical. This is idempotent and runs in ~100ms.
    Fails silently if offline.
    """
    from expense_engine import CEO_CATEGORY_MAP
    from supabase_client import upsert_ceo_categories
    upsert_ceo_categories(CEO_CATEGORY_MAP)


# ---------------------------------------------------------------------------
# Upload pipeline: Storage → ledger_uploads → transactions → monthly_rollups
# ---------------------------------------------------------------------------

def build_and_push_upload(
    df_fixed: pd.DataFrame,
    period_start: date,
    period_end: date,
    filename: str,
    file_bytes: bytes,
) -> dict:
    """Push a confirmed ledger to Supabase. Returns a status dict.

    Step order (strict):
      1. Upload .xlsx to Storage  — if this fails, abort (no DB writes)
      2. Insert ledger_uploads record
      3. Insert transactions rows (batched)
      4. Upsert monthly_rollups (aggregated from df_fixed)

    Returns:
        {"ok": True, "upload_id": "...", "storage_path": "..."}
        {"ok": False, "error": "..."}
    """
    from supabase_client import (
        get_client, insert_ledger_upload, insert_transactions,
        upload_xlsx_to_storage, upsert_monthly_rollups, get_company_id,
    )

    if get_client() is None:
        return {"ok": False, "error": "offline"}

    upload_id = str(uuid.uuid4())

    # --- 1. Upload xlsx to Storage (canonical backup) ---
    storage_path = upload_xlsx_to_storage(upload_id, filename, file_bytes)
    if storage_path is None:
        return {"ok": False, "error": "storage_upload_failed"}

    # --- 2. Insert ledger_uploads record ---
    db_upload_id = insert_ledger_upload(
        period_start=str(period_start),
        period_end=str(period_end),
        filename=filename,
        storage_path=storage_path,
        row_count=len(df_fixed),
    )
    if db_upload_id is None:
        return {"ok": False, "error": "ledger_upload_insert_failed", "storage_path": storage_path}

    # --- 3. Insert transactions rows ---
    txn_rows = _df_to_transaction_rows(df_fixed, db_upload_id)
    ok_txn = insert_transactions(db_upload_id, txn_rows)
    if not ok_txn:
        return {"ok": False, "error": "transactions_insert_failed",
                "upload_id": db_upload_id, "storage_path": storage_path}

    # --- 4. Upsert monthly_rollups ---
    rollup_rows = _build_rollup_rows(df_fixed)
    enriched = []
    company_id_cache: dict[str, str | None] = {}
    for row in rollup_rows:
        name = row["empresa"]
        if name not in company_id_cache:
            company_id_cache[name] = get_company_id(name)
        cid = company_id_cache[name]
        if cid:
            enriched.append({
                "company_id": cid,
                "year": row["year"],
                "month": row["month"],
                "ceo_categoria": row["ceo_categoria"],
                "total": row["total"],
            })

    if enriched:
        upsert_monthly_rollups(enriched)

    return {"ok": True, "upload_id": db_upload_id, "storage_path": storage_path}


def _df_to_transaction_rows(df: pd.DataFrame, upload_id: str) -> list[dict]:
    rows = []
    for _, r in df.iterrows():
        rows.append({
            "upload_id": upload_id,
            "pagto": str(r["Pagto"].date()) if pd.notna(r.get("Pagto")) else None,
            "empresa": str(r.get("Empresa", "")),
            "favorecido": str(r.get("Favorecido", "")),
            "descricao": str(r.get("Descricao", "")),
            "cod_despesa": int(r["CodDespesa"]) if pd.notna(r.get("CodDespesa")) else None,
            "despesas": str(r.get("Despesas", "")),
            "valor": float(abs(r["Valor"])),
            "row_hash": str(r.get("row_id", r.get("row_hash", ""))),
            "is_fixed_auto": bool(r.get("is_fixed_auto", False)),
            "is_fixed": bool(r.get("is_fixed", False)),
            "ceo_categoria": str(r.get("CeoCategoria", "Outros Fixos")),
        })
    return rows


def _build_rollup_rows(df: pd.DataFrame) -> list[dict]:
    if df.empty:
        return []
    agg = (
        df.assign(
            year=df["Pagto"].dt.year,
            month=df["Pagto"].dt.month,
            valor_abs=df["Valor"].abs(),
        )
        .groupby(["Empresa", "year", "month", "CeoCategoria"], as_index=False)["valor_abs"]
        .sum()
        .rename(columns={"valor_abs": "total", "Empresa": "empresa", "CeoCategoria": "ceo_categoria"})
    )
    return agg.to_dict("records")
