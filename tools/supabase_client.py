"""Singleton Supabase client for the Taag Despesas app.

Usage:
    from supabase_client import get_client  # returns None when offline/unconfigured

The service-role key is used for all server-side ops (bypasses RLS).
Never expose this key to a browser — it lives only in .env / Vercel env vars.
"""
from __future__ import annotations

import os
from functools import lru_cache
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_SERVICE_KEY = os.getenv("SUPABASE_SERVICE_KEY", "")


@lru_cache(maxsize=1)
def get_client():
    """Return a Supabase client, or None if not configured."""
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        return None
    try:
        from supabase import create_client
        return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Preset — singleton table (one shared rule set, no per-user separation)
# ---------------------------------------------------------------------------

def load_preset() -> dict:
    """Load the singleton preset from Supabase.

    Raises RuntimeError if Supabase is unreachable. There is no local fallback
    by design: a fallback would create two divergent reads during an outage.
    Loud failure is correct; the dated backup file is a manual restore artifact.
    """
    client = get_client()
    if client is None:
        raise RuntimeError(
            "Cannot reach Supabase to load classification rules — check network connection"
        )
    try:
        resp = client.table("presets").select("*").single().execute()
        return resp.data or {}
    except Exception as exc:
        raise RuntimeError(
            "Cannot reach Supabase to load classification rules — check network connection"
        ) from exc


def upsert_preset(preset_data: dict) -> None:
    """UPSERT the singleton preset. Raises RuntimeError on any failure."""
    client = get_client()
    if client is None:
        raise RuntimeError(
            "Cannot reach Supabase to save classification rules — check network connection"
        )
    try:
        client.table("presets").upsert(
            {
                "id": True,
                "fixed_keywords": preset_data.get("fixed_keywords", []),
                "fixed_codes": preset_data.get("fixed_codes", []),
                "manual_overrides": preset_data.get("manual_overrides", {}),
                "vendor_company_map": preset_data.get("vendor_company_map", {}),
                "updated_at": "now()",
            },
            on_conflict="id",
        ).execute()
    except Exception as exc:
        raise RuntimeError(
            "Cannot reach Supabase to save classification rules — check network connection"
        ) from exc


# ---------------------------------------------------------------------------
# CEO category lookup
# ---------------------------------------------------------------------------

def upsert_ceo_categories(category_map: list[tuple[str, list[str]]]) -> bool:
    """UPSERT CEO_CATEGORY_MAP into ceo_categoria_lookup. Idempotent.

    Called once on app start. Python list is canonical; Supabase table is a
    read-cache for the Next.js CEO view. Fails silently if offline.
    """
    client = get_client()
    if client is None:
        return False
    try:
        rows = [
            {"name": name, "keywords": keywords, "sort_order": i}
            for i, (name, keywords) in enumerate(category_map)
        ]
        client.table("ceo_categoria_lookup").upsert(rows, on_conflict="name").execute()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Ledger upload pipeline
# ---------------------------------------------------------------------------

def insert_ledger_upload(
    period_start: str,
    period_end: str,
    filename: str,
    storage_path: str,
    row_count: int,
) -> Optional[str]:
    """Insert a ledger_uploads record. Returns the new upload UUID, or None on failure."""
    client = get_client()
    if client is None:
        return None
    try:
        resp = (
            client.table("ledger_uploads")
            .insert({
                "period_start": period_start,
                "period_end": period_end,
                "filename": filename,
                "storage_path": storage_path,
                "row_count": row_count,
            })
            .execute()
        )
        return resp.data[0]["id"] if resp.data else None
    except Exception:
        return None


def upload_xlsx_to_storage(upload_id: str, filename: str, file_bytes: bytes) -> Optional[str]:
    """Upload the raw Excel file to Storage. Returns the storage path, or None on failure.

    Must succeed BEFORE any DB writes (Storage is canonical backup).
    """
    client = get_client()
    if client is None:
        return None
    try:
        path = f"{upload_id}/{filename}"
        client.storage.from_("uploads").upload(
            path=path,
            file=file_bytes,
            file_options={"content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
        return f"uploads/{path}"
    except Exception:
        return None


def insert_transactions(upload_id: str, rows: list[dict]) -> bool:
    """Bulk-insert classified transaction rows. Batched in chunks of 500."""
    client = get_client()
    if client is None:
        return False
    try:
        BATCH = 500
        for i in range(0, len(rows), BATCH):
            client.table("transactions").insert(rows[i: i + BATCH]).execute()
        return True
    except Exception:
        return False


def upsert_monthly_rollups(rollup_rows: list[dict]) -> bool:
    """UPSERT pre-aggregated monthly rollups. Conflict target: (company_id, year, month, ceo_categoria)."""
    client = get_client()
    if client is None:
        return False
    try:
        client.table("monthly_rollups").upsert(
            rollup_rows,
            on_conflict="company_id,year,month,ceo_categoria",
        ).execute()
        return True
    except Exception:
        return False


def get_company_id(company_name: str) -> Optional[str]:
    """Look up company UUID by name. Returns None if not found or offline."""
    client = get_client()
    if client is None:
        return None
    try:
        resp = (
            client.table("companies")
            .select("id")
            .eq("name", company_name)
            .single()
            .execute()
        )
        return resp.data["id"] if resp.data else None
    except Exception:
        return None
