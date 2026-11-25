"""Load and normalize contact data from Excel (.xlsx) or CSV files."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pandas as pd


# Canonical column names we expect (lowercase, stripped)
CANONICAL_COLS = {"name", "email", "designation", "company"}


def _normalize_header(header: str) -> str:
    """Lowercase, strip, collapse whitespace and underscores for matching."""
    h = str(header).strip().lower()
    h = re.sub(r"[_\s]+", " ", h)
    return h


def _build_column_map(
    df_columns: list[str], custom_map: dict[str, str] | None = None
) -> dict[str, str]:
    """Return {canonical_name: actual_column_name} for the required columns.

    Uses *custom_map* from config.yaml first, then falls back to
    case-insensitive / whitespace-insensitive matching.
    """
    mapping: dict[str, str] = {}
    remaining = set(CANONICAL_COLS)

    # 1. Apply explicit custom mapping from config
    if custom_map:
        for canon, actual in custom_map.items():
            canon_lower = canon.strip().lower()
            if canon_lower in remaining:
                # Find the actual column in df (case-insensitive)
                for col in df_columns:
                    if _normalize_header(col) == _normalize_header(actual):
                        mapping[canon_lower] = col
                        remaining.discard(canon_lower)
                        break

    # 2. Auto-match remaining by normalized header
    for canon in list(remaining):
        for col in df_columns:
            if _normalize_header(col) == canon:
                mapping[canon] = col
                remaining.discard(canon)
                break

    if remaining:
        raise ValueError(
            f"Could not find required columns: {remaining}. "
            f"Available columns: {list(df_columns)}. "
            f"Add a column_mapping section to config.yaml to map them."
        )
    return mapping


def load_contacts(
    file_path: str | Path,
    sheet: str | None = None,
    column_mapping: dict[str, str] | None = None,
    min_score: int | None = None,
    limit: int | None = None,
    start_row: int = 1,
) -> tuple[pd.DataFrame, dict[str, str]]:
    """Load contacts from an Excel or CSV file.

    Returns:
        (DataFrame with all original columns + normalized canonical columns,
         column_map {canonical -> original column name})
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    ext = file_path.suffix.lower()
    if ext == ".csv":
        df = pd.read_csv(file_path)
    elif ext in (".xlsx", ".xls"):
        sheet_name = sheet if sheet else 0
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    else:
        raise ValueError(f"Unsupported file format: {ext}. Use .xlsx or .csv")

    if df.empty:
        raise ValueError("The input file contains no data rows.")

    col_map = _build_column_map(list(df.columns), column_mapping)

    # Rename canonical columns to lowercase for internal use, keeping originals
    rename_dict = {}
    for canon, actual in col_map.items():
        if actual != canon:
            rename_dict[actual] = canon
    df = df.rename(columns=rename_dict)

    # --- Apply filters ---

    # start_row (1-based, referring to data rows)
    if start_row > 1:
        df = df.iloc[start_row - 1 :].reset_index(drop=True)

    # min_score filter
    score_col = _find_score_column(df.columns)
    if min_score is not None:
        if score_col is None:
            raise ValueError(
                "--min-score requires an 'Intern Interest Score (1-5)' column "
                "in the input file."
            )
        df[score_col] = pd.to_numeric(df[score_col], errors="coerce")
        df = df[df[score_col] >= min_score].reset_index(drop=True)

    # limit
    if limit is not None and limit > 0:
        df = df.head(limit).reset_index(drop=True)

    return df, col_map


def _find_score_column(columns: pd.Index) -> str | None:
    """Find the intern interest score column by normalized name."""
    for col in columns:
        norm = _normalize_header(col)
        if "intern" in norm and "score" in norm:
            return col
    return None


def get_score_for_row(row: pd.Series, columns: pd.Index) -> float | None:
    """Return the intern interest score for a row, or None if not present."""
    score_col = _find_score_column(columns)
    if score_col is None:
        return None
    val = row.get(score_col)
    if pd.isna(val):
        return None
    try:
        return float(val)
    except (ValueError, TypeError):
        return None


def normalize_row(row: pd.Series) -> dict[str, Any]:
    """Clean and normalize the canonical fields of a single row.

    Returns a dict with keys: name, email, company, designation,
    plus first_name and cleaned variants.
    """
    import pandas as _pd

    raw_name_val = row.get("name", "")
    raw_email_val = row.get("email", "")
    raw_company_val = row.get("company", "")
    raw_designation_val = row.get("designation", "")

    # Handle NaN/None from pandas
    raw_name = "" if _pd.isna(raw_name_val) else str(raw_name_val).strip()
    raw_email = "" if _pd.isna(raw_email_val) else str(raw_email_val).strip()
    raw_company = "" if _pd.isna(raw_company_val) else str(raw_company_val).strip()
    raw_designation = "" if _pd.isna(raw_designation_val) else str(raw_designation_val).strip()

    # Collapse whitespace
    name = re.sub(r"\s+", " ", raw_name)
    email = re.sub(r"\s+", "", raw_email).lower()
    company = _clean_company(raw_company)
    designation = re.sub(r"\s+", " ", raw_designation)

    return {
        "name": name,
        "email": email,
        "company": company,
        "designation": designation,
        "raw_name": raw_name,
        "raw_email": raw_email,
        "raw_company": raw_company,
        "raw_designation": raw_designation,
    }


def _clean_company(name: str) -> str:
    """Strip trailing punctuation, collapse whitespace."""
    name = re.sub(r"\s+", " ", name).strip()
    # Remove trailing commas, periods, semicolons, and other stray punctuation
    name = re.sub(r"[,.\;\:!]+$", "", name).strip()
    return name
