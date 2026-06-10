"""
================================================================================
MSCI Annual Update Factual Process Automation — Streamlit Web App
================================================================================
Web edition of annual_update_processor (v19 logic).

Front end (single page, three numbered sections):
    1 · Input files  — upload the extraction CSV and the DVV merged XLSX
    2 · Run          — button enabled once both files are present
    3 · Results      — summary metrics, per-issuer table, ZIP download

Behind the scenes:
    • Templates are loaded automatically from the repo root (next to this file).
    • Issuer ID is read from the DMX_ISSUER_ID column of the extraction CSV.
    • All three output workbooks are built in memory and bundled into one ZIP.

Core processing logic is unchanged from the desktop v19 script — only file
discovery, I/O, and the entry point were adapted for the web.

Requirements: Python 3.11+ | streamlit | pandas | openpyxl
================================================================================
"""

import io
import re
import zipfile
import logging
import datetime
import traceback
from pathlib import Path

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
# Table / TableStyleInfo intentionally not imported:
# openpyxl Table objects cause Excel corruption (orphaned table XML).
# AutoFilter is used instead — same UX, zero corruption.


# ==============================================================================
# SECTION 1: CONFIGURATION
# ==============================================================================
# Paths are no longer hard-coded — input files come from the uploader and
# templates sit beside this script in the repo root. Everything else (column
# mappings, field names, formatting) is identical to the desktop v19 script.

BASE_DIR = Path(__file__).resolve().parent      # repo root (templates live here)

CONFIG = {
    # ── Template file names (must match the files committed to the repo root) ──
    "TEMPLATES": {
        "Position":       "Position_Tab_Split.xlsx",
        "Scalar_Series1": "Scalar_Series1_DP.xlsx",
        "Series2":        "Series2_DP.xlsx",
    },

    # ── Output file name suffixes — final file is  <IssuerID>_<suffix>.xlsx ────
    "OUTPUT_NAMES": {
        "Position":       "Position_Tab_Split",
        "Scalar_Series1": "Scalar_Series1_DP",
        "Series2":        "Series2_DP",
    },

    # ── Name of the scalar (issuer-level) sheet inside Scalar_Series1 ─────────
    "SCALAR_SHEET_NAME": "Scalar",

    # ── DVV TAB filter (Column B in the DVV Merged file) ─────────────────────
    "DVV_TAB_COL": "B",
    "DVV_EXCLUDE_TABS": {
        "director attributes",
        "director data - board",
        "positions",
        "individual",
    },

    # ── DVV UID column (Column S) — committee name for name-match ────────────
    "DVV_UID_COL": "S",

    # ── DVV name-based matching for committee sheets ──────────────────────────
    "DVV_NAME_MATCH_SHEETS": {
        "Committee": "COMMITTEENAME",
    },

    # ── DVV column letters (Excel column letters in the DVV merged file) ───────
    "DVV_DATALIB_COL":    "G",   # Column G  -> DATALIB_TAG  (code like DIRGENDER)
    "DVV_CORRECTVAL_COL": "AH",  # Column AH -> CORRECT_VALUE
    "DVV_SERIES_COL":     "AM",  # Column AM -> SERIALID

    # ── Extraction Series ID ───────────────────────────────────────────────────
    "EXTRACTION_SERIES_COL_INDEX": 29,

    # ── Field names ────────────────────────────────────────────────────────────
    "ISSUER_ID_FIELD": "DMX_ISSUER_ID",

    # ── Individual name column ────────────────────────────────────────────────
    "INDIVIDUAL_NAME_FIELD":      "REL_INDIVID",
    "INDIVIDUAL_NAME_COL_HEADER": "Individual Name",
    "INDIVIDUAL_NAME_SHEETS": {
        "Committee Membership",
        "Director Attributes",
        "Director Ownership",
        "Compensation",
        "CEO Compensation",
        "CEO Compensation CIC,SEV",
    },

    # ── COMMITTEEFUNCTION columns — Yes -> 1 conversion ──────────────────────
    "COMMITTEEFUNCTION_COLS": {
        "COMMITTEEFUNCTIONA",
        "COMMITTEEFUNCTIONC",
        "COMMITTEEFUNCTIONE",
        "COMMITTEEFUNCTIONG",
        "COMMITTEEFUNCTIONN",
        "COMMITTEEFUNCTIONRISK",
        "COMMITFUNCTIONHEALTHSAFETY",
        "COMMITTEEFUNCTIONO",
    },

    # ── Formatting ────────────────────────────────────────────────────────────
    "HEADER_FONT_NAME":    "Times New Roman",
    "HEADER_FONT_SIZE":    12,
    "HEADER_FONT_COLOR":   "FFFFFF",   # White text
    "HEADER_FILL_COLOR":   "0070C0",   # Blue fill
    "DATA_FONT_NAME":      "Times New Roman",
    "DATA_FONT_SIZE":      12,
    "DVV_HIGHLIGHT_COLOR": "E6FDCF",   # Light green for DVV-updated cells
}


# ==============================================================================
# SECTION 2: LOGGING (in-memory, so it can be shown in the UI / bundled in ZIP)
# ==============================================================================

class ListLogHandler(logging.Handler):
    """Collect formatted log lines into a list for display and export."""

    def __init__(self) -> None:
        super().__init__()
        self.lines: list[str] = []

    def emit(self, record: logging.LogRecord) -> None:
        self.lines.append(self.format(record))


def build_logger() -> tuple[logging.Logger, ListLogHandler]:
    """Fresh logger per run — Streamlit reruns the script on every interaction."""
    logger = logging.getLogger("AnnualUpdateWeb")
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG)

    handler = ListLogHandler()
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%H:%M:%S",
    ))
    logger.addHandler(handler)
    return logger, handler


# ==============================================================================
# SECTION 3: EXTRACTION CSV PROCESSING
# ==============================================================================

def standardize_header(raw_header: str) -> str:
    """
    Convert a CSV header to its Data Lib identifier.
    Extracts text inside the LAST pair of parentheses; otherwise unchanged.
        DMX_ISSUER_ID                       -> DMX_ISSUER_ID
        Audit Board Member (REL_AUDIT_BOARD) -> REL_AUDIT_BOARD
    """
    raw = str(raw_header).strip()
    match = re.search(r"\(([^)]+)\)\s*$", raw)
    if match:
        return match.group(1).strip()
    return raw


def load_extraction(csv_source, issuer_id: str,
                    logger: logging.Logger) -> pd.DataFrame:
    """
    Load extraction CSV (path or file-like) and standardise headers.
    All values read as strings to prevent type inference. Column order preserved.
    """
    try:
        df = pd.read_csv(csv_source, dtype=str, encoding="utf-8-sig")
    except UnicodeDecodeError:
        if hasattr(csv_source, "seek"):
            csv_source.seek(0)
        df = pd.read_csv(csv_source, dtype=str, encoding="latin-1")

    if df.empty:
        raise ValueError("Extraction CSV is empty.")

    df.columns = [standardize_header(c) for c in df.columns]

    if len(df.columns) > CONFIG["EXTRACTION_SERIES_COL_INDEX"]:
        series_col_name = df.columns[CONFIG["EXTRACTION_SERIES_COL_INDEX"]]
        logger.info(
            f"Series ID column (index {CONFIG['EXTRACTION_SERIES_COL_INDEX']}) "
            f"= '{series_col_name}'"
        )
    else:
        logger.warning(
            f"Extraction CSV has fewer than "
            f"{CONFIG['EXTRACTION_SERIES_COL_INDEX'] + 1} columns — "
            "Series ID column cannot be read."
        )

    logger.info(
        f"Extraction loaded: {len(df)} rows, {len(df.columns)} columns "
        "after header standardisation"
    )
    return df


def read_issuer_id_from_df(ext_df: pd.DataFrame) -> str:
    """
    Read the Issuer ID directly from the DMX_ISSUER_ID column of the CSV.
    Returns the first non-blank value found.
    """
    field = CONFIG["ISSUER_ID_FIELD"]
    if field not in ext_df.columns:
        raise ValueError(
            f"Column '{field}' not found in the extraction CSV. "
            "Cannot determine the Issuer ID."
        )
    for val in ext_df[field]:
        if pd.notna(val) and str(val).strip() not in ("", "nan"):
            return str(val).strip()
    raise ValueError(
        f"Column '{field}' is present but contains no Issuer ID values."
    )


def build_datalib_to_series_map(ext_df: pd.DataFrame) -> dict[str, int]:
    """Map each data lib code to the nearest preceding 'Series ID' column index."""
    cols = list(ext_df.columns)
    series_id_positions = [
        i for i, c in enumerate(cols)
        if c == "Series ID" or re.match(r"^Series ID(\.\d+)?$", c)
    ]

    datalib_to_series_col: dict[str, int] = {}
    for col_idx, col_name in enumerate(cols):
        if col_name in ("Series ID",) or re.match(r"^Series ID(\.\d+)?$", col_name):
            continue
        before = [p for p in series_id_positions if p < col_idx]
        if before:
            datalib_to_series_col[col_name] = before[-1]
    return datalib_to_series_col


def build_individual_name_map(ext_df: pd.DataFrame) -> dict[str, str]:
    """
    Map serial_id -> individual name (REL_INDIVID), using the Series ID column
    nearest to REL_INDIVID (before it if possible, otherwise after).
    """
    individ_col = CONFIG["INDIVIDUAL_NAME_FIELD"]
    if individ_col not in ext_df.columns:
        return {}

    cols = list(ext_df.columns)
    individ_idx = cols.index(individ_col)
    series_positions = [
        i for i, c in enumerate(cols)
        if c == "Series ID" or re.match(r"^Series ID(\.\d+)?$", c)
    ]

    before = [p for p in series_positions if p < individ_idx]
    after  = [p for p in series_positions if p > individ_idx]

    if before:
        sid_col = cols[before[-1]]
    elif after:
        sid_col = cols[after[0]]
    else:
        return {}

    name_map: dict[str, str] = {}
    for _, row in ext_df.iterrows():
        sid  = row.get(sid_col)
        name = row.get(individ_col)
        if (pd.notna(sid) and pd.notna(name) and
                str(sid).strip() not in ("", "nan") and
                str(name).strip() not in ("", "nan")):
            name_map[str(sid).strip()] = str(name).strip()
    return name_map


def row_has_data(ext_row: pd.Series, template_headers: list[str]) -> bool:
    """True if the row has any non-blank value among the template headers."""
    skip = {CONFIG["ISSUER_ID_FIELD"], "serial_id"}
    for h in template_headers:
        if h in skip:
            continue
        val = ext_row.get(h)
        if pd.notna(val) and str(val).strip() not in ("", "nan"):
            return True
    return False


# ==============================================================================
# SECTION 4: DVV FILE PROCESSING
# ==============================================================================

def load_dvv(dvv_source, logger: logging.Logger) -> pd.DataFrame:
    """
    Load DVV Merged XLSX (path or file-like). Two sequential filters:
        FILTER 1 — keep rows where Correct Value (Col AH) is non-blank.
        FILTER 2 — drop rows whose TAB (Col B) is in the exclude list.
    """
    def col_letter_to_index(letter: str) -> int:
        letter = letter.upper()
        result = 0
        for ch in letter:
            result = result * 26 + (ord(ch) - ord('A') + 1)
        return result - 1

    try:
        dvv_df = pd.read_excel(
            dvv_source, sheet_name=0, dtype=str, header=0, engine="openpyxl"
        )
    except Exception as e:
        raise RuntimeError(f"Cannot read DVV file: {e}")

    if dvv_df.empty:
        raise ValueError("DVV file is empty.")

    tab_idx        = col_letter_to_index(CONFIG["DVV_TAB_COL"])
    uid_idx        = col_letter_to_index(CONFIG["DVV_UID_COL"])
    datalib_idx    = col_letter_to_index(CONFIG["DVV_DATALIB_COL"])
    correctval_idx = col_letter_to_index(CONFIG["DVV_CORRECTVAL_COL"])
    series_idx     = col_letter_to_index(CONFIG["DVV_SERIES_COL"])

    max_needed = max(tab_idx, uid_idx, datalib_idx, correctval_idx, series_idx)
    if dvv_df.shape[1] <= max_needed:
        raise ValueError(
            f"DVV file has only {dvv_df.shape[1]} columns; "
            f"need at least {max_needed + 1} (up to column {CONFIG['DVV_SERIES_COL']})."
        )

    cols = dvv_df.columns.tolist()
    logger.info(
        f"DVV columns -> TAB='{cols[tab_idx]}' UID='{cols[uid_idx]}' "
        f"DataLib='{cols[datalib_idx]}' CorrectVal='{cols[correctval_idx]}' "
        f"SeriesID='{cols[series_idx]}'"
    )

    dvv_df = dvv_df.rename(columns={
        cols[tab_idx]:        "_DVV_TAB",
        cols[uid_idx]:        "_DVV_UID",
        cols[datalib_idx]:    "_DVV_DATALIB",
        cols[correctval_idx]: "_DVV_CORRECT_VALUE",
        cols[series_idx]:     "_DVV_SERIES_ID",
    })

    total_rows = len(dvv_df)

    # FILTER 1 — Correct Value not blank
    dvv_df = dvv_df[
        dvv_df["_DVV_CORRECT_VALUE"].notna() &
        (dvv_df["_DVV_CORRECT_VALUE"].str.strip() != "")
    ].copy()
    after_filter1 = len(dvv_df)
    logger.info(
        f"DVV Filter 1 (Correct Value not blank): {total_rows} -> {after_filter1} rows"
    )

    # FILTER 2 — TAB exclusion (positional Column B)
    exclude_tabs = CONFIG["DVV_EXCLUDE_TABS"]
    dvv_df["_DVV_TAB_NORM"] = (
        dvv_df["_DVV_TAB"].fillna("").str.strip().str.lower()
    )
    excluded = dvv_df[dvv_df["_DVV_TAB_NORM"].isin(exclude_tabs)]
    if not excluded.empty:
        excluded_vals = sorted(excluded["_DVV_TAB_NORM"].unique().tolist())
        logger.debug(f"DVV Filter 2: excluding TAB values: {excluded_vals}")

    dvv_df = dvv_df[~dvv_df["_DVV_TAB_NORM"].isin(exclude_tabs)].copy()
    dvv_df = dvv_df.drop(columns=["_DVV_TAB_NORM"])

    after_filter2 = len(dvv_df)
    logger.info(
        f"DVV Filter 2 (exclude TABs, Column B): {after_filter1} -> {after_filter2} rows"
    )
    logger.info(
        f"DVV loaded: {total_rows} total -> {after_filter2} actionable rows"
    )
    return dvv_df


# ==============================================================================
# SECTION 5: TEMPLATE POPULATION
# ==============================================================================

def populate_template(
    template_path: Path,
    template_key: str,
    ext_df: pd.DataFrame,
    issuer_id: str,
    logger: logging.Logger,
) -> tuple[openpyxl.Workbook, dict]:
    """Populate all sheets in a template workbook from the extraction DataFrame."""
    logger.info(f"Populating '{template_key}': {template_path.name}")
    wb = load_workbook(template_path)
    sheet_meta = {}

    datalib_to_series_map = build_datalib_to_series_map(ext_df)
    individual_name_map = build_individual_name_map(ext_df)
    logger.info(
        f"Individual name map: {len(individual_name_map)} directors loaded from REL_INDIVID"
    )

    comm_fn_cols = CONFIG.get("COMMITTEEFUNCTION_COLS", set())

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        headers = [cell.value for cell in ws[1] if cell.value is not None]
        if not headers:
            logger.warning(f"  Sheet '{sheet_name}' has no headers — skipping.")
            continue

        header_col_map = {
            cell.value: col_idx
            for col_idx, cell in enumerate(ws[1], start=1)
            if cell.value is not None
        }

        is_scalar = (sheet_name == CONFIG["SCALAR_SHEET_NAME"])
        indiv_col_idx = None

        if (not is_scalar
                and "serial_id" in header_col_map
                and sheet_name in CONFIG["INDIVIDUAL_NAME_SHEETS"]):
            indiv_header = CONFIG["INDIVIDUAL_NAME_COL_HEADER"]
            if indiv_header not in header_col_map:
                sid_col_pos = header_col_map["serial_id"]
                insert_at   = sid_col_pos + 1
                ws.insert_cols(insert_at)
                ws.cell(row=1, column=insert_at).value = indiv_header
                header_col_map = {
                    h: (ci + 1 if ci >= insert_at else ci)
                    for h, ci in header_col_map.items()
                }
                header_col_map[indiv_header] = insert_at
                indiv_col_idx = insert_at
                logger.debug(
                    f"  Sheet '{sheet_name}': inserted '{indiv_header}' at column {insert_at}"
                )
            else:
                indiv_col_idx = header_col_map[indiv_header]

        headers = [cell.value for cell in ws[1] if cell.value is not None]

        data_start_row = 2
        populated_rows = 0

        # ── SCALAR: one issuer row ─────────────────────────────────────────
        if is_scalar:
            if CONFIG["ISSUER_ID_FIELD"] in ext_df.columns:
                issuer_rows = ext_df[ext_df[CONFIG["ISSUER_ID_FIELD"]] == issuer_id]
                if issuer_rows.empty:
                    issuer_rows = ext_df
            else:
                issuer_rows = ext_df

            if issuer_rows.empty:
                logger.warning(
                    f"  No rows for issuer {issuer_id}. "
                    f"Scalar sheet '{sheet_name}' left empty."
                )
                continue

            ext_row   = issuer_rows.iloc[0]
            write_row = data_start_row

            for tmpl_header, col_idx in header_col_map.items():
                if tmpl_header == CONFIG["ISSUER_ID_FIELD"]:
                    ws.cell(row=write_row, column=col_idx).value = issuer_id
                elif tmpl_header in ext_df.columns:
                    val = ext_row.get(tmpl_header)
                    ws.cell(row=write_row, column=col_idx).value = (
                        None if (pd.isna(val) or str(val).strip() in ("", "nan"))
                        else str(val).strip()
                    )
            populated_rows = 1

        # ── POSITION / SERIES: one row per extraction row ──────────────────
        else:
            sheet_series_col_idx  = None
            sheet_section_cols    = []

            all_std = list(ext_df.columns)
            series_positions = [
                i for i, c in enumerate(all_std)
                if c == "Series ID" or re.match(r"^Series ID(\.\d+)?$", c)
            ]
            best_match_cnt = 0
            for sp_idx in series_positions:
                nxt = next((p for p in series_positions if p > sp_idx), len(all_std))
                candidate = [
                    all_std[i] for i in range(sp_idx + 1, nxt) if all_std[i] in headers
                ]
                if len(candidate) > best_match_cnt:
                    best_match_cnt       = len(candidate)
                    sheet_series_col_idx = sp_idx
                    sheet_section_cols   = candidate

            for _, ext_row in ext_df.iterrows():

                if sheet_series_col_idx is not None:
                    sid_col_name = ext_df.columns[sheet_series_col_idx]
                    sid_val = ext_row.get(sid_col_name)
                    if pd.isna(sid_val) or str(sid_val).strip() in ("", "nan"):
                        continue
                    section_series_id = str(sid_val).strip()
                else:
                    section_series_id = None

                if sheet_section_cols:
                    section_has_data = any(
                        pd.notna(ext_row.get(c)) and
                        str(ext_row.get(c, "")).strip() not in ("", "nan")
                        for c in sheet_section_cols
                    )
                    if not section_has_data:
                        continue
                elif not row_has_data(ext_row, headers):
                    continue

                write_row = data_start_row + populated_rows

                for tmpl_header, col_idx in header_col_map.items():

                    if tmpl_header == CONFIG["ISSUER_ID_FIELD"]:
                        ws.cell(row=write_row, column=col_idx).value = issuer_id

                    elif tmpl_header == "serial_id":
                        ws.cell(row=write_row, column=col_idx).value = section_series_id
                        if indiv_col_idx and section_series_id:
                            ind_name = individual_name_map.get(section_series_id, "")
                            ws.cell(row=write_row, column=indiv_col_idx).value = (
                                ind_name if ind_name else None
                            )

                    elif tmpl_header == CONFIG["INDIVIDUAL_NAME_COL_HEADER"]:
                        pass  # written above alongside serial_id

                    elif tmpl_header in ext_df.columns:
                        val = ext_row.get(tmpl_header)
                        if pd.isna(val) or str(val).strip() in ("", "nan"):
                            ws.cell(row=write_row, column=col_idx).value = None
                        else:
                            clean_val = str(val).strip()
                            if (tmpl_header in comm_fn_cols and
                                    clean_val.lower() == "yes"):
                                clean_val = "1"
                            ws.cell(row=write_row, column=col_idx).value = clean_val

                populated_rows += 1

        sheet_meta[sheet_name] = {
            "headers":         headers,
            "header_col_map":  header_col_map,
            "data_start_row":  data_start_row,
            "populated_rows":  populated_rows,
        }
        logger.debug(f"  Sheet '{sheet_name}': {populated_rows} row(s) written.")

    return wb, sheet_meta


# ==============================================================================
# SECTION 6: DVV OVERRIDE PROCESSING
# ==============================================================================

def apply_dvv_overrides(
    wb: openpyxl.Workbook,
    sheet_meta: dict,
    dvv_df: pd.DataFrame,
    issuer_id: str,
    template_key: str,
    logger: logging.Logger,
) -> int:
    """Apply DVV Correct Value overrides to the populated workbook (green fill)."""
    dvv_fill = PatternFill(
        "solid",
        start_color=CONFIG["DVV_HIGHLIGHT_COLOR"],
        end_color=CONFIG["DVV_HIGHLIGHT_COLOR"],
    )
    total_overrides = 0
    name_match_sheets = CONFIG.get("DVV_NAME_MATCH_SHEETS", {})

    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_meta:
            continue

        ws             = wb[sheet_name]
        meta           = sheet_meta[sheet_name]
        header_col_map = meta["header_col_map"]
        data_start_row = meta["data_start_row"]

        is_scalar = (sheet_name == CONFIG["SCALAR_SHEET_NAME"])

        serial_id_col = header_col_map.get("serial_id")
        issuer_id_col = header_col_map.get(CONFIG["ISSUER_ID_FIELD"])

        row_map: dict[str, int] = {}
        if not is_scalar and serial_id_col:
            for r in range(data_start_row, data_start_row + meta["populated_rows"]):
                sid = ws.cell(row=r, column=serial_id_col).value
                if sid is not None:
                    row_map[str(sid).strip()] = r

        name_row_map: dict[str, int] = {}
        anchor_col = name_match_sheets.get(sheet_name)
        if anchor_col and anchor_col in header_col_map:
            anchor_col_idx = header_col_map[anchor_col]
            for r in range(data_start_row, data_start_row + meta["populated_rows"]):
                name_val = ws.cell(row=r, column=anchor_col_idx).value
                if name_val is not None:
                    name_row_map[str(name_val).strip().lower()] = r

        sheet_overrides = 0
        comm_fn_cols = CONFIG.get("COMMITTEEFUNCTION_COLS", set())

        for _, dvv_row in dvv_df.iterrows():
            datalib     = str(dvv_row.get("_DVV_DATALIB", "")).strip()
            correct_val = dvv_row.get("_DVV_CORRECT_VALUE")
            dvv_series  = str(dvv_row.get("_DVV_SERIES_ID", "")).strip()
            dvv_tab     = str(dvv_row.get("_DVV_TAB", "")).strip()
            dvv_uid     = str(dvv_row.get("_DVV_UID", "")).strip()

            if dvv_series.lower() in ("nan", "none", ""):
                dvv_series = ""
            if dvv_uid.lower() in ("nan", "none", ""):
                dvv_uid = ""

            if (datalib in comm_fn_cols and
                    correct_val is not None and
                    str(correct_val).strip().lower() == "yes"):
                correct_val = "1"

            if not datalib or datalib not in header_col_map:
                continue

            target_col = header_col_map[datalib]

            # ── SCALAR ──────────────────────────────────────────────────────
            if is_scalar:
                ws.cell(row=data_start_row, column=target_col).value = correct_val
                ws.cell(row=data_start_row, column=target_col).fill  = dvv_fill

            # ── SERIES ID present ────────────────────────────────────────────
            elif dvv_series and serial_id_col:
                if dvv_series in row_map:
                    row_num = row_map[dvv_series]
                    ws.cell(row=row_num, column=target_col).value = correct_val
                    ws.cell(row=row_num, column=target_col).fill  = dvv_fill
                else:
                    matched_by_name = False
                    if sheet_name in name_match_sheets and dvv_uid and name_row_map:
                        lookup_name = dvv_uid.strip().lower()
                        if lookup_name in name_row_map:
                            row_num = name_row_map[lookup_name]
                            ws.cell(row=row_num, column=target_col).value = correct_val
                            ws.cell(row=row_num, column=target_col).fill  = dvv_fill
                            matched_by_name = True
                            logger.info(
                                f"  DVV name-fallback (series miss): sheet='{sheet_name}' "
                                f"UID='{dvv_uid}' DataLib={datalib} -> row {row_num}"
                            )
                    if not matched_by_name:
                        new_row = data_start_row + meta["populated_rows"]
                        if issuer_id_col:
                            ws.cell(row=new_row, column=issuer_id_col).value = issuer_id
                        ws.cell(row=new_row, column=serial_id_col).value = dvv_series
                        ws.cell(row=new_row, column=target_col).value = correct_val
                        ws.cell(row=new_row, column=target_col).fill  = dvv_fill
                        row_map[dvv_series] = new_row
                        meta["populated_rows"] += 1
                        logger.info(
                            f"  DVV new row: sheet='{sheet_name}' "
                            f"SeriesID={dvv_series} DataLib={datalib}"
                        )

            # ── NO SERIES ID: name-match using UID (Column S) ─────────────
            elif not dvv_series and sheet_name in name_match_sheets:
                lookup_name = dvv_uid.strip().lower()
                if lookup_name and lookup_name in name_row_map:
                    row_num = name_row_map[lookup_name]
                    ws.cell(row=row_num, column=target_col).value = correct_val
                    ws.cell(row=row_num, column=target_col).fill  = dvv_fill
                    logger.info(
                        f"  DVV name-match: sheet='{sheet_name}' "
                        f"UID='{dvv_uid}' DataLib={datalib} -> row {row_num}"
                    )
                else:
                    logger.warning(
                        f"  DVV name-match MISS: sheet='{sheet_name}' "
                        f"UID='{dvv_uid}' DataLib={datalib}"
                    )
                    continue

            else:
                continue

            sheet_overrides += 1

        if sheet_overrides:
            logger.info(
                f"  DVV '{template_key}' / '{sheet_name}': {sheet_overrides} override(s)"
            )
        total_overrides += sheet_overrides

    logger.info(f"DVV overrides '{template_key}': {total_overrides} total override(s)")
    return total_overrides


# ==============================================================================
# SECTION 7: OUTPUT FORMATTING
# ==============================================================================

def _thin_border() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def _hdr_font() -> Font:
    return Font(
        name=CONFIG["HEADER_FONT_NAME"], size=CONFIG["HEADER_FONT_SIZE"],
        bold=True, color=CONFIG["HEADER_FONT_COLOR"]
    )

def _hdr_fill() -> PatternFill:
    return PatternFill(
        "solid",
        start_color=CONFIG["HEADER_FILL_COLOR"],
        end_color=CONFIG["HEADER_FILL_COLOR"]
    )

def _dat_font() -> Font:
    return Font(name=CONFIG["DATA_FONT_NAME"], size=CONFIG["DATA_FONT_SIZE"], bold=False)

def _center() -> Alignment:
    return Alignment(horizontal="center", vertical="center")


def format_worksheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    data_start_row: int,
    populated_rows: int,
    logger: logging.Logger,
) -> None:
    """Header styling, top/left data alignment, thin borders, AutoFilter, freeze."""
    max_col = ws.max_column
    if not max_col:
        return

    last_row = data_start_row + populated_rows - 1
    if populated_rows == 0:
        last_row = data_start_row - 1

    border = _thin_border()

    for c in range(1, max_col + 1):
        cell = ws.cell(row=1, column=c)
        if cell.value is not None:
            cell.font      = _hdr_font()
            cell.fill      = _hdr_fill()
            cell.alignment = _center()
            cell.border    = border

    data_align = Alignment(horizontal="left", vertical="top", wrap_text=False)
    for r in range(data_start_row, last_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font      = _dat_font()
            cell.border    = border
            cell.alignment = data_align
        ws.row_dimensions[r].height = 15

    ws.freeze_panes = "A2"

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value), default=8
        )
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

    # AutoFilter over full data range (replaces Excel Table objects → no corruption)
    filter_ref = f"A1:{get_column_letter(max_col)}1"
    if populated_rows > 0:
        filter_ref = f"A1:{get_column_letter(max_col)}{last_row}"
    ws.auto_filter.ref = filter_ref


# ==============================================================================
# SECTION 8: VALIDATION
# ==============================================================================

def build_validation_log(
    all_template_headers: dict,
    ext_df: pd.DataFrame,
    issuer_id: str,
    logger: logging.Logger,
) -> list[dict]:
    """Template Data Libs missing in CSV + CSV Data Libs unused in templates."""
    ext_cols      = set(ext_df.columns.tolist())
    all_tmpl_cols: set[str] = set()
    records = []

    for tmpl_key, sheets in all_template_headers.items():
        for sheet_name, headers in sheets.items():
            for h in headers:
                if h is None:
                    continue
                all_tmpl_cols.add(h)
                if (h not in ext_cols
                        and h != CONFIG["ISSUER_ID_FIELD"]
                        and h != "serial_id"):
                    records.append({
                        "Check Type": "Template DataLib missing in Extraction",
                        "Template":   tmpl_key,
                        "Sheet":      sheet_name,
                        "Data Lib":   h,
                        "Detail":     "Header in template; no matching column in CSV",
                    })

    for col in ext_cols:
        if col not in all_tmpl_cols and col not in {CONFIG["ISSUER_ID_FIELD"], "serial_id"}:
            records.append({
                "Check Type": "Extraction DataLib unused in Templates",
                "Template":   "ALL",
                "Sheet":      "ALL",
                "Data Lib":   col,
                "Detail":     "Column in CSV; not used in any template sheet",
            })

    logger.info(f"Validation: {len(records)} issues found")
    return records


# ==============================================================================
# SECTION 9: ORCHESTRATION (web edition)
# ==============================================================================

def workbook_to_bytes(wb: openpyxl.Workbook) -> bytes:
    """Save an openpyxl workbook to an in-memory bytes object."""
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def run_pipeline(
    csv_bytes: bytes,
    csv_filename: str,
    dvv_bytes: bytes,
    template_dir: Path,
    logger: logging.Logger,
) -> dict:
    """
    Full end-to-end pipeline for one uploaded extraction + DVV pair.

    Returns a dict:
        summary  : metrics dict
        outputs  : { output_filename : xlsx_bytes }
        missing_templates : list of template filenames not found on disk
    """
    start_time = datetime.datetime.now()
    summary = {
        "Issuer ID":         "",
        "Issuer Name":       "",
        "Extraction File":   csv_filename,
        "Records Processed": 0,
        "Fields Populated":  0,
        "DVV Overrides":     0,
        "Validation Issues": 0,
        "Errors":            [],
        "Status":            "PENDING",
    }
    outputs: dict[str, bytes] = {}
    missing_templates: list[str] = []

    # ── STEP 1: Load extraction CSV ────────────────────────────────────────
    logger.info("=" * 70)
    logger.info("STEP 1: Loading extraction CSV")
    ext_df = load_extraction(io.BytesIO(csv_bytes), "", logger)
    summary["Records Processed"] = len(ext_df)

    # ── STEP 2: Issuer ID from CSV; Issuer Name from filename ──────────────
    issuer_id = read_issuer_id_from_df(ext_df)
    summary["Issuer ID"] = issuer_id
    stem = Path(csv_filename).stem
    parts = stem.split("_", 1)
    issuer_name = parts[1].strip() if len(parts) == 2 else issuer_id
    summary["Issuer Name"] = issuer_name
    logger.info(f"Issuer ID (from {CONFIG['ISSUER_ID_FIELD']}): {issuer_id}")
    logger.info(f"Issuer Name (from filename): {issuer_name}")

    # ── STEP 3: Load DVV file ──────────────────────────────────────────────
    logger.info("=" * 70)
    logger.info("STEP 3: Loading DVV file")
    dvv_df = load_dvv(io.BytesIO(dvv_bytes), logger)

    # ── STEP 4: Populate templates + DVV overrides + format + save ─────────
    logger.info("=" * 70)
    logger.info("STEP 4: Populating templates and applying DVV overrides")

    all_template_headers: dict[str, dict[str, list]] = {}
    total_dvv_overrides = 0
    total_fields        = 0

    for tmpl_key, tmpl_filename in CONFIG["TEMPLATES"].items():
        tmpl_path = template_dir / tmpl_filename
        out_name  = f"{issuer_id}_{CONFIG['OUTPUT_NAMES'][tmpl_key]}.xlsx"

        logger.info(f"  ── {tmpl_key}: {tmpl_filename} ──")

        if not tmpl_path.exists():
            msg = f"Template not found in repo: {tmpl_filename}"
            logger.error(msg)
            summary["Errors"].append(msg)
            missing_templates.append(tmpl_filename)
            continue

        try:
            wb, sheet_meta = populate_template(
                tmpl_path, tmpl_key, ext_df, issuer_id, logger
            )
        except Exception as e:
            msg = f"Population failed for '{tmpl_key}': {e}"
            logger.error(msg)
            logger.debug(traceback.format_exc())
            summary["Errors"].append(msg)
            continue

        all_template_headers[tmpl_key] = {
            sn: m["headers"] for sn, m in sheet_meta.items()
        }
        for m in sheet_meta.values():
            total_fields += len(m["headers"]) * m["populated_rows"]

        try:
            overrides = apply_dvv_overrides(
                wb, sheet_meta, dvv_df, issuer_id, tmpl_key, logger
            )
            total_dvv_overrides += overrides
        except Exception as e:
            msg = f"DVV override failed for '{tmpl_key}': {e}"
            logger.error(msg)
            logger.debug(traceback.format_exc())
            summary["Errors"].append(msg)

        try:
            for sn in wb.sheetnames:
                if sn in sheet_meta:
                    m = sheet_meta[sn]
                    format_worksheet(
                        wb[sn], m["data_start_row"], m["populated_rows"], logger
                    )
        except Exception as e:
            logger.error(f"Formatting failed for '{tmpl_key}': {e}")
            logger.debug(traceback.format_exc())

        try:
            outputs[out_name] = workbook_to_bytes(wb)
            logger.info(f"  Built: {out_name}")
        except Exception as e:
            msg = f"Save failed for '{tmpl_key}': {e}"
            logger.error(msg)
            summary["Errors"].append(msg)

    # ── STEP 5: Validation (logged only) ───────────────────────────────────
    logger.info("=" * 70)
    logger.info("STEP 5: Validation check")
    val_records = build_validation_log(all_template_headers, ext_df, issuer_id, logger)

    summary["Validation Issues"] = len(val_records)
    summary["DVV Overrides"]     = total_dvv_overrides
    summary["Fields Populated"]  = total_fields

    elapsed = (datetime.datetime.now() - start_time).total_seconds()
    summary["Status"] = "FAILED" if summary["Errors"] else "SUCCESS"
    logger.info("=" * 70)
    logger.info(f"DONE in {elapsed:.1f}s — status {summary['Status']}")

    return {
        "summary": summary,
        "outputs": outputs,
        "missing_templates": missing_templates,
    }


def build_zip(outputs: dict[str, bytes], log_text: str, issuer_id: str) -> bytes:
    """Bundle all output workbooks (+ execution log) into a single ZIP."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in outputs.items():
            zf.writestr(name, data)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        zf.writestr(f"{issuer_id}_execution_{ts}.log", log_text)
    buf.seek(0)
    return buf.getvalue()


# ==============================================================================
# SECTION 10: STREAMLIT FRONT END
# ==============================================================================

st.set_page_config(
    page_title="Annual Update Factual Process",
    page_icon="🧩",
    layout="centered",
)

st.title("Annual Update — Factual Process Automation")
st.caption(
    "Upload the extraction CSV and the DVV merged file. Templates load "
    "automatically and the Issuer ID is read from the CSV."
)

# Confirm templates are present in the repo root (informational).
present = [
    f for f in CONFIG["TEMPLATES"].values() if (BASE_DIR / f).exists()
]
missing = [
    f for f in CONFIG["TEMPLATES"].values() if not (BASE_DIR / f).exists()
]
if missing:
    st.warning(
        "These template files are expected in the repo root but were not found: "
        + ", ".join(missing)
    )

# ── 1 · Input files ────────────────────────────────────────────────────────
st.subheader("1 · Input files")

col_csv, col_dvv = st.columns(2)
with col_csv:
    csv_upload = st.file_uploader(
        "Extraction CSV",
        type=["csv"],
        help="The raw extraction file, e.g. IID000000002177512_Aurubis AG.csv",
    )
with col_dvv:
    dvv_upload = st.file_uploader(
        "DVV merged file",
        type=["xlsx"],
        help="The DVV merged workbook for the same issuer.",
    )

both_ready = csv_upload is not None and dvv_upload is not None

# ── 2 · Run ──────────────────────────────────────────────────────────────
st.subheader("2 · Run")

if not both_ready:
    st.info("Upload both files above to enable the run button.")

run_clicked = st.button(
    "Run processing",
    type="primary",
    disabled=not both_ready,
    use_container_width=True,
)

if run_clicked and both_ready:
    logger, handler = build_logger()
    try:
        with st.spinner("Processing…"):
            result = run_pipeline(
                csv_bytes=csv_upload.getvalue(),
                csv_filename=csv_upload.name,
                dvv_bytes=dvv_upload.getvalue(),
                template_dir=BASE_DIR,
                logger=logger,
            )
        result["log_text"] = "\n".join(handler.lines)
        st.session_state["result"] = result
    except Exception as e:
        logger.error(f"FATAL: {e}")
        logger.debug(traceback.format_exc())
        st.session_state["result"] = {
            "summary": {"Status": "FAILED", "Errors": [str(e)]},
            "outputs": {},
            "missing_templates": [],
            "log_text": "\n".join(handler.lines),
            "fatal": str(e),
        }

# ── 3 · Results ──────────────────────────────────────────────────────────
st.subheader("3 · Results")

result = st.session_state.get("result")

if not result:
    st.write("Results will appear here after a run.")
else:
    summary = result["summary"]
    outputs = result.get("outputs", {})

    if summary.get("Status") == "FAILED":
        st.error("Run failed. See the errors and execution log below.")
        for err in summary.get("Errors", []):
            st.write(f"• {err}")
    else:
        st.success(f"Completed — {len(outputs)} output file(s) generated.")

    # Summary metrics
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Records processed", summary.get("Records Processed", 0))
    m2.metric("Fields populated",  summary.get("Fields Populated", 0))
    m3.metric("DVV overrides",     summary.get("DVV Overrides", 0))
    m4.metric("Validation issues", summary.get("Validation Issues", 0))

    # Per-issuer table
    issuer_table = pd.DataFrame([{
        "Issuer ID":         summary.get("Issuer ID", ""),
        "Issuer Name":       summary.get("Issuer Name", ""),
        "Extraction File":   summary.get("Extraction File", ""),
        "Records":           summary.get("Records Processed", 0),
        "Fields Populated":  summary.get("Fields Populated", 0),
        "DVV Overrides":     summary.get("DVV Overrides", 0),
        "Validation Issues": summary.get("Validation Issues", 0),
        "Status":            summary.get("Status", ""),
    }])
    st.dataframe(issuer_table, hide_index=True, use_container_width=True)

    if outputs:
        with st.expander("Output files in the ZIP", expanded=False):
            for name in outputs:
                st.write(f"• {name}")

        zip_bytes = build_zip(
            outputs, result.get("log_text", ""), summary.get("Issuer ID", "issuer")
        )
        st.download_button(
            "Download all outputs (ZIP)",
            data=zip_bytes,
            file_name=f"{summary.get('Issuer ID', 'issuer')}_Output.zip",
            mime="application/zip",
            type="primary",
            use_container_width=True,
        )

    with st.expander("Execution log", expanded=False):
        st.code(result.get("log_text", ""), language="text")
