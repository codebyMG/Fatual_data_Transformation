"""
================================================================================
MSCI Annual Update Factual Process Automation  —  Streamlit Web App
================================================================================
Version     : 3.0.0 (Streamlit)
Based on    : annual_update_processor v2.0.0 (command-line)
Changes vs the CLI version:
              - All file system discovery removed. CSV, DVV and template files
                are now UPLOADED through the browser.
              - Issuer ID is no longer typed in. It is read automatically from
                the DMX_ISSUER_ID column of the uploaded extraction CSV.
                * If the CSV contains one issuer  -> that issuer is processed.
                * If it contains several issuers   -> each is processed and all
                  outputs are bundled together.
              - All inputs/outputs are handled in memory (BytesIO). Nothing is
                written to disk. Results are returned as a downloadable ZIP.
              - Logging is captured to an in-memory buffer and shown in the UI.
Core data logic (header standardisation, template population, DVV overrides,
validation, formatting) is preserved unchanged from v2.0.0.
Requirements: streamlit | pandas | openpyxl
================================================================================
"""

import io
import re
import zipfile
import logging
import datetime
import traceback

import pandas as pd
import streamlit as st
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


# ==============================================================================
# SECTION 1: CONFIGURATION
# Only the data-shaping constants remain. All folder paths are gone because
# files now arrive as uploads.
# ==============================================================================

CONFIG = {
    # UI slot label -> internal template key. Each slot is optional.
    "TEMPLATE_SLOTS": {
        "Position":  "Position",
        "Scaler":    "Scaler",
        "Series 1":  "Series1",
        "Series 2":  "Series2",
    },

    # Output file-name suffixes per template key.
    "OUTPUT_NAMES": {
        "Position": "Position_Tab_Split",
        "Scaler":   "Scaler_DP",
        "Series1":  "Series1_DP",
        "Series2":  "Series2_DP",
    },

    # DVV column letters (Excel column letters in the DVV merged file).
    "DVV_DATALIB_COL":    "F",   # Column F  -> Data Lib
    "DVV_CORRECTVAL_COL": "AG",  # Column AG -> Correct Value
    "DVV_SERIES_COL":     "AL",  # Column AL -> Series ID

    # Column E = index 4 (0-based) in the RAW CSV holds the Series ID.
    "EXTRACTION_SERIES_COL_INDEX": 4,

    # Field names.
    "ISSUER_ID_FIELD": "DMX_ISSUER_ID",

    # Formatting.
    "HEADER_FONT_NAME":    "Times New Roman",
    "HEADER_FONT_SIZE":    12,
    "HEADER_FONT_COLOR":   "FFFFFF",   # White text
    "HEADER_FILL_COLOR":   "0070C0",   # Blue fill
    "DATA_FONT_NAME":      "Times New Roman",
    "DATA_FONT_SIZE":      12,
    "DVV_HIGHLIGHT_COLOR": "E6FDCF",   # Light green for DVV-updated cells

    # Excel table style.
    "TABLE_STYLE": "TableStyleMedium2",
}


# ==============================================================================
# SECTION 2: LOGGING (in-memory)
# ==============================================================================

def setup_logging() -> tuple[logging.Logger, io.StringIO]:
    """Log to an in-memory buffer so the messages can be shown in the app."""
    log_buffer = io.StringIO()

    logger = logging.getLogger("AnnualUpdate")
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG)

    handler = logging.StreamHandler(log_buffer)
    handler.setLevel(logging.INFO)
    handler.setFormatter(logging.Formatter("%(levelname)-8s | %(message)s"))
    logger.addHandler(handler)
    logger.propagate = False
    return logger, log_buffer


# ==============================================================================
# SECTION 3: EXTRACTION CSV PROCESSING
# ==============================================================================

def standardize_header(raw_header: str) -> str:
    """
    Convert a CSV header to its Data Lib identifier.

    Format 1 - already a Data Lib (no brackets):
        DMX_ISSUER_ID  ->  DMX_ISSUER_ID
    Format 2 - human label with Data Lib in final parentheses:
        Audit Board Member (REL_AUDIT_BOARD)  ->  REL_AUDIT_BOARD

    Rule: extract text inside the LAST pair of parentheses; otherwise unchanged.
    """
    raw = str(raw_header).strip()
    match = re.search(r"\(([^)]+)\)\s*$", raw)
    if match:
        return match.group(1).strip()
    return raw


def load_extraction(file_bytes: bytes, display_name: str,
                    logger: logging.Logger) -> pd.DataFrame:
    """
    Load extraction CSV (from uploaded bytes) and standardise headers.
    All values read as strings to prevent type inference. No column reordering.
    """
    logger.info(f"Loading extraction CSV: {display_name}")
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, encoding="latin-1")

    if df.empty:
        raise ValueError(f"Extraction CSV is empty: {display_name}")

    df.columns = [standardize_header(c) for c in df.columns]

    idx = CONFIG["EXTRACTION_SERIES_COL_INDEX"]
    if len(df.columns) > idx:
        logger.info(f"Column E (Series ID source) = '{df.columns[idx]}'")
    else:
        logger.warning(
            "Extraction CSV has fewer than 5 columns - "
            "Series ID (Column E) cannot be read."
        )

    logger.info(
        f"Extraction loaded: {len(df)} rows, {len(df.columns)} columns "
        "after header standardisation"
    )
    return df


def get_series_id_from_row(ext_row: pd.Series,
                           ext_df: pd.DataFrame) -> str | None:
    """Read Series ID from Column E (index 4) of the extraction row."""
    idx = CONFIG["EXTRACTION_SERIES_COL_INDEX"]
    if len(ext_df.columns) > idx:
        col_name = ext_df.columns[idx]
        val = ext_row.get(col_name)
        if pd.notna(val) and str(val).strip() not in ("", "nan"):
            return str(val).strip()
    return None


def row_has_data(ext_row: pd.Series, template_headers: list[str]) -> bool:
    """True if the row has at least one non-blank value for a template header
    (excluding DMX_ISSUER_ID and serial_id). Prevents writing empty rows."""
    skip = {CONFIG["ISSUER_ID_FIELD"], "serial_id"}
    for h in template_headers:
        if h in skip:
            continue
        val = ext_row.get(h)
        if pd.notna(val) and str(val).strip() not in ("", "nan"):
            return True
    return False


def detect_issuer_ids(ext_df: pd.DataFrame,
                      logger: logging.Logger) -> list[str]:
    """Return the list of distinct, non-blank issuer IDs in the CSV."""
    col = CONFIG["ISSUER_ID_FIELD"]
    if col not in ext_df.columns:
        raise ValueError(
            f"Column '{col}' not found in the extraction CSV. "
            "The issuer ID cannot be determined automatically."
        )
    series = ext_df[col].astype(str).str.strip()
    ids = [v for v in series.unique().tolist() if v and v.lower() != "nan"]
    if not ids:
        raise ValueError(
            f"No issuer IDs found in column '{col}'. The column is empty."
        )
    logger.info(f"Detected {len(ids)} issuer ID(s): {', '.join(ids)}")
    return ids


# ==============================================================================
# SECTION 4: DVV FILE PROCESSING
# ==============================================================================

def load_dvv(file_bytes: bytes, display_name: str,
             logger: logging.Logger) -> pd.DataFrame:
    """
    Load DVV Merged XLSX (from uploaded bytes). Read as strings, keep only
    rows where Correct Value (Column AG) is non-blank.
    """
    logger.info(f"Loading DVV file: {display_name}")

    def col_letter_to_index(letter: str) -> int:
        letter = letter.upper()
        result = 0
        for ch in letter:
            result = result * 26 + (ord(ch) - ord('A') + 1)
        return result - 1

    try:
        dvv_df = pd.read_excel(
            io.BytesIO(file_bytes), sheet_name=0, dtype=str,
            header=0, engine="openpyxl"
        )
    except Exception as e:
        raise RuntimeError(f"Cannot read DVV file '{display_name}': {e}")

    if dvv_df.empty:
        raise ValueError(f"DVV file is empty: {display_name}")

    datalib_idx    = col_letter_to_index(CONFIG["DVV_DATALIB_COL"])
    correctval_idx = col_letter_to_index(CONFIG["DVV_CORRECTVAL_COL"])
    series_idx     = col_letter_to_index(CONFIG["DVV_SERIES_COL"])

    max_needed = max(datalib_idx, correctval_idx, series_idx)
    if dvv_df.shape[1] <= max_needed:
        raise ValueError(
            f"DVV file has only {dvv_df.shape[1]} columns; "
            f"need at least {max_needed + 1} "
            f"(up to column {CONFIG['DVV_SERIES_COL']})."
        )

    cols = dvv_df.columns.tolist()
    logger.info(
        f"DVV columns -> DataLib='{cols[datalib_idx]}' "
        f"CorrectVal='{cols[correctval_idx]}' SeriesID='{cols[series_idx]}'"
    )

    dvv_df = dvv_df.rename(columns={
        cols[datalib_idx]:    "_DVV_DATALIB",
        cols[correctval_idx]: "_DVV_CORRECT_VALUE",
        cols[series_idx]:     "_DVV_SERIES_ID",
    })

    before = len(dvv_df)
    dvv_df = dvv_df[
        dvv_df["_DVV_CORRECT_VALUE"].notna()
        & (dvv_df["_DVV_CORRECT_VALUE"].str.strip() != "")
    ].copy()

    logger.info(
        f"DVV loaded: {before} total rows, "
        f"{len(dvv_df)} rows with Correct Values"
    )
    return dvv_df


# ==============================================================================
# SECTION 5: TEMPLATE POPULATION
# ==============================================================================

def populate_template(
    template_bytes: bytes,
    template_key: str,
    ext_df: pd.DataFrame,
    issuer_id: str,
    logger: logging.Logger,
) -> tuple[openpyxl.Workbook, dict]:
    """
    Populate all sheets in a template workbook from the extraction DataFrame.

    Scaler      -> one issuer-level row (no Series ID, no blank-row check).
    Position /  -> one output row per extraction row. Series ID from Col E.
    Series 1/2     Rows where all template data fields are blank are skipped.

    Matching: exact Data Lib match only. Unmatched headers -> blank + logged.
    """
    logger.info(f"Populating '{template_key}'")
    wb = load_workbook(io.BytesIO(template_bytes))
    sheet_meta: dict = {}
    is_scalar = (template_key == "Scaler")

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        headers = [c.value for c in ws[1] if c.value is not None]
        if not headers:
            logger.warning(f"  Sheet '{sheet_name}' has no headers - skipping.")
            continue

        header_col_map = {
            cell.value: col_idx
            for col_idx, cell in enumerate(ws[1], start=1)
            if cell.value is not None
        }

        data_start_row = 2
        populated_rows = 0

        # -- SCALAR: one issuer row --------------------------------------------
        if is_scalar:
            if CONFIG["ISSUER_ID_FIELD"] in ext_df.columns:
                issuer_rows = ext_df[
                    ext_df[CONFIG["ISSUER_ID_FIELD"]].astype(str).str.strip()
                    == issuer_id
                ]
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

        # -- POSITION / SERIES: one row per extraction row ---------------------
        else:
            for _, ext_row in ext_df.iterrows():
                if not row_has_data(ext_row, headers):
                    continue

                write_row = data_start_row + populated_rows

                for tmpl_header, col_idx in header_col_map.items():
                    if tmpl_header == CONFIG["ISSUER_ID_FIELD"]:
                        ws.cell(row=write_row, column=col_idx).value = issuer_id
                    elif tmpl_header == "serial_id":
                        series_id = get_series_id_from_row(ext_row, ext_df)
                        ws.cell(row=write_row, column=col_idx).value = series_id
                    elif tmpl_header in ext_df.columns:
                        val = ext_row.get(tmpl_header)
                        ws.cell(row=write_row, column=col_idx).value = (
                            None if (pd.isna(val) or str(val).strip() in ("", "nan"))
                            else str(val).strip()
                        )

                populated_rows += 1

        sheet_meta[sheet_name] = {
            "headers":        headers,
            "header_col_map": header_col_map,
            "data_start_row": data_start_row,
            "populated_rows": populated_rows,
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
) -> list[dict]:
    """
    Apply DVV Correct Value overrides to the populated workbook.
    Scalar -> match on Data Lib. Others -> match on Series ID + Data Lib.
    Updated cells highlighted with DVV_HIGHLIGHT_COLOR. Returns audit records.
    """
    dvv_fill = PatternFill(
        "solid",
        start_color=CONFIG["DVV_HIGHLIGHT_COLOR"],
        end_color=CONFIG["DVV_HIGHLIGHT_COLOR"],
    )
    is_scalar  = (template_key == "Scaler")
    audit_recs: list[dict] = []
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for sheet_name in wb.sheetnames:
        if sheet_name not in sheet_meta:
            continue

        ws             = wb[sheet_name]
        meta           = sheet_meta[sheet_name]
        header_col_map = meta["header_col_map"]
        data_start_row = meta["data_start_row"]
        populated_rows = meta["populated_rows"]

        row_series_map: dict[int, str | None] = {}
        if "serial_id" in header_col_map:
            sc = header_col_map["serial_id"]
            for r in range(data_start_row, data_start_row + populated_rows):
                v = ws.cell(row=r, column=sc).value
                row_series_map[r] = str(v).strip() if v else None
        else:
            for r in range(data_start_row, data_start_row + populated_rows):
                row_series_map[r] = None

        for _, dvv_row in dvv_df.iterrows():
            datalib     = str(dvv_row.get("_DVV_DATALIB", "")).strip()
            correct_val = dvv_row.get("_DVV_CORRECT_VALUE")
            dvv_series  = str(dvv_row.get("_DVV_SERIES_ID", "")).strip()

            if not datalib or datalib not in header_col_map:
                continue

            target_col = header_col_map[datalib]

            if is_scalar:
                target_rows = list(range(
                    data_start_row, data_start_row + populated_rows
                ))
            else:
                if dvv_series:
                    target_rows = [
                        r for r, sid in row_series_map.items()
                        if sid and sid == dvv_series
                    ]
                else:
                    logger.warning(
                        f"DVV row has blank Series ID for DataLib='{datalib}' "
                        f"sheet='{sheet_name}'. Logged as exception."
                    )
                    audit_recs.append({
                        "Issuer ID": issuer_id, "Series ID": "",
                        "Data Lib": datalib, "Old Value": "",
                        "New Value": str(correct_val), "Update Timestamp": ts,
                        "Template Name": template_key, "Worksheet Name": sheet_name,
                        "Status": "EXCEPTION: Blank DVV Series ID",
                    })
                    continue

            if not target_rows:
                logger.warning(
                    f"DVV: No matching row for Issuer={issuer_id}, "
                    f"Series={dvv_series}, DataLib={datalib}, sheet='{sheet_name}'."
                )
                audit_recs.append({
                    "Issuer ID": issuer_id, "Series ID": dvv_series,
                    "Data Lib": datalib, "Old Value": "",
                    "New Value": str(correct_val), "Update Timestamp": ts,
                    "Template Name": template_key, "Worksheet Name": sheet_name,
                    "Status": "EXCEPTION: No matching row found",
                })
                continue

            for row_num in target_rows:
                cell      = ws.cell(row=row_num, column=target_col)
                old_value = cell.value
                cell.value = correct_val
                cell.fill  = dvv_fill
                audit_recs.append({
                    "Issuer ID": issuer_id,
                    "Series ID": row_series_map.get(row_num) or "",
                    "Data Lib": datalib,
                    "Old Value": str(old_value) if old_value is not None else "",
                    "New Value": str(correct_val),
                    "Update Timestamp": ts,
                    "Template Name": template_key,
                    "Worksheet Name": sheet_name,
                    "Status": "SUCCESS",
                })

    success = sum(1 for r in audit_recs if r["Status"] == "SUCCESS")
    excepts = sum(1 for r in audit_recs if "EXCEPTION" in r["Status"])
    logger.info(
        f"DVV overrides '{template_key}': {success} changes, {excepts} exceptions"
    )
    return audit_recs


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
    ws,
    data_start_row: int,
    populated_rows: int,
    table_suffix: str,
    logger: logging.Logger,
) -> None:
    """Header styling, data borders, auto filter, freeze, auto-fit, Excel table.
    DVV green highlights are preserved (data loop sets font/border only)."""
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

    for r in range(data_start_row, last_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font   = _dat_font()
            cell.border = border

    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}1"
    ws.freeze_panes = "A2"

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value), default=8
        )
        ws.column_dimensions[col_letter].width = min(max_len + 4, 60)

    if populated_rows > 0:
        safe_sheet  = re.sub(r"[^A-Za-z0-9_]", "_", ws.title)
        safe_suffix = re.sub(r"[^A-Za-z0-9_]", "_", table_suffix)
        tbl_name    = f"Tbl_{safe_suffix}_{safe_sheet}"[:255]
        tbl_ref     = f"A1:{get_column_letter(max_col)}{last_row}"
        try:
            tbl = Table(displayName=tbl_name, ref=tbl_ref)
            tbl.tableStyleInfo = TableStyleInfo(
                name=CONFIG["TABLE_STYLE"],
                showFirstColumn=False, showLastColumn=False,
                showRowStripes=True, showColumnStripes=False,
            )
            ws.add_table(tbl)
        except Exception as e:
            logger.warning(f"Could not create Excel table in '{ws.title}': {e}")


# ==============================================================================
# SECTION 8: VALIDATION LOG
# ==============================================================================

def build_validation_log(
    all_template_headers: dict,
    ext_df: pd.DataFrame,
    issuer_id: str,
    logger: logging.Logger,
) -> list[dict]:
    """Validation records: template Data Libs missing in extraction, and
    extraction Data Libs unused in any template."""
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
                        "Template": tmpl_key, "Sheet": sheet_name, "Data Lib": h,
                        "Detail": "Header in template; no matching column in CSV",
                    })

    for col in ext_cols:
        if col not in all_tmpl_cols and col not in {CONFIG["ISSUER_ID_FIELD"], "serial_id"}:
            records.append({
                "Check Type": "Extraction DataLib unused in Templates",
                "Template": "ALL", "Sheet": "ALL", "Data Lib": col,
                "Detail": "Column in CSV; not used in any template sheet",
            })

    logger.info(f"Validation: {len(records)} issues found")
    return records


def _write_log_sheet(
    wb: openpyxl.Workbook,
    sheet_name: str,
    columns: list[str],
    records: list[dict],
    logger: logging.Logger,
) -> None:
    """Generic helper to write a formatted log sheet into a workbook."""
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    border = _thin_border()

    for ci, col_name in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.font      = _hdr_font()
        cell.fill      = _hdr_fill()
        cell.alignment = _center()
        cell.border    = border

    for ri, record in enumerate(records, start=2):
        for ci, col_name in enumerate(columns, start=1):
            cell = ws.cell(row=ri, column=ci, value=record.get(col_name, ""))
            cell.font   = _dat_font()
            cell.border = border

    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value), default=10
        )
        ws.column_dimensions[col_letter].width = min(max_len + 4, 80)

    ws.freeze_panes = "A2"
    logger.info(f"'{sheet_name}' written: {len(records)} records.")


def add_validation_sheet(wb, records, logger):
    _write_log_sheet(
        wb, "Validation_Log",
        ["Check Type", "Template", "Sheet", "Data Lib", "Detail"],
        records, logger
    )


def add_dvv_audit_sheet(wb, records, logger):
    _write_log_sheet(
        wb, "DVV_Audit_Log",
        ["Issuer ID", "Series ID", "Data Lib", "Old Value", "New Value",
         "Update Timestamp", "Template Name", "Worksheet Name", "Status"],
        records, logger
    )


# ==============================================================================
# SECTION 9: ORCHESTRATION (in memory)
# ==============================================================================

def _wb_to_bytes(wb: openpyxl.Workbook) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def process_one_issuer(
    issuer_id: str,
    issuer_ext_df: pd.DataFrame,
    full_ext_df: pd.DataFrame,
    dvv_df: pd.DataFrame,
    templates: dict[str, bytes],
    logger: logging.Logger,
) -> tuple[list[tuple[str, bytes]], dict]:
    """
    Run the pipeline for a single issuer.
    Returns (list of (filename, bytes), per-issuer summary dict).
    """
    outputs: list[tuple[str, bytes]] = []
    all_template_headers: dict[str, dict[str, list]] = {}
    all_audit_records: list[dict] = []
    total_dvv_overrides = 0
    total_fields = 0
    errors: list[str] = []

    logger.info("=" * 60)
    logger.info(f"Processing issuer: {issuer_id}  ({len(issuer_ext_df)} rows)")

    for tmpl_key, tmpl_bytes in templates.items():
        logger.info(f"  -- {tmpl_key} --")

        try:
            wb, sheet_meta = populate_template(
                tmpl_bytes, tmpl_key, issuer_ext_df, issuer_id, logger
            )
        except Exception as e:
            msg = f"Population failed for '{tmpl_key}': {e}"
            logger.error(msg)
            logger.debug(traceback.format_exc())
            errors.append(msg)
            continue

        all_template_headers[tmpl_key] = {
            sn: m["headers"] for sn, m in sheet_meta.items()
        }
        for m in sheet_meta.values():
            total_fields += len(m["headers"]) * m["populated_rows"]

        audit_recs: list[dict] = []
        try:
            audit_recs = apply_dvv_overrides(
                wb, sheet_meta, dvv_df, issuer_id, tmpl_key, logger
            )
            all_audit_records.extend(audit_recs)
            total_dvv_overrides += sum(
                1 for r in audit_recs if r["Status"] == "SUCCESS"
            )
        except Exception as e:
            msg = f"DVV override failed for '{tmpl_key}': {e}"
            logger.error(msg)
            logger.debug(traceback.format_exc())
            errors.append(msg)

        try:
            for sn in wb.sheetnames:
                if sn in ("Validation_Log", "DVV_Audit_Log"):
                    continue
                if sn in sheet_meta:
                    m = sheet_meta[sn]
                    format_worksheet(
                        wb[sn], m["data_start_row"], m["populated_rows"],
                        tmpl_key, logger,
                    )
        except Exception as e:
            logger.error(f"Formatting failed for '{tmpl_key}': {e}")
            logger.debug(traceback.format_exc())

        try:
            add_dvv_audit_sheet(wb, audit_recs, logger)
        except Exception as e:
            logger.warning(f"DVV_Audit_Log not written for '{tmpl_key}': {e}")

        out_name = f"{issuer_id}_{CONFIG['OUTPUT_NAMES'][tmpl_key]}.xlsx"
        try:
            outputs.append((out_name, _wb_to_bytes(wb)))
            logger.info(f"  Built: {out_name}")
        except Exception as e:
            msg = f"Save failed for '{tmpl_key}': {e}"
            logger.error(msg)
            errors.append(msg)

    # Validation report
    val_records = build_validation_log(
        all_template_headers, issuer_ext_df, issuer_id, logger
    )
    for rec in all_audit_records:
        if "EXCEPTION" in rec.get("Status", ""):
            val_records.append({
                "Check Type": "DVV Override Exception",
                "Template": rec["Template Name"], "Sheet": rec["Worksheet Name"],
                "Data Lib": rec["Data Lib"], "Detail": rec["Status"],
            })

    try:
        vwb = openpyxl.Workbook()
        vwb.active.title = "Validation_Log"
        add_validation_sheet(vwb, val_records, logger)
        if "Sheet" in vwb.sheetnames:
            del vwb["Sheet"]
        outputs.append(
            (f"{issuer_id}_Validation_Report.xlsx", _wb_to_bytes(vwb))
        )
    except Exception as e:
        logger.warning(f"Validation report not written: {e}")

    # DVV audit report
    try:
        awb = openpyxl.Workbook()
        awb.active.title = "DVV_Audit_Log"
        add_dvv_audit_sheet(awb, all_audit_records, logger)
        if "Sheet" in awb.sheetnames:
            del awb["Sheet"]
        outputs.append(
            (f"{issuer_id}_DVV_Audit_Report.xlsx", _wb_to_bytes(awb))
        )
    except Exception as e:
        logger.warning(f"DVV Audit report not written: {e}")

    summary = {
        "Issuer ID": issuer_id,
        "Records Processed": len(issuer_ext_df),
        "Fields Populated": total_fields,
        "DVV Overrides": total_dvv_overrides,
        "Validation Issues": len(val_records),
        "Errors": errors,
    }
    return outputs, summary


def run_pipeline(
    csv_bytes: bytes, csv_name: str,
    dvv_bytes: bytes, dvv_name: str,
    templates: dict[str, bytes],
    logger: logging.Logger,
) -> tuple[bytes, list[dict], list[tuple[str, bytes]]]:
    """
    Full pipeline. Detects issuer ID(s) from the CSV, processes each, and
    bundles all output files into a single ZIP.
    Returns (zip_bytes, list of per-issuer summaries, list of (name, bytes)).
    """
    logger.info("=" * 60)
    logger.info("STEP 1: Loading extraction CSV")
    ext_df = load_extraction(csv_bytes, csv_name, logger)

    logger.info("=" * 60)
    logger.info("STEP 2: Detecting issuer ID(s)")
    issuer_ids = detect_issuer_ids(ext_df, logger)

    logger.info("=" * 60)
    logger.info("STEP 3: Loading DVV file")
    dvv_df = load_dvv(dvv_bytes, dvv_name, logger)

    issuer_col = CONFIG["ISSUER_ID_FIELD"]
    # One issuer -> use the full CSV (faithful to the original behaviour).
    # Many issuers -> filter the CSV per issuer.
    if len(issuer_ids) == 1:
        jobs = [(issuer_ids[0], ext_df)]
    else:
        jobs = [
            (iid, ext_df[ext_df[issuer_col].astype(str).str.strip() == iid].copy())
            for iid in issuer_ids
        ]

    logger.info("=" * 60)
    logger.info("STEP 4: Populating templates and applying DVV overrides")

    all_outputs: list[tuple[str, bytes]] = []
    summaries: list[dict] = []
    for issuer_id, issuer_df in jobs:
        outs, summ = process_one_issuer(
            issuer_id, issuer_df, ext_df, dvv_df, templates, logger
        )
        # When processing many issuers, keep files apart in per-issuer folders.
        if len(jobs) > 1:
            outs = [(f"{issuer_id}/{name}", data) for name, data in outs]
        all_outputs.extend(outs)
        summaries.append(summ)

    logger.info("=" * 60)
    logger.info("STEP 5: Building ZIP archive")
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in all_outputs:
            zf.writestr(name, data)
    zip_buf.seek(0)

    total_errors = sum(len(s["Errors"]) for s in summaries)
    logger.info("=" * 60)
    logger.info(f"DONE. Issuers: {len(summaries)} | "
                f"Files: {len(all_outputs)} | Errors: {total_errors}")
    return zip_buf.getvalue(), summaries, all_outputs


# ==============================================================================
# SECTION 10: STREAMLIT UI
# ==============================================================================

def main() -> None:
    st.set_page_config(
        page_title="MSCI Annual Update Processor",
        page_icon="📄",
        layout="centered",
    )

    st.title("MSCI Annual Update Factual Process")
    st.caption(
        "Upload the extraction CSV, the DVV merged file and the templates, "
        "then run. The Issuer ID is read automatically from the "
        "`DMX_ISSUER_ID` column of the CSV."
    )

    # ---- Uploads ------------------------------------------------------------
    st.subheader("1 · Input files")
    csv_file = st.file_uploader(
        "Extraction CSV", type=["csv"], key="csv",
        help="The IID..._IssuerName.csv extraction file.",
    )
    dvv_file = st.file_uploader(
        "DVV merged file (XLSX)", type=["xlsx"], key="dvv",
        help="Data Lib in column F, Correct Value in AG, Series ID in AL.",
    )

    st.subheader("2 · Templates")
    st.caption("Upload the templates you want to produce. Each one is optional.")
    template_files: dict[str, object] = {}
    slot_labels = list(CONFIG["TEMPLATE_SLOTS"].keys())
    cols = st.columns(2)
    for i, slot_label in enumerate(slot_labels):
        with cols[i % 2]:
            up = st.file_uploader(
                f"{slot_label} template", type=["xlsx"],
                key=f"tmpl_{slot_label}",
            )
            if up is not None:
                template_files[CONFIG["TEMPLATE_SLOTS"][slot_label]] = up

    # ---- Run ----------------------------------------------------------------
    st.subheader("3 · Run")
    ready = csv_file is not None and dvv_file is not None and len(template_files) > 0
    if not ready:
        st.info("Upload a CSV, a DVV file and at least one template to enable Run.")

    run = st.button("▶ Run", type="primary", disabled=not ready, use_container_width=True)

    if run:
        logger, log_buffer = setup_logging()
        try:
            with st.spinner("Processing…"):
                templates_bytes = {
                    key: up.getvalue() for key, up in template_files.items()
                }
                zip_bytes, summaries, outputs = run_pipeline(
                    csv_file.getvalue(), csv_file.name,
                    dvv_file.getvalue(), dvv_file.name,
                    templates_bytes, logger,
                )
            # Persist in session so the download button survives a rerun.
            st.session_state["zip_bytes"] = zip_bytes
            st.session_state["summaries"] = summaries
            st.session_state["log_text"]  = log_buffer.getvalue()
            st.session_state["n_files"]   = len(outputs)
            st.success("Processing complete.")
        except Exception as e:
            st.session_state.pop("zip_bytes", None)
            st.session_state["log_text"] = log_buffer.getvalue()
            st.error(f"Processing failed: {e}")
            with st.expander("Execution log", expanded=True):
                st.code(log_buffer.getvalue() or "(no log)", language="text")

    # ---- Results ------------------------------------------------------------
    if "zip_bytes" in st.session_state:
        st.subheader("4 · Results")

        summaries = st.session_state.get("summaries", [])
        total_records = sum(s["Records Processed"] for s in summaries)
        total_dvv     = sum(s["DVV Overrides"] for s in summaries)
        total_val     = sum(s["Validation Issues"] for s in summaries)
        total_err     = sum(len(s["Errors"]) for s in summaries)

        m = st.columns(4)
        m[0].metric("Issuers", len(summaries))
        m[1].metric("Records", total_records)
        m[2].metric("DVV overrides", total_dvv)
        m[3].metric("Validation issues", total_val)

        if total_err:
            st.warning(f"{total_err} error(s) occurred — see the per-issuer detail and log.")

        # Per-issuer breakdown
        if summaries:
            table = pd.DataFrame([{
                "Issuer ID": s["Issuer ID"],
                "Records": s["Records Processed"],
                "Fields populated": s["Fields Populated"],
                "DVV overrides": s["DVV Overrides"],
                "Validation issues": s["Validation Issues"],
                "Errors": len(s["Errors"]),
            } for s in summaries])
            st.dataframe(table, use_container_width=True, hide_index=True)

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        if len(summaries) == 1:
            zip_name = f"{summaries[0]['Issuer ID']}_Output_{ts}.zip"
        else:
            zip_name = f"AnnualUpdate_Output_{ts}.zip"

        st.download_button(
            label=f"⬇ Download all outputs ({st.session_state.get('n_files', 0)} files, ZIP)",
            data=st.session_state["zip_bytes"],
            file_name=zip_name,
            mime="application/zip",
            use_container_width=True,
        )

        with st.expander("Execution log"):
            st.code(st.session_state.get("log_text", "") or "(no log)", language="text")


if __name__ == "__main__":
    main()
