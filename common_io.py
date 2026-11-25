"""
Shared Read-Only Excel I/O, Dynamic Header Resolution & QDHS Brand Palette.

All input Excel files are strictly opened read-only and never modified.
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Tuple

import openpyxl


# =============================================================================
# QDHS Official Brand Palette (Pantone 186EC Red & Pantone 281EC Blue)
# =============================================================================
QDHS_RED_HEX = "#C8102E"        # Red Pantone 186EC (UI)
QDHS_RED_HOVER = "#A00D24"
QDHS_BLUE_HEX = "#00205B"       # Blue Pantone 281EC (UI)
QDHS_BLUE_HOVER = "#001640"

QDHS_RED_XL = "C8102E"          # Red Pantone 186EC (Excel)
QDHS_BLUE_XL = "00205B"         # Blue Pantone 281EC (Excel)

MANAGEBAC_HOST = "qibaodwight.managebac.cn"
MANAGEBAC_BEHAVIOR_NOTE_URL_TEMPLATE = (
    f"https://{MANAGEBAC_HOST}/teacher/users/{{mb_id}}/behavior_notes/new"
)


# =============================================================================
# Basic Value & Filename Helpers
# =============================================================================
def clean_str(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip()


def parse_number(val: Any) -> Optional[float]:
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s.upper() in {"N/A", "NA", "NONE", "-", "IN PROGRESS", "TO BE DETERMINED"}:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def sanitize_filename_part(text: str) -> str:
    """Keep Chinese, English, numbers, spaces, hyphens; replace forbidden path chars."""
    s = str(text or "").strip()
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s or "Unknown"


# =============================================================================
# Dynamic Header Parsing (Works with both standard & read_only worksheets)
# =============================================================================
def find_header_row(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    required_keywords: List[str],
    max_scan: int = 15,
    require_all: bool = False,
) -> Tuple[int, Dict[str, int]]:
    """
    Dynamically scan up to `max_scan` rows to find the best header row.
    Returns (1-based header_row_idx, {header_text: 1-based col_idx}).
    """
    best_row: Optional[int] = None
    best_hits = 0
    best_map: Dict[str, int] = {}

    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        row_map: Dict[str, int] = {}
        for c_idx, val in enumerate(row, start=1):
            if val is not None:
                key = str(val).strip()
                if key and key not in row_map:
                    row_map[key] = c_idx
        if not row_map:
            continue

        keys_lower = [k.lower() for k in row_map]
        hits = sum(
            1 for kw in required_keywords
            if any(kw.lower() == kl or kw.lower() in kl for kl in keys_lower)
        )
        if require_all and hits < len(required_keywords):
            continue
        if hits > best_hits:
            best_hits = hits
            best_row = r_idx
            best_map = row_map
            if hits == len(required_keywords):
                break

    if best_row is None or best_hits == 0:
        raise ValueError(
            f"Could not locate header row containing {required_keywords} in sheet '{ws.title}'."
        )
    return best_row, best_map


def resolve_col(
    header_map: Dict[str, int],
    candidates: List[str],
    required: bool = True,
) -> Optional[int]:
    """Resolve a 1-based column index dynamically (exact match first, then substring)."""
    norm_map = {re.sub(r"\s+", " ", k.strip()).lower(): v for k, v in header_map.items()}
    for cand in candidates:
        nc = re.sub(r"\s+", " ", cand.strip()).lower()
        if nc in norm_map:
            return norm_map[nc]
    for cand in candidates:
        nc = re.sub(r"\s+", " ", cand.strip()).lower()
        for k_norm, idx in norm_map.items():
            if nc in k_norm:
                return idx
    if required:
        raise KeyError(
            f"Required column not found. Tried: {candidates}. Available: {list(header_map.keys())}"
        )
    return None


def _row_val(row: Tuple[Any, ...], col_1based: Optional[int]) -> Any:
    if col_1based is None or col_1based < 1:
        return None
    idx = col_1based - 1
    return row[idx] if idx < len(row) else None


# =============================================================================
# Shared Auxiliary Data Loaders (Read-Only)
# =============================================================================
def load_attendance_map(
    attendance_path: Optional[str],
    required: bool = False,
) -> Dict[str, Tuple[str, Optional[int]]]:
    """
    Load Attendance Summary Excel dynamically.
    Returns {student_id: (mapped_tier, total_skip)}.
    Rule: 'Perfect Attendance' and 'Excellent' both map to 'Excellent'; others verbatim.
    """
    if not attendance_path:
        if required:
            raise FileNotFoundError("Attendance file path is empty.")
        return {}
    if not os.path.isfile(attendance_path):
        if required:
            raise FileNotFoundError(f"Attendance file not found: {attendance_path}")
        return {}

    wb = openpyxl.load_workbook(attendance_path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        hdr_row, hmap = find_header_row(ws, ["Student ID", "Total Skip"], max_scan=10, require_all=True)
        col_sid = resolve_col(hmap, ["Student ID"])
        col_skip = resolve_col(hmap, ["Total Skip"])
        col_tier = resolve_col(hmap, ["Tier"], required=False)

        result: Dict[str, Tuple[str, Optional[int]]] = {}
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(_row_val(row, col_sid))
            if not sid:
                continue
            tier_str = clean_str(_row_val(row, col_tier))
            if tier_str.lower() in {"perfect attendance", "excellent"}:
                tier_mapped = "Excellent"
            else:
                tier_mapped = tier_str

            skip_num = parse_number(_row_val(row, col_skip))
            skip_val: Optional[int] = int(skip_num) if skip_num is not None else None
            result[sid] = (tier_mapped, skip_val)
        return result
    finally:
        wb.close()


def load_student_id_map(
    id_map_path: Optional[str],
    required: bool = False,
) -> Dict[str, Tuple[str, str, str]]:
    """
    Load 学生MB系统ID对照表.xlsx dynamically across all sheets.
    Returns {simpo_id: (homeroom, mb_system_id, chinese_name)}.
    """
    if not id_map_path:
        if required:
            raise FileNotFoundError("Student MB ID mapping file path is empty.")
        return {}
    if not os.path.isfile(id_map_path):
        if required:
            raise FileNotFoundError(f"Student MB ID mapping file not found: {id_map_path}")
        return {}

    wb = openpyxl.load_workbook(id_map_path, read_only=True, data_only=True)
    try:
        result: Dict[str, Tuple[str, str, str]] = {}
        for sname in wb.sheetnames:
            ws = wb[sname]
            try:
                hdr_row, hmap = find_header_row(ws, ["SIMPO学号", "行政班", "MB系统ID"], max_scan=10)
            except ValueError:
                continue

            col_simpo = resolve_col(hmap, ["SIMPO学号", "student_id", "Student ID"], required=False)
            col_hr = resolve_col(hmap, ["行政班", "Homeroom"], required=False)
            col_mb = resolve_col(hmap, ["MB系统ID", "MB ID"], required=False)
            col_name = resolve_col(hmap, ["姓名", "Name"], required=False)
            if not col_simpo:
                continue

            for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
                simpo = clean_str(_row_val(row, col_simpo))
                if not simpo:
                    continue
                hr = clean_str(_row_val(row, col_hr))
                mb_raw = _row_val(row, col_mb)
                if isinstance(mb_raw, (int, float)) and float(mb_raw).is_integer():
                    mb_id = str(int(mb_raw))
                else:
                    mb_id = clean_str(mb_raw)
                cname = clean_str(_row_val(row, col_name))
                result[simpo] = (hr, mb_id, cname)

        if required and not result:
            raise ValueError(f"No valid student ID mapping rows found in '{id_map_path}'.")
        return result
    finally:
        wb.close()


def load_student_info_map(
    info_path: Optional[str],
    required: bool = False,
) -> Dict[str, Tuple[str, str, str, str]]:
    """
    Load Student Info Excel (Class of 2030/2029/2028/2027) dynamically.
    Returns {student_id: (student_email, parent1_email, parent2_email, preferred_name)}.
    """
    if not info_path:
        if required:
            raise FileNotFoundError("Student info file path is empty.")
        return {}
    if not os.path.isfile(info_path):
        if required:
            raise FileNotFoundError(f"Student info file not found: {info_path}")
        return {}

    wb = openpyxl.load_workbook(info_path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        hdr_row, hmap = find_header_row(ws, ["Student ID"], max_scan=15, require_all=True)
        col_sid = resolve_col(hmap, ["Student ID"])
        col_semail = resolve_col(hmap, ["E-mail Address", "Student Email"], required=False)
        col_p1 = resolve_col(hmap, ["Parent 1 E-mail address", "Parent 1 Email"], required=False)
        col_p2 = resolve_col(hmap, ["Parent 2 E-mail address", "Parent 2 Email"], required=False)
        col_pref = resolve_col(hmap, ["Preferred Name"], required=False)

        result: Dict[str, Tuple[str, str, str, str]] = {}
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(_row_val(row, col_sid))
            if not sid:
                continue
            s_email = clean_str(_row_val(row, col_semail))
            p1_email = clean_str(_row_val(row, col_p1))
            p2_email = clean_str(_row_val(row, col_p2))
            pref = clean_str(_row_val(row, col_pref))
            result[sid] = (s_email, p1_email, p2_email, pref)
        return result
    finally:
        wb.close()


def load_dwight_diploma_ids(dd_path: Optional[str]) -> set[str]:
    """
    Load Dwight Diploma ('DD') Student IDs from a student roster Excel
    (e.g. 'G11 Dwight Diploma.xlsx', 'G12 Dwight Diploma.xlsx').
    - If an 'Other Name' / 'Status' / 'Program' column exists, selects rows
      containing 'dwight diploma' (or exact 'dd').
    - Otherwise (a dedicated DD-only ID list), includes all non-empty Student IDs.
    """
    if not dd_path or not os.path.isfile(dd_path):
        return set()

    wb = openpyxl.load_workbook(dd_path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        try:
            hdr_row, hmap = find_header_row(ws, ["Student ID"], max_scan=15, require_all=True)
        except ValueError:
            return set()
        col_sid = resolve_col(hmap, ["Student ID"])
        col_flag = resolve_col(hmap, ["Other Name", "Status", "Program"], required=False)

        dd_ids: set[str] = set()
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(_row_val(row, col_sid))
            if not sid:
                continue
            if col_flag is not None:
                flag_val = clean_str(_row_val(row, col_flag)).lower()
                if "dwight diploma" in flag_val or flag_val == "dd":
                    dd_ids.add(sid)
            else:
                dd_ids.add(sid)
        return dd_ids
    finally:
        wb.close()
