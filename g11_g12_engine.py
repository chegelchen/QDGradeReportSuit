"""
G11 & G12 MT1 Grade Report Processor (IB & AL modes).
Generates a 3-sheet Excel workbook:
  1. MT1 Summary
  2. MT1 Overview
  3. Warning

Uses shared read-only loaders from common_io.py. Never modifies input files.
"""

from __future__ import annotations

import os
import re
import statistics
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from common_io import (
    QDHS_BLUE_XL,
    QDHS_RED_XL,
    clean_str,
    find_header_row,
    load_attendance_map,
    load_dwight_diploma_ids,
    load_student_id_map,
    load_student_info_map,
    parse_number,
    resolve_col,
)


# =============================================================================
# Visual Palette (QDHS Pantone Colors)
# =============================================================================
RED_PANTONE_186 = QDHS_RED_XL       # #C8102E
BLUE_PANTONE_281 = QDHS_BLUE_XL     # #00205B
FAIL_RED_FILL = "FFC7CE"
FAIL_RED_FONT = "9C0006"
ALT_GREEN_FILL = "E2EFDA"
WHITE_FILL = "FFFFFF"
SUMMARY_BLUE_FILL = "D9E1F2"

THIN_GRAY = Side(border_style="thin", color="BFBFBF")
BORDER_ALL = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)


# =============================================================================
# Subject Classification & Evaluation Helpers
# =============================================================================
CHINESE_NATIONAL_SUBJECTS = {"Geography", "History", "Politics", "Physical Education"}
AL_PASS_GRADES = {"A*", "A", "B", "C"}
TOK_PASS_GRADES = {"A", "B", "C"}

# Official IB TOK + EE Bonus Matrix (max 3 points; 'E' = Failing Condition -> 0)
TOK_EE_BONUS_MATRIX: Dict[Tuple[str, str], int] = {
    ("A", "A"): 3, ("A", "B"): 3, ("A", "C"): 2, ("A", "D"): 2,
    ("B", "A"): 3, ("B", "B"): 2, ("B", "C"): 2, ("B", "D"): 1,
    ("C", "A"): 2, ("C", "B"): 2, ("C", "C"): 1, ("C", "D"): 0,
    ("D", "A"): 2, ("D", "B"): 1, ("D", "C"): 0, ("D", "D"): 0,
}

# AL letter grade -> 7-point numeric scale (matches MS-Excel SWITCH formula)
AL_SCORE_MAP: Dict[str, int] = {
    "A*": 7,
    "A": 6,
    "B": 5,
    "C": 4,
    "D": 3,
    "E": 2,
    "U": 1,
    "F": 1,
    "G": 1,
}


def compute_tok_ee_bonus(tok_val: Any, ee_val: Any) -> int:
    """Return 0-3 bonus points from TOK & EE letter grades (A-E). Non-letter EE/TOK -> 0."""
    tok = clean_str(tok_val).upper()
    ee = clean_str(ee_val).upper()
    if tok in {"A", "B", "C", "D", "E"} and ee in {"A", "B", "C", "D", "E"}:
        return TOK_EE_BONUS_MATRIX.get((tok, ee), 0)
    return 0


def convert_al_grade_to_score(val: Any) -> Optional[float]:
    """Convert an AL / HS AL grade ('A*'..'G' or numeric 1-7) into a 1-7 numeric score."""
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = clean_str(val).upper()
    if s in AL_SCORE_MAP:
        return float(AL_SCORE_MAP[s])
    return parse_number(s)


@dataclass
class SubjectMeta:
    key: str            # Unique column key
    display_name: str   # Column header display name
    category: str       # 'IB_DP' | 'IB_CORE' | 'AL' | 'HS_CN' | 'HS_AL_IB' | 'HS_AL_ENG'
    sort_order: int     # Ordering priority


@dataclass
class StudentRecord:
    student_id: str
    name: str = ""
    homeroom: str = ""
    mb_id: str = ""
    status: str = ""    # 'DP' | 'DD' | 'AL'
    student_email: str = ""
    parent1_email: str = ""
    parent2_email: str = ""
    attendance_tier: str = ""
    attendance_skip: Optional[int] = None
    grades: Dict[str, Any] = field(default_factory=dict)
    failed_flags: Dict[str, bool] = field(default_factory=dict)


def _normalize_ib_subject(subject: str, level: str, lang_level: str, class_name: str) -> Tuple[str, str]:
    subj_clean = subject.strip()
    cn_lower = class_name.lower()
    lvl = level.strip().upper() if level else ""

    if "theory of knowledge" in subj_clean.lower() or "tok" in subj_clean.lower():
        return "TOK", "IB_CORE"
    if "extended essay" in subj_clean.lower() or subj_clean.upper() == "EE":
        return "EE", "IB_CORE"
    if subj_clean.upper() == "CAS":
        return "CAS", "IB_CORE"

    if subj_clean.lower() == "chinese":
        if "chinese b" in cn_lower or lang_level.strip().upper() == "B":
            label = f"Chinese B {lvl}".strip()
        elif "language & literature" in cn_lower or "language and literature" in cn_lower:
            label = f"Chinese A L&L {lvl}".strip()
        elif "literature" in cn_lower:
            label = f"Chinese A Lit {lvl}".strip()
        else:
            label = f"Chinese {lvl}".strip()
        return f"IB: {label}", "IB_DP"

    if subj_clean.lower() == "english":
        if "english b" in cn_lower or lang_level.strip().upper() == "B":
            label = f"English B {lvl}".strip()
        elif "language & literature" in cn_lower or "language and literature" in cn_lower:
            label = f"English A L&L {lvl}".strip()
        elif "literature" in cn_lower:
            label = f"English A Lit {lvl}".strip()
        else:
            label = f"English {lvl}".strip()
        return f"IB: {label}", "IB_DP"

    if subj_clean.lower() == "mathematics: analysis and approaches":
        return f"IB: {f'Math AA {lvl}'.strip()}", "IB_DP"
    if subj_clean.lower() == "mathematics: applications and interpretation":
        return f"IB: {f'Math AI {lvl}'.strip()}", "IB_DP"
    if subj_clean.lower() == "environmental systems & societies":
        return f"IB: {f'ESS {lvl}'.strip()}", "IB_DP"

    return f"IB: {f'{subj_clean} {lvl}'.strip()}", "IB_DP"


def _normalize_hs_subject(subject: str, class_name: str) -> Tuple[str, str]:
    subj_clean = subject.strip()
    cn_lower = class_name.lower()
    if subj_clean in CHINESE_NATIONAL_SUBJECTS:
        return f"HS {subj_clean}", "HS_CN"
    if subj_clean.lower() == "chinese" or "hs al chn" in cn_lower:
        return "HS AL Chinese", "HS_AL_IB"
    if "english" in subj_clean.lower() or "hs al eng" in cn_lower:
        return "HS AL English", "HS_AL_ENG"
    return f"HS {subj_clean}", "HS_CN"


def _normalize_al_subject(subject: str, level: str) -> Tuple[str, str]:
    return f"AL {subject.strip()}", "AL"


def evaluate_grade(raw_val: Any, category: str, subject_key: str) -> Tuple[Any, bool]:
    """
    Normalize a raw grade value and return (normalized_value, is_fail).
    Rules:
      - IB_DP / HS_AL_IB (1-7): numeric < 4 -> Fail
      - IB_CORE:
          * TOK: A-E letter grade (case-insensitive -> upper); not in {A, B, C} or 'concern' -> Fail
          * EE / CAS: 'concern' (or D/E/F/U) -> Fail; 'To Be Determined', '', 'N/A', 'On-track', 'Excellent' -> NOT Fail
      - AL / HS_AL_ENG: A*-E/U/F; not in {A*, A, B, C} -> Fail
      - HS_CN (Geography / History / Politics / PE): numeric < 65 -> Fail
      - Any subject containing 'concern' -> Fail
    """
    s = clean_str(raw_val)
    if not s or s.upper() in {"N/A", "NA", "NONE", "-", "IN PROGRESS", "TO BE DETERMINED"}:
        return (s if s in {"In Progress", "To Be Determined"} else None), False

    s_lower = s.lower()
    if "concern" in s_lower:
        return "Concern", True

    if category in {"IB_DP", "HS_AL_IB"}:
        num = parse_number(raw_val)
        if num is not None:
            val = int(num) if num.is_integer() else round(num, 2)
            return val, (num < 4)
        return s, False

    if category == "IB_CORE":
        if subject_key == "TOK":
            up = s.upper()
            if up in {"A", "B", "C", "D", "E", "F", "U"}:
                return up, (up not in TOK_PASS_GRADES)
            return s, False
        if "excellent" in s_lower:
            return "Excellent", False
        if "on-track" in s_lower or "on track" in s_lower:
            return "On-track", False
        up = s.upper()
        if up in {"A", "B", "C", "D", "E", "F", "U"}:
            return up, (up not in TOK_PASS_GRADES)
        return s, False

    if category in {"AL", "HS_AL_ENG"}:
        up = s.upper()
        if up in {"A*", "A", "B", "C", "D", "E", "F", "U"}:
            return up, (up not in AL_PASS_GRADES)
        return s, False

    if category == "HS_CN":
        num = parse_number(raw_val)
        if num is not None:
            val = int(num) if num.is_integer() else round(num, 2)
            return val, (num < 65)
        return s, False

    return s, False


# =============================================================================
# Optional IB Core Loaders (CAS & EE)
# =============================================================================
def load_cas_export(path: str) -> Dict[str, str]:
    """Load standalone CAS Export file -> Sheet 'Summary', Col 'Status'."""
    result: Dict[str, str] = {}
    if not path or not os.path.isfile(path):
        return result
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Summary"] if "Summary" in wb.sheetnames else wb[wb.sheetnames[0]]
        try:
            hdr_row, hmap = find_header_row(ws, ["Student ID", "Status"], max_scan=15, require_all=True)
        except ValueError:
            return result
        col_sid = resolve_col(hmap, ["Student ID"])
        col_status = resolve_col(hmap, ["Status"], required=False)
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(row[col_sid - 1] if col_sid - 1 < len(row) else None)
            if not sid:
                continue
            status = clean_str(row[col_status - 1] if col_status and col_status - 1 < len(row) else None)
            if status:
                result[sid] = status
        return result
    finally:
        wb.close()


def load_ee_export(path: str) -> Dict[str, str]:
    """Load standalone EE Export file -> Sheet 'Reflections', Col 'Progress'."""
    result: Dict[str, str] = {}
    if not path or not os.path.isfile(path):
        return result
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb["Reflections"] if "Reflections" in wb.sheetnames else wb[wb.sheetnames[0]]
        try:
            hdr_row, hmap = find_header_row(ws, ["Student ID", "Progress"], max_scan=15, require_all=True)
        except ValueError:
            return result
        col_sid = resolve_col(hmap, ["Student ID"])
        col_prog = resolve_col(hmap, ["Progress"], required=False)
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(row[col_sid - 1] if col_sid - 1 < len(row) else None)
            if not sid:
                continue
            prog = clean_str(row[col_prog - 1] if col_prog and col_prog - 1 < len(row) else None)
            if prog:
                result[sid] = prog
        return result
    finally:
        wb.close()


# =============================================================================
# Core Report Generator
# =============================================================================
def generate_g11_g12_report(
    raw_files: List[str],
    output_path: str,
    mode: str = "IB",
    attendance_path: str = "",
    mb_id_path: str = "",
    student_info_path: str = "",
    cas_path: str = "",
    ee_path: str = "",
    dwight_diploma_path: str = "",
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    Parse all input files and generate the 3-sheet MT1 Excel workbook for G11/G12 (IB or AL).
    """
    def log(msg: str) -> None:
        if progress_cb:
            progress_cb(msg)

    mode = mode.strip().upper()
    if mode not in {"IB", "AL"}:
        raise ValueError(f"Unsupported mode for G11/G12 engine: {mode}. Must be 'IB' or 'AL'.")
    valid_raw_files = [f for f in raw_files if f and os.path.isfile(f)]
    if not valid_raw_files:
        raise ValueError("Please select at least one valid raw grade Excel file.")

    log(f"Loading auxiliary data (Mode: {mode})... / 正在加载辅助数据（模式：{mode}）...")
    mb_map = load_student_id_map(mb_id_path, required=False)
    info_map = load_student_info_map(student_info_path, required=False)
    att_map = load_attendance_map(attendance_path, required=False)
    cas_map = load_cas_export(cas_path) if mode == "IB" else {}
    ee_map = load_ee_export(ee_path) if mode == "IB" else {}
    dd_ids = load_dwight_diploma_ids(dwight_diploma_path) if mode == "IB" else set()
    if mode == "IB" and not dd_ids and student_info_path:
        dd_ids = load_dwight_diploma_ids(student_info_path)

    log(
        f"Loaded mappings -> MB IDs: {len(mb_map)}, Student Info: {len(info_map)}, "
        f"Attendance: {len(att_map)}, CAS: {len(cas_map)}, EE: {len(ee_map)}, DD: {len(dd_ids)}"
    )

    base_student_ids: set[str] = set()
    all_students: Dict[str, StudentRecord] = {}
    subjects_meta: Dict[str, SubjectMeta] = {}
    detected_grades: Dict[str, int] = defaultdict(int)

    for fpath in valid_raw_files:
        wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            try:
                hdr_row, hmap = find_header_row(ws, ["Student ID", "Subject", "Mid-term Grade"], max_scan=10, require_all=True)
            except ValueError:
                log(f"Warning: Could not find grade headers in {os.path.basename(fpath)}, skipping.")
                continue

            col_class = resolve_col(hmap, ["Class Name"], required=False)
            col_subj = resolve_col(hmap, ["Subject"])
            col_lvl = resolve_col(hmap, ["Level"], required=False)
            col_lang = resolve_col(hmap, ["Language Level"], required=False)
            col_gl = resolve_col(hmap, ["Grade Level"], required=False)
            col_sid = resolve_col(hmap, ["Student ID"])
            col_pref = resolve_col(hmap, ["Preferred Name"], required=False)
            col_last = resolve_col(hmap, ["Last Name"], required=False)
            col_first = resolve_col(hmap, ["First Name"], required=False)
            col_mt = resolve_col(hmap, ["Mid-term Grade"])

            # Find exact header string for col_mt to detect IB / AL / HS
            mt_hdr = next((k for k, v in hmap.items() if v == col_mt), "")
            fname_upper = os.path.basename(fpath).upper()

            if "(1-7)" in mt_hdr or "DIPLOMA" in ws.title.upper() or " IB" in fname_upper or "_IB" in fname_upper:
                file_kind = "IB"
            elif "(A*-E)" in mt_hdr or "CAM_A" in ws.title.upper() or " AL" in fname_upper or "_AL" in fname_upper:
                file_kind = "AL"
            else:
                file_kind = "HS"

            log(f"Parsing {os.path.basename(fpath)} (detected: {file_kind}, col: '{mt_hdr}')...")

            def _get(row: Tuple[Any, ...], c: Optional[int]) -> Any:
                return row[c - 1] if c and (c - 1) < len(row) else None

            for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
                sid = clean_str(_get(row, col_sid))
                if not sid:
                    continue

                class_name = clean_str(_get(row, col_class))
                subject = clean_str(_get(row, col_subj))
                if "workshop" in class_name.lower() or "workshop" in subject.lower():
                    continue

                level = clean_str(_get(row, col_lvl))
                lang_level = clean_str(_get(row, col_lang))
                grade_lvl = clean_str(_get(row, col_gl))
                mt_val = _get(row, col_mt)

                if file_kind == mode:
                    base_student_ids.add(sid)
                    m_gl = re.search(r"(\d+)", grade_lvl)
                    if m_gl:
                        detected_grades[m_gl.group(1)] += 1

                if sid not in all_students:
                    pref = clean_str(_get(row, col_pref))
                    if not pref:
                        ln = clean_str(_get(row, col_last))
                        fn = clean_str(_get(row, col_first))
                        pref = f"{fn} {ln}".strip()
                    all_students[sid] = StudentRecord(student_id=sid, name=pref)

                if file_kind == "IB":
                    s_key, cat = _normalize_ib_subject(subject, level, lang_level, class_name)
                    sort_ord = 10 if cat == "IB_DP" else 20
                elif file_kind == "AL":
                    s_key, cat = _normalize_al_subject(subject, level)
                    sort_ord = 15
                else:
                    s_key, cat = _normalize_hs_subject(subject, class_name)
                    sort_ord = 30 if cat == "HS_CN" else 25

                if s_key not in subjects_meta:
                    disp = s_key.replace("IB: ", "")
                    subjects_meta[s_key] = SubjectMeta(key=s_key, display_name=disp, category=cat, sort_order=sort_ord)

                norm_val, is_fail = evaluate_grade(mt_val, cat, s_key)
                if norm_val is not None:
                    all_students[sid].grades[s_key] = norm_val
                    all_students[sid].failed_flags[s_key] = is_fail
        finally:
            wb.close()

    if not base_student_ids:
        raise ValueError(
            f"No students found for mode '{mode}'. "
            f"Please ensure the corresponding {mode} raw grade file is included."
        )

    grade_num = max(detected_grades.items(), key=lambda x: x[1])[0] if detected_grades else "11"
    log(f"Detected Grade Level: G{grade_num} | Base Students ({mode}): {len(base_student_ids)}")

    students: List[StudentRecord] = []
    if mode == "IB":
        if "TOK" not in subjects_meta:
            subjects_meta["TOK"] = SubjectMeta("TOK", "Theory of Knowledge (TOK)", "IB_CORE", 20)
        else:
            subjects_meta["TOK"].display_name = "Theory of Knowledge (TOK)"
        if "EE" not in subjects_meta:
            subjects_meta["EE"] = SubjectMeta("EE", "Extended Essay (EE)", "IB_CORE", 21)
        else:
            subjects_meta["EE"].display_name = "Extended Essay (EE)"
            subjects_meta["EE"].sort_order = 21
        if cas_map:
            subjects_meta["CAS"] = SubjectMeta("CAS", "CAS", "IB_CORE", 22)

    for sid in sorted(base_student_ids):
        rec = all_students[sid]
        if sid in mb_map:
            hr, mb, mb_name = mb_map[sid]
            rec.homeroom = hr
            rec.mb_id = mb
            if not rec.name and mb_name:
                rec.name = mb_name
        if sid in info_map:
            em, p1, p2, pref = info_map[sid]
            rec.student_email = em
            rec.parent1_email = p1
            rec.parent2_email = p2
            if not rec.name and pref:
                rec.name = pref
        if sid in att_map:
            tier, skip = att_map[sid]
            rec.attendance_tier = tier
            rec.attendance_skip = skip

        if mode == "IB":
            rec.status = "DD" if sid in dd_ids else "DP"
            if sid in ee_map:
                norm_val, is_fail = evaluate_grade(ee_map[sid], "IB_CORE", "EE")
                if norm_val is not None:
                    rec.grades["EE"] = norm_val
                    rec.failed_flags["EE"] = is_fail
            if sid in cas_map:
                norm_val, is_fail = evaluate_grade(cas_map[sid], "IB_CORE", "CAS")
                if norm_val is not None:
                    rec.grades["CAS"] = norm_val
                    rec.failed_flags["CAS"] = is_fail
        else:
            rec.status = "AL"

        students.append(rec)

    students.sort(key=lambda r: (r.homeroom or "ZZZ", r.student_id))

    active_subjects: List[SubjectMeta] = []
    for s_key, meta in subjects_meta.items():
        if mode == "IB" and meta.category == "AL":
            continue
        if mode == "AL" and meta.category in {"IB_DP", "IB_CORE"}:
            continue
        has_data = any(s_key in st.grades for st in students)
        if has_data or (mode == "IB" and s_key in {"TOK", "EE", "CAS"} and (s_key != "CAS" or bool(cas_map))):
            active_subjects.append(meta)

    active_subjects.sort(key=lambda m: (m.sort_order, m.display_name))

    wb_out = openpyxl.Workbook()
    ws_summary = wb_out.active
    ws_summary.title = "MT1 Summary"
    ws_overview = wb_out.create_sheet(title="MT1 Overview")
    ws_warning = wb_out.create_sheet(title="Warning")

    warning_students = _build_summary_sheet(ws_summary, students, active_subjects, grade_num, mode)
    _build_overview_sheet(ws_overview, students, active_subjects, grade_num, mode)
    _build_warning_sheet(ws_warning, warning_students)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb_out.save(output_path)
    wb_out.close()

    log(f"Done! G{grade_num} {mode} | Students: {len(students)} | Warnings: {len(warning_students)} / 报表生成完成！")

    return {
        "output_path": output_path,
        "grade_num": grade_num,
        "mode": mode,
        "student_count": len(students),
        "warning_count": len(warning_students),
        "subject_count": len(active_subjects),
    }


# =============================================================================
# Sheet 1: MT1 Summary
# =============================================================================
def _build_summary_sheet(
    ws,
    students: List[StudentRecord],
    subjects: List[SubjectMeta],
    grade_num: str,
    mode: str,
) -> List[Tuple[StudentRecord, int, List[Tuple[str, Any]]]]:
    header_fill = PatternFill("solid", fgColor=BLUE_PANTONE_281)
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    red_fill = PatternFill("solid", fgColor=FAIL_RED_FILL)
    red_font = Font(name="Calibri", size=10, bold=True, color=FAIL_RED_FONT)
    normal_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    summary_fill = PatternFill("solid", fgColor=SUMMARY_BLUE_FILL)

    if mode == "IB":
        fixed_headers = [
            f"G{grade_num} MT1 Grades",
            "Student ID",
            "Candidate Session Number",
            "Personal Code",
            "Student Name\n学生姓名",
            "Homeroom 行政班级",
            "MB系统ID",
            "STATUS",
            "COMMENTS",
            f"G{grade_num} MT1 Total (DP Only)总计",
            f"G{grade_num} MT1 Total 总计",
        ]
        col_name = 5
        col_mb = 7
        col_total = 11
        total_summary_start_col = 10
        stat_lbl_col = 4
        freeze_cell = "F2"
        fixed_widths = {1: 12, 2: 13, 3: 14, 4: 13, 5: 24, 6: 18, 7: 14, 8: 10, 9: 14, 10: 15, 11: 14}
    else:
        fixed_headers = [
            f"G{grade_num} MT1 Grades",
            "Student ID",
            "Student Name\n学生姓名",
            "Homeroom 行政班级",
            "MB系统ID",
            "STATUS",
            "COMMENTS",
            f"G{grade_num} MT1 Total 总计",
        ]
        col_name = 3
        col_mb = 5
        col_total = 8
        total_summary_start_col = 8
        stat_lbl_col = 2
        freeze_cell = "D2"
        fixed_widths = {1: 12, 2: 13, 3: 24, 4: 18, 5: 14, 6: 10, 7: 14, 8: 14}

    subj_headers = [s.display_name for s in subjects]
    tail_headers = [
        "Attendance\n出勤等级",
        "attendance skip\n缺勤次数",
        "Number of Failed Subject\n不合格科目数量",
        "Number of Subjects",
    ]
    all_headers = fixed_headers + subj_headers + tail_headers

    ws.row_dimensions[1].height = 34
    for col_idx, text in enumerate(all_headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=text)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL

    dp_keys = {s.key for s in subjects if s.category == "IB_DP"}
    al_score_keys = {s.key for s in subjects if s.category in {"AL", "HS_AL_IB", "HS_AL_ENG"}}
    subj_col_start = len(fixed_headers) + 1
    col_att_tier = subj_col_start + len(subjects)
    col_att_skip = col_att_tier + 1
    col_fail_cnt = col_att_skip + 1
    col_subj_cnt = col_fail_cnt + 1

    warning_students: List[Tuple[StudentRecord, int, List[Tuple[str, Any]]]] = []
    homeroom_dp_totals: Dict[str, List[float]] = defaultdict(list)
    homeroom_al_grades: Dict[str, Dict[str, int]] = defaultdict(lambda: defaultdict(int))

    for row_offset, st in enumerate(students, start=2):
        ws.row_dimensions[row_offset].height = 20

        if mode == "IB":
            dp_scores = [
                float(st.grades[k])
                for k in dp_keys
                if k in st.grades and isinstance(st.grades[k], (int, float))
            ]
            base_total = sum(dp_scores) if dp_scores else None
            if base_total is not None and base_total.is_integer():
                base_total = int(base_total)

            if st.status == "DD":
                dp_only_total = None
                overall_total = base_total
            else:
                bonus = compute_tok_ee_bonus(st.grades.get("TOK"), st.grades.get("EE"))
                dp_only_total = (base_total + bonus) if base_total is not None else None
                overall_total = dp_only_total
        else:
            al_scores = [
                convert_al_grade_to_score(st.grades[k])
                for k in al_score_keys
                if k in st.grades
            ]
            valid_al_scores = [v for v in al_scores if v is not None]
            al_sum = sum(valid_al_scores) if valid_al_scores else None
            if al_sum is not None and al_sum.is_integer():
                al_sum = int(al_sum)
            dp_only_total = None
            overall_total = al_sum

        failed_list: List[Tuple[str, Any]] = []
        taken_count = 0
        for s in subjects:
            if s.key in st.grades and st.grades[s.key] is not None:
                taken_count += 1
                if st.failed_flags.get(s.key, False):
                    failed_list.append((s.display_name, st.grades[s.key]))
                if mode == "AL" and s.category == "AL":
                    homeroom_al_grades[st.homeroom or "Unknown"][str(st.grades[s.key]).upper()] += 1

        fail_count = len(failed_list)
        if mode == "IB" and overall_total is not None:
            homeroom_dp_totals[st.homeroom or "Unknown"].append(float(overall_total))

        if mode == "IB":
            row_vals = [
                row_offset - 1,
                st.student_id,
                "",
                "",
                st.name,
                st.homeroom,
                st.mb_id,
                st.status,
                "",
                dp_only_total,
                overall_total,
            ]
        else:
            row_vals = [
                row_offset - 1,
                st.student_id,
                st.name,
                st.homeroom,
                st.mb_id,
                st.status,
                "",
                overall_total,
            ]
        for c_i, val in enumerate(row_vals, start=1):
            c = ws.cell(row=row_offset, column=c_i, value=val)
            c.font = normal_font
            c.border = BORDER_ALL
            c.alignment = Alignment(horizontal="center" if c_i != col_name else "left", vertical="center")
            if c_i == col_mb:
                c.number_format = "@"

        for s_idx, s in enumerate(subjects):
            c_i = subj_col_start + s_idx
            val = st.grades.get(s.key, None)
            c = ws.cell(row=row_offset, column=c_i, value=val)
            c.border = BORDER_ALL
            c.alignment = Alignment(horizontal="center", vertical="center")
            if st.failed_flags.get(s.key, False):
                c.fill = red_fill
                c.font = red_font
            else:
                c.font = normal_font

        c_tier = ws.cell(row=row_offset, column=col_att_tier, value=st.attendance_tier)
        c_tier.font = normal_font
        c_tier.border = BORDER_ALL
        c_tier.alignment = Alignment(horizontal="center", vertical="center")

        c_skip = ws.cell(row=row_offset, column=col_att_skip, value=st.attendance_skip)
        c_skip.border = BORDER_ALL
        c_skip.alignment = Alignment(horizontal="center", vertical="center")
        if st.attendance_skip is not None and st.attendance_skip >= 9:
            c_skip.fill = red_fill
            c_skip.font = red_font
        else:
            c_skip.font = normal_font

        c_fc = ws.cell(row=row_offset, column=col_fail_cnt, value=fail_count)
        c_fc.border = BORDER_ALL
        c_fc.alignment = Alignment(horizontal="center", vertical="center")
        if fail_count >= 3:
            c_fc.fill = red_fill
            c_fc.font = red_font
        else:
            c_fc.font = normal_font

        c_sc = ws.cell(row=row_offset, column=col_subj_cnt, value=taken_count)
        c_sc.font = normal_font
        c_sc.border = BORDER_ALL
        c_sc.alignment = Alignment(horizontal="center", vertical="center")

        if fail_count >= 3 or (st.attendance_skip is not None and st.attendance_skip >= 9):
            warning_students.append((st, fail_count, failed_list))

    # Bottom Summary Rows
    last_student_row = len(students) + 1
    row_cnt = last_student_row + 2
    row_avg = last_student_row + 3
    dist_start_row = row_avg + 1
    r_7 = dist_start_row
    r_6 = dist_start_row + 1
    r_5 = dist_start_row + 2
    r_4 = dist_start_row + 3
    r_3 = dist_start_row + 4
    r_2 = dist_start_row + 5
    r_1 = dist_start_row + 6

    ws.cell(row=row_cnt, column=col_name, value="Number of Students 学生人数").font = bold_font
    ws.cell(row=row_avg, column=col_name, value="Subject Average 学科平均分").font = bold_font

    subj_by_col = {subj_col_start + idx: s for idx, s in enumerate(subjects)}
    for c_i in range(total_summary_start_col, col_subj_cnt + 1):
        col_ltr = get_column_letter(c_i)
        c1 = ws.cell(row=row_cnt, column=c_i, value=f"=COUNTA({col_ltr}2:{col_ltr}{last_student_row})")
        c1.font = bold_font
        c1.fill = summary_fill
        c1.border = BORDER_ALL
        c1.alignment = Alignment(horizontal="center", vertical="center")

        subj_meta = subj_by_col.get(c_i)
        if subj_meta and subj_meta.category in {"AL", "HS_AL_IB", "HS_AL_ENG"}:
            avg_formula = (
                f'=IF(SUM({col_ltr}{r_7}:{col_ltr}{r_1})>0, '
                f'(7*{col_ltr}{r_7}+6*{col_ltr}{r_6}+5*{col_ltr}{r_5}+4*{col_ltr}{r_4}'
                f'+3*{col_ltr}{r_3}+2*{col_ltr}{r_2}+1*{col_ltr}{r_1})'
                f'/SUM({col_ltr}{r_7}:{col_ltr}{r_1}), "")'
            )
        else:
            avg_formula = f'=IF(COUNT({col_ltr}2:{col_ltr}{last_student_row})>0, AVERAGE({col_ltr}2:{col_ltr}{last_student_row}), "")'

        c2 = ws.cell(row=row_avg, column=c_i, value=avg_formula)
        c2.font = bold_font
        c2.fill = summary_fill
        c2.border = BORDER_ALL
        c2.number_format = "0.00"
        c2.alignment = Alignment(horizontal="center", vertical="center")

    # 7-to-1 Score Distribution Rows below Subject Average (AL converted 7-1 & IB 7-1)
    total_col_ltr = get_column_letter(col_total)
    dist_specs = [
        (7, "7分 (A*)" if mode == "AL" else "#7 (A*)", ["7", "A~*"], "Median", f"=MEDIAN({total_col_ltr}2:{total_col_ltr}{last_student_row})"),
        (6, "6分 (A)" if mode == "AL" else "#6 (A)", ["6", "A"], "Max", f"=MAX({total_col_ltr}2:{total_col_ltr}{last_student_row})"),
        (5, "5分 (B)" if mode == "AL" else "#5 (B)", ["5", "B"], "Min", f"=MIN({total_col_ltr}2:{total_col_ltr}{last_student_row})"),
        (4, "4分 (C)" if mode == "AL" else "#4 (C)", ["4", "C"], None, None),
        (3, "3分 (D)" if mode == "AL" else "#3 (D)", ["3", "D"], None, None),
        (2, "2分 (E)" if mode == "AL" else "#2 (E)", ["2", "E"], None, None),
        (1, "1分 (U/F/G)" if mode == "AL" else "#1 (U/F/G)", ["1", "U", "F", "G"], None, None),
    ]
    for off, (_g_num, lbl, tokens, stat_lbl, stat_formula) in enumerate(dist_specs):
        r_dist = dist_start_row + off
        if stat_lbl:
            c_sl = ws.cell(row=r_dist, column=stat_lbl_col, value=stat_lbl)
            c_sl.font = bold_font
            c_sl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl = ws.cell(row=r_dist, column=col_name, value=lbl)
        c_lbl.font = bold_font
        c_lbl.fill = summary_fill
        c_lbl.border = BORDER_ALL
        c_lbl.alignment = Alignment(horizontal="left", vertical="center")

        if stat_formula:
            c_sf = ws.cell(row=r_dist, column=col_total, value=stat_formula)
            c_sf.font = bold_font
            c_sf.fill = summary_fill
            c_sf.border = BORDER_ALL
            c_sf.number_format = "0.0"
            c_sf.alignment = Alignment(horizontal="center", vertical="center")

        for s_idx, s in enumerate(subjects):
            c_i = subj_col_start + s_idx
            col_ltr = get_column_letter(c_i)
            if s.category in {"IB_DP", "AL", "HS_AL_IB", "HS_AL_ENG"}:
                expr = "=" + "+".join(f'COUNTIF({col_ltr}2:{col_ltr}{last_student_row},"{tok}")' for tok in tokens)
            else:
                expr = None
            cd = ws.cell(row=r_dist, column=c_i, value=expr)
            cd.font = normal_font
            cd.border = BORDER_ALL
            cd.alignment = Alignment(horizontal="center", vertical="center")

    # Right-Side Homeroom Statistics Table
    stat_col = col_subj_cnt + 2
    if mode == "IB":
        bands = [
            ("40-42", 40, 99), ("36-39", 36, 39.999), ("32-35", 32, 35.999),
            ("28-31", 28, 31.999), ("25-27", 25, 27.999), ("0-24", 0, 24.999),
        ]
        r_cur = 1
        for hr in sorted(homeroom_dp_totals.keys()):
            totals = homeroom_dp_totals[hr]
            for idx_h, txt in enumerate((f"{hr} (N={len(totals)})", "Count / Value")):
                ch = ws.cell(row=r_cur, column=stat_col + idx_h, value=txt)
                ch.fill = header_fill
                ch.font = header_font
                ch.border = BORDER_ALL
                ch.alignment = Alignment(horizontal="center", vertical="center")
            r_cur += 1
            for label, lo, hi in bands:
                cnt = sum(1 for v in totals if lo <= v <= hi)
                for idx_v, val in enumerate((label, cnt)):
                    cx = ws.cell(row=r_cur, column=stat_col + idx_v, value=val)
                    cx.font = normal_font
                    cx.border = BORDER_ALL
                    cx.alignment = Alignment(horizontal="center", vertical="center")
                r_cur += 1
            avg_val = round(sum(totals) / len(totals), 2) if totals else 0
            for idx_v, val in enumerate(("Average", avg_val)):
                cx = ws.cell(row=r_cur, column=stat_col + idx_v, value=val)
                cx.font = bold_font
                cx.fill = summary_fill
                cx.border = BORDER_ALL
                cx.alignment = Alignment(horizontal="center", vertical="center")
            r_cur += 2
    else:
        al_grades_order = ["A*", "A", "B", "C", "D", "E", "U"]
        r_cur = 1
        for hr in sorted(homeroom_al_grades.keys()):
            gmap = homeroom_al_grades[hr]
            for idx_h, txt in enumerate((f"{hr} AL Grades", "Count")):
                ch = ws.cell(row=r_cur, column=stat_col + idx_h, value=txt)
                ch.fill = header_fill
                ch.font = header_font
                ch.border = BORDER_ALL
                ch.alignment = Alignment(horizontal="center", vertical="center")
            r_cur += 1
            for g in al_grades_order:
                for idx_v, val in enumerate((g, gmap.get(g, 0))):
                    cx = ws.cell(row=r_cur, column=stat_col + idx_v, value=val)
                    cx.font = normal_font
                    cx.border = BORDER_ALL
                    cx.alignment = Alignment(horizontal="center", vertical="center")
                r_cur += 1
            r_cur += 1

    for ci, w in fixed_widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci in range(subj_col_start, col_subj_cnt + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 15
    ws.column_dimensions[get_column_letter(stat_col)].width = 20
    ws.column_dimensions[get_column_letter(stat_col + 1)].width = 14
    ws.freeze_panes = freeze_cell

    return warning_students


# =============================================================================
# Sheet 2: MT1 Overview
# =============================================================================
def _style_banner(ws, row: int, start_col: int, end_col: int, text: str, font_size: int = 12) -> None:
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
    fill = PatternFill("solid", fgColor=RED_PANTONE_186)
    font = Font(name="Calibri", size=font_size, bold=True, color="FFFFFF")
    ws.row_dimensions[row].height = 28
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill
        cell.border = BORDER_ALL
        if c == start_col:
            cell.value = text
            cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center")


def _style_subheaders(ws, row: int, headers: List[str], start_col: int = 1) -> None:
    fill = PatternFill("solid", fgColor=BLUE_PANTONE_281)
    font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    ws.row_dimensions[row].height = 26
    for idx, text in enumerate(headers):
        cell = ws.cell(row=row, column=start_col + idx, value=text)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL


def _build_overview_sheet(
    ws,
    students: List[StudentRecord],
    subjects: List[SubjectMeta],
    grade_num: str,
    mode: str,
) -> None:
    white_fill = PatternFill("solid", fgColor=WHITE_FILL)
    alt_fill = PatternFill("solid", fgColor=ALT_GREEN_FILL)
    normal_font = Font(name="Calibri", size=10)
    bold_font = Font(name="Calibri", size=10, bold=True)
    blue_fill = PatternFill("solid", fgColor=BLUE_PANTONE_281)
    blue_font_10 = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    blue_font_8 = Font(name="Calibri", size=8, bold=True, color="FFFFFF")

    # Row 1: Main Title Banner (No blank Row 2)
    title_text = f"2026-2027 G{grade_num} 第一学期期中考试成绩 G{grade_num} First Semester MT1 Grades"
    _style_banner(ws, row=1, start_col=1, end_col=10 if mode == "IB" else 9, text=title_text, font_size=13)
    cur_row = 2

    if mode == "IB":
        ib_subjs = [s for s in subjects if s.category == "IB_DP"]
        _style_subheaders(
            ws,
            cur_row,
            ["Subject 科目", "Number of Students 学生人数", "7", "6", "5", "4", "3", "2", "1", "Average 平均分"],
        )
        cur_row += 1
        for i, s in enumerate(ib_subjs):
            vals = [float(st.grades[s.key]) for st in students if s.key in st.grades and isinstance(st.grades[s.key], (int, float))]
            row_fill = white_fill if i % 2 == 0 else alt_fill
            ws.row_dimensions[cur_row].height = 20
            counts = [sum(1 for v in vals if int(round(v)) == g) for g in range(7, 0, -1)]
            avg = round(sum(vals) / len(vals), 2) if vals else ""
            for c_i, v in enumerate([s.display_name, len(vals)] + counts + [avg], start=1):
                cell = ws.cell(row=cur_row, column=c_i, value=v)
                cell.fill = row_fill
                cell.font = normal_font
                cell.border = BORDER_ALL
                cell.alignment = Alignment(horizontal="left" if c_i == 1 else "center", vertical="center")
                if c_i == 10 and isinstance(v, float):
                    cell.number_format = "0.00"
            cur_row += 1
        cur_row += 1
    else:
        al_subjs = [s for s in subjects if s.category in {"AL", "HS_AL_ENG"}]
        _style_subheaders(
            ws,
            cur_row,
            ["Subject 科目", "Number of Students 学生人数", "A*", "A", "B", "C", "D", "E", "U"],
        )
        cur_row += 1
        grades_al = ["A*", "A", "B", "C", "D", "E", "U"]
        for i, s in enumerate(al_subjs):
            vals = [str(st.grades[s.key]).upper() for st in students if s.key in st.grades and st.grades[s.key] is not None]
            row_fill = white_fill if i % 2 == 0 else alt_fill
            ws.row_dimensions[cur_row].height = 20
            counts = [sum(1 for v in vals if v == g) for g in grades_al]
            for c_i, v in enumerate([s.display_name, len(vals)] + counts, start=1):
                cell = ws.cell(row=cur_row, column=c_i, value=v)
                cell.fill = row_fill
                cell.font = normal_font
                cell.border = BORDER_ALL
                cell.alignment = Alignment(horizontal="left" if c_i == 1 else "center", vertical="center")
            cur_row += 1
        cur_row += 1

    # Chinese National Course 中方课程 (A-F)
    hs_cn_subjs = [s for s in subjects if s.category == "HS_CN"]
    if hs_cn_subjs:
        _style_banner(ws, row=cur_row, start_col=1, end_col=6, text="Chinese National Course 中方课程")
        cur_row += 1
        _style_subheaders(ws, cur_row, ["Subject 科目", "Number of Students 学生人数", ">=85", "75-84", "65-74", "<65"])
        cur_row += 1
        for i, s in enumerate(hs_cn_subjs):
            vals = [float(st.grades[s.key]) for st in students if s.key in st.grades and isinstance(st.grades[s.key], (int, float))]
            row_fill = white_fill if i % 2 == 0 else alt_fill
            ws.row_dimensions[cur_row].height = 20
            row_data = [
                s.display_name,
                len(vals),
                sum(1 for v in vals if v >= 85),
                sum(1 for v in vals if 75 <= v < 85),
                sum(1 for v in vals if 65 <= v < 75),
                sum(1 for v in vals if v < 65),
            ]
            for c_i, v in enumerate(row_data, start=1):
                cell = ws.cell(row=cur_row, column=c_i, value=v)
                cell.fill = row_fill
                cell.font = normal_font
                cell.border = BORDER_ALL
                cell.alignment = Alignment(horizontal="left" if c_i == 1 else "center", vertical="center")
            cur_row += 1
        cur_row += 1

    # IBDP Core 核心课程 (A-G, 2-row merged subheader, Concern 待改进 font size 8)
    if mode == "IB":
        _style_banner(ws, row=cur_row, start_col=1, end_col=7, text="IBDP Core 核心课程")
        cur_row += 1

        r_top, r_bot = cur_row, cur_row + 1
        ws.row_dimensions[r_top].height = 22
        ws.row_dimensions[r_bot].height = 22
        ws.merge_cells(start_row=r_top, start_column=1, end_row=r_bot, end_column=1)
        ws.merge_cells(start_row=r_top, start_column=2, end_row=r_bot, end_column=2)

        top_labels = ["Subject 科目", "Number of Students 学生人数", "A", "B", "C", "D", "E"]
        bot_labels = ["", "", "Excellent 优秀", "On-track 正常", "Concern 待改进", "", ""]

        for c_i in range(1, 8):
            ct = ws.cell(row=r_top, column=c_i, value=top_labels[c_i - 1])
            ct.fill = blue_fill
            ct.font = blue_font_10
            ct.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ct.border = BORDER_ALL

            cb = ws.cell(row=r_bot, column=c_i)
            if c_i >= 3:
                cb.value = bot_labels[c_i - 1]
            cb.fill = blue_fill
            cb.font = blue_font_8 if c_i == 5 else blue_font_10
            cb.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cb.border = BORDER_ALL

        cur_row = r_bot + 1
        core_rows_spec = [
            ("TOK", "Theory of Knowledge (TOK) [A-E]", True),
            ("EE", "Extended Essay (EE) [Excellent/ on-track / concern]", False),
            ("CAS", "CAS [Excellent/ on-track / concern]", False),
        ]
        for idx, (key, label, is_tok) in enumerate(core_rows_spec):
            row_fill = white_fill if idx % 2 == 0 else alt_fill
            ws.row_dimensions[cur_row].height = 20
            vals = [str(st.grades[key]).strip() for st in students if key in st.grades and st.grades[key] is not None]
            if is_tok:
                valid_vals = [v.upper() for v in vals if v.upper() in {"A", "B", "C", "D", "E"}]
                row_data = [
                    label, len(valid_vals),
                    sum(1 for v in valid_vals if v == "A"),
                    sum(1 for v in valid_vals if v == "B"),
                    sum(1 for v in valid_vals if v == "C"),
                    sum(1 for v in valid_vals if v == "D"),
                    sum(1 for v in valid_vals if v == "E"),
                ]
            else:
                c_exc = sum(1 for v in vals if v.lower() == "excellent")
                c_ont = sum(1 for v in vals if v.lower() == "on-track")
                c_con = sum(1 for v in vals if v.lower() == "concern")
                row_data = [label, c_exc + c_ont + c_con, c_exc, c_ont, c_con, "", ""]

            for c_i, v in enumerate(row_data, start=1):
                cell = ws.cell(row=cur_row, column=c_i, value=v)
                cell.fill = row_fill
                cell.font = normal_font
                cell.border = BORDER_ALL
                cell.alignment = Alignment(horizontal="left" if c_i == 1 else "center", vertical="center")
            cur_row += 1
        cur_row += 1

    # Overall Summary 总评 (Total Score)
    _style_banner(ws, row=cur_row, start_col=1, end_col=6, text="Overall Summary 总评 (Total Score)")
    cur_row += 1
    _style_subheaders(ws, cur_row, ["Metric 指标", "Value 数值"], start_col=1)
    cur_row += 1

    if mode == "IB":
        dp_keys = {s.key for s in subjects if s.category == "IB_DP"}
        totals = [
            sum(float(st.grades[k]) for k in dp_keys if k in st.grades and isinstance(st.grades[k], (int, float)))
            for st in students
            if any(k in st.grades and isinstance(st.grades[k], (int, float)) for k in dp_keys)
        ]
        metrics = [
            ("Total Students 学生总人数", len(students)),
            ("Average Total Score 平均总分", round(sum(totals) / len(totals), 2) if totals else 0),
            ("Median Total Score 中位数", round(statistics.median(totals), 2) if totals else 0),
            ("Highest Total Score 最高分", max(totals) if totals else 0),
            ("Lowest Total Score 最低分", min(totals) if totals else 0),
        ]
    else:
        al_keys = {s.key for s in subjects if s.category == "AL"}
        total_entries = 0
        pass_entries = 0
        for st in students:
            for k in al_keys:
                if k in st.grades and st.grades[k] is not None:
                    total_entries += 1
                    if not st.failed_flags.get(k, False):
                        pass_entries += 1
        pass_rate = f"{(pass_entries / total_entries * 100):.1f}%" if total_entries else "0.0%"
        metrics = [
            ("Total Students 学生总人数", len(students)),
            ("Total AL Subject Entries 总科目人次", total_entries),
            ("C and Above Pass Rate (A*-C 合格率)", pass_rate),
        ]

    for i, (m_label, m_val) in enumerate(metrics):
        row_fill = white_fill if i % 2 == 0 else alt_fill
        ws.row_dimensions[cur_row].height = 20
        c1 = ws.cell(row=cur_row, column=1, value=m_label)
        c2 = ws.cell(row=cur_row, column=2, value=m_val)
        for cx in (c1, c2):
            cx.fill = row_fill
            cx.font = bold_font if i == 0 else normal_font
            cx.border = BORDER_ALL
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c2.alignment = Alignment(horizontal="center", vertical="center")
        cur_row += 1

    col_widths = {1: 44, 2: 20, 3: 15, 4: 15, 5: 15, 6: 13, 7: 13, 8: 11, 9: 11, 10: 15}
    for ci, w in col_widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


# =============================================================================
# Sheet 3: Warning
# =============================================================================
def _build_warning_sheet(
    ws,
    warning_students: List[Tuple[StudentRecord, int, List[Tuple[str, Any]]]],
) -> None:
    header_fill = PatternFill("solid", fgColor=BLUE_PANTONE_281)
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    red_fill = PatternFill("solid", fgColor=FAIL_RED_FILL)
    red_font = Font(name="Calibri", size=10, bold=True, color=FAIL_RED_FONT)
    normal_font = Font(name="Calibri", size=10)

    max_fails = max((len(fl) for _, _, fl in warning_students), default=1)
    max_fails = max(max_fails, 1)

    headers = [
        "Student ID",
        "Student Name 学生姓名",
        "Homeroom 行政班",
        "Student Email 学生邮箱",
        "Parent 1 Email 家长1邮箱",
        "Parent 2 Email 家长2邮箱",
        "Number of Failed Subjects 不合格科目数量",
        "Attendance Skip 缺勤次数",
    ] + [f"Failed Subject {i}" for i in range(1, max_fails + 1)]

    ws.row_dimensions[1].height = 30
    for c_i, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=c_i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER_ALL

    for r_i, (st, fail_count, failed_list) in enumerate(warning_students, start=2):
        ws.row_dimensions[r_i].height = 20
        left_vals = [
            st.student_id,
            st.name,
            st.homeroom,
            st.student_email,
            st.parent1_email,
            st.parent2_email,
            fail_count,
            st.attendance_skip,
        ]
        for c_i, val in enumerate(left_vals, start=1):
            c = ws.cell(row=r_i, column=c_i, value=val)
            c.border = BORDER_ALL
            c.alignment = Alignment(horizontal="center" if c_i in {1, 3, 7, 8} else "left", vertical="center")
            if c_i == 7 and fail_count >= 3:
                c.fill = red_fill
                c.font = red_font
            elif c_i == 8 and st.attendance_skip is not None and st.attendance_skip >= 9:
                c.fill = red_fill
                c.font = red_font
            else:
                c.font = normal_font

        for f_idx, (subj_name, grade_val) in enumerate(failed_list):
            c = ws.cell(row=r_i, column=9 + f_idx, value=f"{subj_name}: {grade_val}")
            c.fill = red_fill
            c.font = red_font
            c.border = BORDER_ALL
            c.alignment = Alignment(horizontal="center", vertical="center")

        for f_idx in range(len(failed_list), max_fails):
            ws.cell(row=r_i, column=9 + f_idx, value="").border = BORDER_ALL

    widths = {1: 14, 2: 24, 3: 18, 4: 28, 5: 28, 6: 28, 7: 18, 8: 16}
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci in range(9, 9 + max_fails):
        ws.column_dimensions[get_column_letter(ci)].width = 24
    ws.freeze_panes = "C2"
