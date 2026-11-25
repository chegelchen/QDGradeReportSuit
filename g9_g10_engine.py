"""
G9 (House X) & G10 MT1 Grade Summary, Overview & Warning Processor.
2026-2027 Semester 1 Mid-term 1 (MT1)

Uses shared read-only loaders from common_io.py. Never modifies input files.
"""

from __future__ import annotations

import os
import re
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
    load_student_id_map,
    load_student_info_map,
    resolve_col,
)


# =============================================================================
# Visual Style Constants
# =============================================================================
COLOR_OVERVIEW_RED = QDHS_RED_XL        # #C8102E
COLOR_OVERVIEW_BLUE = QDHS_BLUE_XL      # #00205B
COLOR_ROW_WHITE = "FFFFFF"
COLOR_ROW_GREEN = "D2F2CC"
COLOR_ALERT_RED = "FF9999"
COLOR_TOTAL_BLUE = "DDEBF7"
COLOR_SUMMARY_GRAY = "D9D9D9"
COLOR_HOMEROOM_HDR = "FFC000"
COLOR_PREPARED_GREEN = "00B050"

THIN_SIDE = Side(style="thin", color="BFBFBF")
MED_DARK_SIDE = Side(style="thin", color="808080")
CELL_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_BORDER = Border(left=MED_DARK_SIDE, right=MED_DARK_SIDE, top=MED_DARK_SIDE, bottom=MED_DARK_SIDE)

FONT_FAMILY = "Calibri"


# =============================================================================
# Subject Rules & Ordering
# =============================================================================
PBL_KEYWORDS = (
    "project based learning",
    "hs pbl 1",
    "hs pbl 2",
    "hs pbl 3",
    "hs pbl 4",
    "hs pbl 5",
    "community project",
    "understanding our home",
    "intro to robotics",
    "making decisions using data",
    "creative project in arts",
)

G9_SUBJECT_ORDER = [
    "1-HS Chinese 中文",
    "2-HS English as a Second Language 英文二语",
    "3-HS Humanities - Economics  人文-经济学",
    "4-HS Physics 物理",
    "4-HS Chemistry 化学",
    "4-HS Biology 生物",
    "5-HS Mathematics数学",
    "6-HS Film 电影",
    "6-HS Theater 戏剧",
    "6-HS Visual Art 视觉艺术",
    "6-HS Music 音乐",
    "HS SPARK 升学指导",
    "HS Physical Education 体育",
]

G10_SUBJECT_ORDER = [
    "1-HS Chinese 中文",
    "2-English - First Language  英文一语",
    "2-English as a Second Language 英文二语",
    "2-English as a Second Language Honor 英文二语",
    "2-HS Academic English  学术英文",
    "3-HS Humanities - Economics  人文-经济学",
    "3-HS Humanities - Psychology 人文-心理学",
    "3-HS Humanities - social studies  人文-社会学",
    "4-HS Biology  生物",
    "4-HS Chemistry  化学",
    "4-HS Physics  物理",
    "4-Computer Science  电脑科学",
    "3/4- Environmental Management 环境系统",
    "5-Mathematics  Core 数学",
    "5-Mathematics Standard  数学",
    "5-Mathematics Honor  数学",
    "6-HS Film  电影",
    "6-HS Music  音乐",
    "6-HS Theater  戏剧",
    "6-HS Visual Art   视觉艺术",
    "HS Geography 地理",
    "HS History  历史",
    "HS Politics  政治",
    "HS Physical Education 体育",
]

NATIONAL_AND_PE_SUBJECTS = (
    "HS Geography 地理",
    "HS History  历史",
    "HS Politics  政治",
    "HS Physical Education 体育",
)


def is_excluded_g9_course(subject: str, class_name: str) -> bool:
    s_low = (subject or "").lower()
    c_low = (class_name or "").lower()
    if s_low in ("workshop", "electives", "project based learning"):
        return True
    return any(kw in s_low or kw in c_low for kw in PBL_KEYWORDS)


def parse_grade_value(val: Any) -> Any:
    """Parse raw grade cell into int/float if numeric, clean string if text, or None if N/A."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return int(val) if isinstance(val, float) and val.is_integer() else val
    s = str(val).strip()
    if not s or s.upper() in ("N/A", "NA", "NONE", "-"):
        return None
    if s.endswith("%") and not re.match(r"^\d+(\.\d+)?%$", s):
        s = s[:-1].strip()
    try:
        num = float(s)
        return int(num) if num.is_integer() else num
    except ValueError:
        return s


def map_g9_subject(subject: str, class_name: str) -> Optional[str]:
    if is_excluded_g9_course(subject, class_name):
        return None
    s = (subject or "").strip()
    c = (class_name or "").strip()
    s_low = s.lower()
    c_low = c.lower()

    if "chinese" in s_low or "语文" in c or "中文" in c:
        return "1-HS Chinese 中文"
    if "english" in s_low:
        return "2-HS English as a Second Language 英文二语"
    if "economics" in s_low or "humanities" in s_low:
        return "3-HS Humanities - Economics  人文-经济学"
    if "physics" in s_low:
        return "4-HS Physics 物理"
    if "chemistry" in s_low:
        return "4-HS Chemistry 化学"
    if "biology" in s_low:
        return "4-HS Biology 生物"
    if "math" in s_low:
        return "5-HS Mathematics数学"
    if "film" in s_low:
        return "6-HS Film 电影"
    if "theater" in s_low or "theatre" in s_low:
        return "6-HS Theater 戏剧"
    if "visual art" in s_low or "art" in s_low:
        return "6-HS Visual Art 视觉艺术"
    if "music" in s_low:
        return "6-HS Music 音乐"
    if "college counseling" in s_low or "spark" in c_low:
        return "HS SPARK 升学指导"
    if "physical education" in s_low:
        return "HS Physical Education 体育"
    return s


def map_g10_subject(subject: str, class_name: str) -> Optional[Tuple[str, bool]]:
    s = (subject or "").strip()
    c = (class_name or "").strip()
    s_low = s.lower()
    c_low = c.lower()

    if s_low in ("electives", "workshop", "college counseling") or "spark" in c_low:
        return None

    if s_low == "geography" or "地理" in c:
        return ("HS Geography 地理", True)
    if s_low == "history" and "social studies" not in c_low:
        return ("HS History  历史", True)
    if s_low == "politics" or "政治" in c:
        return ("HS Politics  政治", True)
    if s_low == "physical education":
        return ("HS Physical Education 体育", True)

    if "chinese" in s_low:
        return ("1-HS Chinese 中文", False)
    if "academic english" in c_low:
        return ("2-HS Academic English  学术英文", False)
    if "english" in s_low:
        if "first language" in c_low:
            return ("2-English - First Language  英文一语", False)
        if "honor" in c_low:
            return ("2-English as a Second Language Honor 英文二语", False)
        return ("2-English as a Second Language 英文二语", False)
    if "economics" in s_low or "economics" in c_low:
        return ("3-HS Humanities - Economics  人文-经济学", False)
    if "psychology" in s_low or "psychology" in c_low:
        return ("3-HS Humanities - Psychology 人文-心理学", False)
    if "social studies" in c_low or s_low == "history":
        return ("3-HS Humanities - social studies  人文-社会学", False)
    if "environmental" in s_low or "environmental" in c_low:
        return ("3/4- Environmental Management 环境系统", False)
    if "biology" in s_low:
        return ("4-HS Biology  生物", False)
    if "chemistry" in s_low:
        return ("4-HS Chemistry  化学", False)
    if "physics" in s_low:
        return ("4-HS Physics  物理", False)
    if "computer science" in s_low:
        return ("4-Computer Science  电脑科学", False)
    if "math" in s_low:
        if "honor" in c_low:
            return ("5-Mathematics Honor  数学", False)
        if "core" in c_low:
            return ("5-Mathematics  Core 数学", False)
        return ("5-Mathematics Standard  数学", False)
    if "film" in s_low:
        return ("6-HS Film  电影", False)
    if "music" in s_low:
        return ("6-HS Music  音乐", False)
    if "theat" in s_low:
        return ("6-HS Theater  戏剧", False)
    if "visual art" in s_low:
        return ("6-HS Visual Art   视觉艺术", False)

    return (s, False)


def is_subject_failed(subject: str, val: Any) -> bool:
    """Single source of truth for whether a G9/G10 subject grade is failing."""
    if val is None:
        return False
    if subject == "HS SPARK 升学指导":
        return isinstance(val, str) and "concern" in val.strip().lower()
    if subject in NATIONAL_AND_PE_SUBJECTS:
        return isinstance(val, (int, float)) and val < 65
    if isinstance(val, str) and "concern" in val.strip().lower():
        return True
    return isinstance(val, (int, float)) and val < 4


@dataclass
class StudentRecord:
    student_id: str
    name: str
    homeroom: str = ""
    mb_id: str = ""
    student_email: str = ""
    parent1_email: str = ""
    parent2_email: str = ""
    attendance_tier: str = ""
    attendance_skip: Optional[int] = None
    grades: Dict[str, Any] = field(default_factory=dict)


def _read_raw_grade_rows(
    filepath: str,
    grade_level_filter: Optional[str] = None,
) -> List[Dict[str, Any]]:
    if not filepath or not os.path.isfile(filepath):
        raise FileNotFoundError(f"Grade data file not found: {filepath}")

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        hdr_row, hmap = find_header_row(ws, ["Student ID", "Subject", "Class Name"], max_scan=10)

        col_sid = resolve_col(hmap, ["Student ID"])
        col_subj = resolve_col(hmap, ["Subject"])
        col_cname = resolve_col(hmap, ["Class Name"])
        col_pref = resolve_col(hmap, ["Preferred Name"], required=False)
        col_last = resolve_col(hmap, ["Last Name"], required=False)
        col_first = resolve_col(hmap, ["First Name"], required=False)
        col_gl = resolve_col(hmap, ["Grade Level"], required=False)
        col_mt = resolve_col(
            hmap,
            ["Mid-term Grade 期中成绩 (1-7)", "Mid-term Grade 期中成绩", "Mid-term Grade"],
        )
        col_prog = resolve_col(
            hmap,
            ["Mid-Term Progress Average 期中总评分 (%)", "Mid-Term Progress Average"],
            required=False,
        )

        def _get(row: Tuple[Any, ...], c: Optional[int]) -> Any:
            return row[c - 1] if c and (c - 1) < len(row) else None

        rows: List[Dict[str, Any]] = []
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(_get(row, col_sid))
            if not sid:
                continue

            if grade_level_filter and col_gl:
                gl_val = clean_str(_get(row, col_gl))
                if gl_val.lower() != grade_level_filter.lower():
                    continue

            pref = clean_str(_get(row, col_pref))
            if not pref and (col_first or col_last):
                fn = clean_str(_get(row, col_first))
                ln = clean_str(_get(row, col_last))
                pref = f"{fn} {ln}".strip()

            subj = clean_str(_get(row, col_subj))
            cname = clean_str(_get(row, col_cname))
            mt_val = parse_grade_value(_get(row, col_mt))
            if mt_val is None and "academic english" in cname.lower() and col_prog:
                mt_val = parse_grade_value(_get(row, col_prog))

            rows.append({
                "student_id": sid,
                "name": pref,
                "subject": subj,
                "class_name": cname,
                "grade": mt_val,
            })
        return rows
    finally:
        wb.close()


def _order_subjects(present: set, preferred_order: List[str]) -> List[str]:
    ordered = [s for s in preferred_order if s in present]
    extras = sorted(present - set(ordered))
    return ordered + extras


def _sort_students(students: List[StudentRecord]) -> List[StudentRecord]:
    return sorted(students, key=lambda s: (s.homeroom or "ZZZ", s.name or s.student_id, s.student_id))


def _style_cell(
    cell: openpyxl.cell.cell.Cell,
    font_size: int = 10,
    bold: bool = False,
    font_color: str = "000000",
    fill_color: Optional[str] = None,
    horiz: str = "center",
    vert: str = "center",
    wrap: bool = False,
    rotation: int = 0,
    border: Border = CELL_BORDER,
    num_format: Optional[str] = None,
) -> None:
    cell.font = Font(name=FONT_FAMILY, size=font_size, bold=bold, color=font_color)
    if fill_color:
        cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")
    cell.alignment = Alignment(horizontal=horiz, vertical=vert, wrap_text=wrap, text_rotation=rotation)
    if border:
        cell.border = border
    if num_format:
        cell.number_format = num_format


def generate_g9_g10_report(
    mode: str,
    raw_files: List[str],
    attendance_path: str,
    id_map_path: str,
    output_path: str,
    student_info_path: Optional[str] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    Generate the 3-sheet MT1 report ('MT1 Summary', 'MT1 Overview', 'Warning')
    for G9 (House X) or G10. Order-independent for raw_files.
    """
    mode_norm = mode.strip().upper()
    if mode_norm not in ("G9", "G10"):
        raise ValueError(f"Unsupported mode for G9/G10 engine: {mode}. Must be 'G9' or 'G10'.")

    valid_raw_files = [f for f in raw_files if f and os.path.isfile(f)]
    if not valid_raw_files:
        raise ValueError("Please select at least one valid raw grade Excel file.")

    if progress_cb:
        progress_cb("Loading Attendance, ID mapping & Student Info... / 正在加载考勤、学号对照表与学生信息...")

    att_map = load_attendance_map(attendance_path, required=False)
    id_map = load_student_id_map(id_map_path, required=False)
    info_map = load_student_info_map(student_info_path, required=False)

    students_dict: Dict[str, StudentRecord] = {}
    present_subjects: set = set()

    def _ensure_student(sid: str, name: str) -> StudentRecord:
        if sid not in students_dict:
            hr, mb_id, cname = id_map.get(sid, ("", "", ""))
            tier, skip = att_map.get(sid, ("", None))
            s_email, p1_email, p2_email, pref = info_map.get(sid, ("", "", "", ""))
            students_dict[sid] = StudentRecord(
                student_id=sid,
                name=name or pref or cname,
                homeroom=hr,
                mb_id=mb_id,
                student_email=s_email,
                parent1_email=p1_email,
                parent2_email=p2_email,
                attendance_tier=tier,
                attendance_skip=skip,
            )
        elif not students_dict[sid].name and name:
            students_dict[sid].name = name
        return students_dict[sid]

    if mode_norm == "G9":
        if progress_cb:
            progress_cb("Parsing G9 (House X) raw grades... / 正在解析 G9 (House X) 原始成绩...")
        for rf in valid_raw_files:
            for row in _read_raw_grade_rows(rf):
                mapped_subj = map_g9_subject(row["subject"], row["class_name"])
                if mapped_subj is None:
                    continue
                stu = _ensure_student(row["student_id"], row["name"])
                if row["grade"] is not None:
                    stu.grades[mapped_subj] = row["grade"]
                present_subjects.add(mapped_subj)

        ordered_subjects = _order_subjects(present_subjects, G9_SUBJECT_ORDER)
        ib_1_7_subjects = [
            s for s in ordered_subjects
            if s not in ("HS SPARK 升学指导", "HS Physical Education 体育")
        ]
        total_subjects = list(ib_1_7_subjects)
    else:
        if progress_cb:
            progress_cb("Parsing G10 HS & IB raw grades... / 正在解析 G10 HS 与 IB 原始成绩...")
        for rf in valid_raw_files:
            for row in _read_raw_grade_rows(rf, grade_level_filter="Grade 10"):
                mapped = map_g10_subject(row["subject"], row["class_name"])
                if mapped is None:
                    continue
                mapped_subj, _is_nat = mapped
                stu = _ensure_student(row["student_id"], row["name"])
                if row["grade"] is not None:
                    stu.grades[mapped_subj] = row["grade"]
                present_subjects.add(mapped_subj)

        ordered_subjects = _order_subjects(present_subjects, G10_SUBJECT_ORDER)
        ib_1_7_subjects = [
            s for s in ordered_subjects
            if s not in (
                "HS Geography 地理",
                "HS History  历史",
                "HS Politics  政治",
                "HS Physical Education 体育",
                "HS SPARK 升学指导",
            )
        ]
        # G10 MT1 Total sums Groups 1-5 (6 courses); excludes 6-HS Arts, Chinese National & PE
        total_subjects = [s for s in ib_1_7_subjects if not s.startswith("6-")]

    students = _sort_students(list(students_dict.values()))
    if not students:
        raise ValueError("No student records matched after parsing input files.")

    if progress_cb:
        progress_cb("Building MT1 Summary, MT1 Overview & Warning sheets... / 正在生成汇总表、总览表与预警表...")

    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "MT1 Summary"
    ws_ov = wb.create_sheet("MT1 Overview")
    ws_warn = wb.create_sheet("Warning")

    meta = _build_summary_sheet(ws_sum, mode_norm, students, ordered_subjects, total_subjects, ib_1_7_subjects)
    _build_overview_sheet(ws_ov, mode_norm, ordered_subjects, meta)
    warning_count = _build_warning_sheet(ws_warn, students, ordered_subjects)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()

    if progress_cb:
        progress_cb(
            f"Done! Students: {len(students)} | Subjects: {len(ordered_subjects)} | Warnings: {warning_count} "
            f"/ 报表生成完成！"
        )

    return {
        "output_path": output_path,
        "mode": mode_norm,
        "student_count": len(students),
        "subject_count": len(ordered_subjects),
        "warning_count": warning_count,
        "homerooms": meta["homerooms"],
    }


def _build_summary_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    mode: str,
    students: List[StudentRecord],
    ordered_subjects: List[str],
    total_subjects: List[str],
    ib_1_7_subjects: List[str],
) -> Dict[str, Any]:
    ws.views.sheetView[0].showGridLines = True

    title_a1 = (
        "House X Class 2030\n\n2026-2027 MT1 Grades\nHouse X 第一学期期中考成绩"
        if mode == "G9"
        else "G10 Class 2029\n\n2026-2027 MT1 Grades\nG10 第一学期期中考试成绩"
    )
    headers: List[Tuple[str, Optional[str], str]] = [
        (title_a1, None, "000000"),
        ("Homeroom \n行政班级", None, "000000"),
        ("Student Name\n学生姓名", None, "000000"),
        ("MB系统ID", None, "000000"),
        ("Attendance (1-4: 1: excellent;  2: good;  3: satisfactory;  4: unsatisfactory)", None, "000000"),
        ("attendance skip", None, "000000"),
    ]

    if mode == "G9":
        total42_col_idx = len(headers) + 1
        total56_col_idx = len(headers) + 2
        headers.append(("MT1 Total 总计/42", COLOR_TOTAL_BLUE, "000000"))
        headers.append(("MT1 Total 总计/56", COLOR_TOTAL_BLUE, "000000"))
        primary_total_col = total42_col_idx
    else:
        total_col_idx = len(headers) + 1
        headers.append(("MT1 Total 总计", COLOR_TOTAL_BLUE, "000000"))
        primary_total_col = total_col_idx

    subj_col_map: Dict[str, int] = {}
    for s in ordered_subjects:
        idx = len(headers) + 1
        subj_col_map[s] = idx
        is_nat_or_pe = s in NATIONAL_AND_PE_SUBJECTS
        headers.append((s, "FFF2CC" if is_nat_or_pe else None, "000000"))

    failed_col_idx = len(headers) + 1
    headers.append(("Number of Failed Subjects\n不合格科目数量", "FF0000", "FFFFFF"))

    ws.row_dimensions[1].height = 145
    for c_idx, (h_text, h_fill, h_fc) in enumerate(headers, start=1):
        cell = ws.cell(1, c_idx, h_text)
        _style_cell(
            cell,
            font_size=10,
            bold=True,
            font_color=h_fc,
            fill_color=h_fill or "F2F2F2",
            horiz="center",
            vert="center",
            wrap=True,
            rotation=90 if c_idx >= 4 else 0,
            border=HEADER_BORDER,
        )

    total_cols = sorted(subj_col_map[s] for s in total_subjects if s in subj_col_map)
    total_is_contiguous = bool(total_cols) and (total_cols[-1] - total_cols[0] + 1 == len(total_cols))

    ib_cols = sorted(subj_col_map[s] for s in ib_1_7_subjects if s in subj_col_map)
    ib_is_contiguous = bool(ib_cols) and (ib_cols[-1] - ib_cols[0] + 1 == len(ib_cols))

    spark_col = subj_col_map.get("HS SPARK 升学指导")
    last_stu_row = 1 + len(students)

    for i, stu in enumerate(students):
        r = 2 + i
        row_bg = COLOR_ROW_WHITE if (r % 2 == 0) else COLOR_ROW_GREEN

        _style_cell(ws.cell(r, 1, stu.student_id), fill_color=row_bg, horiz="center")
        _style_cell(ws.cell(r, 2, stu.homeroom), fill_color=row_bg, horiz="center")
        _style_cell(ws.cell(r, 3, stu.name), fill_color=row_bg, horiz="left")
        _style_cell(ws.cell(r, 4, stu.mb_id), fill_color=row_bg, horiz="center", num_format="@")
        _style_cell(ws.cell(r, 5, stu.attendance_tier), fill_color=row_bg, horiz="center")

        skip_bg = COLOR_ALERT_RED if (stu.attendance_skip is not None and stu.attendance_skip >= 9) else row_bg
        _style_cell(ws.cell(r, 6, stu.attendance_skip), fill_color=skip_bg, horiz="center")

        if not total_cols:
            sum_expr = 0
        elif total_is_contiguous:
            sum_expr = f"=SUM({get_column_letter(total_cols[0])}{r}:{get_column_letter(total_cols[-1])}{r})"
        else:
            sum_expr = "=SUM(" + ",".join(f"{get_column_letter(c)}{r}" for c in total_cols) + ")"

        if mode == "G9":
            t56_ltr = get_column_letter(total56_col_idx)
            _style_cell(ws.cell(r, total42_col_idx, f"=(42/56)*{t56_ltr}{r}"), bold=True, fill_color=COLOR_TOTAL_BLUE, num_format="0.0")
            _style_cell(ws.cell(r, total56_col_idx, sum_expr), bold=True, fill_color=COLOR_TOTAL_BLUE, num_format="0")
        else:
            _style_cell(ws.cell(r, primary_total_col, sum_expr), bold=True, fill_color=COLOR_TOTAL_BLUE, num_format="0")

        for s in ordered_subjects:
            c_idx = subj_col_map[s]
            val = stu.grades.get(s)
            cell_bg = COLOR_ALERT_RED if is_subject_failed(s, val) else row_bg
            _style_cell(ws.cell(r, c_idx, val), fill_color=cell_bg, horiz="center")

        fail_parts: List[str] = []
        if ib_cols:
            if ib_is_contiguous:
                fail_parts.append(f'COUNTIF({get_column_letter(ib_cols[0])}{r}:{get_column_letter(ib_cols[-1])}{r},"<4")')
            else:
                fail_parts.extend(f'COUNTIF({get_column_letter(c)}{r},"<4")' for c in ib_cols)
        for nat_pe in NATIONAL_AND_PE_SUBJECTS:
            nat_col = subj_col_map.get(nat_pe)
            if nat_col:
                fail_parts.append(f'COUNTIF({get_column_letter(nat_col)}{r},"<65")')
        if spark_col:
            fail_parts.append(f'COUNTIF({get_column_letter(spark_col)}{r},"Concern")')
        fail_formula = ("=" + "+".join(fail_parts)) if fail_parts else 0
        _style_cell(ws.cell(r, failed_col_idx, fail_formula), bold=True, fill_color="FFFF00")

    # Bottom Summary Section
    prep_row = last_stu_row + 1
    ws.cell(prep_row, 1, "=NOW()").number_format = "yyyy-mm-dd"
    ws.cell(prep_row, 3, "Prepared On")
    for c in range(1, failed_col_idx + 1):
        _style_cell(ws.cell(prep_row, c), bold=True, fill_color=COLOR_PREPARED_GREEN, font_color="FFFFFF")

    homerooms = sorted({s.homeroom for s in students if s.homeroom})
    total_ltr = get_column_letter(primary_total_col)
    cnt_row = prep_row + 1
    avg_row = cnt_row + 1

    _style_cell(ws.cell(cnt_row, 3, "Number of Students \n学生人数"), bold=True, fill_color=COLOR_SUMMARY_GRAY, horiz="left", wrap=True)
    for c in range(1, failed_col_idx + 1):
        if c == 3:
            continue
        ltr = get_column_letter(c)
        if c == primary_total_col or (mode == "G9" and c == primary_total_col + 1):
            f = f'=COUNTIFS({total_ltr}2:{total_ltr}{last_stu_row},"<>0",{total_ltr}2:{total_ltr}{last_stu_row},"<>")'
        elif c == failed_col_idx:
            f = f'=COUNTIF({ltr}2:{ltr}{last_stu_row},">0")'
        elif c in subj_col_map.values():
            if c == spark_col or any("Academic English" in s and subj_col_map[s] == c for s in ordered_subjects):
                f = f"=COUNTA({ltr}2:{ltr}{last_stu_row})"
            else:
                f = f"=COUNT({ltr}2:{ltr}{last_stu_row})"
        else:
            f = None
        is_tot_col = (c == primary_total_col) or (mode == "G9" and c == primary_total_col + 1)
        _style_cell(ws.cell(cnt_row, c, f), bold=True, fill_color=COLOR_TOTAL_BLUE if is_tot_col else None)

    _style_cell(ws.cell(avg_row, 2, "Average"), bold=True)
    _style_cell(ws.cell(avg_row, 3, "Average\n各科平均分"), bold=True, fill_color=COLOR_SUMMARY_GRAY, horiz="left", wrap=True)
    for c in range(4, failed_col_idx + 1):
        ltr = get_column_letter(c)
        if c == primary_total_col or (mode == "G9" and c == primary_total_col + 1):
            f = f"=IFERROR(SUM({ltr}2:{ltr}{last_stu_row})/{ltr}{cnt_row},0)"
            _style_cell(ws.cell(avg_row, c, f), bold=True, fill_color=COLOR_TOTAL_BLUE, num_format="0.00")
        elif c in subj_col_map.values() and c != spark_col:
            f = f'=IFERROR(AVERAGE({ltr}2:{ltr}{last_stu_row}),"")'
            _style_cell(ws.cell(avg_row, c, f), num_format="0.00")
        else:
            _style_cell(ws.cell(avg_row, c, None))

    hr_start_row = avg_row + 1
    for idx_hr, hr in enumerate(homerooms):
        r = hr_start_row + idx_hr
        _style_cell(ws.cell(r, 3, f"{hr} Average"), bold=True, horiz="left")
        for c in range(4, failed_col_idx):
            ltr = get_column_letter(c)
            if c == primary_total_col or (mode == "G9" and c == primary_total_col + 1) or (c in subj_col_map.values() and c != spark_col):
                f = f'=IFERROR(AVERAGEIFS({ltr}2:{ltr}{last_stu_row},B2:B{last_stu_row},"{hr}"),"")'
                _style_cell(ws.cell(r, c, f), fill_color=COLOR_TOTAL_BLUE if c == primary_total_col else None, num_format="0.00")
            else:
                _style_cell(ws.cell(r, c, None))

    dist_start_row = hr_start_row + len(homerooms)
    grade_labels = [
        (7, "#7 (90-100) A*", "Median", f"=MEDIAN({total_ltr}2:{total_ltr}{last_stu_row})"),
        (6, "#6 (80-89) A", "Max", f"=MAX({total_ltr}2:{total_ltr}{last_stu_row})"),
        (5, "#5 (70-79) B", "Min", f"=MIN({total_ltr}2:{total_ltr}{last_stu_row})"),
        (4, "#4 (65-69) C", None, None),
        (3, "#3 (50-64) D", None, None),
        (2, "#2 (40-49) E", None, None),
        (1, "#1 (0-39) F, G, U", None, None),
    ]
    for offset, (g_num, g_label, stat_lbl, stat_formula) in enumerate(grade_labels):
        r = dist_start_row + offset
        if stat_lbl:
            _style_cell(ws.cell(r, 2, stat_lbl), bold=True)
        _style_cell(ws.cell(r, 3, g_label), bold=True, horiz="left")
        if stat_formula:
            _style_cell(ws.cell(r, primary_total_col, stat_formula), bold=True, fill_color=COLOR_TOTAL_BLUE, num_format="0.0")
        else:
            _style_cell(ws.cell(r, primary_total_col, None), fill_color=COLOR_TOTAL_BLUE)

        for s in ordered_subjects:
            c = subj_col_map[s]
            ltr = get_column_letter(c)
            if s == "HS SPARK 升学指导" or "Academic English" in s:
                _style_cell(ws.cell(r, c, None))
            else:
                _style_cell(ws.cell(r, c, f'=COUNTIF({ltr}2:{ltr}{last_stu_row},"={g_num}")'))

    # Rightmost 2 Columns: Per-Homeroom Grade Distribution
    dist_col1 = failed_col_idx + 1
    dist_col2 = failed_col_idx + 2
    cur_r = 2
    for hr in homerooms:
        _style_cell(ws.cell(cur_r, dist_col1, hr), bold=True, fill_color=COLOR_HOMEROOM_HDR)
        _style_cell(ws.cell(cur_r, dist_col2, f'=COUNTIF(B2:B{last_stu_row},"{hr}")'), bold=True, fill_color=COLOR_HOMEROOM_HDR)
        _style_cell(ws.cell(cur_r + 1, dist_col1, "Grade Range"), bold=True, fill_color=COLOR_SUMMARY_GRAY)
        _style_cell(ws.cell(cur_r + 1, dist_col2, "No.of Students"), bold=True, fill_color=COLOR_SUMMARY_GRAY)

        ranges = [
            ("40-42", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},">39")'),
            ("38-39", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},">37",{total_ltr}2:{total_ltr}{last_stu_row},"<40")'),
            ("36-37", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},">35",{total_ltr}2:{total_ltr}{last_stu_row},"<38")'),
            ("30-35", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},">29",{total_ltr}2:{total_ltr}{last_stu_row},"<36")'),
            ("25-29", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},">24",{total_ltr}2:{total_ltr}{last_stu_row},"<30")'),
            ("0-24", f'=COUNTIFS(B2:B{last_stu_row},"{hr}",{total_ltr}2:{total_ltr}{last_stu_row},"<25")'),
            ("Average", f'=IFERROR(AVERAGEIFS({total_ltr}2:{total_ltr}{last_stu_row},B2:B{last_stu_row},"{hr}"),"")'),
        ]
        for idx_rng, (lbl, formula) in enumerate(ranges):
            rr = cur_r + 2 + idx_rng
            _style_cell(ws.cell(rr, dist_col1, lbl), bold=(lbl == "Average"))
            _style_cell(ws.cell(rr, dist_col2, formula), bold=(lbl == "Average"), num_format="0.00" if lbl == "Average" else "0")
        cur_r += 11

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 24
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 16
    ws.column_dimensions["F"].width = 12
    for c in range(7, failed_col_idx + 1):
        ws.column_dimensions[get_column_letter(c)].width = 8.5
    ws.column_dimensions[get_column_letter(dist_col1)].width = 14
    ws.column_dimensions[get_column_letter(dist_col2)].width = 14
    ws.freeze_panes = "D2"

    return {
        "last_stu_row": last_stu_row,
        "cnt_row": cnt_row,
        "avg_row": avg_row,
        "primary_total_col": primary_total_col,
        "subj_col_map": subj_col_map,
        "homerooms": homerooms,
    }


def _build_overview_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    mode: str,
    ordered_subjects: List[str],
    meta: Dict[str, Any],
) -> None:
    ws.views.sheetView[0].showGridLines = True

    last_stu_row = meta["last_stu_row"]
    avg_row = meta["avg_row"]
    total_ltr = get_column_letter(meta["primary_total_col"])
    subj_col_map = meta["subj_col_map"]
    sum_sheet = "'MT1 Summary'"

    title_text = (
        "2026-2027 House X 第一学期期中考成绩\nHouse X First Mid-Term Exam Summary"
        if mode == "G9"
        else "2026-2027 G10 第一学期期中成绩\nG10 First Mid-Term Exam Summary"
    )

    ws.merge_cells("A7:J7")
    ws.row_dimensions[7].height = 42
    for c in range(1, 11):
        _style_cell(
            ws.cell(7, c, title_text if c == 1 else None),
            font_size=12,
            bold=True,
            font_color="FFFFFF",
            fill_color=COLOR_OVERVIEW_RED,
            wrap=True,
            border=HEADER_BORDER,
        )

    ov_headers = [
        "Subject 科目",
        "Number of Students \n学生人数",
        "Average\n各科平均分",
        "#7 (90-100) A*",
        "#6 (80-89) A",
        "#5 (70-79) B",
        "#4 (65-69) C",
        "#3 (50-64) D",
        "#2 (40-49) E",
        "#1 (0-39) F, G, U",
    ]
    ws.row_dimensions[8].height = 34
    for c, h_text in enumerate(ov_headers, start=1):
        _style_cell(
            ws.cell(8, c, h_text),
            font_size=10,
            bold=True,
            font_color="FFFFFF",
            fill_color=COLOR_OVERVIEW_BLUE,
            wrap=True,
            border=HEADER_BORDER,
        )

    if mode == "G10":
        national_subjs = [s for s in ordered_subjects if s in NATIONAL_AND_PE_SUBJECTS]
        main_subjs = [s for s in ordered_subjects if s not in national_subjs]
    else:
        national_subjs = []
        main_subjs = ordered_subjects

    cur_r = 9
    for idx_s, s in enumerate(main_subjs):
        row_bg = COLOR_ROW_WHITE if (idx_s % 2 == 0) else COLOR_ROW_GREEN
        s_ltr = get_column_letter(subj_col_map[s])
        rng = f"{sum_sheet}!{s_ltr}2:{s_ltr}{last_stu_row}"

        _style_cell(ws.cell(cur_r, 1, s), fill_color=row_bg, horiz="left")
        if s == "HS SPARK 升学指导":
            _style_cell(ws.cell(cur_r, 2, f"=COUNTA({rng})"), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 3, None), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 4, f'=COUNTIF({rng},"Excellent")'), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 5, f'=COUNTIF({rng},"On-track")'), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 6, f'=COUNTIF({rng},"Concern")'), fill_color=row_bg)
            for c in range(7, 11):
                _style_cell(ws.cell(cur_r, c, None), fill_color=row_bg)
        elif "Academic English" in s:
            _style_cell(ws.cell(cur_r, 2, f"=COUNTA({rng})"), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 3, None), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 4, f'=COUNTIF({rng},"*Excellent*")'), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 5, f'=COUNTIF({rng},"*On track*")'), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 6, f'=COUNTIF({rng},"*Concern*")'), fill_color=row_bg)
            for c in range(7, 11):
                _style_cell(ws.cell(cur_r, c, None), fill_color=row_bg)
        elif s == "HS Physical Education 体育":
            _style_cell(ws.cell(cur_r, 2, f"=COUNT({rng})"), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 3, f'=IFERROR(AVERAGE({rng}),"")'), fill_color=row_bg, num_format="0.00")
            pe_ranges = [
                f'=COUNTIFS({rng},">=90")',
                f'=COUNTIFS({rng},">=80",{rng},"<90")',
                f'=COUNTIFS({rng},">=70",{rng},"<80")',
                f'=COUNTIFS({rng},">=65",{rng},"<70")',
                f'=COUNTIFS({rng},">=50",{rng},"<65")',
                f'=COUNTIFS({rng},">=40",{rng},"<50")',
                f'=COUNTIFS({rng},"<40")',
            ]
            for c_off, f_str in enumerate(pe_ranges):
                _style_cell(ws.cell(cur_r, 4 + c_off, f_str), fill_color=row_bg)
        else:
            _style_cell(ws.cell(cur_r, 2, f"=COUNT({rng})"), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 3, f'=IFERROR(AVERAGE({rng}),"")'), fill_color=row_bg, num_format="0.00")
            for c_off, g_val in enumerate([7, 6, 5, 4, 3, 2, 1]):
                _style_cell(ws.cell(cur_r, 4 + c_off, f'=COUNTIF({rng},"={g_val}")'), fill_color=row_bg)
        cur_r += 1

    if national_subjs:
        cur_r += 1
        ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=5)
        for c in range(1, 6):
            _style_cell(
                ws.cell(cur_r, c, "Chinese National Course 中方课程" if c == 1 else None),
                font_size=11,
                bold=True,
                font_color="FFFFFF",
                fill_color=COLOR_OVERVIEW_RED,
                border=HEADER_BORDER,
            )
        cur_r += 1
        nat_hdrs = ["Subject 科目", "Number of Students \n学生人数", "Average\n平均分", "Max\n最高分", "Min\n最低分"]
        for c, h_txt in enumerate(nat_hdrs, start=1):
            _style_cell(ws.cell(cur_r, c, h_txt), bold=True, font_color="FFFFFF", fill_color=COLOR_OVERVIEW_BLUE, wrap=True, border=HEADER_BORDER)
        cur_r += 1
        for idx_n, s in enumerate(national_subjs):
            row_bg = COLOR_ROW_WHITE if (idx_n % 2 == 0) else COLOR_ROW_GREEN
            s_ltr = get_column_letter(subj_col_map[s])
            rng = f"{sum_sheet}!{s_ltr}2:{s_ltr}{last_stu_row}"
            _style_cell(ws.cell(cur_r, 1, s), fill_color=row_bg, horiz="left")
            _style_cell(ws.cell(cur_r, 2, f"=COUNT({rng})"), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 3, f'=IFERROR(AVERAGE({rng}),"")'), fill_color=row_bg, num_format="0.00")
            _style_cell(ws.cell(cur_r, 4, f'=IFERROR(MAX({rng}),"")'), fill_color=row_bg)
            _style_cell(ws.cell(cur_r, 5, f'=IFERROR(MIN({rng}),"")'), fill_color=row_bg)
            cur_r += 1

    cur_r += 1
    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=4)
    for c in range(1, 5):
        _style_cell(
            ws.cell(cur_r, c, "Overall Summary 总评" if c == 1 else None),
            font_size=11,
            bold=True,
            font_color="FFFFFF",
            fill_color=COLOR_OVERVIEW_RED,
            border=HEADER_BORDER,
        )
    cur_r += 1
    sum_hdrs = ["平均分\nAverage", "中位数 \nMedian", "最高分 \nMax", "最低分 \nMin"]
    ws.row_dimensions[cur_r].height = 30
    for c, h_txt in enumerate(sum_hdrs, start=1):
        _style_cell(ws.cell(cur_r, c, h_txt), bold=True, font_color="FFFFFF", fill_color=COLOR_OVERVIEW_BLUE, wrap=True, border=HEADER_BORDER)

    cur_r += 1
    tot_rng = f"{sum_sheet}!{total_ltr}2:{total_ltr}{last_stu_row}"
    _style_cell(ws.cell(cur_r, 1, f"={sum_sheet}!{total_ltr}{avg_row}"), bold=True, fill_color=COLOR_ROW_WHITE, num_format="0.00")
    _style_cell(ws.cell(cur_r, 2, f"=MEDIAN({tot_rng})"), bold=True, fill_color=COLOR_ROW_WHITE, num_format="0.0")
    _style_cell(ws.cell(cur_r, 3, f"=MAX({tot_rng})"), bold=True, fill_color=COLOR_ROW_WHITE, num_format="0.0")
    _style_cell(ws.cell(cur_r, 4, f"=MIN({tot_rng})"), bold=True, fill_color=COLOR_ROW_WHITE, num_format="0.0")

    ws.column_dimensions["A"].width = 44
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 16
    for c in range(4, 11):
        ws.column_dimensions[get_column_letter(c)].width = 16


def _build_warning_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    students: List[StudentRecord],
    ordered_subjects: List[str],
) -> int:
    ws.views.sheetView[0].showGridLines = True

    warning_rows: List[Tuple[StudentRecord, int, Dict[str, Any]]] = []
    failed_subjects_any: set = set()
    for stu in students:
        failed_map = {s: stu.grades[s] for s in ordered_subjects if s in stu.grades and is_subject_failed(s, stu.grades[s])}
        fail_cnt = len(failed_map)
        skip_bad = stu.attendance_skip is not None and stu.attendance_skip >= 9
        if fail_cnt >= 3 or skip_bad:
            warning_rows.append((stu, fail_cnt, failed_map))
            failed_subjects_any.update(failed_map.keys())

    failed_subj_cols = [s for s in ordered_subjects if s in failed_subjects_any]

    all_headers = [
        "Student ID",
        "Student Name\n学生姓名",
        "Homeroom\n行政班级",
        "Student Email\n学生邮箱",
        "Parent 1 Email\n家长1邮箱",
        "Parent 2 Email\n家长2邮箱",
        "Number of Failed Subjects\n不合格科目数量",
        "attendance skip",
    ] + failed_subj_cols

    ws.row_dimensions[1].height = 48
    for c_idx, h_text in enumerate(all_headers, start=1):
        _style_cell(
            ws.cell(1, c_idx, h_text),
            font_size=10,
            bold=True,
            font_color="FFFFFF",
            fill_color=COLOR_OVERVIEW_BLUE,
            wrap=True,
            border=HEADER_BORDER,
        )

    if not warning_rows:
        _style_cell(ws.cell(2, 1, "No students meet warning criteria (Failed >= 3 or Skip >= 9)."), horiz="left")
    else:
        for idx_w, (stu, fail_cnt, failed_map) in enumerate(warning_rows):
            r = 2 + idx_w
            row_bg = COLOR_ROW_WHITE if (r % 2 == 0) else COLOR_ROW_GREEN
            _style_cell(ws.cell(r, 1, stu.student_id), fill_color=row_bg, horiz="center")
            _style_cell(ws.cell(r, 2, stu.name), fill_color=row_bg, horiz="left")
            _style_cell(ws.cell(r, 3, stu.homeroom), fill_color=row_bg, horiz="center")
            _style_cell(ws.cell(r, 4, stu.student_email), fill_color=row_bg, horiz="left")
            _style_cell(ws.cell(r, 5, stu.parent1_email), fill_color=row_bg, horiz="left")
            _style_cell(ws.cell(r, 6, stu.parent2_email), fill_color=row_bg, horiz="left")
            _style_cell(ws.cell(r, 7, fail_cnt), bold=True, fill_color="FFFF00" if fail_cnt >= 3 else row_bg, horiz="center")

            skip_bad = stu.attendance_skip is not None and stu.attendance_skip >= 9
            _style_cell(ws.cell(r, 8, stu.attendance_skip), fill_color=COLOR_ALERT_RED if skip_bad else row_bg, horiz="center")

            for s_off, subj in enumerate(failed_subj_cols):
                c_idx = 9 + s_off
                val = failed_map.get(subj)
                _style_cell(ws.cell(r, c_idx, val), fill_color=COLOR_ALERT_RED if val is not None else row_bg, horiz="center")

    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 18
    ws.column_dimensions["D"].width = 30
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["F"].width = 28
    ws.column_dimensions["G"].width = 18
    ws.column_dimensions["H"].width = 14
    for c_idx in range(9, len(all_headers) + 1):
        ws.column_dimensions[get_column_letter(c_idx)].width = 18

    ws.freeze_panes = "D2"
    return len(warning_rows)
