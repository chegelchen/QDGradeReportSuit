"""
Unified Academic Written Warning Letter Generator (PDF + JPG) & ManageBac Link Builder.

Supports BOTH Warning sheet layouts seamlessly:
  - G9/G10 layout: subject names as column headers to the right of 'attendance skip'
  - G11/G12 layout: 'Failed Subject 1..N' columns containing 'Subject: Grade'

Strictly uses `import pymupdf` (never deprecated `import fitz`).
"""

from __future__ import annotations

import html as html_lib
import os
import re
import sys
import webbrowser
from dataclasses import dataclass, field
from typing import Any, Callable, Dict, List, Optional, Tuple

import openpyxl
import pymupdf
from PIL import Image

from common_io import (
    MANAGEBAC_BEHAVIOR_NOTE_URL_TEMPLATE,
    clean_str,
    find_header_row,
    load_student_id_map,
    parse_number,
    resolve_col,
    sanitize_filename_part,
)


MAC_FONTS_DIR = "/System/Library/Fonts/Supplemental"
_RESOURCE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
DEFAULT_LOGO_PATH = os.path.join(_RESOURCE_DIR, "assets", "qdhs_logo.png")


@dataclass
class WarningStudent:
    student_id: str
    name: str
    homeroom: str = ""
    student_email: str = ""
    parent1_email: str = ""
    parent2_email: str = ""
    failed_count: int = 0
    attendance_skip: Optional[int] = None
    failed_subjects: List[Tuple[str, Any]] = field(default_factory=list)
    mb_id: str = ""
    managebac_url: str = ""

    @property
    def has_academic_warning(self) -> bool:
        return self.failed_count >= 3

    @property
    def has_attendance_warning(self) -> bool:
        return self.attendance_skip is not None and self.attendance_skip >= 9

    @property
    def reason_summary(self) -> str:
        parts: List[str] = []
        if self.has_academic_warning:
            parts.append(f"{self.failed_count} Fails / 不合格 {self.failed_count} 门")
        if self.has_attendance_warning:
            parts.append(f"{self.attendance_skip} Skips / 缺勤 {self.attendance_skip} 次")
        return " & ".join(parts) if parts else "-"

    @property
    def failed_subjects_text(self) -> str:
        if not self.failed_subjects:
            return "-"
        return ", ".join(f"{subj}: {grade}" if str(grade) != "" else subj for subj, grade in self.failed_subjects)


def build_managebac_behavior_note_url(mb_id: str) -> str:
    clean = clean_str(mb_id)
    if not clean:
        return ""
    return MANAGEBAC_BEHAVIOR_NOTE_URL_TEMPLATE.format(mb_id=clean)


def _split_subject_grade_text(raw_text: str) -> Tuple[str, str]:
    """Parse 'Subject: Grade' strings from G11/G12 Warning sheets cleanly."""
    if ":" in raw_text:
        left, right = raw_text.rsplit(":", 1)
        return left.strip(), right.strip()
    return raw_text.strip(), ""


def load_warning_students(
    excel_path: str,
    id_map_path: Optional[str] = None,
) -> List[WarningStudent]:
    """
    Dynamically parse the 'Warning' sheet from any QDHS MT1 report (G9/G10/G11/G12)
    and enrich each student with MB System ID and ManageBac Behavior Note URL.
    """
    if not excel_path or not os.path.isfile(excel_path):
        raise FileNotFoundError(f"MT1 report file not found: {excel_path}")

    id_map = load_student_id_map(id_map_path, required=False) if id_map_path else {}

    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        if "Warning" not in wb.sheetnames:
            raise ValueError(
                f"Sheet 'Warning' not found in '{os.path.basename(excel_path)}'. "
                f"Available sheets: {wb.sheetnames}"
            )
        ws = wb["Warning"]
        hdr_row, hmap = find_header_row(ws, ["Student ID", "skip"], max_scan=10)

        col_sid = resolve_col(hmap, ["Student ID"])
        col_name = resolve_col(hmap, ["Student Name", "学生姓名"], required=False)
        col_hr = resolve_col(hmap, ["Homeroom", "行政班级", "行政班"], required=False)
        col_semail = resolve_col(hmap, ["Student Email", "学生邮箱"], required=False)
        col_p1 = resolve_col(hmap, ["Parent 1 Email", "Parent 1", "家长1"], required=False)
        col_p2 = resolve_col(hmap, ["Parent 2 Email", "Parent 2", "家长2"], required=False)
        col_fcnt = resolve_col(hmap, ["Number of Failed Subjects", "Number of Failed", "不合格科目"], required=False)
        col_skip = resolve_col(hmap, ["attendance skip", "Attendance Skip", "缺勤"], required=False)

        # Read header row values verbatim to inspect all columns to the right of skip/failed_count
        header_row_vals = next(ws.iter_rows(min_row=hdr_row, max_row=hdr_row, values_only=True))
        detail_start_col = max(col_skip or 0, col_fcnt or 0) + 1
        detail_cols: List[Tuple[int, str, bool]] = []
        for c_idx in range(detail_start_col, len(header_row_vals) + 1):
            h_txt = clean_str(header_row_vals[c_idx - 1])
            if h_txt:
                is_generic_fail_col = bool(re.match(r"^failed\s+subject\s*\d+$", h_txt, re.IGNORECASE))
                detail_cols.append((c_idx, h_txt, is_generic_fail_col))

        def _get(row: Tuple[Any, ...], c: Optional[int]) -> Any:
            return row[c - 1] if c and (c - 1) < len(row) else None

        students: List[WarningStudent] = []
        for row in ws.iter_rows(min_row=hdr_row + 1, values_only=True):
            sid = clean_str(_get(row, col_sid))
            if not sid or sid.lower().startswith("no students meet"):
                continue

            name = clean_str(_get(row, col_name))
            hr = clean_str(_get(row, col_hr))
            s_email = clean_str(_get(row, col_semail))
            p1_email = clean_str(_get(row, col_p1))
            p2_email = clean_str(_get(row, col_p2))

            skip_num = parse_number(_get(row, col_skip))
            skip_val: Optional[int] = int(skip_num) if skip_num is not None else None

            failed_list: List[Tuple[str, Any]] = []
            for c_idx, h_txt, is_generic in detail_cols:
                raw_v = _get(row, c_idx)
                if raw_v is None or clean_str(raw_v) == "":
                    continue
                if is_generic:
                    subj_name, grade_str = _split_subject_grade_text(clean_str(raw_v))
                    failed_list.append((subj_name, grade_str))
                else:
                    val = int(raw_v) if isinstance(raw_v, float) and raw_v.is_integer() else raw_v
                    failed_list.append((h_txt, val))

            fcnt_num = parse_number(_get(row, col_fcnt))
            failed_count = int(fcnt_num) if fcnt_num is not None else len(failed_list)

            mapped_hr, mb_id, mapped_name = id_map.get(sid, ("", "", ""))
            if not hr and mapped_hr:
                hr = mapped_hr
            if not name and mapped_name:
                name = mapped_name
            mb_url = build_managebac_behavior_note_url(mb_id)

            students.append(
                WarningStudent(
                    student_id=sid,
                    name=name,
                    homeroom=hr,
                    student_email=s_email,
                    parent1_email=p1_email,
                    parent2_email=p2_email,
                    failed_count=failed_count,
                    attendance_skip=skip_val,
                    failed_subjects=failed_list,
                    mb_id=mb_id,
                    managebac_url=mb_url,
                )
            )
        return students
    finally:
        wb.close()


def _build_warning_letter_html(student: WarningStudent) -> Tuple[str, str]:
    """
    Build bilingual HTML + CSS matching 'Academic Written Warning Letter.pdf'
    for the first half, followed by escaped student-specific warning details.
    """
    esc = html_lib.escape
    stu_name = esc(student.name or "-")
    stu_id = esc(student.student_id or "-")
    stu_hr = esc(student.homeroom or "-")

    detail_blocks: List[str] = []

    if student.has_academic_warning and student.failed_subjects:
        items_html = "".join(
            f"<li><b>{esc(str(subj))}</b>: <span class='alert'>{esc(str(grade))}</span></li>"
            if str(grade) != ""
            else f"<li><span class='alert'>{esc(str(subj))}</span></li>"
            for subj, grade in student.failed_subjects
        )
        detail_blocks.append(
            f"""
            <div class="detail-group">
              <div class="detail-label">
                1. Failed Subjects (3 or more) / 不合格科目列表（共 {student.failed_count} 门）：
              </div>
              <ul class="detail-list">{items_html}</ul>
            </div>
            """
        )

    if student.has_attendance_warning:
        idx_num = "2." if detail_blocks else "1."
        detail_blocks.append(
            f"""
            <div class="detail-group">
              <div class="detail-label">
                {idx_num} Attendance Skip / 无故缺勤次数（≥ 9 次）：
                <span class="alert">{esc(str(student.attendance_skip))} 次 / {esc(str(student.attendance_skip))} unexcused skips</span>
              </div>
            </div>
            """
        )

    if not detail_blocks:
        detail_blocks.append(
            "<div class='detail-group'>No specific failing subjects or attendance skips recorded.</div>"
        )

    details_section = "\n".join(detail_blocks)

    css = """
    @font-face {
        font-family: "TimesCustom";
        src: url("Times New Roman.ttf");
        font-weight: normal;
    }
    @font-face {
        font-family: "TimesCustom";
        src: url("Times New Roman Bold.ttf");
        font-weight: bold;
    }
    @font-face {
        font-family: "SongCustom";
        src: url("Songti.ttc");
        font-weight: normal;
    }
    body {
        font-family: "TimesCustom", "SongCustom", serif;
        font-size: 9.4pt;
        line-height: 1.38;
        color: #111111;
        margin: 0;
        padding: 0;
    }
    h1 {
        text-align: center;
        font-size: 12.5pt;
        font-weight: bold;
        margin: 0 0 8pt 0;
    }
    .salutation {
        text-align: left;
        font-weight: bold;
        margin: 0 0 6pt 0;
    }
    p {
        margin: 0 0 5.5pt 0;
    }
    ul {
        margin: 2.5pt 0 5.5pt 0;
        padding-left: 18pt;
    }
    ul ul {
        margin: 1.5pt 0 1.5pt 0;
        padding-left: 16pt;
    }
    li {
        margin: 1.2pt 0;
    }
    .cn-block {
        margin-top: 6pt;
    }
    .details-box {
        margin-top: 10pt;
        border-top: 1.4pt solid #00205B;
        padding-top: 7pt;
    }
    .details-title {
        font-size: 10.5pt;
        font-weight: bold;
        color: #C8102E;
        margin-bottom: 4pt;
    }
    .student-meta {
        background-color: #F4F6F9;
        border-left: 3pt solid #C8102E;
        border-top: 0.6pt solid #D0D7DE;
        border-right: 0.6pt solid #D0D7DE;
        border-bottom: 0.6pt solid #D0D7DE;
        padding: 4.5pt 8pt;
        margin-bottom: 5pt;
        font-size: 8.8pt;
    }
    .detail-group {
        margin: 3.5pt 0;
    }
    .detail-label {
        font-weight: bold;
        margin-bottom: 2pt;
    }
    .detail-list {
        margin: 2pt 0 3.5pt 0;
        padding-left: 18pt;
    }
    .alert {
        color: #C8102E;
        font-weight: bold;
    }
    .signature-block {
        text-align: right;
        margin-top: 14pt;
        line-height: 1.42;
        font-size: 9.4pt;
    }
    """

    html_content = f"""
    <h1>Academic Written Warning Letter – 学术警告信</h1>

    <p class="salutation">Dear {stu_name},</p>

    <p>This warning letter is issued to you in accordance with QDHS Assessment and Attendance Policies,
    as you have met one or more of the criteria listed below:</p>

    <ul>
      <li>Failing grades in three or more subjects, including:
        <ul>
          <li>An overall mark below 4 or a C grade</li>
          <li>A score lower than 65 in Chinese National Curriculum courses (Politics, Geography, History and PE)</li>
          <li>A &quot;Concern&quot; rating in other programmes (e.g., Core, Spark)</li>
        </ul>
      </li>
      <li>Nine or more unexcused skips recorded in your attendance record</li>
    </ul>

    <p>It is crucial to understand that not fulfilling the necessary requirements for the subjects of concern
    might lead to you not achieving the credits required to obtain the Dwight Diploma. Your consistent effort
    and improvement in these areas are essential for your academic progression.<br/>
    This letter thus constitutes a written warning to you to improve in all these areas. Your parents will
    receive a copy of this letter.</p>

    <div class="cn-block">
      <p>根据 QDHS 评估政策和出勤政策，你收到了这封警告信，因为你符合了以下一项或多项条件：</p>
      <ul>
        <li>在三个或更多科目中不及格，包括：
          <ul>
            <li>平均成绩低于 4 分或 C 级</li>
            <li>中国国家课程（政治、地理、历史和体育）中取得低于 65 分的成绩</li>
            <li>其他课程（如 Core、Spark）中为 concern</li>
          </ul>
        </li>
        <li>出勤中有 9 次或以上翘课（skip）行为</li>
      </ul>
      <p>请重点留意：倘若无法满足相关学科的要求，你可能无法取得德怀特文凭对应的学分。你在这些方面坚持不懈的努力与进步，对学业发展至关重要。<br/>
      因此，本函为一份书面警示，告知你需要在以上领域做出改进，信件副本会抄送你的家长。</p>
    </div>

    <div class="details-box">
      <div class="details-title">Specific Warning Details / 具体警告内容</div>
      <div class="student-meta">
        <b>Student / 学生：</b>{stu_name} &nbsp;|&nbsp;
        <b>ID / 学号：</b>{stu_id} &nbsp;|&nbsp;
        <b>Homeroom / 行政班：</b>{stu_hr}
      </div>
      {details_section}
    </div>

    <div class="signature-block">
      上海七宝德怀特高级中学<br/>
      Shanghai Qibao Dwight High School<br/>
      2026年11月23日<br/>
      23 November 2026
    </div>
    """
    return html_content, css


def generate_warning_letter_pdf_and_jpg(
    student: WarningStudent,
    output_dir: str,
    dpi: int = 200,
) -> Tuple[str, str]:
    """
    Generate both PDF and JPG warning letters for a single student.
    Filename format: {Homeroom}_{Student ID}_{Student Name}.pdf and .jpg
    """
    os.makedirs(output_dir, exist_ok=True)
    base_name = (
        f"{sanitize_filename_part(student.homeroom or 'NoHomeroom')}_"
        f"{sanitize_filename_part(student.student_id)}_"
        f"{sanitize_filename_part(student.name)}"
    )
    pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
    jpg_path = os.path.join(output_dir, f"{base_name}.jpg")

    html_content, css = _build_warning_letter_html(student)
    archive = pymupdf.Archive(MAC_FONTS_DIR) if os.path.isdir(MAC_FONTS_DIR) else None

    def _render_page(target_doc: pymupdf.Document, active_css: str) -> float:
        pg = target_doc.new_page(width=595, height=842)  # A4 portrait
        top_y = 44
        if os.path.isfile(DEFAULT_LOGO_PATH):
            logo_w = 270.0
            logo_h = 54.4
            x0 = (595.0 - logo_w) / 2.0
            logo_rect = pymupdf.Rect(x0, 24.0, x0 + logo_w, 24.0 + logo_h)
            pg.insert_image(logo_rect, filename=DEFAULT_LOGO_PATH, keep_proportion=True)
            top_y = 90
        rect = pymupdf.Rect(54, top_y, 595 - 54, 842 - 34)
        spare_h, _scale = pg.insert_htmlbox(rect, html_content, css=active_css, archive=archive)
        return spare_h

    doc = pymupdf.open()
    try:
        spare_height = _render_page(doc, css)

        if spare_height < 0:
            doc.close()
            doc = pymupdf.open()
            compact_css = (
                css.replace("font-size: 9.4pt;", "font-size: 8.4pt;")
                .replace("line-height: 1.38;", "line-height: 1.24;")
                .replace("margin-top: 14pt;", "margin-top: 8pt;")
            )
            _render_page(doc, compact_css)

        doc.save(pdf_path, deflate=True)

        # Render to high-res JPG (supports single-page or vertical stitch if multi-page)
        if len(doc) == 1:
            pix = doc[0].get_pixmap(dpi=dpi, alpha=False)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.save(jpg_path, "JPEG", quality=95)
        else:
            imgs = []
            for p in doc:
                pix = p.get_pixmap(dpi=dpi, alpha=False)
                imgs.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
            combined = Image.new("RGB", (max(i.width for i in imgs), sum(i.height for i in imgs)), (255, 255, 255))
            y_off = 0
            for im in imgs:
                combined.paste(im, (0, y_off))
                y_off += im.height
            combined.save(jpg_path, "JPEG", quality=94)
    finally:
        doc.close()

    return pdf_path, jpg_path


def batch_generate_warning_letters(
    students: List[WarningStudent],
    output_dir: str,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    if not students:
        raise ValueError("No students selected for warning letter generation.")

    os.makedirs(output_dir, exist_ok=True)
    generated: List[Tuple[str, str]] = []
    total = len(students)

    for idx, stu in enumerate(students, start=1):
        if progress_cb:
            progress_cb(
                f"[{idx}/{total}] Generating PDF & JPG / 正在生成: "
                f"{stu.homeroom} {stu.student_id} {stu.name}..."
            )
        pdf_p, jpg_p = generate_warning_letter_pdf_and_jpg(stu, output_dir)
        generated.append((pdf_p, jpg_p))

    if progress_cb:
        progress_cb(f"Done! Generated {len(generated)} PDF + JPG pairs in: {output_dir}")

    return {
        "output_dir": output_dir,
        "count": len(generated),
        "files": generated,
    }


def open_managebac_batch(
    students: List[WarningStudent],
    start_index: int = 0,
    batch_size: int = 5,
    log_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    """
    Open a slice of ManageBac Behavior Note URLs in the default browser.
    Returns {opened_count, next_index, total_with_url, missing_students}.
    """
    with_url = [s for s in students if s.managebac_url]
    missing = [s for s in students if not s.managebac_url]

    batch = with_url[start_index: start_index + max(1, batch_size)]
    for st in batch:
        webbrowser.open_new_tab(st.managebac_url)
        if log_cb:
            log_cb(f"Opened ManageBac / 已打开 -> {st.homeroom} {st.student_id} {st.name}: {st.managebac_url}")

    next_index = start_index + len(batch)
    return {
        "opened_count": len(batch),
        "next_index": next_index,
        "total_with_url": len(with_url),
        "missing_students": missing,
    }
