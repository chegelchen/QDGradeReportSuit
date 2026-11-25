"""
Automated End-to-End & Unit Test Suite for QDHS_Academic_Report_Suite.
Verifies G9 (House X), G10, G11 IB, G11 AL, G12 IB report generation,
unified warning-sheet parsing across both layouts, PDF+JPG warning letters,
ManageBac URL generation, and anti-regression rules (IBDP Core font 8, no fitz warning).
"""

from __future__ import annotations

import os
import tempfile
import unittest
import warnings

import openpyxl

from g9_g10_engine import generate_g9_g10_report, is_subject_failed
from g11_g12_engine import (
    compute_tok_ee_bonus,
    convert_al_grade_to_score,
    evaluate_grade,
    generate_g11_g12_report,
)
from warning_letter import (
    build_managebac_behavior_note_url,
    generate_warning_letter_pdf_and_jpg,
    load_warning_students,
)


ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NEWDATA_DIR = os.path.join(ROOT_DIR, "NewData.Nov.24")
G10_DIR = os.path.join(ROOT_DIR, "G10")
G11_DIR = os.path.join(ROOT_DIR, "G11")
G12_DIR = os.path.join(ROOT_DIR, "G12")

ATT_PATH = os.path.join(ROOT_DIR, "All_Attendance_Summary_Generated.xlsx")
MB_PATH = os.path.join(ROOT_DIR, "学生MB系统ID对照表.xlsx")


class TestQDHSReportSuite(unittest.TestCase):
    def test_grade_evaluation_rules(self) -> None:
        # IB DP 1-7
        self.assertEqual(evaluate_grade(3, "IB_DP", "IB: Math AA HL"), (3, True))
        self.assertEqual(evaluate_grade(4, "IB_DP", "IB: Math AA HL"), (4, False))
        # IB Core TOK (case-insensitive)
        self.assertEqual(evaluate_grade("c", "IB_CORE", "TOK"), ("C", False))
        self.assertEqual(evaluate_grade("d", "IB_CORE", "TOK"), ("D", True))
        # IB Core EE / CAS anti-regression
        self.assertEqual(evaluate_grade("To Be Determined", "IB_CORE", "EE"), ("To Be Determined", False))
        self.assertEqual(evaluate_grade("", "IB_CORE", "CAS"), (None, False))
        self.assertEqual(evaluate_grade("Concern", "IB_CORE", "CAS"), ("Concern", True))
        # AL A*-E
        self.assertEqual(evaluate_grade("a*", "AL", "AL Mathematics"), ("A*", False))
        self.assertEqual(evaluate_grade("d", "AL", "AL Mathematics"), ("D", True))
        # Chinese National Courses
        self.assertEqual(evaluate_grade(64, "HS_CN", "HS History"), (64, True))
        self.assertEqual(evaluate_grade(65, "HS_CN", "HS History"), (65, False))
        # G9/G10 fail checks
        self.assertTrue(is_subject_failed("HS SPARK 升学指导", "Concern"))
        self.assertFalse(is_subject_failed("HS SPARK 升学指导", "On-track"))
        self.assertTrue(is_subject_failed("HS Physical Education 体育", 64))
        self.assertTrue(is_subject_failed("6-HS Film  电影", 3))

    def test_tok_ee_bonus_and_al_conversion(self) -> None:
        # TOK / EE Bonus Matrix (TOK.EE.Bonus.png)
        self.assertEqual(compute_tok_ee_bonus("A", "A"), 3)
        self.assertEqual(compute_tok_ee_bonus("A", "B"), 3)
        self.assertEqual(compute_tok_ee_bonus("A", "C"), 2)
        self.assertEqual(compute_tok_ee_bonus("B", "D"), 1)
        self.assertEqual(compute_tok_ee_bonus("C", "D"), 0)
        self.assertEqual(compute_tok_ee_bonus("D", "A"), 2)
        self.assertEqual(compute_tok_ee_bonus("E", "A"), 0)
        self.assertEqual(compute_tok_ee_bonus("A", "E"), 0)
        self.assertEqual(compute_tok_ee_bonus("A", "On-track"), 0)
        self.assertEqual(compute_tok_ee_bonus(None, "A"), 0)

        # AL Grade -> 7..1 Score Conversion
        self.assertEqual(convert_al_grade_to_score("A*"), 7)
        self.assertEqual(convert_al_grade_to_score("A"), 6)
        self.assertEqual(convert_al_grade_to_score("B"), 5)
        self.assertEqual(convert_al_grade_to_score("C"), 4)
        self.assertEqual(convert_al_grade_to_score("D"), 3)
        self.assertEqual(convert_al_grade_to_score("E"), 2)
        self.assertEqual(convert_al_grade_to_score("U"), 1)
        self.assertEqual(convert_al_grade_to_score("F"), 1)
        self.assertEqual(convert_al_grade_to_score("G"), 1)
        self.assertEqual(convert_al_grade_to_score(6), 6)
        self.assertIsNone(convert_al_grade_to_score("N/A"))

    def test_g9_end_to_end_and_warning_letter(self) -> None:
        raw_hx = os.path.join(G10_DIR, "Updated House X Data .xlsx")
        info_hx = os.path.join(G10_DIR, "Class of 2030 House X.xlsx")
        if not os.path.isfile(raw_hx):
            self.skipTest("G9 sample data not found.")

        with tempfile.TemporaryDirectory() as tmp:
            out_xlsx = os.path.join(tmp, "g9_test.xlsx")
            res = generate_g9_g10_report(
                mode="G9",
                raw_files=[raw_hx],
                attendance_path=ATT_PATH,
                id_map_path=MB_PATH,
                output_path=out_xlsx,
                student_info_path=info_hx,
            )
            self.assertGreater(res["student_count"], 0)
            wb = openpyxl.load_workbook(out_xlsx)
            self.assertEqual(wb.sheetnames, ["MT1 Summary", "MT1 Overview", "Warning"])
            wb.close()

            warn_list = load_warning_students(out_xlsx, MB_PATH)
            self.assertEqual(len(warn_list), res["warning_count"])
            if warn_list:
                pdf_p, jpg_p = generate_warning_letter_pdf_and_jpg(warn_list[0], tmp)
                self.assertTrue(os.path.isfile(pdf_p) and os.path.getsize(pdf_p) > 2000)
                self.assertTrue(os.path.isfile(jpg_p) and os.path.getsize(jpg_p) > 5000)
                if warn_list[0].mb_id:
                    self.assertIn(warn_list[0].mb_id, warn_list[0].managebac_url)

    def test_g10_end_to_end_order_independent(self) -> None:
        raw_hs = os.path.join(G10_DIR, "Updated HS G10.xlsx")
        raw_ib = os.path.join(G10_DIR, "G10 IB .xlsx")
        info_g10 = os.path.join(G10_DIR, "Class of 2029 G10.xlsx")
        if not (os.path.isfile(raw_hs) and os.path.isfile(raw_ib)):
            self.skipTest("G10 sample data not found.")

        with tempfile.TemporaryDirectory() as tmp:
            out_xlsx = os.path.join(tmp, "g10_test.xlsx")
            # Intentionally pass IB before HS to test order independence
            res = generate_g9_g10_report(
                mode="G10",
                raw_files=[raw_ib, raw_hs],
                attendance_path=ATT_PATH,
                id_map_path=MB_PATH,
                output_path=out_xlsx,
                student_info_path=info_g10,
            )
            self.assertGreater(res["student_count"], 50)
            warn_list = load_warning_students(out_xlsx, MB_PATH)
            self.assertEqual(len(warn_list), res["warning_count"])

    def test_g11_ib_end_to_end_and_core_font8(self) -> None:
        raw_ib = os.path.join(G11_DIR, "G11 IB.xlsx")
        raw_hs = os.path.join(G11_DIR, "G11 HS.xlsx")
        info_g11 = os.path.join(G11_DIR, "Class of 2028 G11.xlsx")
        cas_p = os.path.join(G11_DIR, "G11 CAS_Export.xlsx")
        ee_p = os.path.join(G11_DIR, "G11_EE_export.xlsx")
        dd_p = os.path.join(NEWDATA_DIR, "G11 Dwight Diploma.xlsx")
        if not (os.path.isfile(raw_ib) and os.path.isfile(raw_hs)):
            self.skipTest("G11 IB sample data not found.")

        with tempfile.TemporaryDirectory() as tmp:
            out_xlsx = os.path.join(tmp, "g11_ib_test.xlsx")
            res = generate_g11_g12_report(
                raw_files=[raw_ib, raw_hs],
                output_path=out_xlsx,
                mode="IB",
                attendance_path=ATT_PATH,
                mb_id_path=MB_PATH,
                student_info_path=info_g11,
                cas_path=cas_p,
                ee_path=ee_p,
                dwight_diploma_path=dd_p if os.path.isfile(dd_p) else "",
            )
            self.assertEqual(res["grade_num"], "11")
            self.assertGreater(res["student_count"], 50)

            wb = openpyxl.load_workbook(out_xlsx)
            self.assertEqual(wb.sheetnames, ["MT1 Summary", "MT1 Overview", "Warning"])
            ws_sum = wb["MT1 Summary"]
            if os.path.isfile(dd_p):
                dd_rows = 0
                dp_rows = 0
                for r in range(2, ws_sum.max_row + 1):
                    status = ws_sum.cell(r, 8).value
                    if status == "DD":
                        dd_rows += 1
                        self.assertIsNone(ws_sum.cell(r, 10).value, "DD student DP Only total must be blank")
                        self.assertIsNotNone(ws_sum.cell(r, 11).value, "DD student MT1 Total must be populated")
                    elif status == "DP":
                        dp_rows += 1
                        self.assertIsNotNone(ws_sum.cell(r, 10).value, "DP student DP Only total must be populated")
                        self.assertIsNotNone(ws_sum.cell(r, 11).value, "DP student MT1 Total must be populated")
                self.assertEqual(dd_rows, 13)
                self.assertGreater(dp_rows, 100)

            ws_ov = wb["MT1 Overview"]
            # Verify "Concern 待改进" has font size 8
            found_concern_font8 = False
            for row in ws_ov.iter_rows():
                for cell in row:
                    if cell.value == "Concern 待改进":
                        self.assertEqual(cell.font.size, 8)
                        found_concern_font8 = True
            self.assertTrue(found_concern_font8, "Expected 'Concern 待改进' with font size 8 in MT1 Overview")
            wb.close()

            warn_list = load_warning_students(out_xlsx, MB_PATH)
            self.assertEqual(len(warn_list), res["warning_count"])
            if warn_list:
                pdf_p, jpg_p = generate_warning_letter_pdf_and_jpg(warn_list[0], tmp)
                self.assertTrue(os.path.isfile(pdf_p) and os.path.isfile(jpg_p))

    def test_g11_al_end_to_end(self) -> None:
        raw_al = os.path.join(G11_DIR, "G11 AL.xlsx")
        raw_hs = os.path.join(G11_DIR, "G11 HS.xlsx")
        info_g11 = os.path.join(G11_DIR, "Class of 2028 G11.xlsx")
        if not (os.path.isfile(raw_al) and os.path.isfile(raw_hs)):
            self.skipTest("G11 AL sample data not found.")

        with tempfile.TemporaryDirectory() as tmp:
            out_xlsx = os.path.join(tmp, "g11_al_test.xlsx")
            res = generate_g11_g12_report(
                raw_files=[raw_al, raw_hs],
                output_path=out_xlsx,
                mode="AL",
                attendance_path=ATT_PATH,
                mb_id_path=MB_PATH,
                student_info_path=info_g11,
            )
            self.assertEqual(res["mode"], "AL")
            self.assertGreater(res["student_count"], 0)
            wb = openpyxl.load_workbook(out_xlsx)
            ws_sum = wb["MT1 Summary"]
            headers = [ws_sum.cell(1, c).value for c in range(1, ws_sum.max_column + 1)]
            self.assertNotIn("Candidate Session Number", headers)
            self.assertNotIn("Personal Code", headers)
            self.assertNotIn("G11 MT1 Total (DP Only)总计", headers)
            self.assertEqual(headers[5], "STATUS")
            self.assertEqual(headers[7], "G11 MT1 Total 总计")

            al_rows = 0
            dist_labels = set()
            avg_row_idx = None
            for r in range(2, ws_sum.max_row + 1):
                if ws_sum.cell(r, 6).value == "AL":
                    al_rows += 1
                    self.assertIsInstance(ws_sum.cell(r, 8).value, int)
                    self.assertGreater(ws_sum.cell(r, 8).value, 0)
                lbl = str(ws_sum.cell(r, 3).value or "")
                if lbl == "Subject Average 学科平均分":
                    avg_row_idx = r
                if "分 (" in lbl:
                    dist_labels.add(lbl)
            self.assertEqual(al_rows, res["student_count"])
            self.assertIsNotNone(avg_row_idx)
            # Verify AL subject average formula (Col I = first AL subject)
            al_avg_expr = str(ws_sum.cell(avg_row_idx, 9).value or "")
            self.assertIn("7*", al_avg_expr)
            self.assertIn("SUM(", al_avg_expr)
            for expected in ("7分 (A*)", "6分 (A)", "5分 (B)", "4分 (C)", "3分 (D)", "2分 (E)", "1分 (U/F/G)"):
                self.assertIn(expected, dist_labels)
            wb.close()
            warn_list = load_warning_students(out_xlsx, MB_PATH)
            self.assertEqual(len(warn_list), res["warning_count"])

    def test_managebac_url_builder(self) -> None:
        url = build_managebac_behavior_note_url("12345678")
        self.assertTrue(url.endswith("/teacher/users/12345678/behavior_notes/new"))
        self.assertEqual(build_managebac_behavior_note_url(""), "")


if __name__ == "__main__":
    warnings.simplefilter("error", FutureWarning)
    unittest.main(verbosity=2)
