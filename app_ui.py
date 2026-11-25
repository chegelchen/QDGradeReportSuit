"""
QDHS Grade Report Generator / QDHS七德高中成绩报告生成器
Unified 2-Step Bilingual Desktop GUI (G9 House X · G10 · G11/G12 IB & AL).

Design Highlights:
  - Title Banner: QDHS Red Pantone 186EC (#C8102E) with white bilingual title
  - Official Colors: Red #C8102E (Pantone 186EC) & Blue #00205B (Pantone 281EC)
  - All UI labels are English + Chinese bilingual
  - All buttons display English on top and Chinese on the bottom
  - Single-level 4-way mode selector: G9 (House X), G10, G11/G12 (IB), G11/G12 (AL)
  - Thread-safe background execution via queue.Queue + root.after() polling
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from typing import Any, Callable, Dict, List, Optional

from common_io import QDHS_BLUE_HEX, QDHS_BLUE_HOVER, QDHS_RED_HEX, QDHS_RED_HOVER
from g9_g10_engine import generate_g9_g10_report
from g11_g12_engine import generate_g11_g12_report
from warning_letter import (
    WarningStudent,
    batch_generate_warning_letters,
    load_warning_students,
    open_managebac_batch,
)


# =============================================================================
# UI Colors & Constants
# =============================================================================
BG_PAGE = "#F4F6F9"
BG_CARD = "#FFFFFF"
BORDER_SUBTLE = "#D0D7DE"
TEXT_DARK = "#1B2430"
TEXT_MUTED = "#4E5D70"
BTN_SEC_BG = "#FFFFFF"
BTN_SEC_HOVER = "#EDF1F7"
BTN_DISABLED_BG = "#C8CFD8"
BTN_DISABLED_FG = "#6B7280"

MODE_SPECS = [
    ("G9", "G9 (House X)", "九年级 (House X)"),
    ("G10", "G10", "十年级"),
    ("G11_12_IB", "G11 / G12 (IB)", "高二/高三 IB"),
    ("G11_12_AL", "G11 / G12 (AL)", "高二/高三 AL"),
]

MODE_HINTS = {
    "G9": (
        "G9 (House X): Add House X raw grade Excel (1 file). PBL/Workshop/Electives are excluded automatically.\n"
        "G9 (House X)：请添加 House X 原始成绩表（1 份），程序自动排除 PBL、Workshop 与 Electives。"
    ),
    "G10": (
        "G10: Add both G10 HS and G10 IB raw grade Excels (2 files, any order). Arts (6-HS), Chinese National & PE excluded from total.\n"
        "G10：请添加 G10 HS 与 G10 IB 两份原始成绩表（不限顺序），艺术、中方课程与体育不计入总分但计入不合格科目数。"
    ),
    "G11_12_IB": (
        "G11/G12 (IB): Add IB + HS raw grade Excels (grade 11 or 12 auto-detected). Optional CAS, EE & Dwight Diploma lists below.\n"
        "G11/G12 (IB)：请添加 IB 与 HS 原始成绩表（自动识别 11/12 年级），下方可选填 CAS、EE 导出表及 Dwight Diploma (DD) 名单。"
    ),
    "G11_12_AL": (
        "G11/G12 (AL): Add AL + HS raw grade Excels (grade 11 or 12 auto-detected). Converts A*-U/F/G to 7-1 total & distribution.\n"
        "G11/G12 (AL)：请添加 AL 与 HS 原始成绩表（自动识别 11/12 年级），自动按 A*-U/F/G 折算 7-1 总分及统计各分值人数。"
    ),
}


# =============================================================================
# macOS-Safe Bilingual Two-Line Button (English Top / Chinese Bottom)
# =============================================================================
class BilingualButton(tk.Frame):
    """
    Custom macOS-safe clickable button with English on top and Chinese on bottom.
    Renders exact QDHS Pantone #C8102E / #00205B colors even under macOS Aqua.
    """

    def __init__(
        self,
        parent: tk.Misc,
        en_text: str,
        zh_text: str,
        command: Callable[[], None],
        style: str = "blue",  # "blue" | "red" | "secondary"
        padx: int = 14,
        pady: int = 6,
        en_size: int = 10,
        zh_size: int = 9,
        width_px: Optional[int] = None,
    ) -> None:
        if style == "red":
            self._bg = QDHS_RED_HEX
            self._hover = QDHS_RED_HOVER
            self._fg = "#FFFFFF"
            border_color = QDHS_RED_HEX
        elif style == "blue":
            self._bg = QDHS_BLUE_HEX
            self._hover = QDHS_BLUE_HOVER
            self._fg = "#FFFFFF"
            border_color = QDHS_BLUE_HEX
        else:
            self._bg = BTN_SEC_BG
            self._hover = BTN_SEC_HOVER
            self._fg = QDHS_BLUE_HEX
            border_color = QDHS_BLUE_HEX

        super().__init__(
            parent,
            bg=self._bg,
            highlightbackground=border_color,
            highlightthickness=1,
            cursor="hand2",
        )
        self._command = command
        self._enabled = True

        self._inner = tk.Frame(self, bg=self._bg)
        self._inner.pack(fill="both", expand=True, padx=padx, pady=pady)

        self._lbl_en = tk.Label(
            self._inner,
            text=en_text,
            bg=self._bg,
            fg=self._fg,
            font=("Helvetica Neue", en_size, "bold"),
            cursor="hand2",
        )
        self._lbl_en.pack(anchor="center")

        self._lbl_zh = tk.Label(
            self._inner,
            text=zh_text,
            bg=self._bg,
            fg=self._fg,
            font=("PingFang SC", zh_size, "bold" if style != "secondary" else "normal"),
            cursor="hand2",
        )
        self._lbl_zh.pack(anchor="center")

        self._spacer: Optional[tk.Frame] = None
        if width_px:
            spacer_w = max(10, width_px - 2 * padx)
            self._spacer = tk.Frame(self._inner, bg=self._bg, width=spacer_w, height=0)
            self._spacer.pack()

        for widget in (self, self._inner, self._lbl_en, self._lbl_zh):
            widget.bind("<Button-1>", self._on_click)
            widget.bind("<Enter>", self._on_enter)
            widget.bind("<Leave>", self._on_leave)

    def _apply_colors(self, bg: str, fg: str) -> None:
        self.configure(bg=bg)
        self._inner.configure(bg=bg)
        self._lbl_en.configure(bg=bg, fg=fg)
        self._lbl_zh.configure(bg=bg, fg=fg)
        if self._spacer is not None:
            self._spacer.configure(bg=bg)

    def _on_click(self, _event: tk.Event) -> None:
        if self._enabled and self._command:
            self._command()

    def _on_enter(self, _event: tk.Event) -> None:
        if self._enabled:
            self._apply_colors(self._hover, self._fg)

    def _on_leave(self, _event: tk.Event) -> None:
        if self._enabled:
            self._apply_colors(self._bg, self._fg)

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        cur = "hand2" if enabled else "arrow"
        for w in (self, self._inner, self._lbl_en, self._lbl_zh):
            w.configure(cursor=cur)
        if enabled:
            self._apply_colors(self._bg, self._fg)
        else:
            self._apply_colors(BTN_DISABLED_BG, BTN_DISABLED_FG)


# =============================================================================
# Helper to Locate Default Workspace Files (Read-Only Convenience)
# =============================================================================
def _find_first_existing(base_dirs: List[str], candidates: List[str]) -> str:
    for d in base_dirs:
        if not d or not os.path.isdir(d):
            continue
        for name in candidates:
            p = os.path.join(d, name)
            if os.path.isfile(p):
                return p
    return ""


# =============================================================================
# Main Application Window
# =============================================================================
class QDHSGradeReportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("QDHS Grade Report Generator | QDHS七德高中成绩报告生成器")
        screen_h = self.winfo_screenheight()
        init_h = min(1020, screen_h - 60) if screen_h > 600 else 1020
        self.geometry(f"1160x{init_h}")
        self.minsize(980, min(680, init_h))
        self.configure(bg=BG_PAGE)

        # Base directories for smart default file discovery
        self._app_dir = os.path.dirname(os.path.abspath(__file__))
        self._parent_dir = os.path.dirname(self._app_dir)
        self._search_dirs = [
            self._parent_dir,
            os.path.join(self._parent_dir, "NewData.Nov.24"),
            os.path.join(self._parent_dir, "G10"),
            os.path.join(self._parent_dir, "G11"),
            os.path.join(self._parent_dir, "G12"),
            os.path.join(self._parent_dir, "HouseX"),
        ]

        # Thread-safe UI event queue
        self._queue: queue.Queue = queue.Queue()
        self._busy = False

        # State variables
        self.mode_var = tk.StringVar(value="G9")
        self.raw_files: List[str] = []
        self.attendance_var = tk.StringVar()
        self.mb_id_var = tk.StringVar()
        self.student_info_var = tk.StringVar()
        self.cas_var = tk.StringVar()
        self.ee_var = tk.StringVar()
        self.dwight_diploma_var = tk.StringVar()
        self.output_excel_var = tk.StringVar()

        # Step 2 state
        self.warning_excel_var = tk.StringVar()
        self.warning_mb_var = tk.StringVar()
        self.warning_out_dir_var = tk.StringVar(
            value=os.path.join(self._app_dir, "Generated_Warning_Letters")
        )
        self.batch_size_var = tk.IntVar(value=5)
        self.warning_students: List[WarningStudent] = []
        self.checked_map: Dict[str, bool] = {}  # tree item_id -> bool
        self._item_to_student: Dict[str, WarningStudent] = {}
        self._mb_batch_offset = 0

        self._configure_ttk_styles()
        self._build_ui()
        self._apply_mode_defaults("G9")
        self.after(100, self._poll_queue)

    def _configure_ttk_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TNotebook", background=BG_PAGE, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            font=("Helvetica Neue", 10, "bold"),
            padding=(16, 8),
            background="#E5EAF0",
            foreground=QDHS_BLUE_HEX,
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", QDHS_BLUE_HEX)],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "QDHS.Treeview",
            font=("Helvetica Neue", 10),
            rowheight=26,
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
        )
        style.configure(
            "QDHS.Treeview.Heading",
            font=("Helvetica Neue", 10, "bold"),
            background=QDHS_BLUE_HEX,
            foreground="#FFFFFF",
            relief="flat",
        )
        style.map("QDHS.Treeview.Heading", background=[("active", QDHS_BLUE_HOVER)])

    # -------------------------------------------------------------------------
    # UI Layout
    # -------------------------------------------------------------------------
    def _build_ui(self) -> None:
        # 1. Top Red Header Banner (#C8102E Pantone 186EC) — compact single-row title
        banner = tk.Frame(self, bg=QDHS_RED_HEX, height=52)
        banner.pack(fill="x", side="top")
        banner.pack_propagate(False)

        banner_inner = tk.Frame(banner, bg=QDHS_RED_HEX)
        banner_inner.pack(fill="both", expand=True, padx=22, pady=8)

        tk.Label(
            banner_inner,
            text="QDHS Grade Report Generator  |  QDHS七德高中成绩报告生成器",
            bg=QDHS_RED_HEX,
            fg="#FFFFFF",
            font=("Helvetica Neue", 17, "bold"),
        ).pack(anchor="w", expand=True)

        # Thin QDHS Blue accent strip under banner
        tk.Frame(self, bg=QDHS_BLUE_HEX, height=4).pack(fill="x", side="top")

        # 2. Custom Dynamic 2-Step Tab Bar (Selected: larger & bold; Unselected: smaller & faded; Chinese below English)
        tab_bar_outer = tk.Frame(self, bg=BG_PAGE)
        tab_bar_outer.pack(fill="x", padx=14, pady=(8, 0))

        self._step_tabs: Dict[int, Dict[str, Any]] = {}
        tab_specs = [
            (1, "Step 1: Grade & Attendance Report", "第一步：生成成绩及出勤报告"),
            (2, "Step 2: Warning Letter & ManageBac", "第二步：警告信与 ManageBac"),
        ]
        for step_num, en_txt, zh_txt in tab_specs:
            tab_frame = tk.Frame(
                tab_bar_outer,
                bg="#E6EAF0",
                cursor="hand2",
                highlightthickness=1,
                highlightbackground=BORDER_SUBTLE,
            )
            tab_frame.pack(side="left", anchor="s", padx=(0, 8))

            inner = tk.Frame(tab_frame, bg="#E6EAF0", cursor="hand2")
            inner.pack(fill="both", expand=True)

            lbl_en = tk.Label(
                inner,
                text=en_txt,
                bg="#E6EAF0",
                fg="#8A94A6",
                font=("Helvetica Neue", 9, "normal"),
                cursor="hand2",
            )
            lbl_en.pack(anchor="center")

            lbl_zh = tk.Label(
                inner,
                text=zh_txt,
                bg="#E6EAF0",
                fg="#9AA3B2",
                font=("PingFang SC", 8, "normal"),
                cursor="hand2",
            )
            lbl_zh.pack(anchor="center", pady=(1, 0))

            indicator = tk.Frame(tab_frame, bg="#E6EAF0", height=4)
            indicator.pack(fill="x", side="bottom")

            for w in (tab_frame, inner, lbl_en, lbl_zh, indicator):
                w.bind("<Button-1>", lambda _e, s=step_num: self._select_step(s))

            self._step_tabs[step_num] = {
                "frame": tab_frame,
                "inner": inner,
                "en": lbl_en,
                "zh": lbl_zh,
                "indicator": indicator,
            }

        # Bottom divider line under the step tab bar
        tk.Frame(self, bg=QDHS_BLUE_HEX, height=2).pack(fill="x", padx=14)

        # Vertically Scrollable Step Content Container (supports smaller screens)
        scroll_outer = tk.Frame(self, bg=BG_PAGE)
        scroll_outer.pack(fill="both", expand=True, padx=14, pady=(6, 10))

        self._main_canvas = tk.Canvas(scroll_outer, bg=BG_PAGE, highlightthickness=0, bd=0)
        self._main_canvas.pack(side="left", fill="both", expand=True)

        v_scroll = ttk.Scrollbar(scroll_outer, orient="vertical", command=self._main_canvas.yview)
        v_scroll.pack(side="right", fill="y")
        self._main_canvas.configure(yscrollcommand=v_scroll.set)

        self._step_container = tk.Frame(self._main_canvas, bg=BG_PAGE)
        self._canvas_win_id = self._main_canvas.create_window((0, 0), window=self._step_container, anchor="nw")

        self._step_container.bind(
            "<Configure>",
            lambda _e: self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all")),
        )
        self._main_canvas.bind("<Configure>", self._on_canvas_configure)
        self.bind_all("<MouseWheel>", self._on_global_mousewheel)
        self.bind_all("<Button-4>", self._on_global_mousewheel)
        self.bind_all("<Button-5>", self._on_global_mousewheel)

        self.tab_step1 = tk.Frame(self._step_container, bg=BG_PAGE)
        self.tab_step2 = tk.Frame(self._step_container, bg=BG_PAGE)

        self._build_step1_tab()
        self._build_step2_tab()
        self._current_step = 1
        self._select_step(1)

    def _on_canvas_configure(self, event: tk.Event) -> None:
        self._main_canvas.itemconfigure(self._canvas_win_id, width=event.width)
        self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all"))

    def _on_global_mousewheel(self, event: tk.Event) -> None:
        w = event.widget
        # Let inner scrollable widgets (Text, Listbox, Treeview) handle their own wheel events
        cur = w
        while cur is not None:
            if isinstance(cur, (tk.Text, tk.Listbox, ttk.Treeview)):
                return
            cur = getattr(cur, "master", None)
        if getattr(event, "num", None) == 4:
            self._main_canvas.yview_scroll(-3, "units")
        elif getattr(event, "num", None) == 5:
            self._main_canvas.yview_scroll(3, "units")
        elif getattr(event, "delta", 0):
            step = -1 if event.delta > 0 else 1
            if abs(event.delta) >= 120:
                step = int(-event.delta / 120)
            self._main_canvas.yview_scroll(step, "units")

    def _select_step(self, step_num: int) -> None:
        self._current_step = step_num
        for s_idx, parts in self._step_tabs.items():
            if s_idx == step_num:
                # Selected: larger size, bold crisp text, QDHS Blue bg + Red bottom bar
                parts["frame"].configure(bg=QDHS_BLUE_HEX, highlightbackground=QDHS_BLUE_HEX, highlightthickness=1)
                parts["inner"].configure(bg=QDHS_BLUE_HEX)
                parts["inner"].pack_configure(padx=26, pady=10)
                parts["en"].configure(bg=QDHS_BLUE_HEX, fg="#FFFFFF", font=("Helvetica Neue", 12, "bold"))
                parts["zh"].configure(bg=QDHS_BLUE_HEX, fg="#FCE8EC", font=("PingFang SC", 10, "bold"))
                parts["indicator"].configure(bg=QDHS_RED_HEX, height=4)
            else:
                # Unselected: smaller size, faded muted non-bold text
                parts["frame"].configure(bg="#E6EAF0", highlightbackground="#D4DAE3", highlightthickness=1)
                parts["inner"].configure(bg="#E6EAF0")
                parts["inner"].pack_configure(padx=14, pady=5)
                parts["en"].configure(bg="#E6EAF0", fg="#8A94A6", font=("Helvetica Neue", 9, "normal"))
                parts["zh"].configure(bg="#E6EAF0", fg="#9AA3B2", font=("PingFang SC", 8, "normal"))
                parts["indicator"].configure(bg="#E6EAF0", height=2)

        if step_num == 1:
            self.tab_step2.pack_forget()
            self.tab_step1.pack(fill="both", expand=True)
        else:
            self.tab_step1.pack_forget()
            self.tab_step2.pack(fill="both", expand=True)
        self._main_canvas.after_idle(
            lambda: (
                self._main_canvas.configure(scrollregion=self._main_canvas.bbox("all")),
                self._main_canvas.yview_moveto(0.0),
            )
        )

    # -------------------------------------------------------------------------
    # Step 1: Grade & Attendance Report
    # -------------------------------------------------------------------------
    def _build_step1_tab(self) -> None:
        container = tk.Frame(self.tab_step1, bg=BG_PAGE)
        container.pack(fill="both", expand=True, padx=6, pady=6)

        # Card 1: Single-Level 4-Way Mode Selection
        card_mode = self._make_card(
            container,
            "1. Select Grade & Curriculum Mode  /  选择年级与课程模式",
        )
        mode_row = tk.Frame(card_mode, bg=BG_CARD)
        mode_row.pack(fill="x", padx=12, pady=(4, 6))

        self._mode_cards: Dict[str, tk.Frame] = {}
        self._mode_labels: Dict[str, tuple[tk.Label, tk.Label]] = {}
        for code, en_lbl, zh_lbl in MODE_SPECS:
            box = tk.Frame(
                mode_row,
                bg="#FFFFFF",
                highlightthickness=2,
                highlightbackground=BORDER_SUBTLE,
                cursor="hand2",
            )
            box.pack(side="left", fill="x", expand=True, padx=5)
            l_en = tk.Label(box, text=en_lbl, bg="#FFFFFF", fg=QDHS_BLUE_HEX, font=("Helvetica Neue", 11, "bold"), cursor="hand2")
            l_en.pack(pady=(6, 0))
            l_zh = tk.Label(box, text=zh_lbl, bg="#FFFFFF", fg=TEXT_MUTED, font=("PingFang SC", 9), cursor="hand2")
            l_zh.pack(pady=(0, 6))
            for w in (box, l_en, l_zh):
                w.bind("<Button-1>", lambda _e, c=code: self._on_mode_selected(c))
            self._mode_cards[code] = box
            self._mode_labels[code] = (l_en, l_zh)

        self.mode_hint_lbl = tk.Label(
            card_mode,
            text="",
            bg="#F0F4FA",
            fg=QDHS_BLUE_HEX,
            font=("PingFang SC", 9),
            justify="left",
            anchor="w",
            padx=10,
            pady=5,
        )
        self.mode_hint_lbl.pack(fill="x", padx=12, pady=(0, 8))

        # Card 2: Raw Grade Files
        card_raw = self._make_card(
            container,
            "2. Raw Grade Excel Files (Read-Only)  /  原始成绩 Excel 文件（只读，不修改原文件）",
        )
        raw_body = tk.Frame(card_raw, bg=BG_CARD)
        raw_body.pack(fill="x", padx=12, pady=(4, 8))

        list_frame = tk.Frame(raw_body, bg=BG_CARD)
        list_frame.pack(side="left", fill="both", expand=True)

        self.raw_listbox = tk.Listbox(
            list_frame,
            height=4,
            font=("Helvetica Neue", 10),
            selectmode=tk.EXTENDED,
            bg="#FAFBFC",
            fg=TEXT_DARK,
            highlightthickness=1,
            highlightbackground=BORDER_SUBTLE,
        )
        self.raw_listbox.pack(side="left", fill="both", expand=True)
        sb_raw = ttk.Scrollbar(list_frame, orient="vertical", command=self.raw_listbox.yview)
        sb_raw.pack(side="right", fill="y")
        self.raw_listbox.config(yscrollcommand=sb_raw.set)

        raw_btns = tk.Frame(raw_body, bg=BG_CARD)
        raw_btns.pack(side="right", padx=(10, 0), anchor="n")

        self.btn_add_raw = BilingualButton(
            raw_btns, "Add Files", "添加文件", self._add_raw_files, style="blue", padx=10, pady=3, en_size=9, zh_size=8
        )
        self.btn_add_raw.pack(fill="x", pady=(0, 4))
        self.btn_rm_raw = BilingualButton(
            raw_btns, "Remove Selected", "移除选中", self._remove_raw_files, style="secondary", padx=10, pady=3, en_size=9, zh_size=8
        )
        self.btn_rm_raw.pack(fill="x", pady=(0, 4))
        self.btn_clr_raw = BilingualButton(
            raw_btns, "Clear All", "清空列表", self._clear_raw_files, style="secondary", padx=10, pady=3, en_size=9, zh_size=8
        )
        self.btn_clr_raw.pack(fill="x")

        # Card 3: Auxiliary Files & Output Path (Single unified grid for exact left label & Entry alignment)
        card_aux = self._make_card(
            container,
            "3. Auxiliary Files & Output Report Path  /  辅助数据文件与输出报表路径",
        )
        aux_grid = tk.Frame(card_aux, bg=BG_CARD)
        aux_grid.pack(fill="x", padx=12, pady=(4, 8))
        aux_grid.columnconfigure(0, minsize=215)
        aux_grid.columnconfigure(1, weight=1)
        aux_grid.columnconfigure(2, minsize=155)

        self._add_file_row(
            aux_grid, 0,
            "Attendance Summary\n考勤汇总表:",
            self.attendance_var,
            lambda: self._pick_open_file(self.attendance_var, "Select Attendance Excel / 选择考勤汇总表"),
        )
        self._add_file_row(
            aux_grid, 1,
            "Student MB ID Map\n行政班与 MB ID 对照表:",
            self.mb_id_var,
            lambda: self._pick_open_file(self.mb_id_var, "Select MB ID Map / 选择学生MB系统ID对照表"),
        )
        self._add_file_row(
            aux_grid, 2,
            "Student Info (Emails)\n学生信息表（邮箱）:",
            self.student_info_var,
            lambda: self._pick_open_file(self.student_info_var, "Select Student Info / 选择学生信息表"),
        )

        # IB-Only Optional Rows (CAS, EE & Dwight Diploma) inside the SAME grid so Entry boxes align identically
        cas_widgets = self._add_file_row(
            aux_grid, 3,
            "CAS Export (Optional)\nCAS 独立成绩表（可选）:",
            self.cas_var,
            lambda: self._pick_open_file(self.cas_var, "Select CAS Export / 选择 CAS 导出文件"),
            allow_clear=True,
        )
        ee_widgets = self._add_file_row(
            aux_grid, 4,
            "EE Export (Optional)\nEE 独立成绩表（可选）:",
            self.ee_var,
            lambda: self._pick_open_file(self.ee_var, "Select EE Export / 选择 EE 导出文件"),
            allow_clear=True,
        )
        dd_widgets = self._add_file_row(
            aux_grid, 5,
            "Dwight Diploma List (Opt.)\nDwight Diploma 名单（可选）:",
            self.dwight_diploma_var,
            lambda: self._pick_open_file(self.dwight_diploma_var, "Select Dwight Diploma List / 选择 Dwight Diploma 名单"),
            allow_clear=True,
        )
        self._ib_core_widgets = list(cas_widgets) + list(ee_widgets) + list(dd_widgets)

        # Output Path Row inside the SAME grid
        self._add_file_row(
            aux_grid, 6,
            "Output Report (.xlsx)\n输出 Excel 报表路径:",
            self.output_excel_var,
            self._pick_save_excel,
            btn_en="Save As",
            btn_zh="选择路径",
        )

        # Action Bar
        action_bar = tk.Frame(container, bg=BG_PAGE)
        action_bar.pack(fill="x", pady=(6, 6))

        self.btn_generate_report = BilingualButton(
            action_bar,
            "Generate Grade & Attendance Report",
            "开始生成成绩及出勤报告",
            self._start_generate_report,
            style="blue",
            padx=22,
            pady=8,
            en_size=11,
            zh_size=10,
        )
        self.btn_generate_report.pack(side="left", padx=(0, 12))

        self.btn_goto_step2 = BilingualButton(
            action_bar,
            "Go to Step 2: Warning & ManageBac",
            "前往第二步：警告信与 ManageBac",
            lambda: self._select_step(2),
            style="red",
            padx=18,
            pady=8,
            en_size=10,
            zh_size=9,
        )
        self.btn_goto_step2.pack(side="left", padx=(0, 12))

        self.btn_open_excel = BilingualButton(
            action_bar,
            "Open Output Folder",
            "打开报表所在目录",
            self._open_report_folder,
            style="secondary",
            padx=14,
            pady=8,
            en_size=9,
            zh_size=9,
        )
        self.btn_open_excel.pack(side="right")

        # Log Box (expands cleanly so the full Execution Log is visible)
        log_card = self._make_card(container, "Execution Log  /  运行日志", expand=True)
        self.log_step1 = scrolledtext.ScrolledText(
            log_card,
            height=7,
            font=("Menlo", 9),
            bg="#FAFBFC",
            fg=TEXT_DARK,
            state="disabled",
            relief="flat",
        )
        self.log_step1.pack(fill="both", expand=True, padx=10, pady=(2, 8))

    # -------------------------------------------------------------------------
    # Step 2: Warning Letters & ManageBac
    # -------------------------------------------------------------------------
    def _build_step2_tab(self) -> None:
        container = tk.Frame(self.tab_step2, bg=BG_PAGE)
        container.pack(fill="both", expand=True, padx=6, pady=6)

        # Card 1: Load Source Report & MB ID Map
        card_src = self._make_card(
            container,
            "1. Load Warning List from Report  /  从成绩报告加载警告信名单",
        )
        src_grid = tk.Frame(card_src, bg=BG_CARD)
        src_grid.pack(fill="x", padx=12, pady=(4, 8))
        src_grid.columnconfigure(1, weight=1)

        self._add_file_row(
            src_grid, 0,
            "Generated Report (.xlsx)\n成绩报告 Excel 文件:",
            self.warning_excel_var,
            lambda: self._pick_open_file(self.warning_excel_var, "Select MT1 Report / 选择成绩报告 Excel"),
        )
        self._add_file_row(
            src_grid, 1,
            "Student MB ID Map\n学生 MB ID 对照表:",
            self.warning_mb_var,
            lambda: self._pick_open_file(self.warning_mb_var, "Select MB ID Map / 选择学生MB系统ID对照表"),
        )

        load_row = tk.Frame(card_src, bg=BG_CARD)
        load_row.pack(fill="x", padx=12, pady=(0, 8))

        self.btn_load_warning = BilingualButton(
            load_row,
            "Load Warning List",
            "加载警告名单",
            self._load_warning_list,
            style="blue",
            padx=16,
            pady=5,
            en_size=10,
            zh_size=9,
        )
        self.btn_load_warning.pack(side="left", padx=(0, 10))

        self.btn_sel_all = BilingualButton(
            load_row, "Select All", "全选", lambda: self._set_all_checked(True), style="secondary", padx=12, pady=5, en_size=9, zh_size=8
        )
        self.btn_sel_all.pack(side="left", padx=(0, 6))

        self.btn_desel_all = BilingualButton(
            load_row, "Deselect All", "取消全选", lambda: self._set_all_checked(False), style="secondary", padx=12, pady=5, en_size=9, zh_size=8
        )
        self.btn_desel_all.pack(side="left", padx=(0, 12))

        self.sel_count_lbl = tk.Label(
            load_row,
            text="Selected: 0 / 0  |  已选：0 / 0 人  (Click any row to toggle / 点击行切换勾选)",
            bg=BG_CARD,
            fg=QDHS_BLUE_HEX,
            font=("PingFang SC", 10, "bold"),
        )
        self.sel_count_lbl.pack(side="left")

        # Card 2: Selectable Warning Student Table
        card_tbl = self._make_card(
            container,
            "2. Select Students for Warning Letters & ManageBac  /  勾选需要处理警告信与 ManageBac 的学生",
            expand=True,
        )
        tbl_frame = tk.Frame(card_tbl, bg=BG_CARD)
        tbl_frame.pack(fill="both", expand=True, padx=12, pady=(4, 8))

        cols = ("chk", "sid", "name", "homeroom", "fails", "skips", "mb_id", "details")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", style="QDHS.Treeview", height=10)
        col_specs = [
            ("chk", "Sel / 选", 62, "center"),
            ("sid", "Student ID / 学号", 110, "center"),
            ("name", "Name / 姓名", 150, "w"),
            ("homeroom", "Homeroom / 班级", 110, "center"),
            ("fails", "Fails / 不合格", 95, "center"),
            ("skips", "Skips / 缺勤", 85, "center"),
            ("mb_id", "MB ID", 100, "center"),
            ("details", "Failed Subjects / 不合格科目详情", 340, "w"),
        ]
        for cid, heading, width, anchor in col_specs:
            self.tree.heading(cid, text=heading)
            self.tree.column(cid, width=width, anchor=anchor)

        self.tree.pack(side="left", fill="both", expand=True)
        sb_tree = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.tree.yview)
        sb_tree.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=sb_tree.set)
        self.tree.bind("<Button-1>", self._on_tree_click)

        # Card 3: Output Warning Letters & Batch Open ManageBac
        card_act = self._make_card(
            container,
            "3. Generate Warning Letters (PDF + JPG) & Batch Open ManageBac  /  生成警告信与分批打开 ManageBac",
        )
        out_dir_grid = tk.Frame(card_act, bg=BG_CARD)
        out_dir_grid.pack(fill="x", padx=12, pady=(4, 6))
        out_dir_grid.columnconfigure(1, weight=1)
        self._add_file_row(
            out_dir_grid, 0,
            "Warning Output Folder\n警告信输出文件夹:",
            self.warning_out_dir_var,
            self._pick_warning_dir,
            btn_en="Choose Folder",
            btn_zh="选择文件夹",
            btn_width=118,
        )

        act_row = tk.Frame(card_act, bg=BG_CARD)
        act_row.pack(fill="x", padx=12, pady=(4, 10))

        self.btn_gen_letters = BilingualButton(
            act_row,
            "Generate Warning Letters (PDF + JPG)",
            "批量生成警告信 (PDF + JPG)",
            self._start_generate_letters,
            style="red",
            padx=18,
            pady=7,
            en_size=10,
            zh_size=9,
        )
        self.btn_gen_letters.pack(side="left", padx=(0, 14))

        # ManageBac Batch Controls
        mb_box = tk.Frame(act_row, bg="#F0F4FA", highlightbackground=QDHS_BLUE_HEX, highlightthickness=1)
        mb_box.pack(side="left", fill="x", expand=True, padx=(0, 8))

        tk.Label(
            mb_box,
            text="Batch Size\n每批人数:",
            bg="#F0F4FA",
            fg=QDHS_BLUE_HEX,
            font=("PingFang SC", 9, "bold"),
            justify="center",
        ).pack(side="left", padx=(8, 4), pady=4)

        tk.Spinbox(
            mb_box,
            from_=1,
            to=30,
            width=4,
            textvariable=self.batch_size_var,
            font=("Helvetica Neue", 11, "bold"),
            justify="center",
        ).pack(side="left", padx=(0, 8))

        self.btn_open_mb = BilingualButton(
            mb_box,
            "Open ManageBac Behavior Notes",
            "分批打开 ManageBac 链接",
            self._open_next_mb_batch,
            style="blue",
            padx=14,
            pady=6,
            en_size=10,
            zh_size=9,
        )
        self.btn_open_mb.pack(side="left", padx=(0, 8), pady=4)

        self.btn_reset_mb = BilingualButton(
            mb_box,
            "Reset Batch",
            "重置批次",
            self._reset_mb_batch,
            style="secondary",
            padx=10,
            pady=6,
            en_size=9,
            zh_size=8,
        )
        self.btn_reset_mb.pack(side="left", padx=(0, 8), pady=4)

        self.mb_progress_lbl = tk.Label(
            mb_box,
            text="Batch Progress / 进度:\n0 / 0",
            bg="#F0F4FA",
            fg=QDHS_RED_HEX,
            font=("Helvetica Neue", 9, "bold"),
            justify="center",
        )
        self.mb_progress_lbl.pack(side="left", padx=(4, 8))

        self.btn_open_letter_dir = BilingualButton(
            act_row,
            "Open Folder",
            "打开输出目录",
            self._open_warning_folder,
            style="secondary",
            padx=10,
            pady=7,
            en_size=9,
            zh_size=8,
            width_px=118,
        )
        self.btn_open_letter_dir.pack(side="right")

        # Step 2 Log
        log_card2 = self._make_card(container, "Step 2 Log  /  警告信与 ManageBac 日志")
        self.log_step2 = scrolledtext.ScrolledText(
            log_card2,
            height=5,
            font=("Menlo", 9),
            bg="#FAFBFC",
            fg=TEXT_DARK,
            state="disabled",
            relief="flat",
        )
        self.log_step2.pack(fill="both", expand=True, padx=10, pady=(2, 8))

    # -------------------------------------------------------------------------
    # Reusable Layout Builders
    # -------------------------------------------------------------------------
    def _make_card(self, parent: tk.Misc, title: str, expand: bool = False) -> tk.Frame:
        outer = tk.Frame(
            parent,
            bg=BG_CARD,
            highlightbackground=BORDER_SUBTLE,
            highlightthickness=1,
        )
        outer.pack(fill="both" if expand else "x", expand=expand, pady=(0, 8))

        hdr = tk.Frame(outer, bg="#EDF1F7")
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text=title,
            bg="#EDF1F7",
            fg=QDHS_BLUE_HEX,
            font=("Helvetica Neue", 10, "bold"),
            anchor="w",
            padx=12,
            pady=5,
        ).pack(fill="x")
        return outer

    def _add_file_row(
        self,
        parent: tk.Frame,
        row: int,
        label_text: str,
        var: tk.StringVar,
        browse_cmd: Callable[[], None],
        btn_en: str = "Browse",
        btn_zh: str = "选择文件",
        allow_clear: bool = False,
        btn_width: Optional[int] = None,
    ) -> tuple[tk.Widget, tk.Widget, tk.Widget]:
        lbl = tk.Label(
            parent,
            text=label_text,
            bg=BG_CARD,
            fg=TEXT_DARK,
            font=("PingFang SC", 9, "bold"),
            justify="left",
            anchor="w",
            width=24,
        )
        lbl.grid(row=row, column=0, sticky="w", padx=(0, 10), pady=3)

        entry = tk.Entry(
            parent,
            textvariable=var,
            font=("Helvetica Neue", 10),
            bg="#FAFBFC",
            fg=TEXT_DARK,
            relief="solid",
            bd=1,
        )
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 8), pady=3, ipady=4)

        def _scroll_entry_right(*_args: Any) -> None:
            try:
                if self.focus_get() != entry:
                    entry.after_idle(lambda e=entry: e.xview_moveto(1.0))
            except tk.TclError:
                pass

        var.trace_add("write", _scroll_entry_right)
        entry.bind("<Configure>", lambda _e: _scroll_entry_right())
        entry.bind("<FocusOut>", lambda _e: entry.after_idle(lambda e=entry: e.xview_moveto(1.0)))
        entry.after_idle(lambda e=entry: e.xview_moveto(1.0))

        btn_wrap = tk.Frame(parent, bg=BG_CARD)
        btn_wrap.grid(row=row, column=2, sticky="e", pady=3)
        if allow_clear:
            BilingualButton(
                btn_wrap, "Clear", "清空", lambda: var.set(""), style="secondary", padx=8, pady=2, en_size=8, zh_size=8
            ).pack(side="left", padx=(0, 4))

        BilingualButton(
            btn_wrap, btn_en, btn_zh, browse_cmd, style="secondary", padx=10, pady=2, en_size=8, zh_size=8, width_px=btn_width
        ).pack(side="left")

        return lbl, entry, btn_wrap

    # -------------------------------------------------------------------------
    # Mode Switching & Smart Defaults
    # -------------------------------------------------------------------------
    def _on_mode_selected(self, mode_code: str) -> None:
        if self._busy:
            return
        self.mode_var.set(mode_code)
        self._apply_mode_defaults(mode_code)

    def _apply_mode_defaults(self, mode_code: str) -> None:
        # Highlight selected mode card in QDHS Blue / Red
        for code, box in self._mode_cards.items():
            l_en, l_zh = self._mode_labels[code]
            if code == mode_code:
                box.configure(bg=QDHS_BLUE_HEX, highlightbackground=QDHS_RED_HEX, highlightthickness=2)
                l_en.configure(bg=QDHS_BLUE_HEX, fg="#FFFFFF")
                l_zh.configure(bg=QDHS_BLUE_HEX, fg="#E0E8F5")
            else:
                box.configure(bg="#FFFFFF", highlightbackground=BORDER_SUBTLE, highlightthickness=1)
                l_en.configure(bg="#FFFFFF", fg=QDHS_BLUE_HEX)
                l_zh.configure(bg="#FFFFFF", fg=TEXT_MUTED)

        self.mode_hint_lbl.configure(text=MODE_HINTS[mode_code])

        # Show CAS & EE rows ONLY in G11/G12 IB mode (keeping unified grid alignment)
        for w in self._ib_core_widgets:
            if mode_code == "G11_12_IB":
                w.grid()
            else:
                w.grid_remove()

        # Smart auto-fill of default files if present in workspace folders
        att = _find_first_existing(self._search_dirs, ["All_Attendance_Summary_Generated.xlsx"])
        mb = _find_first_existing(self._search_dirs, ["学生MB系统ID对照表.xlsx"])
        if att and not self.attendance_var.get():
            self.attendance_var.set(att)
        if mb:
            if not self.mb_id_var.get():
                self.mb_id_var.set(mb)
            if not self.warning_mb_var.get():
                self.warning_mb_var.set(mb)

        raw_defaults: List[str] = []
        info_default = ""
        out_name = "QDHS_MT1_Report.xlsx"

        if mode_code == "G9":
            f1 = _find_first_existing(self._search_dirs, ["Updated House X Data .xlsx", "Updated House X Data.xlsx"])
            if f1:
                raw_defaults.append(f1)
            info_default = _find_first_existing(self._search_dirs, ["Class of 2030 House X.xlsx"])
            out_name = "G9_HouseX_MT1_Report_2026-2027.xlsx"
        elif mode_code == "G10":
            f1 = _find_first_existing(self._search_dirs, ["Updated HS G10.xlsx"])
            f2 = _find_first_existing(self._search_dirs, ["G10 IB .xlsx", "G10 IB.xlsx"])
            for f in (f1, f2):
                if f:
                    raw_defaults.append(f)
            info_default = _find_first_existing(self._search_dirs, ["Class of 2029 G10.xlsx"])
            out_name = "G10_MT1_Report_2026-2027.xlsx"
        elif mode_code == "G11_12_IB":
            f1 = _find_first_existing(self._search_dirs, ["G11 IB.xlsx"])
            f2 = _find_first_existing(self._search_dirs, ["G11 HS.xlsx"])
            for f in (f1, f2):
                if f:
                    raw_defaults.append(f)
            info_default = _find_first_existing(self._search_dirs, ["Class of 2028 G11.xlsx", "Class of 2027 G12.xlsx"])
            cas_f = _find_first_existing(self._search_dirs, ["G11 CAS_Export.xlsx"])
            ee_f = _find_first_existing(self._search_dirs, ["G11_EE_export.xlsx"])
            dd_f = _find_first_existing(self._search_dirs, ["G11 Dwight Diploma.xlsx", "G12 Dwight Diploma.xlsx"])
            self.cas_var.set(cas_f)
            self.ee_var.set(ee_f)
            self.dwight_diploma_var.set(dd_f)
            out_name = "G11_G12_IB_MT1_Report_2026-2027.xlsx"
        elif mode_code == "G11_12_AL":
            f1 = _find_first_existing(self._search_dirs, ["G11 AL.xlsx"])
            f2 = _find_first_existing(self._search_dirs, ["G11 HS.xlsx"])
            for f in (f1, f2):
                if f:
                    raw_defaults.append(f)
            info_default = _find_first_existing(self._search_dirs, ["Class of 2028 G11.xlsx", "Class of 2027 G12.xlsx"])
            out_name = "G11_G12_AL_MT1_Report_2026-2027.xlsx"

        self.raw_files = raw_defaults
        self._refresh_raw_listbox()
        if info_default:
            self.student_info_var.set(info_default)
        self.output_excel_var.set(os.path.join(self._app_dir, "Generated_Reports", out_name))
        # Isolate warning letter output per mode so G9/G10/G11/G12 files never mix in the same folder
        mode_subfolders = {
            "G9": "G9_HouseX",
            "G10": "G10",
            "G11_12_IB": "G11_G12_IB",
            "G11_12_AL": "G11_G12_AL",
        }
        self.warning_out_dir_var.set(
            os.path.join(self._app_dir, "Generated_Warning_Letters", mode_subfolders.get(mode_code, mode_code))
        )

    def _refresh_raw_listbox(self) -> None:
        self.raw_listbox.delete(0, tk.END)
        for p in self.raw_files:
            self.raw_listbox.insert(tk.END, p)
        self.raw_listbox.after_idle(lambda: self.raw_listbox.xview_moveto(1.0))

    # -------------------------------------------------------------------------
    # File Pickers
    # -------------------------------------------------------------------------
    def _add_raw_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Raw Grade Excel Files / 选择原始成绩 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        for p in paths:
            if p and p not in self.raw_files:
                self.raw_files.append(p)
        self._refresh_raw_listbox()

    def _remove_raw_files(self) -> None:
        sel = list(self.raw_listbox.curselection())
        for idx in reversed(sel):
            if 0 <= idx < len(self.raw_files):
                self.raw_files.pop(idx)
        self._refresh_raw_listbox()

    def _clear_raw_files(self) -> None:
        self.raw_files.clear()
        self._refresh_raw_listbox()

    def _pick_open_file(self, var: tk.StringVar, title: str) -> None:
        p = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if p:
            var.set(p)

    def _pick_save_excel(self) -> None:
        p = filedialog.asksaveasfilename(
            title="Save Output Report / 选择报表保存路径",
            defaultextension=".xlsx",
            initialfile=os.path.basename(self.output_excel_var.get() or "QDHS_MT1_Report.xlsx"),
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if p:
            self.output_excel_var.set(p)

    def _pick_warning_dir(self) -> None:
        d = filedialog.askdirectory(title="Select Warning Letters Output Folder / 选择警告信输出文件夹")
        if d:
            self.warning_out_dir_var.set(d)

    def _open_folder_in_os(self, folder_path: str) -> None:
        if not folder_path or not os.path.exists(folder_path):
            messagebox.showwarning(
                "Folder Not Found / 目录不存在",
                f"Folder does not exist yet / 目录尚未创建:\n{folder_path}",
            )
            return
        if sys.platform == "darwin":
            subprocess.Popen(["open", folder_path])
        elif sys.platform == "win32":
            os.startfile(folder_path)  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", folder_path])

    def _open_report_folder(self) -> None:
        out_p = self.output_excel_var.get().strip()
        self._open_folder_in_os(os.path.dirname(out_p) if out_p else self._app_dir)

    def _open_warning_folder(self) -> None:
        self._open_folder_in_os(self.warning_out_dir_var.get().strip())

    # -------------------------------------------------------------------------
    # Step 1 Execution (Background Thread)
    # -------------------------------------------------------------------------
    def _start_generate_report(self) -> None:
        if self._busy:
            return
        if not self.raw_files:
            messagebox.showerror(
                "Missing Input / 缺少输入文件",
                "Please add at least one raw grade Excel file.\n请至少添加一份原始成绩 Excel 文件。",
            )
            return
        out_path = self.output_excel_var.get().strip()
        if not out_path:
            messagebox.showerror(
                "Missing Output Path / 缺少输出路径",
                "Please specify an output Excel file path.\n请指定输出 Excel 报表路径。",
            )
            return

        mode_code = self.mode_var.get()
        att_p = self.attendance_var.get().strip()
        mb_p = self.mb_id_var.get().strip()
        info_p = self.student_info_var.get().strip()
        cas_p = self.cas_var.get().strip()
        ee_p = self.ee_var.get().strip()
        dd_p = self.dwight_diploma_var.get().strip()
        raw_copy = list(self.raw_files)

        self._set_busy(True)
        self._append_log(self.log_step1, f"=== Starting Report Generation ({mode_code}) / 开始生成报表 ===")

        def worker() -> None:
            try:
                cb = lambda m: self._queue.put(("log1", m))
                if mode_code in ("G9", "G10"):
                    res = generate_g9_g10_report(
                        mode=mode_code,
                        raw_files=raw_copy,
                        attendance_path=att_p,
                        id_map_path=mb_p,
                        output_path=out_path,
                        student_info_path=info_p or None,
                        progress_cb=cb,
                    )
                else:
                    sub_mode = "IB" if mode_code == "G11_12_IB" else "AL"
                    res = generate_g11_g12_report(
                        raw_files=raw_copy,
                        output_path=out_path,
                        mode=sub_mode,
                        attendance_path=att_p,
                        mb_id_path=mb_p,
                        student_info_path=info_p,
                        cas_path=cas_p if sub_mode == "IB" else "",
                        ee_path=ee_p if sub_mode == "IB" else "",
                        dwight_diploma_path=dd_p if sub_mode == "IB" else "",
                        progress_cb=cb,
                    )
                self._queue.put(("report_done", res))
            except Exception as exc:
                self._queue.put(("error", f"Report generation failed / 报表生成失败:\n{exc}"))

        threading.Thread(target=worker, daemon=True).start()

    # -------------------------------------------------------------------------
    # Step 2: Warning List, PDF/JPG & Batch ManageBac
    # -------------------------------------------------------------------------
    def _load_warning_list(self, silent: bool = False) -> None:
        excel_p = self.warning_excel_var.get().strip()
        mb_p = self.warning_mb_var.get().strip()
        if not excel_p or not os.path.isfile(excel_p):
            if not silent:
                messagebox.showerror(
                    "File Not Found / 文件未找到",
                    "Please select a valid generated MT1 Report Excel file.\n请选择有效的成绩报告 Excel 文件。",
                )
            return

        try:
            students = load_warning_students(excel_p, mb_p or None)
        except Exception as exc:
            if not silent:
                messagebox.showerror("Load Failed / 加载失败", str(exc))
            return

        self.warning_students = students
        self.checked_map.clear()
        self._item_to_student.clear()
        self._mb_batch_offset = 0

        for iid in self.tree.get_children():
            self.tree.delete(iid)

        for st in students:
            skip_str = "" if st.attendance_skip is None else str(st.attendance_skip)
            iid = self.tree.insert(
                "",
                "end",
                values=(
                    "☑",
                    st.student_id,
                    st.name,
                    st.homeroom,
                    st.failed_count,
                    skip_str,
                    st.mb_id or "Missing",
                    st.failed_subjects_text,
                ),
            )
            self.checked_map[iid] = True
            self._item_to_student[iid] = st

        self._update_selection_indicators()
        self._append_log(
            self.log_step2,
            f"Loaded {len(students)} warning student(s) from {os.path.basename(excel_p)} / 已加载 {len(students)} 名预警学生。",
        )

    def _on_tree_click(self, event: tk.Event) -> None:
        region = self.tree.identify_region(event.x, event.y)
        if region not in ("cell", "tree"):
            return
        iid = self.tree.identify_row(event.y)
        if not iid or iid not in self.checked_map:
            return
        new_state = not self.checked_map[iid]
        self.checked_map[iid] = new_state
        vals = list(self.tree.item(iid, "values"))
        vals[0] = "☑" if new_state else "☐"
        self.tree.item(iid, values=vals)
        self._mb_batch_offset = 0
        self._update_selection_indicators()

    def _set_all_checked(self, checked: bool) -> None:
        mark = "☑" if checked else "☐"
        for iid in list(self.checked_map.keys()):
            self.checked_map[iid] = checked
            vals = list(self.tree.item(iid, "values"))
            vals[0] = mark
            self.tree.item(iid, values=vals)
        self._mb_batch_offset = 0
        self._update_selection_indicators()

    def _get_selected_students(self) -> List[WarningStudent]:
        return [
            self._item_to_student[iid]
            for iid in self.tree.get_children()
            if self.checked_map.get(iid, False)
        ]

    def _update_selection_indicators(self) -> None:
        total = len(self.warning_students)
        selected = self._get_selected_students()
        with_url = sum(1 for s in selected if s.managebac_url)
        self.sel_count_lbl.configure(
            text=f"Selected: {len(selected)} / {total}  |  已选：{len(selected)} / {total} 人  (Click row to toggle / 点击行切换勾选)"
        )
        self.mb_progress_lbl.configure(
            text=f"Batch Progress / 进度:\n{min(self._mb_batch_offset, with_url)} / {with_url}"
        )

    def _start_generate_letters(self) -> None:
        if self._busy:
            return
        selected = self._get_selected_students()
        if not selected:
            messagebox.showwarning(
                "No Students Selected / 未勾选学生",
                "Please select at least one student from the warning list.\n请至少勾选一名预警学生。",
            )
            return
        out_dir = self.warning_out_dir_var.get().strip()
        if not out_dir:
            messagebox.showerror(
                "Missing Output Folder / 缺少输出文件夹",
                "Please choose an output folder for warning letters.\n请选择警告信输出文件夹。",
            )
            return

        self._set_busy(True)
        self._append_log(
            self.log_step2,
            f"=== Generating {len(selected)} Warning Letters (PDF + JPG) / 开始批量生成 {len(selected)} 份警告信 ===",
        )

        def worker() -> None:
            try:
                res = batch_generate_warning_letters(
                    selected,
                    out_dir,
                    progress_cb=lambda m: self._queue.put(("log2", m)),
                )
                self._queue.put(("letters_done", res))
            except Exception as exc:
                self._queue.put(("error", f"Warning letter generation failed / 警告信生成失败:\n{exc}"))

        threading.Thread(target=worker, daemon=True).start()

    def _open_next_mb_batch(self) -> None:
        selected = self._get_selected_students()
        if not selected:
            messagebox.showwarning(
                "No Students Selected / 未勾选学生",
                "Please select at least one student first.\n请先勾选至少一名学生。",
            )
            return

        try:
            batch_sz = max(1, int(self.batch_size_var.get()))
        except (ValueError, tk.TclError):
            batch_sz = 5
            self.batch_size_var.set(5)

        with_url = [s for s in selected if s.managebac_url]
        if not with_url:
            messagebox.showwarning(
                "No ManageBac IDs / 无可用 MB ID",
                "None of the selected students have a matched MB System ID.\n"
                "所选学生均未匹配到 MB 系统 ID，请检查学生MB系统ID对照表。",
            )
            return

        if self._mb_batch_offset >= len(with_url):
            messagebox.showinfo(
                "All Opened / 已全部打开",
                f"All {len(with_url)} ManageBac links have been opened. Click 'Reset Batch' to start over.\n"
                f"全部 {len(with_url)} 位学生的 ManageBac 链接已打开，如需重开请点击“重置批次”。",
            )
            return

        res = open_managebac_batch(
            selected,
            start_index=self._mb_batch_offset,
            batch_size=batch_sz,
            log_cb=lambda m: self._append_log(self.log_step2, m),
        )
        self._mb_batch_offset = res["next_index"]
        self._update_selection_indicators()

        if res["missing_students"]:
            miss_names = ", ".join(f"{s.student_id}({s.name})" for s in res["missing_students"][:6])
            self._append_log(
                self.log_step2,
                f"Notice: {len(res['missing_students'])} student(s) missing MB ID / 缺少 MB ID: {miss_names}",
            )

    def _reset_mb_batch(self) -> None:
        self._mb_batch_offset = 0
        self._update_selection_indicators()
        self._append_log(self.log_step2, "ManageBac batch progress reset to 0 / 已重置 ManageBac 批次进度。")

    # -------------------------------------------------------------------------
    # Thread-Safe Queue Polling & Logging
    # -------------------------------------------------------------------------
    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for btn in (self.btn_generate_report, self.btn_gen_letters, self.btn_load_warning):
            btn.set_enabled(not busy)

    def _append_log(self, widget: scrolledtext.ScrolledText, msg: str) -> None:
        widget.configure(state="normal")
        widget.insert(tk.END, msg + "\n")
        widget.see(tk.END)
        widget.configure(state="disabled")

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log1":
                    self._append_log(self.log_step1, str(payload))
                elif kind == "log2":
                    self._append_log(self.log_step2, str(payload))
                elif kind == "report_done":
                    self._set_busy(False)
                    res: Dict[str, Any] = payload
                    out_p = res["output_path"]
                    self.warning_excel_var.set(out_p)
                    if self.mb_id_var.get().strip():
                        self.warning_mb_var.set(self.mb_id_var.get().strip())
                    self._load_warning_list(silent=True)
                    messagebox.showinfo(
                        "Report Generated / 报表生成成功",
                        f"MT1 Report generated successfully!\n成绩及出勤报告已生成！\n\n"
                        f"Students / 学生数: {res.get('student_count', 0)}\n"
                        f"Subjects / 科目数: {res.get('subject_count', 0)}\n"
                        f"Warning Students / 预警人数: {res.get('warning_count', 0)}\n\n"
                        f"Output / 文件路径:\n{out_p}\n\n"
                        f"Warning list has been loaded into Step 2.\n预警名单已自动载入“第二步”。",
                    )
                elif kind == "letters_done":
                    self._set_busy(False)
                    res = payload
                    messagebox.showinfo(
                        "Warning Letters Generated / 警告信生成成功",
                        f"Generated {res['count']} PDF + JPG warning letter pairs!\n"
                        f"已成功生成 {res['count']} 份 PDF + JPG 警告信！\n\n"
                        f"Folder / 输出目录:\n{res['output_dir']}",
                    )
                elif kind == "error":
                    self._set_busy(False)
                    self._append_log(self.log_step1, f"[ERROR] {payload}")
                    messagebox.showerror("Error / 错误", str(payload))
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main() -> None:
    app = QDHSGradeReportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
