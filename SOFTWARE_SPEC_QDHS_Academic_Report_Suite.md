# QDHS 七德高中成绩报告生成器 — 完整软件需求与架构规格说明书
# QDHS Grade Report Generator (`QDHS_Academic_Report_Suite`) — Complete Software Requirements & Architecture Specification

> **文档定位**：本文档为可直接提供给其他大语言模型或开发者的**端到端可重建规格说明书**，完整定义了覆盖 **G9 (House X)、G10、G11/G12 (IB & AL)** 的成绩及出勤报告生成、学业/考勤警告信批量生成（PDF + JPG）、ManageBac Behavior Note 分批打开功能，以及双语红蓝主题桌面 UI 的全部交互与防回归约束。

---

## 1. 系统定位、目录结构与运行环境

### 1.1 适用范围与核心原则
- **适用学校与学期**：上海七宝德怀特高级中学（QDHS）2026-2027 学年第一学期期中考试（MT1），统一覆盖 **G9 (House X)**、**G10**、**G11 / G12 (IB)** 与 **G11 / G12 (AL)**。
- **核心工作流（两步式）**：
  1. **Step 1（第一步）**：根据所选年级与课程模式，读取原始成绩表与辅助数据，生成包含 `MT1 Summary`、`MT1 Overview`、`Warning` 三个工作表的格式化 `.xlsx` 报表。
  2. **Step 2（第二步）**：自动（或手动）载入 `Warning` 预警名单，允许勾选目标学生，批量生成中英双语学术警告信（`.pdf` + `.jpg`），并按批次在浏览器中打开对应学生的 ManageBac Behavior Note 创建页面。
- **只读原则（硬性约束）**：所有输入 `.xlsx` / `.xlsm` 文件仅以只读方式解析（`read_only=True, data_only=True`），严禁修改任何原始输入文件。
- **动态表头原则（硬性约束）**：所有输入表格（原始成绩表、考勤汇总表、学生 MB ID 对照表、学生信息表、CAS/EE 独立表、`Warning` 表）严禁写死列号，全部通过动态扫描表头关键词定位列索引。

### 1.2 目录结构
```text
QDHS_Academic_Report_Suite/
├── main.py                                     # 程序入口，调用 app_ui.main()
├── common_io.py                                # 公共只读 Excel 解析、动态表头定位、QDHS 校色常量
├── g9_g10_engine.py                            # G9 (House X) & G10 三表生成引擎（无 UI 依赖）
├── g11_g12_engine.py                           # G11 & G12 (IB / AL) 三表生成引擎（无 UI 依赖）
├── warning_letter.py                           # 统一警告信解析（兼容双格式）、PDF+JPG 渲染、ManageBac 链接模块（无 UI 依赖）
├── app_ui.py                                   # 双语红蓝主题两步式 Tkinter 桌面应用
├── test_suite.py                               # unittest 单元测试与端到端样本验证套件
├── assets/
│   └── qdhs_logo.png                           # 警告信顶部学校 Logo（814x164 RGBA）
├── 启动QDHS成绩报告生成器.command                  # macOS 双击启动脚本
└── SOFTWARE_SPEC_QDHS_Academic_Report_Suite.md # 本软件规格说明书
```

### 1.3 技术栈与依赖约束
- **Python 版本**：兼容 Python 3.9 – 3.14。
- **第三方依赖**：
  - `openpyxl`：只读解析原始 Excel，生成带样式与动态公式的 3-Sheet `.xlsx`。
  - `pymupdf`：直接渲染中英 HTML 生成 A4 PDF 并导出 200 DPI JPG。**硬性要求：统一使用 `import pymupdf`，禁止使用已弃用的 `import fitz`**（避免 `FutureWarning`）。
  - `Pillow`（`PIL`）：高分辨率 JPG 导出及多页 PDF 纵向拼接长图。
  - 标准库：`tkinter`、`queue`、`threading`、`webbrowser`、`html`、`dataclasses`、`statistics`。
  - **无 `reportlab` 依赖**，保证跨 Python 版本兼容性。

---

## 2. 品牌视觉与 UI 交互规范

### 2.1 官方色板（QDHS Pantone Colors）
- **校色红（Red Pantone 186EC）**：`#C8102E`（悬停 `#A00D24`，Excel 十六进制 `C8102E`）
- **校色蓝（Blue Pantone 281EC）**：`#00205B`（悬停 `#001640`，Excel 十六进制 `00205B`）
- **工作区底色**：浅灰 `#F4F6F9` 配白色卡片 `#FFFFFF` 与细边框 `#D0D7DE`，保证长时间核对数据清晰不刺眼。

### 2.2 标题栏与双语按钮规范
- **软件窗口标题**：`QDHS Grade Report Generator | QDHS七德高中成绩报告生成器`
- **顶部主标题栏（Header Banner）**：采用 `#C8102E`（Red Pantone 186EC）红色背景、单行紧凑高度 `52px`，居中显示白色粗体双语标题（不设副标题行以节约垂直空间），下方衔接 4px `#00205B` 蓝色装饰条。
- **全界面双语要求**：所有标题、提示、表头、状态标签、弹窗均采用英中双语。
- **按钮排版硬性规范**：全界面**所有按钮**采用自定义 macOS 兼容组件 `BilingualButton(tk.Frame)`，严格实现**英文在上、中文在下**双行排版（解决 macOS Aqua 忽略原生 `tk.Button` 背景色及多行排版受限的问题），并支持通过 `width_px` 固定等宽：
  - 主操作按钮：蓝底白字（`#00205B`）或红底白字（`#C8102E`）。
  - 辅助按钮：白底蓝框蓝字（`#FFFFFF` / `#00205B`）。

### 2.3 两步式导航标签（Step Tabs）与垂直滚动容器规范
顶层两步导航采用自定义双行标签栏（替代静态 `ttk.Notebook` 标签），严格满足以下视觉状态切换：
- **标签文案（中文置于英文下方）**：
  - Step 1：上行 `Step 1: Grade & Attendance Report`，下行 `第一步：生成成绩及出勤报告`
  - Step 2：上行 `Step 2: Warning Letter & ManageBac`，下行 `第二步：警告信与 ManageBac`
- **选中状态（Active）**：尺寸放大（内边距 `padx=26, pady=10`）、字体增大且**加粗**（英文 `12pt bold` / 中文 `10pt bold`）、深蓝底 `#00205B` 配纯白/浅粉白字，底部附 4px `#C8102E` 红色高亮指示条。
- **未选中状态（Inactive）**：尺寸缩小（内边距 `padx=14, pady=5`）、字体缩小且不加粗（英文 `9pt normal` / 中文 `8pt normal`）、文字变虚（低饱和灰蓝 `#8A94A6` / `#9AA3B2` 配浅灰底 `#E6EAF0`），底部对齐使选中标签视觉上明显突出。
- **垂直滚动侧边栏**：Step 1 与 Step 2 内容区均置于 `Canvas + ttk.Scrollbar` 垂直滚动容器内，支持鼠标滚轮滚动（且不抢占内部 `Text`/`Listbox`/`Treeview` 的滚轮事件），切换 Step 时自动回顶；最小窗口尺寸设为 `980x680`（启动高度 `min(1020, screen_h - 60)`），保证小屏幕也能完整浏览与操作。

### 2.4 单层 4 选 1 模式切换、路径右对齐与按钮等宽
第一步顶部提供单层 4 选 1 卡片（操作一步到位，无需二级菜单）：
1. `G9 (House X)` / `九年级 (House X)`
2. `G10` / `十年级`
3. `G11 / G12 (IB)` / `高二/高三 IB`
4. `G11 / G12 (AL)` / `高二/高三 AL`

切换模式时界面自动联动：
- 仅在 `G11 / G12 (IB)` 模式下显示可选的 `CAS Export`、`EE Export` 与 `Dwight Diploma List` 文件输入行，其他模式自动隐藏。
- 自动更新模式说明提示、默认输出文件名，并在父级/同级数据目录（含 `NewData.Nov.24`）中智能探测默认输入文件（如存在）。
- 警告信输出目录按模式隔离至 `Generated_Warning_Letters/{G9_HouseX|G10|G11_G12_IB|G11_G12_AL}`，防止跨年级文件混杂。
- **区域 3 对齐与路径优先从右显示**：区域 3（`3. Auxiliary Files & Output Report Path`）所有标签定格左对齐、文件路径输入框同列对齐、右侧 `Browse` / `Save As（选择路径）` 按钮全部右对齐；Step 1 与 Step 2 的所有路径 `Entry` 及原始文件 `Listbox` 在赋值、失焦或窗口调整时自动右滚（`xview_moveto(1.0)`），优先展示末级文件夹与文件名。Step 2 的 `Open Folder` 按钮宽度与上方 `Choose Folder` 按钮保持等宽（`width_px=118`）。
- **macOS 启动脚本自适应 Python 环境**：`启动QDHS成绩报告生成器.command` 依次探测 `/usr/bin/python3`、系统 `PATH`、Homebrew 等 Python 3，自动选用同时具备 `openpyxl`、`pymupdf`、`PIL`、`tkinter` 的解释器启动 `main.py`。

### 2.5 macOS Tkinter 线程安全规范
- 报表生成与批量警告信生成均在后台 `threading.Thread(..., daemon=True)` 中执行。
- 后台线程严禁直接调用 Tkinter 控件或 `root.after()`；统一向 `queue.Queue` 投递事件，由主线程 `root.after(100, self._poll_queue)` 轮询刷新日志、恢复按钮状态并弹出完成提示。

---

## 3. 输入数据契约与公共解析规则（`common_io.py`）

| 输入类型 | 适用模式 | 默认/样本文件名 | 动态匹配关键表头 | 清洗与边界规则 |
| :--- | :--- | :--- | :--- | :--- |
| **原始成绩表（多选）** | 全部 | G9: `Updated House X Data .xlsx`<br>G10: `Updated HS G10.xlsx` + `G10 IB .xlsx`<br>G11/12: `G11/G12 IB/AL/HS.xlsx` | `Student ID`, `Subject`, `Class Name`, `Mid-term Grade*`, `Grade Level`, `Preferred Name` | 多文件不限添加顺序；G11/G12 自动识别文件类型（IB/AL/HS）及年级（11/12）；忽略 `Workshop` |
| **考勤汇总表** | 全部 | `All_Attendance_Summary_Generated.xlsx` | `Student ID`, `Tier`, `Total Skip` | `Perfect Attendance` 与 `Excellent` 均映射为 `Excellent`，其余原样保留 |
| **行政班与 MB ID 对照表** | 全部 | `学生MB系统ID对照表.xlsx` | `SIMPO学号`, `行政班`, `MB系统ID`, `姓名` | 跨 Sheet 扫描；浮点 MB ID 去 `.0` 转字符串；输出表强制文本格式 `@` |
| **学生信息表（邮箱）** | 全部 | `Class of 2030/2029/2028/2027 ...xlsx` | `Student ID`, `E-mail Address`, `Parent 1 E-mail address`, `Parent 2 E-mail address` | 支持表头位于前 15 行任意行（`E-mail Address` 可选） |
| **CAS 独立表（可选）** | 仅 G11/G12 IB | `G11/G12 CAS_Export.xlsx` | Sheet `Summary`，列 `Student ID`, `Status` | 可留空；空值不标红、不计入 Fail |
| **EE 独立表（可选）** | 仅 G11/G12 IB | `G11/G12 EE_export.xlsx` | Sheet `Reflections`，列 `Student ID`, `Progress` | 可留空；覆盖 IB 原表中 EE 为 N/A 的记录；`To Be Determined` 不标红、不计入 Fail |
| **Dwight Diploma 名单（可选）** | 仅 G11/G12 IB | `G11 Dwight Diploma.xlsx` / `G12 Dwight Diploma.xlsx` | `Student ID`, `Other Name` / `Status` / `Program` | 标记 `Dwight Diploma` / `DD` 的学生标为 `DD`，其余 IB 学生标为 `DP` |

---

## 4. 各年级核心业务规则

### 4.1 G9 (House X) 与 G10 规则（`g9_g10_engine.py`）
- **G9 (House X)**：
  - 完全排除 5 门 PBL 课程（`HS PBL 1`–`5` / `Project Based Learning`）及 `Workshop`、`Electives`。
  - 总分列包含 `MT1 Total 总计/56`（1–7 分制科目 `SUM`）与 `MT1 Total 总计/42`（`=(42/56)*总计/56`）。
  - 标红与不合格科目数：`1–7` 分制 `< 4`、`HS SPARK 升学指导` 为 `"Concern"`、`HS Physical Education 体育` `< 65` 均标红（`#FF9999`）并计入 `Number of Failed Subjects`。
- **G10**：
  - 按 `Grade Level == "Grade 10"` 过滤，排除 `Electives`、`Workshop`、`College Counseling (HS SPARK)`；`Academic English` 在 `Mid-term Grade` 为 N/A 时回退读取 `Mid-Term Progress Average`。
  - 总分列 `MT1 Total 总计` 仅汇总前 5 类（Groups 1–5，6 门课），**排除 `6-HS` 艺术类、中方课程（`Geography`/`History`/`Politics`）与 `PE`**。
  - 不合格科目数与标红：1–7 分制学科（**含 `6-HS` 艺术类**）`< 4`、中方课程 `< 65`、`PE < 65` 均标红并计入 `Number of Failed Subjects`。
- **考勤缺勤**：`attendance skip >= 9` 标红，**绝不计入 `Number of Failed Subjects`**，仅作为 `Warning` 筛选条件之一。

### 4.2 G11 与 G12 (IB / AL) 规则（`g11_g12_engine.py`）
- **学生基准名单**：仅以当前模式主课程文件（IB 或 AL）中出现的学生为基准，仅在 HS 表中出现的学生不纳入。
- **科目消歧**：IB 中 `Chinese`/`English` 拆分为 `Chinese A Lit` / `Chinese A L&L` / `Chinese B` / `English A L&L` / `English B` + `HL/SL`；IB `History`（1–7）与 HS `History`（0–100）通过内部键严格隔离；AL 课程前缀 `AL `；HS 中方课程前缀 `HS `，AL 班中文/英文分别识别为 `HS AL Chinese`（1–7）与 `HS AL English`（A*–E）。
- **成绩评估与标红（`evaluate_grade`）**：
  - `IB_DP` / `HS_AL_IB`：数值 `< 4` → Fail（红底 `#FFC7CE`、红字 `#9C0006`）。
  - `IB_CORE (TOK)`：转大写，非 `A`/`B`/`C`（或含 `Concern`）→ Fail。
  - `IB_CORE (EE / CAS)`：含 `Concern` 或等级 `< C` → Fail；`To Be Determined`、`N/A`、空值、`Excellent`、`On-track` 均**不标红、不计 Fail**。
  - `AL` / `HS_AL_ENG`：转大写，非 `A*`/`A`/`B`/`C` → Fail。
  - `HS_CN`（地理/历史/政治/体育）：`< 65` → Fail。
  - `attendance skip >= 9`：标红，**不计入 `Number of Failed Subject`**。
- **IB Dwight Diploma (`DD`) vs IB DP (`DP`) 与 TOK+EE 加分矩阵**：
  - `STATUS` 列：`Dwight Diploma` 名单内学生标 `"DD"`，其余 IB 学生标 `"DP"`；AL 学生标 `"AL"`。
  - `DP` 学生：依据官方 `TOK + EE` 矩阵计算综合加分（0–3 分：`A+A/B=3, A+C/D=2, B+A=3, B+B/C=2, B+D=1, C+A/B=2, C+C=1, D+A=2, D+B=1`，含 `E` 或非等级状态记 0 分），加总至 `G{grade} MT1 Total (DP Only)总计`（Col J）与 `G{grade} MT1 Total 总计`（Col K）。
  - `DD` 学生：不计算 TOK+EE 综合分，`G{grade} MT1 Total (DP Only)总计`（Col J）**留空**，Col K 仅汇总 1–7 分制科目基础总分。
- **AL 7–1 折算总分与汇总分布**：
  - AL 模式在 `MT1 Summary` 中**删去** `Candidate Session Number`、`Personal Code` 与 `G{grade} MT1 Total (DP Only)总计` 三列（保留 8 个固定前置列 Col A–H，冻结窗格 `D2`）。
  - `G{grade} MT1 Total 总计`（Col H）按 `A*=7, A=6, B=5, C=4, D=3, E=2, U/F/G=1` 对 `AL`、`HS AL English`、`HS AL Chinese`（保持 1–7 数值）求和（不含中方 `HS_CN` 课程）。
  - `MT1 Summary` 下方 `Subject Average 学科平均分` 行对所有 AL / HS AL 学科列采用 7 分制加权平均公式：`=IF(SUM(col_r7:col_r1)>0, (7*col_r7+6*col_r6+5*col_r5+4*col_r4+3*col_r3+2*col_r2+1*col_r1)/SUM(col_r7:col_r1), "")`，数值总分与中方课程使用 `AVERAGE(...)`。
  - 下方附 7–1 分值分布统计行（`7分 (A*)` 至 `1分 (U/F/G)`，使用 `COUNTIF(...,"A~*")` 转义通配符 `*`）及 Col H 的 `Median`/`Max`/`Min`。

---

## 5. 输出报表、警告信与 ManageBac 规范

### 5.1 Excel 3-Sheet 输出规范
1. **`MT1 Summary`**：
   - G9/G10：Row 1 自 Col D 起竖向排版；学生行白/浅绿交替；底部含动态公式统计人数、各科/各行政班平均分、中位数/最高/最低分与 7–1 分布；最右两列按行政班输出分数段分布。
   - G11/G12：IB 模式固定前置列 Col A–K（含 `STATUS`、`DP Only` 与 `Total`，冻结 `F2`），AL 模式固定前置列 Col A–H（已移除 3 个 DP 专用列，冻结 `D2`）；动态学科列、尾部考勤与不合格数、底部 `COUNTA`/`AVERAGE`/7 分制加权平均与 7–1 分布统计行、最右侧各行政班分布（IB 分数段 / AL `A*–U` 人次）。
2. **`MT1 Overview`**：
   - 统一采用 `#C8102E` 红底白字大表头、`#00205B` 蓝底白字列标题、白/浅绿斑马纹。
   - G11/G12 Row 1 主表头后无空行直接衔接 Row 2；IB 模式包含 `Chinese National Course 中方课程`（Col A–F）、`IBDP Core 核心课程`（Col A–G，**两行合并蓝底子表头，其中 `Concern 待改进` 字号强制为 `8`**）与 `Overall Summary 总评 (Total Score)`。
3. **`Warning`**：
   - 统一筛选条件：**`Number of Failed Subjects >= 3` OR `attendance skip >= 9`**。
   - 包含学生 ID、姓名、行政班、学生邮箱、家长 1/2 邮箱、不合格科目数、缺勤次数及不合格科目明细。

### 5.2 统一警告信生成与双格式解析（`warning_letter.py`）
- **双格式兼容解析**：`load_warning_students(excel_path, id_map_path)` 自动识别并兼容 G9/G10（科目名为列头、单元格为分数）与 G11/G12（`Failed Subject 1..N` 列存 `科目: 成绩`）两种 `Warning` 表结构。
- **警告信内容与排版**：
  - 标题上方居中放置学校 Logo（`assets/qdhs_logo.png`，嵌入 Base64 PNG）；标题下方左对齐加入 `Dear {Student Name},`。
  - 上半部分完整保留中英双语政策声明（英文准确措辞：`"as you have met one or more of the criteria listed below"`），所有动态字段均经过 `html.escape()` 转义。
  - 下半部分 `Specific Warning Details / 具体警告内容`：列出学生姓名/学号/行政班；若 `failed_count >= 3` 逐项列出不合格科目及成绩（`#C8102E` 标红）；若 `attendance_skip >= 9` 列出缺勤次数（`#C8102E` 标红）；两者均满足则同时列出。
  - 结尾右下角右对齐署名为：`上海七宝德怀特高级中学` / `Shanghai Qibao Dwight High School` / `2026年11月23日` / `23 November 2026`。
  - 使用 `pymupdf`（Times New Roman + 宋体）渲染 A4 PDF（内容超出一页时自动紧凑缩放，若多页则用 Pillow 纵向拼接），同步输出 200 DPI `.jpg`。
  - 文件命名格式：**`{Homeroom}_{Student ID}_{Student Name}.pdf`** 与同名 **`.jpg`**。

### 5.3 ManageBac Behavior Note 分批打开
- 域名配置：`MANAGEBAC_HOST = "qibaodwight.managebac.cn"`
- 链接模板：`https://{MANAGEBAC_HOST}/teacher/users/{MB系统ID}/behavior_notes/new`
- GUI 支持对勾选学生按自定义批次大小（默认每批 5 人）分批在浏览器新标签页打开，实时展示 `Batch Progress / 进度: X / Y` 并提供 `Reset Batch / 重置批次` 按钮。

---

## 6. 核心模块公开接口（API 契约）

```python
# common_io.py
find_header_row(ws, required_keywords: List[str], max_scan: int = 15, require_all: bool = False) -> Tuple[int, Dict[str, int]]
resolve_col(header_map: Dict[str, int], candidates: List[str], required: bool = True) -> Optional[int]
load_attendance_map(attendance_path: Optional[str], required: bool = False) -> Dict[str, Tuple[str, Optional[int]]]
load_student_id_map(id_map_path: Optional[str], required: bool = False) -> Dict[str, Tuple[str, str, str]]
load_student_info_map(info_path: Optional[str], required: bool = False) -> Dict[str, Tuple[str, str, str, str]]
load_dwight_diploma_ids(dd_path: Optional[str]) -> set[str]

# g9_g10_engine.py
generate_g9_g10_report(
    mode: str,                                  # "G9" 或 "G10"
    raw_files: List[str],                       # 不限顺序
    attendance_path: str,
    id_map_path: str,
    output_path: str,
    student_info_path: Optional[str] = None,
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]

# g11_g12_engine.py
compute_tok_ee_bonus(tok_val: Any, ee_val: Any) -> int
convert_al_grade_to_score(val: Any) -> Optional[float]
evaluate_grade(raw_val: Any, category: str, subject_key: str) -> Tuple[Any, bool]
generate_g11_g12_report(
    raw_files: List[str],
    output_path: str,
    mode: str = "IB",                           # "IB" 或 "AL"
    attendance_path: str = "",
    mb_id_path: str = "",
    student_info_path: str = "",
    cas_path: str = "",                         # 可选
    ee_path: str = "",                          # 可选
    dwight_diploma_path: str = "",              # 可选（仅 IB 模式）
    progress_cb: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]

# warning_letter.py
load_warning_students(excel_path: str, id_map_path: Optional[str] = None) -> List[WarningStudent]
build_managebac_behavior_note_url(mb_id: str) -> str
generate_warning_letter_pdf_and_jpg(student: WarningStudent, output_dir: str, dpi: int = 200) -> Tuple[str, str]
batch_generate_warning_letters(students: List[WarningStudent], output_dir: str, progress_cb=None) -> Dict[str, Any]
open_managebac_batch(students: List[WarningStudent], start_index: int = 0, batch_size: int = 5, log_cb=None) -> Dict[str, Any]
```

---

## 7. 防回归检查清单（供二次开发/跨模型复用）
1. 严禁修改任何原始输入 Excel 文件；严禁使用 `import fitz`，统一 `import pymupdf`。
2. 保持 `MT1 Overview`（IB 模式）中 `IBDP Core` 两行合并表头及 `Concern 待改进` **8 号字体**。
3. 保持 `EE` 的 `To Be Determined` 与 `CAS` 空值**不标红、不计入 Fail**；保持 `CAS`/`EE` 文件输入为独立可选项。
4. 保持 `attendance skip >= 9` **标红但不计入 `Number of Failed Subject(s)`**，预警阈值统一为 `不合格 >= 3 OR 缺勤 >= 9`。
5. 保持 UI 顶部标题栏 `#C8102E` 红底白字、按钮全双语（英文在上、中文在下）与 `#00205B`/`#C8102E` 校色，以及 Step 1 / Step 2 标签“选中放大加粗、未选中缩小变虚、中文在英文下方”的动态交互效果。
