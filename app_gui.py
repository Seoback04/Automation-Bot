from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from typing import Any

from config import (
    APPLICATION_PIPELINE_PATH,
    APP_STATE_PATH,
    PROFILE_PATH,
    REPORTS_DIR,
    RESUMES_DIR,
    RUN_HISTORY_PATH,
    SCREENSHOTS_DIR,
)
from core.assistant_brain import AssistantBrain, SearchDirection
from core.browser import BrowserSession
from core.job_search import JobSearchEngine
from core.profile_store import ProfileStore
from flows.easy_apply import EasyApplyBot, EasyApplyFlow
from flows.external_apply import ExternalApplyFlow
from xml.sax.saxutils import escape
import zipfile


DASHBOARD_COLUMNS = [
    ("timestamp", "Timestamp", 150),
    ("role", "Role", 180),
    ("location", "Location", 120),
    ("source", "Source", 120),
    ("result", "Result", 140),
    ("filled_count", "Filled", 70),
    ("detected_count", "Detected", 80),
    ("url", "Job URL", 280),
    ("summary", "Summary", 260),
    ("screenshot", "Screenshot", 260),
]

PIPELINE_COLUMNS = [
    ("timestamp", "Added", 150),
    ("company", "Company", 150),
    ("role", "Role", 150),
    ("stage", "Stage", 120),
    ("priority", "Priority", 90),
]


def _clean_user_path(raw: str) -> str:
    text = (raw or "").strip()
    if len(text) >= 2 and text[0] == text[-1] and text[0] in {'"', "'"}:
        text = text[1:-1].strip()
    text = text.strip('"').strip("'")
    text = os.path.expandvars(text)
    if text.lower().startswith("file://"):
        text = text[7:]
    return text


def _resolve_existing_resume_path(path_text: str) -> Path | None:
    cleaned = _clean_user_path(path_text)
    if not cleaned:
        return None

    primary = Path(cleaned).expanduser()
    candidates = [primary]
    if not primary.is_absolute():
        candidates.append((Path.cwd() / primary).resolve())
    candidates.append((RESUMES_DIR / primary.name).resolve())

    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return candidate.resolve()
    return None


def _write_simple_xlsx(file_path: Path, sheet_name: str, headers: list[str], rows: list[list[Any]]) -> None:
    safe_sheet_name = (sheet_name or "Sheet1")[:31]

    def inline_cell(value: Any) -> str:
        text = "" if value is None else str(value)
        return (
            '<c t="inlineStr"><is><t xml:space="preserve">'
            f"{escape(text)}"
            "</t></is></c>"
        )

    row_xml: list[str] = []
    for row_index, row in enumerate([headers, *rows], start=1):
        cells = "".join(inline_cell(value) for value in row)
        row_xml.append(f'<row r="{row_index}">{cells}</row>')

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(row_xml)}</sheetData>"
        "</worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        "<sheets>"
        f'<sheet name="{escape(safe_sheet_name)}" sheetId="1" r:id="rId1"/>'
        "</sheets>"
        "</workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Aptos"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    with zipfile.ZipFile(file_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", root_rels)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        workbook.writestr("xl/styles.xml", styles_xml)


class JobBotApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Job Bot v2 - Professional Dashboard")
        self.root.geometry("1280x840")
        self.root.minsize(1080, 720)

        self.profile_var = tk.StringVar(value=str(PROFILE_PATH))
        self.resume_var = tk.StringVar(value="")
        self.role_var = tk.StringVar(value="")
        self.location_var = tk.StringVar(value="")
        self.source_var = tk.StringVar(value="1")
        self.url_var = tk.StringVar(value="")
        self.headless_var = tk.BooleanVar(value=False)
        self.url_entry: ttk.Entry | None = None

        self.dashboard_tree: ttk.Treeview | None = None
        self.dashboard_editor: ttk.Entry | ttk.Combobox | None = None
        self.dashboard_status_var = tk.StringVar(value="Dashboard ready")
        self.dashboard_total_var = tk.StringVar(value="0")
        self.dashboard_pending_var = tk.StringVar(value="0")
        self.dashboard_pipeline_var = tk.StringVar(value="0")
        self.dashboard_success_var = tk.StringVar(value="0")
        self.history_tree: ttk.Treeview | None = None
        self.pipeline_tree: ttk.Treeview | None = None
        self.assistant_output: tk.Text | None = None
        self.assistant_input: tk.Text | None = None
        self.assistant_results_tree: ttk.Treeview | None = None
        self.assistant_matches: list[Any] = []
        self.assistant_strategy_text: tk.Text | None = None
        self.assistant_sequence_var = tk.StringVar(value="")
        self.assistant_provider_var = tk.StringVar(value="")
        self.assistant_brain = AssistantBrain()
        self.search_direction = SearchDirection()

        self._apply_theme()

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=14, pady=14)

        self._create_dashboard_tab()
        self._create_apply_tab()
        self._create_assistant_tab()
        self._create_history_tab()
        self._create_settings_tab()
        self._create_reports_tab()
        self._create_pipeline_tab()

        self._load_defaults()
        self._refresh_all_views()

    def _apply_theme(self) -> None:
        colors = {
            "bg": "#F4F7FB",
            "surface": "#FFFFFF",
            "surface_alt": "#EAF1FB",
            "text": "#1F2A44",
            "muted": "#5F6C86",
            "accent": "#0F766E",
            "accent_alt": "#F97316",
            "highlight": "#2563EB",
            "border": "#D7E2F0",
            "success": "#0F9D58",
            "chat_bg": "#0F172A",
            "chat_surface": "#111827",
            "chat_surface_alt": "#1F2937",
            "chat_text": "#E5EEF9",
            "chat_muted": "#93A4BD",
            "chat_user": "#1D4ED8",
            "chat_assistant": "#1E293B",
        }

        self.root.configure(bg=colors["bg"])
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".", background=colors["bg"], foreground=colors["text"], fieldbackground=colors["surface"])
        style.configure("TFrame", background=colors["bg"])
        style.configure("TLabelframe", background=colors["surface"], bordercolor=colors["border"], relief="solid")
        style.configure("TLabelframe.Label", background=colors["surface"], foreground=colors["text"], font=("Segoe UI", 11, "bold"))
        style.configure("TLabel", background=colors["bg"], foreground=colors["text"], font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=colors["bg"], foreground=colors["muted"], font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=colors["bg"], foreground=colors["text"], font=("Segoe UI Semibold", 20, "bold"))
        style.configure("Hero.TLabel", background=colors["bg"], foreground=colors["highlight"], font=("Segoe UI Semibold", 11, "bold"))
        style.configure("Section.TLabelframe", background=colors["surface"], bordercolor=colors["border"], padding=12)
        style.configure("TNotebook", background=colors["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", padding=(18, 10), background="#DCE7F7", foreground=colors["text"], font=("Segoe UI", 10, "bold"))
        style.map("TNotebook.Tab", background=[("selected", colors["surface"]), ("active", colors["surface_alt"])])
        style.configure("TButton", background=colors["accent"], foreground="#FFFFFF", borderwidth=0, focusthickness=0, padding=(12, 8), font=("Segoe UI", 10, "bold"))
        style.map("TButton", background=[("active", "#0B5E58"), ("disabled", "#A9B8C9")], foreground=[("disabled", "#F8FAFC")])
        style.configure("Secondary.TButton", background=colors["accent_alt"], foreground="#FFFFFF")
        style.map("Secondary.TButton", background=[("active", "#EA580C")])
        style.configure("Treeview", background=colors["surface"], foreground=colors["text"], fieldbackground=colors["surface"], rowheight=28, bordercolor=colors["border"])
        style.configure("Treeview.Heading", background="#D9E7FB", foreground=colors["text"], font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#CFE3FF")], foreground=[("selected", colors["text"])])
        style.configure("TEntry", fieldbackground=colors["surface"], bordercolor=colors["border"], insertcolor=colors["text"], padding=6)
        style.configure("TCombobox", fieldbackground=colors["surface"], bordercolor=colors["border"], padding=4)
        style.configure("TCheckbutton", background=colors["surface"], foreground=colors["text"])
        style.configure("TRadiobutton", background=colors["surface"], foreground=colors["text"])
        style.configure("Horizontal.TProgressbar", background=colors["highlight"], troughcolor="#DDE7F5", bordercolor=colors["border"], lightcolor=colors["highlight"], darkcolor=colors["highlight"])

        self.colors = colors

    def _create_dashboard_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Dashboard")
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        header = ttk.Frame(tab)
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="Job Application Dashboard", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Track, edit, and export your applications like a spreadsheet.", style="Hero.TLabel").grid(row=1, column=0, sticky="w", pady=(4, 0))

        stats = ttk.Frame(tab)
        stats.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 10))
        for index in range(4):
            stats.columnconfigure(index, weight=1)

        self._create_stat_card(stats, 0, "Applications", self.dashboard_total_var, self.colors["highlight"])
        self._create_stat_card(stats, 1, "Pending", self.dashboard_pending_var, self.colors["accent_alt"])
        self._create_stat_card(stats, 2, "Submitted", self.dashboard_success_var, self.colors["success"])
        self._create_stat_card(stats, 3, "Pipeline Jobs", self.dashboard_pipeline_var, self.colors["accent"])

        sheet_frame = ttk.LabelFrame(tab, text="Dashboard Sheet", padding=12)
        sheet_frame.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 10))
        sheet_frame.columnconfigure(0, weight=1)
        sheet_frame.rowconfigure(0, weight=1)

        tree_frame = ttk.Frame(sheet_frame)
        tree_frame.grid(row=0, column=0, sticky="nsew")
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.dashboard_tree = ttk.Treeview(tree_frame, columns=[name for name, _, _ in DASHBOARD_COLUMNS], show="headings")
        for name, label, width in DASHBOARD_COLUMNS:
            self.dashboard_tree.heading(name, text=label)
            self.dashboard_tree.column(name, width=width, stretch=(name in {"role", "url", "summary", "screenshot"}))

        y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.dashboard_tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.dashboard_tree.xview)
        self.dashboard_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.dashboard_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        self.dashboard_tree.bind("<Double-1>", self._begin_dashboard_edit)

        actions = ttk.Frame(sheet_frame)
        actions.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(actions, text="Add Row", command=self._add_dashboard_entry).pack(side=tk.LEFT)
        ttk.Button(actions, text="Delete Selected", command=self._delete_dashboard_entry).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(actions, text="Save Changes", command=self._save_dashboard_changes).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(actions, text="Refresh", command=self._refresh_all_views).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(actions, text="Export to Excel", style="Secondary.TButton", command=self._export_dashboard_to_excel).pack(side=tk.RIGHT)

        footer = ttk.Frame(tab)
        footer.grid(row=3, column=0, sticky="ew", padx=18, pady=(0, 18))
        ttk.Label(footer, textvariable=self.dashboard_status_var, style="Muted.TLabel").pack(anchor="w")

    def _create_stat_card(self, parent: ttk.Frame, column: int, title: str, value_var: tk.StringVar, accent: str) -> None:
        card = tk.Frame(parent, bg=self.colors["surface"], highlightbackground=self.colors["border"], highlightthickness=1)
        card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 8, 0))
        tk.Frame(card, bg=accent, height=6).pack(fill=tk.X)
        body = tk.Frame(card, bg=self.colors["surface"], padx=14, pady=12)
        body.pack(fill=tk.BOTH, expand=True)
        tk.Label(body, text=title, bg=self.colors["surface"], fg=self.colors["muted"], font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(body, textvariable=value_var, bg=self.colors["surface"], fg=self.colors["text"], font=("Segoe UI Semibold", 20, "bold")).pack(anchor="w", pady=(6, 0))

    def _load_json_list(self, file_path: Path) -> list[dict[str, Any]]:
        try:
            with file_path.open("r", encoding="utf-8") as file:
                data = json.load(file)
            if isinstance(data, list):
                return [item for item in data if isinstance(item, dict)]
        except Exception:
            pass
        return []

    def _save_json_list(self, file_path: Path, rows: list[dict[str, Any]]) -> None:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with file_path.open("w", encoding="utf-8") as file:
            json.dump(rows, file, indent=2)

    def _load_run_history(self) -> list[dict[str, Any]]:
        return self._load_json_list(RUN_HISTORY_PATH)

    def _save_run_history(self, rows: list[dict[str, Any]]) -> None:
        self._save_json_list(RUN_HISTORY_PATH, rows)

    def _load_pipeline(self) -> list[dict[str, Any]]:
        return self._load_json_list(APPLICATION_PIPELINE_PATH)

    def _save_pipeline(self, rows: list[dict[str, Any]]) -> None:
        self._save_json_list(APPLICATION_PIPELINE_PATH, rows)

    def _refresh_all_views(self) -> None:
        self._load_dashboard_tree()
        self._load_history_tree()
        self._load_pipeline_tree()
        self._update_dashboard_stats()

    def _load_dashboard_tree(self) -> None:
        if self.dashboard_tree is None:
            return
        self._cancel_dashboard_editor()
        for item in self.dashboard_tree.get_children():
            self.dashboard_tree.delete(item)
        for row in self._load_run_history():
            self.dashboard_tree.insert("", tk.END, values=[row.get(name, "") for name, _, _ in DASHBOARD_COLUMNS])
        self.dashboard_status_var.set(f"Dashboard ready. {len(self.dashboard_tree.get_children())} row(s) loaded.")

    def _update_dashboard_stats(self) -> None:
        run_history = self._load_run_history()
        pipeline = self._load_pipeline()
        self.dashboard_total_var.set(str(len(run_history)))
        self.dashboard_pending_var.set(str(sum(1 for row in run_history if "pending" in str(row.get("result", "")).lower())))
        self.dashboard_success_var.set(str(sum(1 for row in run_history if "submitted" in str(row.get("result", "")).lower())))
        self.dashboard_pipeline_var.set(str(len(pipeline)))

    def _tree_rows_to_dicts(self, tree: ttk.Treeview, columns: list[str]) -> list[dict[str, str]]:
        rows: list[dict[str, str]] = []
        for item_id in tree.get_children():
            values = tree.item(item_id, "values")
            rows.append({column: (values[index] if index < len(values) else "") for index, column in enumerate(columns)})
        return rows

    def _create_apply_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Apply")

        style = ttk.Style(tab)
        style.configure("Title.TLabel", font=("Segoe UI", 16, "bold"))
        style.configure("Section.TLabelframe", padding=10)

        outer = ttk.Frame(tab, padding=12)
        outer.pack(fill=tk.BOTH, expand=True)

        ttk.Label(outer, text="Job Application", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            outer,
            text="Fill details once, press Start, review before submit.",
        ).pack(anchor=tk.W, pady=(0, 10))

        top_row = ttk.Frame(outer)
        top_row.pack(fill=tk.X)

        left_col = ttk.Frame(top_row)
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        right_col = ttk.Frame(top_row)
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, padx=(12, 0))

        profile_box = ttk.LabelFrame(left_col, text="Profile", style="Section.TLabelframe")
        profile_box.pack(fill=tk.X, pady=(0, 10))
        self._row_with_browse(profile_box, "Profile JSON", self.profile_var, self._browse_profile)
        self._row_with_browse(profile_box, "Resume File", self.resume_var, self._browse_resume)

        pref_box = ttk.LabelFrame(left_col, text="Job Preferences", style="Section.TLabelframe")
        pref_box.pack(fill=tk.X, pady=(0, 10))
        self._row(pref_box, "Role", self.role_var)
        self._row(pref_box, "Location", self.location_var)

        source_box = ttk.LabelFrame(left_col, text="Job Source", style="Section.TLabelframe")
        source_box.pack(fill=tk.X, pady=(0, 10))
        ttk.Radiobutton(
            source_box,
            text="Paste Job Link",
            variable=self.source_var,
            value="1",
            command=self._toggle_source,
        ).pack(anchor=tk.W)
        ttk.Radiobutton(
            source_box,
            text="Auto Search Jobs",
            variable=self.source_var,
            value="2",
            command=self._toggle_source,
        ).pack(anchor=tk.W)
        self._row(source_box, "Job URL", self.url_var)

        options_box = ttk.LabelFrame(right_col, text="Options", style="Section.TLabelframe")
        options_box.pack(fill=tk.X)
        ttk.Checkbutton(options_box, text="Headless Browser", variable=self.headless_var).pack(anchor=tk.W)

        action_box = ttk.LabelFrame(right_col, text="Run", style="Section.TLabelframe")
        action_box.pack(fill=tk.X, pady=(10, 0))
        self.start_btn = ttk.Button(action_box, text="Start", command=self._start)
        self.start_btn.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(action_box, textvariable=self.status_var).pack(anchor=tk.W, pady=(8, 0))

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(action_box, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(5, 0))

        log_box = ttk.LabelFrame(outer, text="Execution Log", style="Section.TLabelframe")
        log_box.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.log_text = tk.Text(log_box, height=16, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        self._toggle_source()

    def _create_assistant_tab(self) -> None:
        tab = tk.Frame(self.notebook, bg=self.colors["chat_bg"])
        self.notebook.add(tab, text="AI Assistant")
        tab.grid_columnconfigure(0, weight=0)
        tab.grid_columnconfigure(1, weight=1)
        tab.grid_rowconfigure(0, weight=1)

        sidebar = tk.Frame(tab, bg=self.colors["chat_surface"], width=330, padx=18, pady=18)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        main = tk.Frame(tab, bg=self.colors["chat_bg"], padx=18, pady=18)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        sidebar_header = tk.Frame(sidebar, bg=self.colors["chat_surface"])
        sidebar_header.pack(fill=tk.X)
        tk.Label(
            sidebar_header,
            text="JobGPT",
            bg=self.colors["chat_surface"],
            fg=self.colors["chat_text"],
            font=("Segoe UI Semibold", 22, "bold"),
        ).pack(anchor="w")
        tk.Label(
            sidebar_header,
            text="AI-guided job search workspace",
            bg=self.colors["chat_surface"],
            fg=self.colors["chat_muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        quick_actions = tk.Frame(sidebar, bg=self.colors["chat_surface"])
        quick_actions.pack(fill=tk.X, pady=(18, 0))
        ttk.Button(quick_actions, text="Find Matches", command=self._assistant_find_matches).pack(fill=tk.X)
        ttk.Button(quick_actions, text="Use Selected Match", command=self._use_selected_assistant_result).pack(fill=tk.X, pady=(10, 0))
        ttk.Button(quick_actions, text="Start Search", style="Secondary.TButton", command=self._start).pack(fill=tk.X, pady=(10, 0))

        strategy_card = tk.Frame(sidebar, bg=self.colors["chat_surface_alt"], padx=14, pady=14)
        strategy_card.pack(fill=tk.BOTH, expand=False, pady=(18, 0))
        tk.Label(
            strategy_card,
            text="Search Sequence",
            bg=self.colors["chat_surface_alt"],
            fg=self.colors["chat_text"],
            font=("Segoe UI Semibold", 12, "bold"),
        ).pack(anchor="w")
        tk.Label(
            strategy_card,
            textvariable=self.assistant_sequence_var,
            justify=tk.LEFT,
            bg=self.colors["chat_surface_alt"],
            fg=self.colors["chat_muted"],
            font=("Consolas", 10),
        ).pack(anchor="w", pady=(8, 0))

        tk.Label(
            sidebar,
            text="Current Strategy",
            bg=self.colors["chat_surface"],
            fg=self.colors["chat_text"],
            font=("Segoe UI Semibold", 12, "bold"),
        ).pack(anchor="w", pady=(18, 8))

        strategy_box = tk.Text(
            sidebar,
            height=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg=self.colors["chat_surface_alt"],
            fg=self.colors["chat_text"],
            relief="flat",
            padx=10,
            pady=10,
            insertbackground=self.colors["chat_text"],
        )
        strategy_box.pack(fill=tk.X)
        self.assistant_strategy_text = strategy_box

        tk.Label(
            sidebar,
            text="Smart Matches",
            bg=self.colors["chat_surface"],
            fg=self.colors["chat_text"],
            font=("Segoe UI Semibold", 12, "bold"),
        ).pack(anchor="w", pady=(18, 8))

        matches_wrap = tk.Frame(sidebar, bg=self.colors["chat_surface"])
        matches_wrap.pack(fill=tk.BOTH, expand=True)
        matches_wrap.grid_columnconfigure(0, weight=1)
        matches_wrap.grid_rowconfigure(0, weight=1)

        self.assistant_results_tree = ttk.Treeview(matches_wrap, columns=("rank", "title", "url"), show="headings", height=18)
        self.assistant_results_tree.heading("rank", text="#")
        self.assistant_results_tree.heading("title", text="Match")
        self.assistant_results_tree.heading("url", text="URL")
        self.assistant_results_tree.column("rank", width=40, stretch=False)
        self.assistant_results_tree.column("title", width=190)
        self.assistant_results_tree.column("url", width=220)
        self.assistant_results_tree.grid(row=0, column=0, sticky="nsew")
        self.assistant_results_tree.bind("<Double-1>", lambda _event: self._use_selected_assistant_result())

        hero = tk.Frame(main, bg=self.colors["chat_bg"])
        hero.grid(row=0, column=0, sticky="ew")
        tk.Label(
            hero,
            text="AI Job Search Assistant",
            bg=self.colors["chat_bg"],
            fg=self.colors["chat_text"],
            font=("Segoe UI Semibold", 24, "bold"),
        ).pack(anchor="w")
        tk.Label(
            hero,
            textvariable=self.assistant_provider_var,
            bg=self.colors["chat_bg"],
            fg=self.colors["chat_muted"],
            font=("Segoe UI", 10),
        ).pack(anchor="w", pady=(4, 0))

        chat_shell = tk.Frame(main, bg=self.colors["chat_surface"], padx=1, pady=1)
        chat_shell.grid(row=1, column=0, sticky="nsew", pady=(14, 0))
        chat_shell.grid_columnconfigure(0, weight=1)
        chat_shell.grid_rowconfigure(0, weight=1)

        self.assistant_output = tk.Text(
            chat_shell,
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg=self.colors["chat_surface"],
            fg=self.colors["chat_text"],
            insertbackground=self.colors["chat_text"],
            relief="flat",
            bd=0,
            padx=24,
            pady=24,
            spacing1=8,
            spacing3=14,
        )
        self.assistant_output.grid(row=0, column=0, sticky="nsew")
        self.assistant_output.tag_configure("assistant_name", foreground="#7DD3FC", font=("Segoe UI Semibold", 10, "bold"))
        self.assistant_output.tag_configure("assistant_body", foreground=self.colors["chat_text"], font=("Segoe UI", 11), lmargin1=18, lmargin2=18, rmargin=40, background=self.colors["chat_assistant"])
        self.assistant_output.tag_configure("user_name", foreground="#FDE68A", font=("Segoe UI Semibold", 10, "bold"), justify="right")
        self.assistant_output.tag_configure("user_body", foreground="#FFFFFF", font=("Segoe UI", 11), lmargin1=180, lmargin2=180, rmargin=18, background=self.colors["chat_user"], justify="right")
        self.assistant_output.tag_configure("spacer", spacing1=10, spacing3=18)

        composer_shell = tk.Frame(main, bg=self.colors["chat_bg"], pady=14)
        composer_shell.grid(row=2, column=0, sticky="ew")
        composer_shell.grid_columnconfigure(0, weight=1)

        composer = tk.Frame(composer_shell, bg=self.colors["chat_surface_alt"], padx=14, pady=14)
        composer.grid(row=0, column=0, sticky="ew")
        composer.grid_columnconfigure(0, weight=1)

        self.assistant_input = tk.Text(
            composer,
            height=4,
            wrap=tk.WORD,
            bg=self.colors["chat_surface_alt"],
            fg=self.colors["chat_text"],
            insertbackground=self.colors["chat_text"],
            relief="flat",
            bd=0,
            padx=6,
            pady=6,
        )
        self.assistant_input.grid(row=0, column=0, sticky="ew")
        self.assistant_input.bind("<Control-Return>", lambda _event: self._handle_assistant_send())

        action_row = tk.Frame(composer, bg=self.colors["chat_surface_alt"])
        action_row.grid(row=0, column=1, sticky="ns", padx=(12, 0))
        ttk.Button(action_row, text="Send", command=self._handle_assistant_send).pack(fill=tk.X)
        ttk.Button(action_row, text="Apply", style="Secondary.TButton", command=self._start).pack(fill=tk.X, pady=(10, 0))

        footer = tk.Label(
            composer_shell,
            text="Enter to add a new line. Ctrl+Enter to send.",
            bg=self.colors["chat_bg"],
            fg=self.colors["chat_muted"],
            font=("Segoe UI", 9),
        )
        footer.grid(row=1, column=0, sticky="w", pady=(8, 0))

        self.assistant_provider_var.set(f"{self.assistant_brain.provider_label()} active. Guide the search in sequence, then ask me to search or apply.")
        self._refresh_assistant_strategy_view()
        self._assistant_append(
            "assistant",
            "We’ll work in sequence: set role, set location, add focus keywords or target sites, search, pick a match, then start the application.",
        )

    def _assistant_append(self, speaker: str, message: str) -> None:
        if self.assistant_output is None:
            return
        self.assistant_output.configure(state=tk.NORMAL)
        if speaker == "user":
            self.assistant_output.insert(tk.END, "You\n", ("user_name",))
            self.assistant_output.insert(tk.END, f"{message}\n", ("user_body", "spacer"))
        else:
            self.assistant_output.insert(tk.END, "JobGPT\n", ("assistant_name",))
            self.assistant_output.insert(tk.END, f"{message}\n", ("assistant_body", "spacer"))
        self.assistant_output.see(tk.END)
        self.assistant_output.configure(state=tk.DISABLED)

    def _refresh_assistant_strategy_view(self) -> None:
        sequence = [
            f"1. Role: {'Done' if self.search_direction.role else 'Waiting'}",
            f"2. Location: {'Done' if self.search_direction.location else 'Waiting'}",
            f"3. Direction: {'Done' if self.search_direction.include_keywords or self.search_direction.target_sites or self.search_direction.exclude_keywords else 'Optional'}",
            f"4. Search: {'Ready' if self.search_direction.role else 'Blocked'}",
            f"5. Pick match: {'Ready' if self.assistant_matches else 'Waiting'}",
            "6. Start application",
        ]
        self.assistant_sequence_var.set("\n".join(sequence))

        if self.assistant_strategy_text is not None:
            self.assistant_strategy_text.configure(state=tk.NORMAL)
            self.assistant_strategy_text.delete("1.0", tk.END)
            self.assistant_strategy_text.insert(tk.END, self.search_direction.summary())
            self.assistant_strategy_text.configure(state=tk.DISABLED)

    def _handle_assistant_send(self) -> None:
        if self.assistant_input is None:
            return
        message = self.assistant_input.get("1.0", tk.END).strip()
        if not message:
            return
        self.assistant_input.delete("1.0", tk.END)
        self._assistant_append("user", message)
        reply = self._process_assistant_message(message)
        self._assistant_append("assistant", reply)

    def _process_assistant_message(self, message: str) -> str:
        lowered = message.lower()
        url_match = re.search(r"https?://\S+", message)
        if url_match:
            url = url_match.group(0).rstrip(".,)")
            self.source_var.set("1")
            self.url_var.set(url)
            self._toggle_source()
            self.notebook.select(1)
            return f"I set the job link to {url}. Press Start when you're ready."

        self.search_direction, updates = self.assistant_brain.update_direction(message, self.search_direction)
        if self.search_direction.role:
            self.role_var.set(self.search_direction.role)
        if self.search_direction.location:
            self.location_var.set(self.search_direction.location)
        self._refresh_assistant_strategy_view()

        if "headless" in lowered:
            self.headless_var.set("on" in lowered or "true" in lowered or "enable" in lowered)
            updates.append(f"headless = {self.headless_var.get()}")

        if any(term in lowered for term in ("find matches", "search now", "run search", "look for", "find jobs", "search jobs", "match")):
            results = self._assistant_find_matches()
            summary = f"I found {len(results)} strong matches." if results else "I couldn't find strong matches right now."
            if updates:
                summary = f"I updated {', '.join(updates)}. {summary}"
            return summary + " Double-click a result or use 'Use Selected Match'."

        if any(term in lowered for term in ("start", "apply", "run")):
            self.notebook.select(1)
            self._start()
            if updates:
                return f"I updated {', '.join(updates)} and started the workflow."
            return "I started the workflow."

        if updates:
            self.notebook.select(1)
            return f"I updated {', '.join(updates)}. {self.assistant_brain.next_step_guidance(self.search_direction)}"

        return self.assistant_brain.next_step_guidance(self.search_direction)

    def _assistant_find_matches(self) -> list[Any]:
        preferences = self.search_direction.to_preferences()
        if not preferences["role"]:
            self._assistant_append("assistant", "Set a role first, for example: Find software tester jobs in Auckland.")
            return []

        engine = JobSearchEngine()
        results = engine.pick_smart_matches(engine.search(preferences=preferences))
        self.assistant_matches = results

        if self.assistant_results_tree is not None:
            for item in self.assistant_results_tree.get_children():
                self.assistant_results_tree.delete(item)
            for index, item in enumerate(results, start=1):
                self.assistant_results_tree.insert("", tk.END, values=(index, item.title, item.url))

        self._refresh_assistant_strategy_view()
        return results

    def _use_selected_assistant_result(self) -> None:
        if self.assistant_results_tree is None:
            return
        selected = self.assistant_results_tree.selection()
        if not selected:
            messagebox.showerror("AI Assistant", "Select a smart match first.")
            return
        values = self.assistant_results_tree.item(selected[0], "values")
        if len(values) < 3:
            return
        self.source_var.set("1")
        self.url_var.set(str(values[2]))
        self._toggle_source()
        self.notebook.select(1)
        self._assistant_append("assistant", f"I moved the selected match into the Apply tab: {values[2]}")

    def _create_history_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="History")

        ttk.Label(tab, text="Application History", style="Title.TLabel").pack(anchor="w", padx=20, pady=(18, 6))

        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.history_tree = ttk.Treeview(tree_frame, columns=("timestamp", "role", "location", "result", "url"), show="headings", height=20)
        self.history_tree.heading("timestamp", text="Timestamp")
        self.history_tree.heading("role", text="Role")
        self.history_tree.heading("location", text="Location")
        self.history_tree.heading("result", text="Result")
        self.history_tree.heading("url", text="URL")
        self.history_tree.column("timestamp", width=150)
        self.history_tree.column("role", width=180)
        self.history_tree.column("location", width=120)
        self.history_tree.column("result", width=140)
        self.history_tree.column("url", width=360, stretch=True)

        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scroll.set)

        self.history_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 16))
        ttk.Button(btn_frame, text="Export to CSV", command=self._export_history).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="View Screenshot", command=self._view_history_screenshot).pack(side=tk.LEFT)

    def _export_history(self) -> None:
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            import csv
            run_history = self._load_run_history()
            with open(file_path, "w", newline="") as f:
                writer = csv.DictWriter(f, fieldnames=["timestamp", "role", "location", "result", "url", "summary"])
                writer.writeheader()
                writer.writerows(run_history)
            messagebox.showinfo("Export", "History exported successfully!")

    def _view_history_screenshot(self) -> None:
        if self.history_tree is None:
            return
        selected = self.history_tree.selection()
        if not selected:
            messagebox.showerror("History", "Select a history row first.")
            return

        item = self.history_tree.item(selected[0])
        values = item["values"]
        for run in self._load_run_history():
            if run.get("timestamp") == values[0]:
                screenshot = run.get("screenshot")
                if screenshot and os.path.exists(screenshot):
                    os.startfile(screenshot)
                else:
                    messagebox.showerror("History", "Screenshot not found")
                break

    def _create_settings_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Settings")

        ttk.Label(tab, text="Settings", style="Title.TLabel").pack(anchor="w", padx=20, pady=(18, 6))

        # Profile settings
        profile_frame = ttk.LabelFrame(tab, text="Profile Settings", padding=10)
        profile_frame.pack(fill=tk.X, padx=20, pady=10)
        self._row_with_browse(profile_frame, "Profile JSON", self.profile_var, self._browse_profile)
        self._row_with_browse(profile_frame, "Resume File", self.resume_var, self._browse_resume)

        # Preferences
        pref_frame = ttk.LabelFrame(tab, text="Job Preferences", padding=10)
        pref_frame.pack(fill=tk.X, padx=20, pady=10)
        self._row(pref_frame, "Default Role", self.role_var)
        self._row(pref_frame, "Default Location", self.location_var)

        # Options
        options_frame = ttk.LabelFrame(tab, text="Options", padding=10)
        options_frame.pack(fill=tk.X, padx=20, pady=10)
        ttk.Checkbutton(options_frame, text="Headless Browser", variable=self.headless_var).pack(anchor=tk.W)

        # Save button
        ttk.Button(tab, text="Save Settings", command=self._save_settings).pack(anchor="w", padx=20, pady=10)

    def _save_settings(self) -> None:
        # Save to app_state.json or similar
        state = {
            "profile_path": self.profile_var.get(),
            "resume_path": self.resume_var.get(),
            "role": self.role_var.get(),
            "location": self.location_var.get(),
            "headless": self.headless_var.get()
        }
        APP_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
        with APP_STATE_PATH.open("w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
        messagebox.showinfo("Settings", "Settings saved!")

    def _create_reports_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Reports")

        ttk.Label(tab, text="Reports & Analytics", style="Title.TLabel").pack(anchor="w", padx=20, pady=(18, 6))

        ttk.Button(tab, text="Generate Application Report", command=self._generate_report).pack(anchor="w", padx=20, pady=10)

        self.report_text = tk.Text(tab, height=20, wrap=tk.WORD, bg=self.colors["surface"], fg=self.colors["text"], insertbackground=self.colors["text"], relief="flat")
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        self.report_text.configure(state=tk.DISABLED)

    def _generate_report(self) -> None:
        run_history = self._load_run_history()
        pipeline = self._load_pipeline()

        total_apps = len(run_history)
        successful = sum(1 for r in run_history if "submitted" in r.get("result", "").lower())
        pending = sum(1 for r in run_history if "pending" in r.get("result", "").lower())
        pipeline_count = len(pipeline)

        report = f"""
Application Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Total Applications: {total_apps}
Successful Submissions: {successful}
Pending Reviews: {pending}
Pipeline Jobs: {pipeline_count}

Recent Activity:
"""
        for run in run_history[-5:]:
            report += f"- {run.get('timestamp')}: {run.get('role')} in {run.get('location')} - {run.get('result')}\n"

        self.report_text.configure(state=tk.NORMAL)
        self.report_text.delete(1.0, tk.END)
        self.report_text.insert(tk.END, report)
        self.report_text.configure(state=tk.DISABLED)

    def _create_pipeline_tab(self) -> None:
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Pipeline")

        ttk.Label(tab, text="Application Pipeline", style="Title.TLabel").pack(anchor="w", padx=20, pady=(18, 6))

        # Treeview for pipeline
        tree_frame = ttk.Frame(tab)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        self.pipeline_tree = ttk.Treeview(tree_frame, columns=[name for name, _, _ in PIPELINE_COLUMNS], show="headings", height=15)
        for name, label, width in PIPELINE_COLUMNS:
            self.pipeline_tree.heading(name, text=label)
            self.pipeline_tree.column(name, width=width, stretch=(name in {"company", "role"}))

        scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.pipeline_tree.yview)
        self.pipeline_tree.configure(yscrollcommand=scroll.set)

        self.pipeline_tree.grid(row=0, column=0, sticky="nsew")
        scroll.grid(row=0, column=1, sticky="ns")

        btn_frame = ttk.Frame(tab)
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 16))
        ttk.Button(btn_frame, text="Add Job", command=self._add_pipeline_job).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Edit Selected", command=self._edit_pipeline_job).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Remove Selected", command=self._remove_pipeline_job).pack(side=tk.LEFT)

    def _load_pipeline_tree(self) -> None:
        if self.pipeline_tree is None:
            return
        for item in self.pipeline_tree.get_children():
            self.pipeline_tree.delete(item)
        pipeline = self._load_pipeline()
        for job in pipeline:
            self.pipeline_tree.insert("", tk.END, values=(job.get("timestamp"), job.get("company"), job.get("role"), job.get("stage"), job.get("priority")))

    def _add_pipeline_job(self) -> None:
        # Simple dialog to add job
        company = simpledialog.askstring("Add Job", "Company:")
        if not company:
            return
        role = simpledialog.askstring("Add Job", "Role:")
        if not role:
            return
        stage = "Ready to Apply"
        priority = "High"

        new_job = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "company": company,
            "role": role,
            "location": "",
            "url": "",
            "stage": stage,
            "priority": priority,
            "notes": "",
            "source_value": "1"
        }
        pipeline = self._load_pipeline()
        pipeline.append(new_job)
        self._save_pipeline(pipeline)
        self._refresh_all_views()

    def _edit_pipeline_job(self) -> None:
        tree = self.pipeline_tree
        if tree is None:
            return
        selected = tree.selection()
        if not selected:
            messagebox.showerror("Error", "Select a job to edit")
            return
        item = tree.item(selected[0])
        values = item["values"]
        # For simplicity, edit stage
        new_stage = simpledialog.askstring("Edit Stage", "New Stage:", initialvalue=values[3])
        if new_stage:
            pipeline = self._load_pipeline()
            for job in pipeline:
                if job.get("timestamp") == values[0] and job.get("role") == values[2]:
                    job["stage"] = new_stage
                    break
            self._save_pipeline(pipeline)
            self._refresh_all_views()

    def _remove_pipeline_job(self) -> None:
        tree = self.pipeline_tree
        if tree is None:
            return
        selected = tree.selection()
        if not selected:
            messagebox.showerror("Error", "Select a job to remove")
            return
        if messagebox.askyesno("Confirm", "Remove selected job?"):
            item = tree.item(selected[0])
            values = item["values"]
            pipeline = self._load_pipeline()
            pipeline = [j for j in pipeline if not (j.get("timestamp") == values[0] and j.get("role") == values[2])]
            self._save_pipeline(pipeline)
            self._refresh_all_views()

    def _load_history_tree(self) -> None:
        if self.history_tree is None:
            return
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for run in self._load_run_history():
            self.history_tree.insert("", tk.END, values=(run.get("timestamp"), run.get("role"), run.get("location"), run.get("result"), run.get("url")))

    def _add_dashboard_entry(self) -> None:
        if self.dashboard_tree is None:
            return
        self.dashboard_tree.insert(
            "",
            tk.END,
            values=(
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "New Role",
                "",
                "Manual",
                "Draft",
                "0",
                "0",
                "",
                "",
                "",
            ),
        )
        self.dashboard_status_var.set("New dashboard row added. Double-click any cell to edit.")

    def _delete_dashboard_entry(self) -> None:
        if self.dashboard_tree is None:
            return
        selected = self.dashboard_tree.selection()
        if not selected:
            messagebox.showerror("Dashboard", "Select a dashboard row to delete.")
            return
        for item_id in selected:
            self.dashboard_tree.delete(item_id)
        self.dashboard_status_var.set("Selected dashboard row removed. Save changes to persist it.")

    def _begin_dashboard_edit(self, event: tk.Event) -> None:
        if self.dashboard_tree is None:
            return
        if self.dashboard_tree.identify("region", event.x, event.y) != "cell":
            return

        item_id = self.dashboard_tree.identify_row(event.y)
        column_id = self.dashboard_tree.identify_column(event.x)
        if not item_id or not column_id:
            return

        column_name = DASHBOARD_COLUMNS[int(column_id.replace("#", "")) - 1][0]
        self._show_dashboard_editor(item_id, column_name)

    def _show_dashboard_editor(self, item_id: str, column_name: str) -> None:
        if self.dashboard_tree is None:
            return

        self._cancel_dashboard_editor()
        column_index = next(index for index, (name, _, _) in enumerate(DASHBOARD_COLUMNS) if name == column_name)
        bbox = self.dashboard_tree.bbox(item_id, f"#{column_index + 1}")
        if not bbox:
            return

        x, y, width, height = bbox
        values = list(self.dashboard_tree.item(item_id, "values"))
        current_value = values[column_index] if column_index < len(values) else ""

        options_map = {
            "source": ["Manual", "Auto Search", "Paste Link"],
            "result": ["Draft", "Review pending", "Submitted", "Manual submit required", "Rejected", "Interview", "Offer"],
        }

        if column_name in options_map:
            editor = ttk.Combobox(self.dashboard_tree, values=options_map[column_name], state="readonly")
            editor.set(str(current_value))
            editor.bind("<<ComboboxSelected>>", lambda _event: self._commit_dashboard_editor(item_id, column_name))
        else:
            editor = ttk.Entry(self.dashboard_tree)
            editor.insert(0, str(current_value))

        editor.place(x=x, y=y, width=width, height=height)
        editor.focus_set()
        editor.bind("<Return>", lambda _event: self._commit_dashboard_editor(item_id, column_name))
        editor.bind("<Escape>", lambda _event: self._cancel_dashboard_editor())
        editor.bind("<FocusOut>", lambda _event: self._commit_dashboard_editor(item_id, column_name))
        self.dashboard_editor = editor

    def _commit_dashboard_editor(self, item_id: str, column_name: str) -> None:
        if self.dashboard_tree is None or self.dashboard_editor is None:
            return

        column_index = next(index for index, (name, _, _) in enumerate(DASHBOARD_COLUMNS) if name == column_name)
        values = list(self.dashboard_tree.item(item_id, "values"))
        while len(values) < len(DASHBOARD_COLUMNS):
            values.append("")
        values[column_index] = self.dashboard_editor.get().strip()
        self.dashboard_tree.item(item_id, values=values)
        self.dashboard_status_var.set(f"Updated {column_name.replace('_', ' ')}. Save changes to persist.")
        self._cancel_dashboard_editor()

    def _cancel_dashboard_editor(self) -> None:
        if self.dashboard_editor is not None:
            self.dashboard_editor.destroy()
            self.dashboard_editor = None

    def _save_dashboard_changes(self) -> None:
        if self.dashboard_tree is None:
            return
        self._cancel_dashboard_editor()
        rows = self._tree_rows_to_dicts(self.dashboard_tree, [name for name, _, _ in DASHBOARD_COLUMNS])
        self._save_run_history(rows)
        self._refresh_all_views()
        self.dashboard_status_var.set("Dashboard changes saved.")
        messagebox.showinfo("Dashboard", "Dashboard changes saved successfully.")

    def _export_dashboard_to_excel(self) -> None:
        if self.dashboard_tree is None:
            return
        self._cancel_dashboard_editor()
        REPORTS_DIR.mkdir(parents=True, exist_ok=True)
        default_name = f"job_application_dashboard_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        file_path = filedialog.asksaveasfilename(
            title="Export Dashboard to Excel",
            initialdir=str(REPORTS_DIR),
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not file_path:
            return

        rows = self._tree_rows_to_dicts(self.dashboard_tree, [name for name, _, _ in DASHBOARD_COLUMNS])
        _write_simple_xlsx(
            Path(file_path),
            "Dashboard",
            [label for _, label, _ in DASHBOARD_COLUMNS],
            [[row.get(name, "") for name, _, _ in DASHBOARD_COLUMNS] for row in rows],
        )
        self.dashboard_status_var.set(f"Dashboard exported to {file_path}")
        messagebox.showinfo("Dashboard", "Dashboard exported to Excel successfully.")

    def _prompt_job_match_choice(self, matches: list[Any]) -> Any | None:
        if not matches:
            return None
        if len(matches) == 1:
            return matches[0]

        dialog = tk.Toplevel(self.root)
        dialog.title("Choose Job Match")
        dialog.geometry("900x420")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.colors["bg"])

        selected_match: dict[str, Any] = {"value": None}

        ttk.Label(dialog, text="Pick the best job match", style="Title.TLabel").pack(anchor="w", padx=16, pady=(16, 4))
        ttk.Label(dialog, text="The app found several likely job pages. Choose the one you want to open.", style="Hero.TLabel").pack(anchor="w", padx=16, pady=(0, 12))

        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 12))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        tree = ttk.Treeview(frame, columns=("rank", "title", "url"), show="headings")
        tree.heading("rank", text="#")
        tree.heading("title", text="Title")
        tree.heading("url", text="URL")
        tree.column("rank", width=40, stretch=False)
        tree.column("title", width=240)
        tree.column("url", width=520)
        tree.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        scroll.grid(row=0, column=1, sticky="ns")

        for index, match in enumerate(matches, start=1):
            tree.insert("", tk.END, values=(index, match.title, match.url))

        def choose() -> None:
            selection = tree.selection()
            if not selection:
                return
            values = tree.item(selection[0], "values")
            rank = int(values[0]) - 1
            if 0 <= rank < len(matches):
                selected_match["value"] = matches[rank]
            dialog.destroy()

        tree.bind("<Double-1>", lambda _event: choose())

        actions = ttk.Frame(dialog)
        actions.pack(fill=tk.X, padx=16, pady=(0, 16))
        ttk.Button(actions, text="Use Selected", command=choose).pack(side=tk.LEFT)
        ttk.Button(actions, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=(10, 0))

        dialog.wait_window()
        return selected_match["value"]

    def _row(self, parent: ttk.Widget, label: str, var: tk.StringVar) -> None:
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        entry = ttk.Entry(row, textvariable=var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        if var is self.url_var:
            self.url_entry = entry

    def _row_with_browse(
        self,
        parent: ttk.Widget,
        label: str,
        var: tk.StringVar,
        callback,
    ) -> None:
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text=label, width=12).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(row, text="Browse", command=callback).pack(side=tk.LEFT, padx=(8, 0))

    def _browse_profile(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Profile JSON",
            filetypes=[("JSON", "*.json"), ("All Files", "*.*")],
        )
        if path:
            self.profile_var.set(path)

    def _browse_resume(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Resume File",
            filetypes=[("Documents", "*.pdf *.doc *.docx *.txt"), ("All Files", "*.*")],
        )
        if path:
            self.resume_var.set(path)

    def _toggle_source(self) -> None:
        enabled = self.source_var.get() == "1"
        if self.url_entry is not None:
            self.url_entry.configure(state=("normal" if enabled else "disabled"))

    def _load_defaults(self) -> None:
        try:
            store = ProfileStore(Path(self.profile_var.get()))
            profile = store.load()
            basics = profile.get("basics", {})
            prefs = profile.get("job_preferences", {})
            self.resume_var.set(str(basics.get("resume_path", "")).strip())
            self.role_var.set(str(prefs.get("role", "")))
            self.location_var.set(str(prefs.get("location", "")))
        except Exception:
            pass

        # Load from app_state.json
        try:
            with APP_STATE_PATH.open("r", encoding="utf-8") as f:
                state = json.load(f)
            self.profile_var.set(state.get("profile_path", self.profile_var.get()))
            self.resume_var.set(state.get("resume_path", self.resume_var.get()))
            self.role_var.set(state.get("role", self.role_var.get()))
            self.location_var.set(state.get("location", self.location_var.get()))
            self.headless_var.set(state.get("headless", self.headless_var.get()))
        except Exception:
            pass

        self.search_direction.role = self.role_var.get().strip()
        self.search_direction.location = self.location_var.get().strip()
        self._refresh_assistant_strategy_view()

    def _log(self, message: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{stamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update_idletasks()

    def _append_run_history(self, entry: dict[str, Any]) -> None:
        history = self._load_run_history()
        history.append(entry)
        self._save_run_history(history)

    def _start(self) -> None:
        self.start_btn.configure(state=tk.DISABLED)
        self.status_var.set("Running...")
        self.progress_var.set(0)
        self._log("Starting workflow")
        try:
            self._run_workflow()
            self.progress_var.set(100)
        except Exception as exc:
            self._log(f"Error: {exc}")
            messagebox.showerror("Job Bot", str(exc))
        finally:
            self.start_btn.configure(state=tk.NORMAL)
            self.status_var.set("Ready")
            self._refresh_all_views()

    def _run_workflow(self) -> None:
        self.progress_var.set(10)
        profile_path = Path(_clean_user_path(self.profile_var.get()) or PROFILE_PATH)
        store = ProfileStore(profile_path)
        profile = store.load()

        RESUMES_DIR.mkdir(parents=True, exist_ok=True)
        SCREENSHOTS_DIR.mkdir(parents=True, exist_ok=True)

        resume_source = _resolve_existing_resume_path(self.resume_var.get())
        if resume_source is None:
            raise FileNotFoundError("Resume file not found. Use Browse and pick your resume file.")

        destination = RESUMES_DIR / resume_source.name
        if resume_source.resolve() != destination.resolve():
            shutil.copy2(resume_source, destination)
        final_resume = str(destination.resolve())
        store.remember_answer(profile, "resume_path", final_resume, "Resume")
        store.remember_document(profile, "resume", final_resume, "Resume")

        preferences = self.search_direction.to_preferences()
        preferences["role"] = self.role_var.get().strip() or preferences.get("role", "")
        preferences["location"] = self.location_var.get().strip() or preferences.get("location", "")
        if not preferences["role"]:
            raise ValueError("Role is required.")
        profile["job_preferences"] = preferences
        store.save(profile)

        self.progress_var.set(20)
        choice = self.source_var.get()
        source_label = "Paste Link"
        if choice == "1":
            job_url = self.url_var.get().strip()
            if not job_url:
                raise ValueError("Job URL is required when 'Paste Job Link' is selected.")
            if "://" not in job_url:
                job_url = f"https://{job_url}"
        else:
            source_label = "Auto Search"
            self._log("Auto-searching for jobs...")
            search_engine = JobSearchEngine()
            raw_results = search_engine.search(preferences=preferences)
            results = search_engine.pick_smart_matches(raw_results)
            if not results:
                raise RuntimeError("No jobs found by auto-search. Try manual URL mode.")
            for index, item in enumerate(results[:5], start=1):
                self._log(f"{index}. {item.title} -> {item.url}")
            selected_match = self._prompt_job_match_choice(results)
            if selected_match is None:
                raise RuntimeError("Job search was cancelled before opening a result.")
            self._log(f"Picked match: {selected_match.url}")
            job_url = selected_match.url

        self.progress_var.set(30)
        browser = BrowserSession(headless=self.headless_var.get())
        bot = EasyApplyBot(profile_store=store)
        easy_flow = EasyApplyFlow(bot=bot)
        external_flow = ExternalApplyFlow(bot=bot)

        def prompt_missing(missing_fields: list[dict[str, Any]]) -> bool:
            changed = False
            seen: set[str] = set()
            for field in missing_fields:
                key = str(field.get("key", "")).strip()
                if not key or key in seen:
                    continue
                seen.add(key)
                label = str(field.get("label") or key)
                section = str(field.get("section") or "").strip()
                options = field.get("options", []) or []
                field_type = str(field.get("type", "")).lower()
                suffix = f"\nSection: {section}" if section else ""
                options_text = f"\nOptions: {', '.join(map(str, options[:8]))}" if options else ""

                if field_type == "file":
                    picked = filedialog.askopenfilename(
                        title=f"Choose file for {label}",
                        parent=self.root,
                    )
                    value = picked.strip()
                    if value:
                        store.remember_document(profile, key or "document", value, label)
                else:
                    value = simpledialog.askstring(
                        "Missing Field",
                        f"Enter value for: {label}{suffix}{options_text}",
                        parent=self.root,
                    ) or ""

                if value and value.strip():
                    store.remember_field_answer(profile=profile, field=field, value=value.strip())
                    changed = True
            if changed:
                store.save(profile)
            return changed

        self.progress_var.set(40)
        final_result = "Review pending"
        detected_count = 0
        filled_count = 0
        screenshot_path = ""
        try:
            self._log("Opening browser...")
            attached = browser.start(attach_to_existing=not self.headless_var.get())
            if attached:
                self._log("Attached to existing Brave window. Opening job in a new tab.")
            else:
                self._log(
                    "Could not attach to an existing Brave window. "
                    "Started a separate automation window instead."
                )
            browser.prepare_for_job_search(job_url)
            browser.wait_for_page_settle()

            self.progress_var.set(50)
            apply_type = browser.detect_apply_type()
            self._log(f"Detected apply type: {apply_type}")
            if apply_type == "easy_apply":
                result = easy_flow.run(
                    browser=browser,
                    profile=profile,
                    resume_path=final_resume,
                    prompt_missing=prompt_missing,
                )
            else:
                result = external_flow.run(
                    browser=browser,
                    profile=profile,
                    resume_path=final_resume,
                    prompt_missing=prompt_missing,
                )

            self.progress_var.set(80)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_file = SCREENSHOTS_DIR / f"review_{timestamp}.png"
            browser.save_screenshot(screenshot_file)
            screenshot_path = str(screenshot_file)
            detected_count = len(result.get("fill_plan", []))
            filled_count = sum(1 for item in result.get("fill_plan", []) if item.get("applied"))
            store.save(profile)

            self._log(f"Review screenshot: {screenshot_file}")
            self._log(f"Fields detected: {detected_count}")
            self._log(f"Fields filled: {filled_count}")

            self.progress_var.set(90)
            do_submit = messagebox.askyesno(
                "Review Before Submit",
                "Form filling is complete.\n\nSubmit now?",
                parent=self.root,
            )
            if do_submit:
                clicked = browser.click_submit()
                if clicked:
                    final_result = "Submitted"
                    self._log("Submit clicked. Check browser for final confirmation.")
                else:
                    final_result = "Manual submit required"
                    self._log("Could not auto-click submit. Please submit manually.")
            else:
                final_result = "Review pending"
                self._log("Submission skipped by user.")

            messagebox.showinfo("Job Bot", "Workflow completed.", parent=self.root)
        finally:
            browser.stop()
            self._append_run_history(
                {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "role": preferences.get("role", ""),
                    "location": preferences.get("location", ""),
                    "url": job_url,
                    "source": source_label,
                    "source_value": choice,
                    "result": final_result,
                    "summary": f"Filled {filled_count}/{detected_count} fields on {source_label}",
                    "filled_count": filled_count,
                    "detected_count": detected_count,
                    "screenshot": screenshot_path,
                }
            )


def main() -> None:
    root = tk.Tk()
    app = JobBotApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
