from __future__ import annotations

import csv
import importlib
import re
import subprocess
import sys
import tkinter as tk
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from tkinter import font
from tkinter import filedialog, messagebox, ttk

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

try:
    from tksheet import Sheet
except ModuleNotFoundError:
    Sheet = None


VIRTUAL_VIEW_ROWS = 200
FULL_RENDER_ROW_LIMIT = 5000
MOUSE_WHEEL_ROWS = 20
AUTO_SPLIT_DELIMITERS = [",", ";", "\t", "|"]
AUTO_SPLIT_MIN_RATIO = 0.6
ZOOM_MIN = 0.5
ZOOM_MAX = 2.0
ZOOM_STEP = 0.1
UNDO_LIMIT = 50


class DataAnalysisApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("LSPR Data Analysis Software")
        self.geometry("980x680")
        self.minsize(820, 560)

        self.data = None
        self.current_file: Path | None = None
        self.load_note = ""
        self.source_data_rows: int | None = None
        self.selected_columns: set[int] = set()
        self.selected_rows: set[int] = set()
        self.average_columns: set[int] = set()
        self.drag_mode: str | None = None
        self.drag_anchor: int | None = None
        self.undo_stack: list[dict] = []
        self.table_offset = 0
        self.full_render_mode = False
        self.zoom_scale = 1.0
        self.average_editable_cell: tuple[int, int] | None = None

        self._build_ui()

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(self, padding=(12, 12, 12, 8))
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.columnconfigure(9, weight=1)

        ttk.Button(toolbar, text="Load Data", command=self.load_data).grid(
            row=0, column=0, padx=(0, 8)
        )
        ttk.Button(toolbar, text="Auto Analyze", command=self.analyze_data).grid(
            row=0, column=1, padx=(0, 8)
        )
        ttk.Button(
            toolbar,
            text="Average Selected Columns",
            command=self.average_selected_columns,
        ).grid(row=0, column=2, padx=(0, 8))
        ttk.Button(toolbar, text="Undo", command=self.undo_last_action).grid(
            row=0, column=3, padx=(0, 8)
        )
        ttk.Button(toolbar, text="Install Dependencies", command=self.install_dependencies).grid(
            row=0, column=4, padx=(0, 12)
        )
        ttk.Button(toolbar, text="Zoom -", command=lambda: self._change_zoom(-ZOOM_STEP)).grid(
            row=0, column=5, padx=(0, 6)
        )
        self.zoom_label = ttk.Label(toolbar, text="100%")
        self.zoom_label.grid(row=0, column=6, padx=(0, 6))
        ttk.Button(toolbar, text="Zoom +", command=lambda: self._change_zoom(ZOOM_STEP)).grid(
            row=0, column=7, padx=(0, 6)
        )
        ttk.Button(toolbar, text="Reset Zoom", command=self._reset_zoom).grid(
            row=0, column=8, padx=(0, 12)
        )

        self.file_label = ttk.Label(toolbar, text="No data loaded")
        self.file_label.grid(row=0, column=9, sticky="w")

        content = ttk.Frame(self, padding=(12, 0, 12, 12))
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(content)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.table_frame = ttk.Frame(self.notebook)
        self.report_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.table_frame, text="Data Table")
        self.notebook.add(self.report_frame, text="Report")

        self.table_frame.columnconfigure(0, weight=1)
        self.table_frame.rowconfigure(0, weight=1)
        self.report_frame.columnconfigure(0, weight=1)
        self.report_frame.rowconfigure(0, weight=1)

        self.style = ttk.Style(self)
        self.report_font = font.Font(family="Consolas", size=10)
        self.sheet = self._create_sheet()

        self.output = tk.Text(self.report_frame, wrap="none", font=self.report_font)
        self.output.grid(row=0, column=0, sticky="nsew")
        self.output.bind("<Control-MouseWheel>", self._on_zoom_mousewheel)
        self.output.bind("<Control-z>", lambda event: self._undo_from_event())
        self.output.bind("<Control-Z>", lambda event: self._undo_from_event())

        report_y_scroll = ttk.Scrollbar(
            self.report_frame, orient="vertical", command=self.output.yview
        )
        report_y_scroll.grid(row=0, column=1, sticky="ns")

        report_x_scroll = ttk.Scrollbar(
            self.report_frame, orient="horizontal", command=self.output.xview
        )
        report_x_scroll.grid(row=1, column=0, sticky="ew")

        self.output.configure(
            yscrollcommand=report_y_scroll.set,
            xscrollcommand=report_x_scroll.set,
        )
        self._apply_zoom()
        self._write_output("Click \"Load Data\" to select a CSV or Excel file.\n")
        self.notebook.select(self.report_frame)

    def load_data(self) -> None:
        if not self._ensure_dependencies():
            return

        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                ("Data files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*"),
            ],
        )
        if not file_path:
            return

        path = Path(file_path)
        try:
            self.data = self._read_file(path)
        except Exception as exc:
            messagebox.showerror("Load Failed", f"Unable to load file:\n{exc}")
            return

        self.current_file = path
        self.undo_stack.clear()
        self.average_columns.clear()
        self.average_editable_cell = None
        self.file_label.configure(text=f"Current file: {path.name}")
        self._show_data_table()

    def analyze_data(self) -> None:
        if not self._ensure_dependencies():
            return

        if self.data is None:
            messagebox.showwarning("No Data", "Please click \"Load Data\" first.")
            return

        df = self.data
        lines: list[str] = []
        lines.append("Auto Analysis Result")
        lines.append("=" * 80)
        lines.append(f"File name: {self.current_file.name if self.current_file else 'Unknown'}")
        lines.append(f"Data size: {df.shape[0]} rows x {df.shape[1]} columns")
        lines.append(f"Duplicate rows: {df.duplicated().sum()}")
        lines.append("")

        lines.append("Column Types")
        lines.append("-" * 80)
        lines.append(df.dtypes.astype(str).to_string())
        lines.append("")

        missing = df.isna().sum()
        missing_rate = (df.isna().mean() * 100).round(2)
        missing_report = pd.DataFrame(
            {"Missing Count": missing, "Missing Rate (%)": missing_rate}
        )
        lines.append("Missing Value Analysis")
        lines.append("-" * 80)
        lines.append(missing_report.to_string())
        lines.append("")

        numeric_df = df.select_dtypes(include="number")
        if not numeric_df.empty:
            lines.append("Numeric Column Statistics")
            lines.append("-" * 80)
            lines.append(numeric_df.describe().round(4).to_string())
            lines.append("")

            if numeric_df.shape[1] >= 2:
                lines.append("Numeric Column Correlation Matrix")
                lines.append("-" * 80)
                lines.append(numeric_df.corr().round(4).to_string())
                lines.append("")
        else:
            lines.append("No numeric columns detected.")
            lines.append("")

        categorical_df = df.select_dtypes(exclude="number")
        if not categorical_df.empty:
            lines.append("Text/Categorical Column Top 5 Values")
            lines.append("-" * 80)
            for column in categorical_df.columns:
                lines.append(f"[{column}]")
                lines.append(categorical_df[column].value_counts(dropna=False).head(5).to_string())
                lines.append("")

        self._write_output("\n".join(lines))
        self.notebook.select(self.report_frame)

    def average_selected_columns(self) -> None:
        if not self._ensure_dependencies():
            return

        if self.data is None:
            messagebox.showwarning("No Data", "Please click \"Load Data\" first.")
            return

        self._sync_sheet_selection()
        if not self.selected_columns:
            messagebox.showwarning(
                "No Columns Selected",
                "Click one or more column headers in the Data Table tab first.",
            )
            return

        selected_indexes = sorted(
            index for index in self.selected_columns if 0 <= index < len(self.data.columns)
        )
        if not selected_indexes:
            messagebox.showwarning("No Columns Selected", "The selected columns are no longer valid.")
            self.selected_columns.clear()
            self._apply_sheet_highlights()
            return

        selected_data = self.data.iloc[:, selected_indexes].apply(pd.to_numeric, errors="coerce")
        averages = selected_data.mean(axis=1, skipna=True)
        if averages.isna().all():
            messagebox.showwarning(
                "No Numeric Data",
                "The selected columns do not contain numeric values.",
            )
            return

        insert_at = max(selected_indexes) + 1
        column_name = self._next_average_column_name()
        self._save_undo_snapshot()
        self.data.insert(
            insert_at,
            column_name,
            averages.map(self._format_decimal_places),
            allow_duplicates=True,
        )
        self.average_columns = {index + 1 if index >= insert_at else index for index in self.average_columns}
        self.average_columns.add(insert_at)
        self.average_editable_cell = (0, insert_at)
        self.load_note = (
            f"Inserted '{column_name}' at column {insert_at + 1} from "
            f"{len(selected_indexes)} selected columns. Values are rounded to 3 decimal places."
        )
        self._show_data_table()
        self._activate_average_first_cell(insert_at)

    def undo_last_action(self) -> None:
        if not self.undo_stack:
            messagebox.showinfo("Nothing to Undo", "There is no previous calculation to undo.")
            return

        snapshot = self.undo_stack.pop()
        self.data = snapshot["data"]
        self.load_note = snapshot["load_note"]
        self.selected_columns = set(snapshot["selected_columns"])
        self.selected_rows = set(snapshot["selected_rows"])
        self.average_columns = set(snapshot["average_columns"])
        self.average_editable_cell = snapshot.get("average_editable_cell")
        self.table_offset = snapshot["table_offset"]

        self._show_data_table(keep_selection=True)
        self._write_output(
            "\n".join(
                [
                    "Undo Complete",
                    "=" * 80,
                    "The previous step was restored.",
                    f"Remaining undo steps: {len(self.undo_stack)}",
                ]
            )
        )

    def _undo_from_event(self) -> str:
        self.undo_last_action()
        return "break"

    def _save_undo_snapshot(self) -> None:
        if self.data is None:
            return

        self.undo_stack.append({
            "data": self.data.copy(deep=True),
            "load_note": self.load_note,
            "selected_columns": set(self.selected_columns),
            "selected_rows": set(self.selected_rows),
            "average_columns": set(self.average_columns),
            "average_editable_cell": self.average_editable_cell,
            "table_offset": self.table_offset,
        })
        if len(self.undo_stack) > UNDO_LIMIT:
            self.undo_stack.pop(0)

    def install_dependencies(self) -> None:
        requirements_path = Path(__file__).with_name("requirements.txt")
        if not requirements_path.exists():
            messagebox.showerror("Install Failed", "requirements.txt was not found.")
            return

        self._write_output(
            "Installing dependencies. Please wait...\n\n"
            f"Python: {sys.executable}\n"
            f"Command: {sys.executable} -m pip install -r {requirements_path}\n"
        )
        self.update_idletasks()

        try:
            result = subprocess.run(
                [sys.executable, "-m", "pip", "install", "-r", str(requirements_path)],
                capture_output=True,
                text=True,
                check=False,
            )
        except Exception as exc:
            messagebox.showerror("Install Failed", f"Unable to start pip:\n{exc}")
            return

        output = "\n".join(part for part in [result.stdout, result.stderr] if part)
        if result.returncode != 0:
            self._write_output("Dependency installation failed:\n\n" + output)
            messagebox.showerror(
                "Install Failed",
                "Dependency installation failed. See the output panel for details.",
            )
            return

        self._load_pandas()
        self._load_tksheet()
        self._write_output(
            "Dependencies installed successfully.\n\n"
            "You can now click \"Load Data\" to continue.\n\n"
            + output
        )
        messagebox.showinfo(
            "Install Complete",
            "Dependencies have been installed. You can now load data.",
        )

    def _read_file(self, path: Path):
        if pd is None:
            raise RuntimeError("Missing pandas dependency. Run: pip install -r requirements.txt")

        self.load_note = ""
        self.source_data_rows = None
        suffix = path.suffix.lower()
        if suffix == ".csv":
            data = self._read_csv_file(path)
            return self._clean_column_names(self._auto_split_single_column(data))
        if suffix in {".xlsx", ".xls"}:
            data = pd.read_excel(path)
            return self._clean_column_names(self._auto_split_single_column(data))
        raise ValueError("Only CSV, XLSX, and XLS files are supported.")

    def _read_csv_file(self, path: Path):
        self.load_note = ""
        text_lines = self._read_text_lines(path)
        expected_data_rows = max(len(text_lines) - 1, 0)
        self.source_data_rows = expected_data_rows
        attempts = [
            {"sep": None, "engine": "python"},
            {"sep": ",", "engine": "python"},
        ]

        last_error = None
        for options in attempts:
            try:
                data = pd.read_csv(path, **options)
                if self._csv_row_count_is_valid(data, expected_data_rows):
                    return data
                last_error = RuntimeError(
                    f"Parsed only {len(data)} rows from {expected_data_rows} expected data rows."
                )
            except Exception as exc:
                last_error = exc

        try:
            self.load_note = (
                "CSV parser fallback was used because the parsed row count did not match "
                f"the file line count ({expected_data_rows} expected data rows). "
                "Loaded each line as one column before auto-splitting."
            )
            if not text_lines:
                return pd.DataFrame()
            return pd.DataFrame({text_lines[0]: text_lines[1:]})
        except Exception as fallback_error:
            raise RuntimeError(
                "Unable to read CSV with automatic, comma, or raw-line parsing. "
                f"Last error: {last_error}; fallback error: {fallback_error}"
            ) from fallback_error

    def _csv_row_count_is_valid(self, data, expected_data_rows: int) -> bool:
        if expected_data_rows <= 0:
            return True

        missing_rows = expected_data_rows - len(data)
        tolerance = max(3, round(expected_data_rows * 0.02))
        return missing_rows <= tolerance

    def _read_text_lines(self, path: Path) -> list[str]:
        for encoding in ("utf-8-sig", "utf-8", "gbk", "latin1"):
            try:
                return path.read_text(encoding=encoding).splitlines()
            except UnicodeDecodeError:
                continue

        return path.read_text(errors="replace").splitlines()

    def _ensure_dependencies(self) -> bool:
        self._load_pandas()
        self._load_tksheet()
        if pd is not None and Sheet is not None:
            return True

        missing = []
        if pd is None:
            missing.append("pandas")
        if Sheet is None:
            missing.append("tksheet")
        message = (
            f"The current Python environment is missing dependencies: {', '.join(missing)}.\n\n"
            "Click the \"Install Dependencies\" button above, or run this command "
            "in the project folder:\n"
            f"{sys.executable} -m pip install -r requirements.txt\n\n"
            "After installation, click RUN again."
        )
        messagebox.showerror("Missing Dependencies", message)
        self._write_output(message)
        return False

    def _ensure_pandas(self) -> bool:
        return self._ensure_dependencies()

    def _load_pandas(self) -> None:
        global pd
        if pd is not None:
            return

        try:
            pd = importlib.import_module("pandas")
        except ModuleNotFoundError:
            pd = None

    def _load_tksheet(self) -> None:
        global Sheet
        if Sheet is not None:
            return

        try:
            from tksheet import Sheet as LoadedSheet
        except ModuleNotFoundError:
            Sheet = None
        else:
            Sheet = LoadedSheet

    def _create_sheet(self):
        if Sheet is None:
            label = ttk.Label(
                self.table_frame,
                text="Missing tksheet dependency. Click Install Dependencies.",
            )
            label.grid(row=0, column=0, sticky="nsew")
            return None

        sheet = Sheet(
            self.table_frame,
            data=[],
            headers=[],
            row_index=[],
            theme="light blue",
            show_row_index=True,
            show_header=True,
            font=("Segoe UI", 9, "normal"),
            header_font=("Segoe UI", 9, "bold"),
            index_font=("Segoe UI", 9, "normal"),
            default_row_height=24,
            default_column_width=110,
            table_selected_columns_bg="#cfe8ff",
            table_selected_rows_bg="#cfe8ff",
            header_selected_columns_bg="#0B57D0",
            header_selected_columns_fg="#FFFFFF",
            index_selected_rows_bg="#0B57D0",
            index_selected_rows_fg="#FFFFFF",
        )
        sheet.grid(row=0, column=0, sticky="nsew")
        sheet.enable_bindings(
            "single_select",
            "drag_select",
            "column_select",
            "row_select",
            "arrowkeys",
            "copy",
            "rc_select",
        )
        sheet.bind("<ButtonRelease-1>", self._sync_sheet_selection_later)
        sheet.bind("<KeyRelease>", self._sync_sheet_selection_later)
        sheet.bind("<Double-Button-1>", self._on_sheet_double_click, add="+")
        sheet.bind("<Control-MouseWheel>", self._on_zoom_mousewheel)
        sheet.bind("<Control-z>", lambda event: self._undo_from_event())
        sheet.bind("<Control-Z>", lambda event: self._undo_from_event())
        return sheet

    def _auto_split_single_column(self, data):
        existing_note = self.load_note
        if data.shape[1] != 1:
            return data

        column_name = str(data.columns[0])
        values = data.iloc[:, 0].dropna().astype(str)
        sample = [column_name, *values.head(200).tolist()]
        split_rule = self._detect_split_rule(sample)
        if split_rule is None:
            return data

        delimiter, expected_columns = split_rule
        parsed_rows = [
            self._split_cell(value, delimiter, expected_columns)
            for value in data.iloc[:, 0].map(self._format_cell)
        ]

        header_values = self._split_cell(column_name, delimiter, expected_columns)
        if len(header_values) == expected_columns and len(set(header_values)) == expected_columns:
            columns = [value.strip() for value in header_values]
        else:
            columns = [""] * expected_columns

        split_note = (
            f"Auto-split one-column data into {expected_columns} columns "
            f"using {self._delimiter_label(delimiter)}."
        )
        self.load_note = f"{existing_note}\n{split_note}" if existing_note else split_note
        return pd.DataFrame(parsed_rows, columns=columns)

    def _clean_column_names(self, data):
        cleaned_columns = []
        for column in data.columns:
            column_name = "" if column is None else str(column).strip()
            if re.fullmatch(r"Unnamed:\s*\d+(?:_level_\d+)?", column_name):
                column_name = ""
            cleaned_columns.append(column_name)

        data = data.copy()
        data.columns = cleaned_columns
        return data

    def _detect_split_rule(self, sample: list[str]) -> tuple[str, int] | None:
        best_rule = None
        best_score = 0

        for delimiter in AUTO_SPLIT_DELIMITERS:
            counts = [
                len(self._split_cell(value, delimiter))
                for value in sample
                if delimiter in value
            ]
            rule = self._score_split_counts(delimiter, counts, len(sample))
            if rule and rule[2] > best_score:
                best_rule = (rule[0], rule[1])
                best_score = rule[2]

        if best_rule is not None:
            return best_rule

        whitespace_counts = [
            len(re.split(r"\s+", value.strip()))
            for value in sample
            if re.search(r"\s{2,}|\t", value.strip())
        ]
        rule = self._score_split_counts("whitespace", whitespace_counts, len(sample))
        if rule:
            return rule[0], rule[1]

        return None

    def _score_split_counts(
        self, delimiter: str, counts: list[int], sample_size: int
    ) -> tuple[str, int, int] | None:
        valid_counts = [count for count in counts if count > 1]
        if not valid_counts:
            return None

        expected_columns = max(valid_counts)
        matching_count = valid_counts.count(expected_columns)
        split_ratio = len(valid_counts) / max(sample_size, 1)
        if sample_size > 3 and split_ratio < AUTO_SPLIT_MIN_RATIO:
            return None

        return delimiter, expected_columns, len(valid_counts) * expected_columns

    def _split_cell(
        self, value: str, delimiter: str, expected_columns: int | None = None
    ) -> list[str]:
        text = str(value)
        if delimiter == "whitespace":
            parts = re.split(r"\s+", text.strip()) if text.strip() else [""]
        else:
            parts = next(csv.reader([text], delimiter=delimiter))

        cleaned_parts = [part.strip() for part in parts]
        if expected_columns is None:
            return cleaned_parts

        if len(cleaned_parts) < expected_columns:
            return [*cleaned_parts, *([""] * (expected_columns - len(cleaned_parts)))]
        return cleaned_parts

    def _delimiter_label(self, delimiter: str) -> str:
        labels = {
            ",": "comma delimiter",
            ";": "semicolon delimiter",
            "\t": "tab delimiter",
            "|": "pipe delimiter",
            "whitespace": "whitespace delimiter",
        }
        return labels.get(delimiter, f"{delimiter!r} delimiter")

    def _show_data_table(self, keep_selection: bool = False) -> None:
        if self.data is None:
            return
        if self.sheet is None:
            self._load_tksheet()
            self.sheet = self._create_sheet()
            if self.sheet is None:
                messagebox.showerror("Missing Dependency", "tksheet is required for the data table.")
                return

        df = self.data
        if not keep_selection:
            self.table_offset = 0
            self.selected_columns.clear()
            self.selected_rows.clear()
        self.drag_mode = None
        self.drag_anchor = None
        self.full_render_mode = True
        self._set_sheet_data()
        self._apply_sheet_highlights()
        self._restore_average_editable_cell()

        self.notebook.select(self.table_frame)
        self._write_output(
            "\n".join(
                [
                    "Data Loaded Successfully",
                    "=" * 80,
                    f"File name: {self.current_file.name if self.current_file else 'Unknown'}",
                    f"Data size: {df.shape[0]} rows x {df.shape[1]} columns",
                    *( [f"Source data rows: {self.source_data_rows}"] if self.source_data_rows is not None else [] ),
                    "",
                    *( [self.load_note, ""] if self.load_note else [] ),
                    "The dataset is shown in the Data Table tab.",
                    "All loaded rows are rendered in the table.",
                    "Selected rows and columns are highlighted in blue.",
                    "Generated Average columns are highlighted in yellow.",
                ]
            )
        )

    def _set_sheet_data(self) -> None:
        if self.data is None or self.sheet is None:
            return

        sheet_data = [
            [self._format_cell(value) for value in row]
            for row in self.data.itertuples(index=False, name=None)
        ]
        headers = [self._display_column_name(index) for index in range(len(self.data.columns))]
        row_index = [str(index) for index in range(1, len(self.data) + 1)]
        self.sheet.set_sheet_data(
            sheet_data,
            reset_col_positions=True,
            reset_row_positions=True,
            reset_highlights=True,
        )
        self.sheet.headers(headers, reset_col_positions=False)
        self.sheet.row_index(row_index, reset_row_positions=False)
        self.sheet.set_all_column_widths(width=self._scale_size(110), redraw=False)
        self.sheet.set_all_row_heights(height=self._scale_size(24), redraw=False)
        self.sheet.redraw()

    def _apply_sheet_highlights(self) -> None:
        if self.sheet is None:
            return

        self.sheet.dehighlight_all(redraw=False)
        if self.average_columns:
            self.sheet.highlight_columns(
                sorted(self.average_columns),
                bg="#fff2a8",
                fg="black",
                highlight_header=True,
                redraw=False,
            )
        if self.selected_columns:
            self.sheet.highlight_columns(
                sorted(self.selected_columns),
                bg="#cfe8ff",
                fg="black",
                highlight_header=True,
                redraw=False,
            )
        if self.selected_rows:
            self.sheet.highlight_rows(
                sorted(self.selected_rows),
                bg="#cfe8ff",
                fg="black",
                highlight_index=True,
                redraw=False,
            )
        self.sheet.redraw()

    def _activate_average_first_cell(self, average_column_index: int) -> None:
        if self.sheet is None or self.data is None:
            return

        if not (0 <= average_column_index < len(self.data.columns)):
            return

        self.sheet.readonly_cells(cells=[(0, average_column_index)], readonly=False, redraw=False)
        self.sheet.select_cell(0, average_column_index, redraw=False)
        self.sheet.redraw()
        try:
            self.sheet.open_cell()
        except Exception:
            pass

    def _restore_average_editable_cell(self) -> None:
        if self.sheet is None or self.data is None or self.average_editable_cell is None:
            return

        row, column = self.average_editable_cell
        if not (0 <= row < len(self.data) and 0 <= column < len(self.data.columns)):
            return

        self.sheet.readonly_cells(cells=[(row, column)], readonly=False, redraw=False)
        self.sheet.redraw()

    def _on_sheet_double_click(self, event: tk.Event) -> str | None:
        if self.sheet is None or self.data is None or self.average_editable_cell is None:
            return None

        if self.sheet.identify_region(event) != "table":
            return None

        row = self.sheet.identify_row(event)
        column = self.sheet.identify_column(event)
        if row is None or column is None:
            return None

        if (row, column) != self.average_editable_cell:
            return None

        self.sheet.readonly_cells(cells=[(row, column)], readonly=False, redraw=False)
        self.sheet.select_cell(row, column, redraw=False)
        self.sheet.redraw()
        self.after_idle(lambda: self._open_average_editable_cell(row, column))
        return "break"

    def _open_average_editable_cell(self, row: int, column: int) -> None:
        if self.sheet is None or self.data is None:
            return

        if self.average_editable_cell != (row, column):
            return

        if not (0 <= row < len(self.data) and 0 <= column < len(self.data.columns)):
            return

        self.sheet.readonly_cells(cells=[(row, column)], readonly=False, redraw=False)
        self.sheet.select_cell(row, column, redraw=False)
        try:
            self.sheet.open_cell()
        except Exception:
            pass

    def _sync_sheet_selection_later(self, event=None) -> str:
        self.after_idle(self._sync_sheet_selection)
        return "break"

    def _sync_sheet_selection(self) -> None:
        if self.sheet is None or self.data is None:
            return

        selected_columns = set(self.sheet.get_selected_columns())
        selected_rows = set(self.sheet.get_selected_rows())
        selected_columns = {index for index in selected_columns if 0 <= index < len(self.data.columns)}
        selected_rows = {index for index in selected_rows if 0 <= index < len(self.data)}
        if not selected_columns and not selected_rows:
            return
        if selected_columns != self.selected_columns or selected_rows != self.selected_rows:
            self.selected_columns = selected_columns
            self.selected_rows = selected_rows
            self._apply_sheet_highlights()

    def _toggle_column_selection(self, column_index: int) -> None:
        if column_index in self.selected_columns:
            self.selected_columns.remove(column_index)
        else:
            self.selected_columns.add(column_index)
        self._apply_sheet_highlights()

    def _on_table_click(self, event: tk.Event) -> str | None:
        return None

    def _on_table_drag(self, event: tk.Event) -> str:
        return "break"

    def _on_table_release(self, event: tk.Event) -> str:
        self.drag_mode = None
        self.drag_anchor = None
        return "break"

    def _table_column_to_data_index(self, column_id: str) -> int | None:
        if not column_id.startswith("#"):
            return None

        try:
            tree_column_index = int(column_id[1:]) - 1
        except ValueError:
            return None

        if tree_column_index <= 0:
            return None

        data_column_index = tree_column_index - 1
        if self.data is None or data_column_index >= len(self.data.columns):
            return None
        return data_column_index

    def _event_to_data_row_index(self, event: tk.Event) -> int | None:
        return None

    def _select_column_range(self, start: int, end: int) -> None:
        if self.data is None:
            return

        low, high = sorted((start, end))
        self.selected_columns = set(range(low, high + 1))
        self._apply_sheet_highlights()

    def _select_row_range(self, start: int, end: int) -> None:
        if self.data is None:
            return

        low, high = sorted((start, end))
        self.selected_rows = set(range(low, high + 1))
        self._apply_sheet_highlights()

    def _update_table_headings(self) -> None:
        self._apply_sheet_highlights()

    def _display_column_name(self, column_index: int) -> str:
        if self.data is None:
            return ""
        return str(self.data.columns[column_index])

    def _next_average_column_name(self) -> str:
        if self.data is None:
            return "Average"

        existing_names = {str(column) for column in self.data.columns}
        if "Average" not in existing_names:
            return "Average"

        suffix = 2
        while f"Average {suffix}" in existing_names:
            suffix += 1
        return f"Average {suffix}"

    def _render_table_rows(self) -> None:
        self._set_sheet_data()
        self._apply_sheet_highlights()

    def _update_visible_row_tags(self) -> None:
        self._apply_sheet_highlights()

    def _on_table_scroll(self, action: str, value: str, unit: str | None = None) -> None:
        return None

    def _on_table_mousewheel(self, event: tk.Event) -> str:
        return None

    def _on_zoom_mousewheel(self, event: tk.Event) -> str:
        if event.delta > 0:
            self._change_zoom(ZOOM_STEP)
        elif event.delta < 0:
            self._change_zoom(-ZOOM_STEP)
        return "break"

    def _change_zoom(self, delta: float) -> None:
        self.zoom_scale = max(ZOOM_MIN, min(ZOOM_MAX, self.zoom_scale + delta))
        self._apply_zoom()

    def _reset_zoom(self) -> None:
        self.zoom_scale = 1.0
        self._apply_zoom()

    def _apply_zoom(self) -> None:
        table_size = max(7, round(9 * self.zoom_scale))
        report_size = max(7, round(10 * self.zoom_scale))
        row_height = max(18, round(24 * self.zoom_scale))

        self.report_font.configure(size=report_size)
        self.zoom_label.configure(text=f"{round(self.zoom_scale * 100)}%")

        if self.sheet is not None:
            self.sheet.set_options(
                font=("Segoe UI", table_size, "normal"),
                header_font=("Segoe UI", table_size, "bold"),
                index_font=("Segoe UI", table_size, "normal"),
                row_height=row_height,
                redraw=False,
            )
            self.sheet.set_all_row_heights(height=row_height, redraw=False)
            self.sheet.set_all_column_widths(width=self._scale_size(110), redraw=True)

    def _refresh_table_column_widths(self) -> None:
        if self.data is None or self.sheet is None:
            return

        self.sheet.set_all_column_widths(width=self._scale_size(110), redraw=True)

    def _move_table(self, delta: int) -> str:
        return ""

    def _set_table_offset(self, offset: int) -> str:
        self.table_offset = max(0, min(offset, self._max_table_offset()))
        return "break"

    def _max_table_offset(self) -> int:
        if self.data is None:
            return 0
        return max(len(self.data) - VIRTUAL_VIEW_ROWS, 0)

    def _calculate_column_width(self, series, column_name: str) -> int:
        sample_values = series.head(100).map(self._format_cell).tolist()
        sample_values.append(column_name)
        max_length = max((len(value) for value in sample_values), default=len(column_name))
        base_width = min(max(max_length * 8 + 24, 100), 320)
        return self._scale_size(base_width)

    def _scale_size(self, value: int) -> int:
        return max(1, round(value * self.zoom_scale))

    def _format_cell(self, value) -> str:
        if pd is not None and pd.isna(value):
            return ""
        return str(value)

    def _format_table_cell(self, value, column_index: int) -> str:
        return self._format_cell(value)

    def _format_decimal_places(self, value) -> str:
        if pd is not None and pd.isna(value):
            return ""
        try:
            number = Decimal(str(value))
        except InvalidOperation:
            return ""
        return str(number.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP))

    def _write_output(self, text: str) -> None:
        self.output.configure(state="normal")
        self.output.delete("1.0", tk.END)
        self.output.insert(tk.END, text)
        self.output.configure(state="disabled")


if __name__ == "__main__":
    app = DataAnalysisApp()
    app.mainloop()
