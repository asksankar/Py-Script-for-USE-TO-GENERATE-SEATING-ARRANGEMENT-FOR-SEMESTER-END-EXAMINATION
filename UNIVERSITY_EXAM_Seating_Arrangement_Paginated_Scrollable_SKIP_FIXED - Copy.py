import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import re

import pandas as pd

try:
    import pdfplumber
except ImportError:
    pdfplumber = None

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class SeatingArrangementApp:

    # ==========================================================
    # REGISTER NUMBER HEADINGS
    # ==========================================================

    REGISTER_HEADER_PATTERNS = [
        r"REGISTER\s*NUMBER",
        r"REGISTER\s*NO",
        r"REGISTRATION\s*NUMBER",
        r"REGISTRATION\s*NO",
        r"REG\s*NUMBER",
        r"REG\s*NO",
        r"REG\.?\s*NUMBER",
        r"REG\.?\s*NO",
    ]

    def __init__(self, root):

        self.root = root

        self.root.title(
            "Register Number - Multi Allotment Seating Arrangement"
        )

        self.root.geometry("1450x950")

        # ======================================================
        # DATA
        # ======================================================

        self.register_numbers = []

        self.source_records = []

        self.selected_files = []

        self.current_arrangement = []

        self.all_allotments = []

        self.blocked_cells = set()

        self.cell_buttons = {}
        self.page_buttons = {}
        self.current_page = 1
        self.page_info_var = tk.StringVar(value="Page 1 / 1")

        # ======================================================
        # VARIABLES
        # ======================================================

        self.rows_var = tk.IntVar(value=20)

        self.cols_var = tk.IntVar(value=6)

        self.direction_var = tk.StringVar(
            value="Left → Right"
        )

        self.pattern_var = tk.StringVar(
            value="Row Wise"
        )

        self.reverse_var = tk.BooleanVar(
            value=False
        )

        self.snake_var = tk.BooleanVar(
            value=False
        )

        self.skip_rows_var = tk.StringVar()

        self.skip_cols_var = tk.StringVar()

        self.status_var = tk.StringVar(
            value="No files loaded."
        )

        # ======================================================
        # BUILD GUI
        # ======================================================

        self.build_gui()

    # ==========================================================
    # GUI
    # ==========================================================

    def build_gui(self):

        title = tk.Label(
            self.root,
            text=(
                "REGISTER NUMBER BASED "
                "MULTI-ALLOTMENT SEATING ARRANGEMENT"
            ),
            font=("Arial", 18, "bold")
        )

        title.pack(pady=8)

        # ======================================================
        # FILE FRAME
        # ======================================================

        file_frame = ttk.LabelFrame(
            self.root,
            text="1. Import Register Numbers"
        )

        file_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        ttk.Button(
            file_frame,
            text="Select Excel / PDF Files",
            command=self.select_files
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=6
        )

        ttk.Button(
            file_frame,
            text="Extract Again",
            command=self.extract_again
        ).grid(
            row=0,
            column=1,
            padx=5
        )

        ttk.Button(
            file_frame,
            text="Clear All",
            command=self.clear_all
        ).grid(
            row=0,
            column=2,
            padx=5
        )

        ttk.Label(
            file_frame,
            text=(
                "PDF: Register numbers are extracted ONLY "
                "from detected table cells under a Register Number heading."
            )
        ).grid(
            row=1,
            column=0,
            columnspan=10,
            sticky="w",
            padx=5,
            pady=3
        )

        # ======================================================
        # REGISTER NUMBER LIST
        # ======================================================

        list_frame = ttk.LabelFrame(
            self.root,
            text="2. Imported Register Numbers - Source Order"
        )

        list_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        self.number_list = tk.Listbox(
            list_frame,
            height=7,
            font=("Consolas", 11)
        )

        self.number_list.pack(
            side="left",
            fill="both",
            expand=True,
            padx=5,
            pady=5
        )

        scroll = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.number_list.yview
        )

        scroll.pack(
            side="right",
            fill="y"
        )

        self.number_list.configure(
            yscrollcommand=scroll.set
        )

        # ======================================================
        # SETTINGS
        # ======================================================

        setting_frame = ttk.LabelFrame(
            self.root,
            text="3. Seating Configuration"
        )

        setting_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        # ROWS

        ttk.Label(
            setting_frame,
            text="Rows (1-20):"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )

        tk.Spinbox(
            setting_frame,
            from_=1,
            to=20,
            width=5,
            textvariable=self.rows_var
        ).grid(
            row=0,
            column=1,
            padx=5
        )

        # COLUMNS

        ttk.Label(
            setting_frame,
            text="Columns (1-10):"
        ).grid(
            row=0,
            column=2,
            padx=5
        )

        tk.Spinbox(
            setting_frame,
            from_=1,
            to=10,
            width=5,
            textvariable=self.cols_var
        ).grid(
            row=0,
            column=3,
            padx=5
        )

        # DIRECTION

        ttk.Label(
            setting_frame,
            text="Direction:"
        ).grid(
            row=0,
            column=4,
            padx=5
        )

        direction_box = ttk.Combobox(
            setting_frame,
            textvariable=self.direction_var,
            values=[
                "Left → Right",
                "Right → Left",
                "Top → Bottom",
                "Bottom → Top"
            ],
            state="readonly",
            width=18
        )

        direction_box.grid(
            row=0,
            column=5,
            padx=5
        )

        # PATTERN

        ttk.Label(
            setting_frame,
            text="Pattern:"
        ).grid(
            row=0,
            column=10,
            padx=5
        )

        pattern_box = ttk.Combobox(
            setting_frame,
            textvariable=self.pattern_var,
            values=[
                "Row Wise",
                "Column Wise"
            ],
            state="readonly",
            width=15
        )

        pattern_box.grid(
            row=0,
            column=7,
            padx=5
        )

        # REVERSE

        ttk.Checkbutton(
            setting_frame,
            text="Reverse Register Order",
            variable=self.reverse_var
        ).grid(
            row=0,
            column=8,
            padx=5
        )

        # ZIGZAG

        ttk.Checkbutton(
            setting_frame,
            text="Zig-Zag",
            variable=self.snake_var
        ).grid(
            row=0,
            column=9,
            padx=5
        )

        # SKIP ROWS

        ttk.Label(
            setting_frame,
            text="Skip Rows:"
        ).grid(
            row=1,
            column=0,
            padx=5,
            pady=5
        )

        ttk.Entry(
            setting_frame,
            textvariable=self.skip_rows_var,
            width=18
        ).grid(
            row=1,
            column=1,
            columnspan=2,
            padx=5
        )

        ttk.Label(
            setting_frame,
            text="Example: 3,7-9"
        ).grid(
            row=1,
            column=3,
            padx=5
        )

        # SKIP COLUMNS

        ttk.Label(
            setting_frame,
            text="Skip Columns:"
        ).grid(
            row=1,
            column=4,
            padx=5
        )

        ttk.Entry(
            setting_frame,
            textvariable=self.skip_cols_var,
            width=18
        ).grid(
            row=1,
            column=5,
            columnspan=2,
            padx=5
        )

        ttk.Label(
            setting_frame,
            text="Example: 2,5"
        ).grid(
            row=1,
            column=7,
            padx=5
        )

        # ======================================================
        # BUTTONS
        # ======================================================

        button_frame = tk.Frame(
            self.root
        )

        button_frame.pack(
            fill="x",
            padx=10,
            pady=5
        )

        ttk.Button(
            button_frame,
            text="Create / Refresh Grid",
            command=self.create_grid
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            button_frame,
            text="GENERATE ALL ALLOTMENTS",
            command=self.generate_all_allotments
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            button_frame,
            text="Clear Blocked Seats",
            command=self.clear_blocked
        ).pack(
            side="left",
            padx=5
        )

        ttk.Button(
            button_frame,
            text="EXPORT ALL PAGES TO EXCEL",
            command=self.export_excel
        ).pack(
            side="left",
            padx=5
        )

        # ======================================================
        # ALLOTMENT PAGES / PAGINATION
        # ======================================================

        page_outer = ttk.LabelFrame(
            self.root,
            text="4. Allotment Pages"
        )
        page_outer.pack(
            fill="x",
            padx=10,
            pady=5
        )

        page_nav = tk.Frame(page_outer)
        page_nav.pack(fill="x", padx=5, pady=4)

        ttk.Button(
            page_nav,
            text="<< Previous",
            command=self.previous_page
        ).pack(side="left", padx=3)

        self.page_buttons_frame = tk.Frame(page_nav)
        self.page_buttons_frame.pack(
            side="left",
            fill="x",
            expand=True
        )

        ttk.Button(
            page_nav,
            text="Next >>",
            command=self.next_page
        ).pack(side="left", padx=3)

        tk.Label(
            page_nav,
            textvariable=self.page_info_var,
            font=("Arial", 10, "bold")
        ).pack(side="left", padx=10)

        # ======================================================
        # GRID
        # ======================================================

        self.grid_frame = ttk.LabelFrame(
            self.root,
            text=(
                "5. Seating Grid - "
                "Click any seat to Block / Unblock"
            )
        )

        self.grid_frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        # ------------------------------------------------------
        # SAFE SCROLLABLE SEATING AREA
        # The pagination system remains unchanged.
        # Only the rows/columns inside the current page scroll.
        # ------------------------------------------------------
        self.grid_canvas = tk.Canvas(
            self.grid_frame,
            highlightthickness=0
        )

        self.grid_vscroll = ttk.Scrollbar(
            self.grid_frame,
            orient="vertical",
            command=self.grid_canvas.yview
        )

        self.grid_hscroll = ttk.Scrollbar(
            self.grid_frame,
            orient="horizontal",
            command=self.grid_canvas.xview
        )

        self.grid_canvas.configure(
            yscrollcommand=self.grid_vscroll.set,
            xscrollcommand=self.grid_hscroll.set
        )

        self.grid_vscroll.pack(
            side="right",
            fill="y"
        )

        self.grid_hscroll.pack(
            side="bottom",
            fill="x"
        )

        self.grid_canvas.pack(
            side="left",
            fill="both",
            expand=True
        )

        self.grid_content = tk.Frame(
            self.grid_canvas
        )

        self.grid_window = self.grid_canvas.create_window(
            (0, 0),
            window=self.grid_content,
            anchor="nw"
        )

        self.grid_content.bind(
            "<Configure>",
            self._update_grid_scrollregion
        )

        self.grid_canvas.bind(
            "<Configure>",
            self._fit_grid_width
        )

        # Mouse wheel scrolling only when the pointer is
        # inside the seating-grid area.
        self.grid_canvas.bind(
            "<MouseWheel>",
            self._grid_mousewheel
        )

        self.grid_canvas.bind(
            "<Shift-MouseWheel>",
            self._grid_shift_mousewheel
        )

        # ======================================================
        # STATUS
        # ======================================================

        tk.Label(
            self.root,
            textvariable=self.status_var,
            anchor="w",
            relief="sunken",
            font=("Arial", 10, "bold")
        ).pack(
            fill="x",
            padx=10,
            pady=5
        )

    # ==========================================================
    # NORMALIZE HEADER
    # ==========================================================

    def normalize_header(self, value):

        if value is None:
            return ""

        text = str(value).upper()

        text = text.replace(
            "\n",
            " "
        )

        text = re.sub(
            r"[^A-Z0-9]+",
            " ",
            text
        )

        text = re.sub(
            r"\s+",
            " ",
            text
        ).strip()

        return text

    # ==========================================================
    # CHECK REGISTER HEADER
    # ==========================================================

    def is_register_header(self, value):

        text = self.normalize_header(
            value
        )

        if not text:
            return False

        for pattern in self.REGISTER_HEADER_PATTERNS:

            if re.search(
                r"\b" + pattern + r"\b",
                text
            ):
                return True

        return False

    # ==========================================================
    # VALID REGISTER NUMBER
    # ==========================================================

    def valid_register_number(self, value):

        if value is None:
            return None

        if isinstance(
            value,
            float
        ):

            if pd.isna(value):
                return None

            if value.is_integer():

                value = int(value)

        text = str(
            value
        ).strip()

        if not text:
            return None

        text = text.replace(
            " ",
            ""
        )

        text = text.replace(
            ",",
            ""
        )

        # Register number must be numeric
        # and at least 4 digits.
        #
        # This prevents:
        # page 1
        # S.No 1
        # year 2026
        # etc.
        #
        # from being accepted as register numbers.

        if not re.fullmatch(
            r"\d{4,20}",
            text
        ):
            return None

        return text

    # ==========================================================
    # ADD REGISTER
    # ==========================================================

    def add_register(
        self,
        number,
        filename,
        location,
        position
    ):

        number = self.valid_register_number(
            number
        )

        if not number:
            return

        self.register_numbers.append(
            number
        )

        self.source_records.append(
            {
                "register_number": number,
                "file": filename,
                "location": location,
                "position": position
            }
        )

    # ==========================================================
    # SELECT FILES
    # ==========================================================

    def select_files(self):

        files = filedialog.askopenfilenames(
            title="Select Excel / PDF Files",
            filetypes=[
                (
                    "Excel / PDF",
                    "*.xlsx *.xls *.pdf"
                ),
                (
                    "Excel",
                    "*.xlsx *.xls"
                ),
                (
                    "PDF",
                    "*.pdf"
                )
            ]
        )

        if not files:
            return

        self.selected_files = list(
            files
        )

        self.extract_all()

    # ==========================================================
    # EXTRACT ALL FILES
    # ==========================================================

    def extract_all(self):

        self.register_numbers = []

        self.source_records = []

        for filepath in self.selected_files:

            try:

                extension = os.path.splitext(
                    filepath
                )[1].lower()

                if extension in (
                    ".xlsx",
                    ".xls"
                ):

                    self.extract_excel(
                        filepath
                    )

                elif extension == ".pdf":

                    self.extract_pdf_table_only(
                        filepath
                    )

            except Exception as error:

                messagebox.showwarning(
                    "Extraction Error",
                    "{}\n\n{}".format(
                        os.path.basename(filepath),
                        str(error)
                    )
                )

        self.remove_duplicates()

        self.display_registers()

        self.create_grid()

        self.status_var.set(
            "TOTAL REGISTER NUMBERS: {}".format(
                len(
                    self.register_numbers
                )
            )
        )

        if not self.register_numbers:

            messagebox.showwarning(
                "No Register Numbers",
                (
                    "No register numbers were found.\n\n"
                    "PDF extraction searches ONLY detected "
                    "tables and ONLY the column below a "
                    "Register Number heading."
                )
            )

    # ==========================================================
    # EXCEL EXTRACTION
    # ==========================================================

    def extract_excel(
        self,
        filepath
    ):

        sheets = pd.read_excel(
            filepath,
            sheet_name=None,
            header=None,
            dtype=object
        )

        for sheet_name, df in sheets.items():

            if df.empty:
                continue

            found = False

            header_row = None

            register_col = None

            # --------------------------------------------------
            # Find Register Number heading
            # --------------------------------------------------

            for r in range(
                len(df)
            ):

                for c in range(
                    len(df.columns)
                ):

                    value = df.iloc[
                        r,
                        c
                    ]

                    if self.is_register_header(
                        value
                    ):

                        header_row = r

                        register_col = c

                        found = True

                        break

                if found:
                    break

            if not found:
                continue

            # --------------------------------------------------
            # Extract ONLY same column below heading
            # --------------------------------------------------

            for r in range(
                header_row + 1,
                len(df)
            ):

                value = df.iloc[
                    r,
                    register_col
                ]

                number = self.valid_register_number(
                    value
                )

                if number:

                    self.add_register(
                        number,
                        os.path.basename(
                            filepath
                        ),
                        "Excel Sheet: {}".format(
                            sheet_name
                        ),
                        "Row {}".format(
                            r + 1
                        )
                    )

    # ==========================================================
    # PDF TABLE ONLY
    # ==========================================================

    def extract_pdf_table_only(
        self,
        filepath
    ):

        if pdfplumber is None:

            raise Exception(
                "pdfplumber is not installed.\n\n"
                "Install it with:\n"
                "pip install pdfplumber"
            )

        with pdfplumber.open(
            filepath
        ) as pdf:

            for page_number, page in enumerate(
                pdf.pages,
                start=1
            ):

                tables = []

                # ------------------------------------------------
                # First method - ruled tables
                # ------------------------------------------------

                try:

                    tables = page.extract_tables(
                        {
                            "vertical_strategy":
                                "lines",

                            "horizontal_strategy":
                                "lines",

                            "intersection_tolerance":
                                5,

                            "snap_tolerance":
                                3,

                            "join_tolerance":
                                3
                        }
                    )

                except Exception:
                    tables = []

                # ------------------------------------------------
                # Second method - borderless tables
                #
                # Still table extraction.
                # We DO NOT call page.extract_text().
                # ------------------------------------------------

                if not tables:

                    try:

                        tables = page.extract_tables(
                            {
                                "vertical_strategy":
                                    "text",

                                "horizontal_strategy":
                                    "text",

                                "text_tolerance":
                                    3,

                                "intersection_tolerance":
                                    5
                            }
                        )

                    except Exception:
                        tables = []

                # ------------------------------------------------
                # Process tables only
                # ------------------------------------------------

                for table_number, table in enumerate(
                    tables,
                    start=1
                ):

                    self.process_pdf_table(
                        table,
                        filepath,
                        page_number,
                        table_number
                    )

    # ==========================================================
    # PROCESS PDF TABLE
    # ==========================================================

    def process_pdf_table(
        self,
        table,
        filepath,
        page_number,
        table_number
    ):

        if not table:
            return

        cleaned = []

        for row in table:

            if not row:
                continue

            new_row = []

            for cell in row:

                if cell is None:
                    new_row.append("")
                else:
                    new_row.append(
                        str(cell).strip()
                    )

            if any(
                x != ""
                for x in new_row
            ):

                cleaned.append(
                    new_row
                )

        if not cleaned:
            return

        # ------------------------------------------------------
        # Find register-number heading
        # ------------------------------------------------------

        header_row = None

        register_column = None

        for r, row in enumerate(
            cleaned
        ):

            for c, cell in enumerate(
                row
            ):

                if self.is_register_header(
                    cell
                ):

                    header_row = r

                    register_column = c

                    break

            if register_column is not None:
                break

        # ------------------------------------------------------
        # If table does not contain Register Number heading,
        # completely ignore table.
        # ------------------------------------------------------

        if register_column is None:
            return

        # ------------------------------------------------------
        # Extract ONLY exact column
        # below heading.
        # ------------------------------------------------------

        for r in range(
            header_row + 1,
            len(cleaned)
        ):

            row = cleaned[r]

            if register_column >= len(row):
                continue

            cell = row[
                register_column
            ]

            number = self.extract_number_from_cell(
                cell
            )

            if number:

                self.add_register(
                    number,
                    os.path.basename(
                        filepath
                    ),
                    "PDF Page {} / Table {}".format(
                        page_number,
                        table_number
                    ),
                    "Table Row {}".format(
                        r + 1
                    )
                )

    # ==========================================================
    # EXACT TABLE CELL
    # ==========================================================

    def extract_number_from_cell(
        self,
        cell
    ):

        if cell is None:
            return None

        text = str(
            cell
        ).strip()

        if not text:
            return None

        # Exact cell first

        number = self.valid_register_number(
            text
        )

        if number:
            return number

        # PDF may place newline inside cell

        pieces = re.split(
            r"[\r\n]+",
            text
        )

        for piece in pieces:

            piece = piece.strip()

            number = self.valid_register_number(
                piece
            )

            if number:
                return number

        # Remove harmless punctuation

        cleaned = re.sub(
            r"[^0-9]",
            "",
            text
        )

        return self.valid_register_number(
            cleaned
        )

    # ==========================================================
    # REMOVE DUPLICATES
    # ==========================================================

    def remove_duplicates(self):

        seen = set()

        new_numbers = []

        new_records = []

        for number, record in zip(
            self.register_numbers,
            self.source_records
        ):

            if number in seen:
                continue

            seen.add(
                number
            )

            new_numbers.append(
                number
            )

            new_records.append(
                record
            )

        self.register_numbers = (
            new_numbers
        )

        self.source_records = (
            new_records
        )

    # ==========================================================
    # DISPLAY NUMBERS
    # ==========================================================

    def display_registers(self):

        self.number_list.delete(
            0,
            tk.END
        )

        for i, number in enumerate(
            self.register_numbers,
            start=1
        ):

            self.number_list.insert(
                tk.END,
                "{:4d}. {}".format(
                    i,
                    number
                )
            )

    # ==========================================================
    # EXTRACT AGAIN
    # ==========================================================

    def extract_again(self):

        if not self.selected_files:

            messagebox.showinfo(
                "No Files",
                "Please select Excel/PDF files first."
            )

            return

        self.extract_all()

    # ==========================================================
    # CLEAR
    # ==========================================================

    def clear_all(self):

        self.register_numbers = []

        self.source_records = []

        self.selected_files = []

        self.current_arrangement = []

        self.all_allotments = []

        self.blocked_cells = set()

        self.number_list.delete(
            0,
            tk.END
        )

        self.create_grid()

        self.status_var.set(
            "All data cleared."
        )

    # ==========================================================
    # PARSE SKIP ROW / COLUMN
    # ==========================================================

    def parse_skip_values(
        self,
        text
    ):

        result = set()

        if not text.strip():
            return result

        for item in text.split(","):

            item = item.strip()

            if not item:
                continue

            if "-" in item:

                try:

                    a, b = item.split(
                        "-",
                        1
                    )

                    a = int(
                        a.strip()
                    )

                    b = int(
                        b.strip()
                    )

                    for x in range(
                        min(a, b),
                        max(a, b) + 1
                    ):

                        result.add(x)

                except Exception:
                    pass

            else:

                try:

                    result.add(
                        int(item)
                    )

                except Exception:
                    pass

        return result

    # ==========================================================
    # ALLOTMENT PAGE NAVIGATION
    # ==========================================================

    def rebuild_page_buttons(self):
        for widget in self.page_buttons_frame.winfo_children():
            widget.destroy()

        self.page_buttons = {}

        total_pages = max(1, len(self.all_allotments))

        if self.current_page < 1:
            self.current_page = 1
        if self.current_page > total_pages:
            self.current_page = total_pages

        # Show every page as a visible button.
        # If there are many pages, the frame can scroll horizontally.
        for page_no in range(1, total_pages + 1):
            btn = tk.Button(
                self.page_buttons_frame,
                text="Page {}".format(page_no),
                width=9,
                font=("Arial", 10, "bold"),
                command=lambda p=page_no: self.show_page(p)
            )
            btn.pack(side="left", padx=2)
            self.page_buttons[page_no] = btn

        self.update_page_buttons()

    def update_page_buttons(self):
        total_pages = max(1, len(self.all_allotments))
        self.page_info_var.set(
            "Page {} / {}".format(self.current_page, total_pages)
        )

        for page_no, button in self.page_buttons.items():
            if page_no == self.current_page:
                button.configure(relief="sunken")
            else:
                button.configure(relief="raised")

    def show_page(self, page_number):
        if not self.all_allotments:
            self.current_page = 1
            self.page_info_var.set("Page 1 / 1")
            return

        if page_number < 1:
            page_number = 1

        if page_number > len(self.all_allotments):
            page_number = len(self.all_allotments)

        self.current_page = page_number
        self.current_arrangement = self.all_allotments[
            page_number - 1
        ]

        self.display_current_allotment()
        self.update_page_buttons()

    def previous_page(self):
        if self.all_allotments and self.current_page > 1:
            self.show_page(self.current_page - 1)

    def next_page(self):
        if self.all_allotments and self.current_page < len(self.all_allotments):
            self.show_page(self.current_page + 1)

    # ==========================================================
    # CREATE GRID
    # ==========================================================

    def _update_grid_scrollregion(self, event=None):
        try:
            self.grid_canvas.configure(
                scrollregion=self.grid_canvas.bbox("all")
            )
        except Exception:
            pass

    def _fit_grid_width(self, event=None):
        # Keep the inner frame at least as wide as its contents.
        # This allows horizontal scrolling when columns are wider.
        try:
            requested = self.grid_content.winfo_reqwidth()
            visible = self.grid_canvas.winfo_width()
            self.grid_canvas.itemconfigure(
                self.grid_window,
                width=max(requested, visible)
            )
            self.grid_canvas.configure(
                scrollregion=self.grid_canvas.bbox("all")
            )
        except Exception:
            pass

    def _grid_mousewheel(self, event):
        try:
            self.grid_canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )
        except Exception:
            pass

    def _grid_shift_mousewheel(self, event):
        try:
            self.grid_canvas.xview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )
        except Exception:
            pass

    def create_grid(self):

        try:

            rows = int(
                self.rows_var.get()
            )

            cols = int(
                self.cols_var.get()
            )

        except Exception:

            messagebox.showerror(
                "Invalid Layout",
                "Rows and columns must be numbers."
            )

            return

        if rows < 1 or rows > 20:

            messagebox.showerror(
                "Invalid Rows",
                "Rows must be between 1 and 20."
            )

            return

        if cols < 1 or cols > 10:

            messagebox.showerror(
                "Invalid Columns",
                "Columns must be between 1 and 10."
            )

            return

        # ------------------------------------------------------
        # Preserve manually blocked seats,
        # but remove seats that no longer exist.
        # ------------------------------------------------------

        new_blocked = set()

        for r, c in self.blocked_cells:

            if (
                r <= rows
                and
                c <= cols
            ):

                new_blocked.add(
                    (r, c)
                )

        self.blocked_cells = new_blocked

        # ------------------------------------------------------
        # Clear grid
        # ------------------------------------------------------

        # Remove only the old seating cells.
        # Do NOT destroy the canvas or scrollbars.
        for widget in self.grid_content.winfo_children():
            widget.destroy()

        self.cell_buttons = {}
        self.page_buttons = {}
        self.current_page = 1
        self.page_info_var.set("Page 1 / 1")

        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )

        # ------------------------------------------------------
        # Header
        # ------------------------------------------------------

        tk.Label(
            self.grid_content,
            text="ROW / COL",
            width=13,
            relief="ridge",
            font=("Arial", 10, "bold")
        ).grid(
            row=0,
            column=0,
            padx=2,
            pady=2
        )

        for c in range(
            1,
            cols + 1
        ):

            tk.Label(
                self.grid_content,
                text=(
                    "C{} - SKIP".format(c)
                    if c in skip_cols
                    else "C{}".format(c)
                ),
                width=15,
                relief="ridge",
                font=("Arial", 10, "bold")
            ).grid(
                row=0,
                column=c,
                padx=2,
                pady=2
            )

        # ------------------------------------------------------
        # Seats
        # ------------------------------------------------------

        for r in range(
            1,
            rows + 1
        ):

            tk.Label(
                self.grid_content,
                text=(
                    "ROW {} - SKIP".format(r)
                    if r in skip_rows
                    else "ROW {}".format(r)
                ),
                width=13,
                relief="ridge",
                font=("Arial", 10, "bold")
            ).grid(
                row=r,
                column=0,
                padx=2,
                pady=2
            )

            for c in range(
                1,
                cols + 1
            ):

                button = tk.Button(
                    self.grid_content,
                    text="",
                    width=15,
                    height=3,
                    font=("Arial", 10, "bold"),
                    command=lambda rr=r, cc=c:
                    self.toggle_seat(
                        rr,
                        cc
                    )
                )

                button.grid(
                    row=r,
                    column=c,
                    padx=2,
                    pady=2
                )

                self.cell_buttons[
                    (r, c)
                ] = button

        self.update_grid()
        self.rebuild_page_buttons()

        # If imported data exists,
        # automatically show first allotment.

        if self.register_numbers:

            self.generate_all_allotments(
                show_message=False
            )

    # ==========================================================
    # TOGGLE SEAT
    # ==========================================================

    def toggle_seat(
        self,
        row,
        col
    ):

        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )

        if row in skip_rows:
            return

        if col in skip_cols:
            return

        key = (
            row,
            col
        )

        if key in self.blocked_cells:

            self.blocked_cells.remove(
                key
            )

        else:

            self.blocked_cells.add(
                key
            )

        self.update_grid()

        if self.register_numbers:

            self.generate_all_allotments(
                show_message=False
            )

    # ==========================================================
    # UPDATE GRID
    # ==========================================================

    def update_grid(self):

        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )

        for (
            r,
            c
        ), button in self.cell_buttons.items():

            if r in skip_rows:

                button.configure(
                    text=""
                )

            elif c in skip_cols:

                button.configure(
                    text=""
                )

            elif (
                r,
                c
            ) in self.blocked_cells:

                button.configure(
                    text=""
                )

            else:

                button.configure(
                    text=""
                )

    # ==========================================================
    # CLEAR BLOCKED
    # ==========================================================

    def clear_blocked(self):

        self.blocked_cells = set()

        self.update_grid()

        if self.register_numbers:

            self.generate_all_allotments(
                show_message=False
            )

    # ==========================================================
    # BUILD SEAT ORDER
    # ==========================================================
    #
    # IMPORTANT SKIP LOGIC
    # --------------------
    # A skipped row/column is completely FORGOTTEN when creating
    # the numbering sequence.
    #
    # Example:
    #
    # Rows = 3, Columns = 3
    # Skip Rows = 2
    #
    # Logical usable rows are only:
    #     Row 1
    #     Row 3
    #
    # Therefore the sequence is:
    #
    # Row 1 : 560023  560024  560025
    # Row 3 : 560026  560027  560028
    #
    # If Zig-Zag is ON, the second usable row is reversed:
    #
    # Row 1 : 560023  560024  560025
    # Row 3 : 560028  560027  560026
    #
    # The skipped row never receives a number and does not affect
    # the zig-zag alternation.
    #
    # The same rule applies to skipped columns.
    # ==========================================================

    def build_seat_order(self):

        rows = int(
            self.rows_var.get()
        )

        cols = int(
            self.cols_var.get()
        )

        direction = self.direction_var.get()

        pattern = self.pattern_var.get()

        snake = self.snake_var.get()

        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )

        # ------------------------------------------------------
        # FORGET skipped rows and columns completely.
        #
        # These are logical seat lists.  Their indexes are used
        # for zig-zag, NOT the original physical row/column
        # numbers.  This is the important correction.
        # ------------------------------------------------------

        usable_rows = [
            r
            for r in range(1, rows + 1)
            if r not in skip_rows
        ]

        usable_cols = [
            c
            for c in range(1, cols + 1)
            if c not in skip_cols
        ]

        cells = []

        # ======================================================
        # ROW WISE
        # ======================================================

        if pattern == "Row Wise":

            row_order = list(
                usable_rows
            )

            # Bottom -> Top changes the order of the remaining
            # physical rows only.
            if direction == "Bottom → Top":

                row_order.reverse()

            for logical_row_index, r in enumerate(
                row_order
            ):

                # Start with the remaining columns only.
                col_order = list(
                    usable_cols
                )

                # Right -> Left is the base horizontal direction.
                if direction == "Right → Left":

                    col_order.reverse()

                # Zig-Zag reverses every SECOND USABLE row.
                # A skipped row is not counted.
                if (
                    snake
                    and
                    logical_row_index % 2 == 1
                ):

                    col_order.reverse()

                for c in col_order:

                    # Extra protection in case blocked cells are
                    # present. Blocked cells also consume no number.
                    if (
                        r,
                        c
                    ) in self.blocked_cells:

                        continue

                    cells.append(
                        (
                            r,
                            c
                        )
                    )

        # ======================================================
        # COLUMN WISE
        # ======================================================

        else:

            col_order = list(
                usable_cols
            )

            if direction == "Right → Left":

                col_order.reverse()

            for logical_col_index, c in enumerate(
                col_order
            ):

                row_order = list(
                    usable_rows
                )

                if direction == "Bottom → Top":

                    row_order.reverse()

                # Zig-Zag reverses every SECOND USABLE column.
                # A skipped column is not counted.
                if (
                    snake
                    and
                    logical_col_index % 2 == 1
                ):

                    row_order.reverse()

                for r in row_order:

                    if (
                        r,
                        c
                    ) in self.blocked_cells:

                        continue

                    cells.append(
                        (
                            r,
                            c
                        )
                    )

        return cells

    # ==========================================================
    # GENERATE ALL ALLOTMENTS
    #
    # THIS IS THE IMPORTANT PART
    #
    # It repeatedly fills a new allotment until EVERY
    # register number has been assigned.
    # ==========================================================

    def generate_all_allotments(
        self,
        show_message=True
    ):

        if not self.register_numbers:

            if show_message:

                messagebox.showwarning(
                    "No Register Numbers",
                    "Please import register numbers first."
                )

            return

        # ------------------------------------------------------
        # Build usable seat sequence
        # ------------------------------------------------------

        seat_order = self.build_seat_order()

        usable_seats = len(
            seat_order
        )

        # ------------------------------------------------------
        # IMPORTANT SAFETY CHECK
        # ------------------------------------------------------

        if usable_seats == 0:

            self.all_allotments = []

            self.current_arrangement = []

            self.status_var.set(
                "NO USABLE SEATS. "
                "Cannot allot register numbers."
            )

            if show_message:

                messagebox.showerror(
                    "No Usable Seats",
                    (
                        "All rows/columns are skipped "
                        "or blocked.\n\n"
                        "At least one usable seat is "
                        "required."
                    )
                )

            return

        # ------------------------------------------------------
        # Register numbers
        # ------------------------------------------------------

        numbers = list(
            self.register_numbers
        )

        # Reverse once before creating allotments
        if self.reverse_var.get():

            numbers.reverse()

        # ------------------------------------------------------
        # CLEAR OLD ALLOTMENTS
        # ------------------------------------------------------

        self.all_allotments = []

        # ------------------------------------------------------
        # KEEP CARRYING REMAINING NUMBERS
        # TO NEXT SHEET
        # ------------------------------------------------------

        remaining = list(
            numbers
        )

        allotment_number = 1

        while remaining:

            this_batch = remaining[
                :usable_seats
            ]

            remaining = remaining[
                usable_seats:
            ]

            arrangement = []

            for i, number in enumerate(
                this_batch
            ):

                r, c = seat_order[i]

                arrangement.append(
                    {
                        "allotment":
                            allotment_number,

                        "allotment_order":
                            i + 1,

                        "overall_order":
                            (
                                len(
                                    numbers
                                )
                                -
                                len(
                                    remaining
                                )
                                -
                                len(
                                    this_batch
                                )
                                +
                                i
                                +
                                1
                            ),

                        "register_number":
                            number,

                        "row":
                            r,

                        "column":
                            c,

                        "seat":
                            "R{}C{}".format(
                                r,
                                c
                            )
                    }
                )

            self.all_allotments.append(
                arrangement
            )

            allotment_number += 1

        # ------------------------------------------------------
        # GUI PAGE DISPLAY
        # ------------------------------------------------------

        if self.all_allotments:
            self.current_page = 1
            self.current_arrangement = self.all_allotments[0]
        else:
            self.current_page = 1
            self.current_arrangement = []

        self.rebuild_page_buttons()
        self.display_current_allotment()

        # ------------------------------------------------------
        # COUNTS
        # ------------------------------------------------------

        total = len(
            self.register_numbers
        )

        allotted = sum(
            len(x)
            for x in self.all_allotments
        )

        number_of_allotments = len(
            self.all_allotments
        )

        # ------------------------------------------------------
        # SAFETY CHECK
        # ------------------------------------------------------

        if allotted != total:

            messagebox.showerror(
                "Internal Allotment Error",
                (
                    "The program could not allot all "
                    "register numbers.\n\n"
                    "Total = {}\n"
                    "Allotted = {}"
                ).format(
                    total,
                    allotted
                )
            )

            return

        # ------------------------------------------------------
        # STATUS
        # ------------------------------------------------------

        self.status_var.set(
            "TOTAL REGISTERS: {}   |   "
            "USABLE SEATS / ALLOTMENT: {}   |   "
            "TOTAL ALLOTTED: {}   |   "
            "ALLOTMENT PAGES: {}   |   "
            "STATUS: ALL ALLOTTED".format(
                total,
                usable_seats,
                allotted,
                number_of_allotments
            )
        )

        if show_message:

            messagebox.showinfo(
                "Allotment Completed",
                (
                    "ALL REGISTER NUMBERS HAVE BEEN ALLOTTED.\n\n"
                    "Total Register Numbers : {}\n"
                    "Seats per Allotment    : {}\n"
                    "Number of Allotments   : {}\n"
                    "Total Allotted         : {}\n\n"
                    "No register number remains unallotted."
                ).format(
                    total,
                    usable_seats,
                    number_of_allotments,
                    allotted
                )
            )

    # ==========================================================
    # DISPLAY FIRST ALLOTMENT
    # ==========================================================

    def display_current_allotment(self):

        mapping = {}

        for item in self.current_arrangement:

            mapping[
                (
                    item["row"],
                    item["column"]
                )
            ] = item[
                "register_number"
            ]

        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )

        for (
            r,
            c
        ), button in self.cell_buttons.items():

            if r in skip_rows:

                button.configure(
                    text=""
                )

            elif c in skip_cols:

                button.configure(
                    text=""
                )

            elif (
                r,
                c
            ) in self.blocked_cells:

                button.configure(
                    text=""
                )

            elif (
                r,
                c
            ) in mapping:

                button.configure(
                    text=str(
                        mapping[
                            (r, c)
                        ]
                    )
                )

            else:

                button.configure(
                    text=""
                )

    # ==========================================================
    # EXPORT EXCEL
    # ==========================================================

    def export_excel(self):

        if not self.register_numbers:

            messagebox.showwarning(
                "No Data",
                "Please import register numbers first."
            )

            return

        # Always regenerate
        self.generate_all_allotments(
            show_message=False
        )

        if not self.all_allotments:

            messagebox.showerror(
                "No Allotment",
                "No usable seats are available."
            )

            return

        filename = filedialog.asksaveasfilename(
            title="Save All Seating Allotments",
            defaultextension=".xlsx",
            filetypes=[
                (
                    "Excel Workbook",
                    "*.xlsx"
                )
            ]
        )

        if not filename:
            return

        wb = Workbook()

        # ======================================================
        # CREATE ONE SHEET FOR EVERY ALLOTMENT
        # ======================================================

        for allotment_index, arrangement in enumerate(
            self.all_allotments,
            start=1
        ):

            if allotment_index == 1:

                ws = wb.active

                ws.title = (
                    "Allotment 1"
                )

            else:

                ws = wb.create_sheet(
                    "Allotment {}".format(
                        allotment_index
                    )
                )

            self.write_allotment_sheet(
                ws,
                allotment_index,
                arrangement
            )

        # ======================================================
        # COMPLETE CANDIDATE LIST
        # ======================================================

        ws_all = wb.create_sheet(
            "Complete Candidate List"
        )

        ws_all.append(
            [
                "Overall Order",
                "Register Number",
                "Allotment",
                "Row",
                "Column",
                "Seat"
            ]
        )

        for cell in ws_all[1]:

            cell.font = Font(
                bold=True
            )

        overall = 1

        for allotment_index, arrangement in enumerate(
            self.all_allotments,
            start=1
        ):

            for item in arrangement:

                ws_all.append(
                    [
                        overall,
                        item[
                            "register_number"
                        ],
                        "Allotment {}".format(
                            allotment_index
                        ),
                        item["row"],
                        item["column"],
                        item["seat"]
                    ]
                )

                overall += 1

        # ======================================================
        # IMPORTED REGISTERS
        # ======================================================

        ws_source = wb.create_sheet(
            "Imported Registers"
        )

        ws_source.append(
            [
                "Source Order",
                "Register Number",
                "File",
                "Location",
                "Position"
            ]
        )

        for cell in ws_source[1]:

            cell.font = Font(
                bold=True
            )

        for i, record in enumerate(
            self.source_records,
            start=1
        ):

            ws_source.append(
                [
                    i,
                    record[
                        "register_number"
                    ],
                    record[
                        "file"
                    ],
                    record[
                        "location"
                    ],
                    record[
                        "position"
                    ]
                ]
            )

        # ======================================================
        # SETTINGS
        # ======================================================

        ws_settings = wb.create_sheet(
            "Settings"
        )

        rows = int(
            self.rows_var.get()
        )

        cols = int(
            self.cols_var.get()
        )

        usable_seats = len(
            self.build_seat_order()
        )

        settings = [

            (
                "Total Register Numbers",
                len(
                    self.register_numbers
                )
            ),

            (
                "Rows",
                rows
            ),

            (
                "Columns",
                cols
            ),

            (
                "Usable Seats per Allotment",
                usable_seats
            ),

            (
                "Number of Allotments",
                len(
                    self.all_allotments
                )
            ),

            (
                "Total Allotted",
                sum(
                    len(x)
                    for x in self.all_allotments
                )
            ),

            (
                "Direction",
                self.direction_var.get()
            ),

            (
                "Pattern",
                self.pattern_var.get()
            ),

            (
                "Reverse Register Order",
                "YES"
                if self.reverse_var.get()
                else "NO"
            ),

            (
                "Zig-Zag",
                "YES"
                if self.snake_var.get()
                else "NO"
            ),

            (
                "Skipped Rows",
                self.skip_rows_var.get()
            ),

            (
                "Skipped Columns",
                self.skip_cols_var.get()
            ),

            (
                "Final Status",
                "ALL REGISTER NUMBERS ALLOTTED"
            )
        ]

        for r, (
            name,
            value
        ) in enumerate(
            settings,
            start=1
        ):

            ws_settings.cell(
                row=r,
                column=1,
                value=name
            )

            ws_settings.cell(
                row=r,
                column=2,
                value=value
            )

        # ======================================================
        # FORMATTING ALL SHEETS
        # ======================================================

        thin = Side(
            style="thin"
        )

        border = Border(
            left=thin,
            right=thin,
            top=thin,
            bottom=thin
        )

        for ws in wb.worksheets:

            for row in ws.iter_rows():

                for cell in row:

                    cell.alignment = Alignment(
                        horizontal="center",
                        vertical="center"
                    )

                    cell.border = border

            # Set reasonable widths

            for column_cells in ws.columns:

                try:

                    column_letter = get_column_letter(
                        column_cells[0].column
                    )

                    ws.column_dimensions[
                        column_letter
                    ].width = 20

                except Exception:
                    pass

        # ======================================================
        # SAVE
        # ======================================================

        try:

            wb.save(
                filename
            )

            total = len(
                self.register_numbers
            )

            allotted = sum(
                len(x)
                for x in self.all_allotments
            )

            messagebox.showinfo(
                "Excel Export Completed",
                (
                    "SEATING ARRANGEMENT EXPORTED SUCCESSFULLY.\n\n"
                    "Total Register Numbers : {}\n"
                    "Total Allotted         : {}\n"
                    "Allotment Sheets       : {}\n\n"
                    "Every register number has been allotted.\n\n"
                    "File:\n{}"
                ).format(
                    total,
                    allotted,
                    len(
                        self.all_allotments
                    ),
                    filename
                )
            )

        except Exception as error:

            messagebox.showerror(
                "Excel Export Error",
                str(error)
            )

     # ==========================================================
    # WRITE ONE ALLOTMENT SHEET
    # ==========================================================

    def write_allotment_sheet(
        self,
        ws,
        allotment_number,
        arrangement
    ):

        rows = int(
            self.rows_var.get()
        )

        cols = int(
            self.cols_var.get()
        )

        # --- CORRECTION: Parse skip values at the very beginning ---
        skip_rows = self.parse_skip_values(
            self.skip_rows_var.get()
        )

        skip_cols = self.parse_skip_values(
            self.skip_cols_var.get()
        )
        # -----------------------------------------------------------

        ws["A1"] = (
            "SEATING ALLOTMENT {}".format(
                allotment_number
            )
        )

        ws["A1"].font = Font(
            bold=True,
            size=16
        )

        ws.merge_cells(
            start_row=1,
            start_column=1,
            end_row=1,
            end_column=cols + 1
        )

        # ------------------------------------------------------
        # Summary
        # ------------------------------------------------------

        ws["A2"] = "Candidates"

        ws["B2"] = len(
            arrangement
        )

        ws["C2"] = "Rows"

        ws["D2"] = rows

        ws["E2"] = "Columns"

        ws["F2"] = cols

        ws["G2"] = "Allotment"

        ws["H2"] = allotment_number

        # ------------------------------------------------------
        # Grid header
        # ------------------------------------------------------

        ws.cell(
            row=4,
            column=1,
            value="ROW / COL"
        )

        for c in range(
            1,
            cols + 1
        ):

            ws.cell(
                row=4,
                column=c + 1,
                value=(
                    "C{} - SKIP".format(c)
                    if c in skip_cols
                    else "C{}".format(c)
                )
            )

        # ------------------------------------------------------
        # Mapping
        # ------------------------------------------------------

        mapping = {}

        for item in arrangement:

            mapping[
                (
                    item["row"],
                    item["column"]
                )
            ] = item[
                "register_number"
            ]

        # ------------------------------------------------------
        # Write grid
        # ------------------------------------------------------

        for r in range(
            1,
            rows + 1
        ):

            ws.cell(
                row=r + 4,
                column=1,
                value=(
                    "ROW {} - SKIP".format(r)
                    if r in skip_rows
                    else "ROW {}".format(r)
                )
            )

            for c in range(
                1,
                cols + 1
            ):

                cell = ws.cell(
                    row=r + 4,
                    column=c + 1
                )

                if r in skip_rows:

                    cell.value = None

                elif c in skip_cols:

                    cell.value = None

                elif (
                    r,
                    c
                ) in self.blocked_cells:

                    cell.value = None

                elif (
                    r,
                    c
                ) in mapping:

                    cell.value = mapping[
                        (r, c)
                    ]

                else:

                    cell.value = None

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

        # ------------------------------------------------------
        # Candidate list below grid
        # ------------------------------------------------------

        start_row = rows + 8

        ws.cell(
            row=start_row,
            column=1,
            value="Allotment Order"
        )

        ws.cell(
            row=start_row,
            column=2,
            value="Register Number"
        )

        ws.cell(
            row=start_row,
            column=3,
            value="Row"
        )

        ws.cell(
            row=start_row,
            column=4,
            value="Column"
        )

        ws.cell(
            row=start_row,
            column=5,
            value="Seat"
        )

        for cell in ws[start_row]:

            if cell.column <= 5:

                cell.font = Font(
                    bold=True
                )

        for i, item in enumerate(
            arrangement,
            start=1
        ):

            rr = start_row + i

            ws.cell(
                row=rr,
                column=1,
                value=i
            )

            ws.cell(
                row=rr,
                column=2,
                value=item[
                    "register_number"
                ]
            )

            ws.cell(
                row=rr,
                column=3,
                value=item["row"]
            )

            ws.cell(
                row=rr,
                column=4,
                value=item["column"]
            )

            ws.cell(
                row=rr,
                column=5,
                value=item["seat"]
            )

# ==============================================================
# MAIN
# ==============================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = SeatingArrangementApp(
        root
    )

    root.mainloop()