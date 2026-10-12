"""Student Attendance Management System

A simple Tkinter mini project that stores attendance records in an Excel workbook.
Run with: python "student attandannce system.py"
Demo login: onkar@69 / admin123
"""

from __future__ import annotations

import hashlib
import os
import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox, ttk

from openpyxl import Workbook, load_workbook

APP_TITLE = "Student Attendance Management System"
DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attendance_records.xlsx")
SHEET_NAME = "Attendance"
HEADERS = ["Record ID", "Student ID", "Student Name", "Class / Section", "Subject", "Date", "Status", "Remarks"]
STUDENT_SHEET = "Students"
STUDENT_HEADERS = ["Student ID", "Student Name", "Class / Section", "Phone / Email"]
SAMPLE_STUDENTS = [
    {"Student ID": "STU1001", "Student Name": "Aarav Sharma", "Class / Section": "10-A", "Phone / Email": "aarav@example.com"},
    {"Student ID": "STU1002", "Student Name": "Diya Patel", "Class / Section": "10-A", "Phone / Email": "diya@example.com"},
    {"Student ID": "STU1003", "Student Name": "Kabir Mehta", "Class / Section": "10-B", "Phone / Email": "kabir@example.com"},
    {"Student ID": "STU1004", "Student Name": "Anaya Singh", "Class / Section": "10-B", "Phone / Email": "anaya@example.com"},
    {"Student ID": "STU1005", "Student Name": "Ishaan Rao", "Class / Section": "9-A", "Phone / Email": "ishaan@example.com"},
]

# Demonstration credentials. Change these before using the project beyond a classroom demo.
USERNAME = "onkar@69"
PASSWORD_SHA256 = hashlib.sha256(b"admin123").hexdigest()

COLORS = {
    "navy": "#17324D",
    "blue": "#2864A5",
    "light": "#F2F6FA",
    "white": "#FFFFFF",
    "text": "#203040",
    "muted": "#64748B",
    "green": "#198754",
    "red": "#C0392B",
    "line": "#D8E1EA",
}


def ensure_workbook() -> None:
    """Create the Excel file and header row the first time the app is launched."""
    if not os.path.exists(DATA_FILE):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = SHEET_NAME
        sheet.append(HEADERS)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = f"A1:{__import__('openpyxl').utils.get_column_letter(len(HEADERS))}1"
        for cell in sheet[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="17324D")
        widths = [13, 15, 25, 20, 22, 15, 14, 32]
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[__import__("openpyxl").utils.get_column_letter(index)].width = width
        students = workbook.create_sheet(STUDENT_SHEET)
        students.append(STUDENT_HEADERS)
        students.freeze_panes = "A2"
        students.auto_filter.ref = "A1:D1"
        for cell in students[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor="17324D")
        for index, width in enumerate([16, 28, 20, 24], start=1):
            students.column_dimensions[__import__("openpyxl").utils.get_column_letter(index)].width = width
        for student in SAMPLE_STUDENTS:
            students.append([student[header] for header in STUDENT_HEADERS])
        for student_id, name, section, subject, status in [
            ("STU1001", "Aarav Sharma", "10-A", "Mathematics", "Present"),
            ("STU1002", "Diya Patel", "10-A", "Science", "Present"),
            ("STU1003", "Kabir Mehta", "10-B", "English", "Absent"),
            ("STU1004", "Anaya Singh", "10-B", "Mathematics", "Late"),
            ("STU1005", "Ishaan Rao", "9-A", "Science", "Present"),
        ]:
            sheet.append([str(sheet.max_row), student_id, name, section, subject,
                          date.today().isoformat(), status, "Sample demonstration record"])
        workbook.save(DATA_FILE)
    else:
        workbook = load_workbook(DATA_FILE)
        if SHEET_NAME not in workbook.sheetnames:
            sheet = workbook.create_sheet(SHEET_NAME)
            sheet.append(HEADERS)
            workbook.save(DATA_FILE)
        else:
            sheet = workbook[SHEET_NAME]
            old_headers = [str(cell.value or "") for cell in sheet[1]]
            if old_headers != HEADERS:
                old_rows = list(sheet.iter_rows(min_row=2, values_only=True))
                old_index = {header: index for index, header in enumerate(old_headers)}
                saved_rows = []
                for row in old_rows:
                    values = list(row)
                    if not any(value is not None for value in values):
                        continue
                    saved_rows.append([values[old_index[header]] if header in old_index and old_index[header] < len(values) else ""
                                       for header in HEADERS])
                sheet.delete_rows(1, sheet.max_row)
                sheet.append(HEADERS)
                for row in saved_rows:
                    sheet.append(row)
                sheet.freeze_panes = "A2"
                sheet.auto_filter.ref = f"A1:{__import__('openpyxl').utils.get_column_letter(len(HEADERS))}{max(sheet.max_row, 1)}"
                workbook.save(DATA_FILE)
        if STUDENT_SHEET not in workbook.sheetnames:
            students = workbook.create_sheet(STUDENT_SHEET)
            students.append(STUDENT_HEADERS)
            students.freeze_panes = "A2"
            students.auto_filter.ref = "A1:D1"
        students = workbook[STUDENT_SHEET]
        if students.max_row == 1:
            for student in SAMPLE_STUDENTS:
                students.append([student[header] for header in STUDENT_HEADERS])
        if sheet.max_row == 1:
            for student_id, name, section, subject, status in [
                ("STU1001", "Aarav Sharma", "10-A", "Mathematics", "Present"),
                ("STU1002", "Diya Patel", "10-A", "Science", "Present"),
                ("STU1003", "Kabir Mehta", "10-B", "English", "Absent"),
                ("STU1004", "Anaya Singh", "10-B", "Mathematics", "Late"),
                ("STU1005", "Ishaan Rao", "9-A", "Science", "Present"),
            ]:
                sheet.append([str(sheet.max_row), student_id, name, section, subject,
                              date.today().isoformat(), status, "Sample demonstration record"])
        workbook.save(DATA_FILE)
        workbook.close()


def read_students() -> list[dict[str, str]]:
    ensure_workbook()
    workbook = load_workbook(DATA_FILE, data_only=True)
    sheet = workbook[STUDENT_SHEET]
    students = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or all(value is None for value in row):
            continue
        values = list(row) + [None] * (len(STUDENT_HEADERS) - len(row))
        students.append({header: str(values[index] or "") for index, header in enumerate(STUDENT_HEADERS)})
    workbook.close()
    return students


def write_students(students: list[dict[str, str]]) -> None:
    ensure_workbook()
    workbook = load_workbook(DATA_FILE)
    sheet = workbook[STUDENT_SHEET]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row)
    for student in students:
        sheet.append([student.get(header, "") for header in STUDENT_HEADERS])
    workbook.save(DATA_FILE)
    workbook.close()


def read_records() -> list[dict[str, str]]:
    """Read attendance rows from Excel, returning display-friendly values."""
    ensure_workbook()
    workbook = load_workbook(DATA_FILE, data_only=True)
    sheet = workbook[SHEET_NAME]
    records: list[dict[str, str]] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if not row or all(value is None for value in row):
            continue
        values = list(row) + [None] * (len(HEADERS) - len(row))
        records.append({
            "Record ID": str(values[0] or ""),
            "Student ID": str(values[1] or ""),
            "Student Name": str(values[2] or ""),
            "Class / Section": str(values[3] or ""),
            "Subject": str(values[4] or ""),
            "Date": values[5].strftime("%Y-%m-%d") if isinstance(values[5], (date, datetime)) else str(values[5] or ""),
            "Status": str(values[6] or ""),
            "Remarks": str(values[7] or ""),
        })
    workbook.close()
    return records


def write_records(records: list[dict[str, str]]) -> None:
    """Replace the worksheet data with the supplied records."""
    ensure_workbook()
    workbook = load_workbook(DATA_FILE)
    sheet = workbook[SHEET_NAME]
    if sheet.max_row > 1:
        sheet.delete_rows(2, sheet.max_row)
    for record in records:
        sheet.append([record.get(header, "") for header in HEADERS])
    workbook.save(DATA_FILE)
    workbook.close()


class AttendanceApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x730")
        self.minsize(940, 630)
        self.configure(bg=COLORS["light"])
        self.selected_record_id: str | None = None
        self.search_var = tk.StringVar()
        self.status_filter = tk.StringVar(value="All statuses")
        self._configure_style()
        ensure_workbook()
        self.show_login()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground=COLORS["text"],
                        rowheight=30, fieldbackground="white", bordercolor=COLORS["line"])
        style.configure("Treeview.Heading", background=COLORS["navy"], foreground="white",
                        font=("Segoe UI", 9, "bold"), padding=8)
        style.map("Treeview", background=[("selected", "#DCEBFA")], foreground=[("selected", COLORS["text"])])
        style.configure("TCombobox", padding=6)

    def _clear_window(self) -> None:
        for child in self.winfo_children():
            child.destroy()

    def show_login(self) -> None:
        self._clear_window()
        self.geometry("520x500")
        outer = tk.Frame(self, bg=COLORS["light"])
        outer.pack(fill="both", expand=True)
        card = tk.Frame(outer, bg=COLORS["white"], padx=42, pady=38,
                        highlightbackground=COLORS["line"], highlightthickness=1)
        card.place(relx=0.5, rely=0.5, anchor="center", width=410, height=410)
        tk.Label(card, text="ATTENDANCE", bg=COLORS["white"], fg=COLORS["blue"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        tk.Label(card, text="Sign in", bg=COLORS["white"], fg=COLORS["navy"],
                 font=("Segoe UI", 25, "bold")).pack(anchor="w", pady=(5, 3))
        tk.Label(card, text="Student Attendance Management System", bg=COLORS["white"],
                 fg=COLORS["muted"], font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 24))

        tk.Label(card, text="Username", bg=COLORS["white"], fg=COLORS["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        username = ttk.Entry(card, font=("Segoe UI", 11))
        username.pack(fill="x", ipady=7, pady=(6, 16))
        tk.Label(card, text="Password", bg=COLORS["white"], fg=COLORS["text"],
                 font=("Segoe UI", 10, "bold")).pack(anchor="w")
        password = ttk.Entry(card, show="•", font=("Segoe UI", 11))
        password.pack(fill="x", ipady=7, pady=(6, 20))

        def login(_event: object = None) -> None:
            entered_user = username.get().strip()
            entered_hash = hashlib.sha256(password.get().encode("utf-8")).hexdigest()
            if entered_user == USERNAME and entered_hash == PASSWORD_SHA256:
                self.show_dashboard()
            else:
                messagebox.showerror("Login failed", "The username or password is incorrect.", parent=self)
                password.delete(0, "end")
                password.focus_set()

        tk.Button(card, text="Login", command=login, bg=COLORS["blue"], fg="white",
                  activebackground=COLORS["navy"], activeforeground="white", relief="flat",
                  font=("Segoe UI", 11, "bold"), cursor="hand2").pack(fill="x", ipady=9)
        tk.Label(card, text="Demo account: onkar@69  |  Password: admin123", bg=COLORS["white"],
                 fg=COLORS["muted"], font=("Segoe UI", 9)).pack(pady=(16, 0))
        password.bind("<Return>", login)
        username.bind("<Return>", lambda _e: password.focus_set())
        username.focus_set()

    def show_dashboard(self, page: str = "attendance") -> None:
        self._clear_window()
        self.geometry("1120x730")
        header = tk.Frame(self, bg=COLORS["navy"], padx=24, pady=16)
        header.pack(fill="x")
        title_area = tk.Frame(header, bg=COLORS["navy"])
        title_area.pack(side="left")
        tk.Label(title_area, text="Student Attendance", bg=COLORS["navy"], fg="white",
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(title_area, text="Management System  •  Excel-backed register", bg=COLORS["navy"],
                 fg="#C8D7E6", font=("Segoe UI", 9)).pack(anchor="w", pady=(2, 0))
        nav = tk.Frame(self, bg=COLORS["white"], padx=22, pady=11,
                       highlightbackground=COLORS["line"], highlightthickness=1)
        nav.pack(fill="x")
        menu = [("Add student", lambda: self.show_dashboard("student_add")),
                ("View / Search students", lambda: self.show_dashboard("students")),
                ("Delete student", lambda: self.show_dashboard("student_delete")),
                ("Attendance", lambda: self.show_dashboard("attendance"))]
        for label, command in menu:
            tk.Button(nav, text=label, command=command, bg=COLORS["white"], fg=COLORS["blue"],
                      activebackground=COLORS["light"], relief="flat", font=("Segoe UI", 10, "bold"),
                      cursor="hand2", padx=12).pack(side="left", padx=(0, 8))
        tk.Button(nav, text="Logout", command=self.logout, bg="#FCEAE8", fg=COLORS["red"],
                  activebackground="#F8D7D4", relief="flat", font=("Segoe UI", 10, "bold"),
                  cursor="hand2", padx=12).pack(side="right")

        body = tk.Frame(self, bg=COLORS["light"], padx=22, pady=18)
        body.pack(fill="both", expand=True)
        self.page_host = tk.Frame(body, bg=COLORS["light"])
        self.page_host.pack(fill="both", expand=True)
        if page == "attendance":
            self._build_summary(self.page_host)
            content = tk.Frame(self.page_host, bg=COLORS["light"])
            content.pack(fill="both", expand=True, pady=(17, 0))
            self._build_form(content)
            self._build_records(content)
            self.refresh_records()
        elif page == "student_add":
            self._show_student_add()
        else:
            self._show_student_directory(delete_mode=page == "student_delete")

    def _show_student_add(self) -> None:
        panel = tk.Frame(self.page_host, bg="white", padx=28, pady=24,
                         highlightbackground=COLORS["line"], highlightthickness=1)
        panel.place(relx=0.5, rely=0.08, relwidth=0.58, anchor="n")
        tk.Label(panel, text="Add student", bg="white", fg=COLORS["navy"],
                 font=("Segoe UI", 18, "bold")).pack(anchor="w")
        tk.Label(panel, text="Student details are saved to the Students sheet in the Excel workbook.",
                 bg="white", fg=COLORS["muted"], font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 20))
        self.student_vars = {key: tk.StringVar() for key in ("id", "name", "class", "contact")}
        for label, key in [("Student ID *", "id"), ("Student name *", "name"),
                           ("Class / Section *", "class"), ("Phone / Email", "contact")]:
            tk.Label(panel, text=label, bg="white", fg=COLORS["text"],
                     font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 5))
            ttk.Entry(panel, textvariable=self.student_vars[key], font=("Segoe UI", 11)).pack(
                fill="x", ipady=6, pady=(0, 13))
        tk.Button(panel, text="Save student", command=self.save_student, bg=COLORS["blue"], fg="white",
                  relief="flat", font=("Segoe UI", 10, "bold"), cursor="hand2", padx=15, pady=9).pack(anchor="w")
        tk.Label(panel, text="* Required fields", bg="white", fg=COLORS["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(12, 0))

    def save_student(self) -> None:
        student = {
            "Student ID": self.student_vars["id"].get().strip(),
            "Student Name": self.student_vars["name"].get().strip(),
            "Class / Section": self.student_vars["class"].get().strip(),
            "Phone / Email": self.student_vars["contact"].get().strip(),
        }
        if not all(student[key] for key in ("Student ID", "Student Name", "Class / Section")):
            messagebox.showwarning("Missing information", "Enter the student ID, name, and class / section.", parent=self)
            return
        students = read_students()
        if any(item["Student ID"].casefold() == student["Student ID"].casefold() for item in students):
            messagebox.showwarning("Student already exists", "A student with this ID is already in the register.", parent=self)
            return
        students.append(student)
        try:
            write_students(students)
        except PermissionError:
            messagebox.showerror("Workbook is open", "Close attendance_records.xlsx in Excel, then try again.", parent=self)
            return
        for variable in self.student_vars.values():
            variable.set("")
        messagebox.showinfo("Saved", "Student added to the Excel register.", parent=self)

    def _show_student_directory(self, delete_mode: bool = False) -> None:
        panel = tk.Frame(self.page_host, bg="white", padx=20, pady=18,
                         highlightbackground=COLORS["line"], highlightthickness=1)
        panel.pack(fill="both", expand=True)
        heading = "Delete student" if delete_mode else "View / Search students"
        tk.Label(panel, text=heading, bg="white", fg=COLORS["navy"],
                 font=("Segoe UI", 17, "bold")).pack(anchor="w")
        hint = ("Select a student and confirm to remove them from the student register. Attendance history is retained."
                if delete_mode else "Search by student ID, name, class, or contact details.")
        tk.Label(panel, text=hint, bg="white", fg=COLORS["muted"],
                 font=("Segoe UI", 10)).pack(anchor="w", pady=(4, 14))
        toolbar = tk.Frame(panel, bg="white")
        toolbar.pack(fill="x", pady=(0, 12))
        self.student_search_var = tk.StringVar()
        search = ttk.Entry(toolbar, textvariable=self.student_search_var, font=("Segoe UI", 10))
        search.pack(side="left", fill="x", expand=True, ipady=5)
        search.bind("<KeyRelease>", lambda _event: self.refresh_students())
        tk.Button(toolbar, text="Refresh", command=self.refresh_students, bg="#E8EEF4", fg=COLORS["text"],
                  relief="flat", padx=12, pady=6, cursor="hand2").pack(side="left", padx=8)
        self.student_tree = ttk.Treeview(panel, columns=STUDENT_HEADERS, show="headings", selectmode="browse")
        widths = [150, 240, 180, 220]
        for header, width in zip(STUDENT_HEADERS, widths):
            self.student_tree.heading(header, text=header)
            self.student_tree.column(header, width=width, minwidth=90, anchor="w")
        yscroll = ttk.Scrollbar(panel, orient="vertical", command=self.student_tree.yview)
        self.student_tree.configure(yscrollcommand=yscroll.set)
        self.student_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="left", fill="y")
        if delete_mode:
            tk.Button(panel, text="Delete selected student", command=self.delete_student,
                      bg="#FCEAE8", fg=COLORS["red"], relief="flat", padx=14, pady=9,
                      font=("Segoe UI", 10, "bold"), cursor="hand2").pack(side="bottom", anchor="e", pady=(12, 0))
        self.refresh_students()

    def refresh_students(self) -> None:
        if not hasattr(self, "student_tree") or not self.student_tree.winfo_exists():
            return
        query = self.student_search_var.get().strip().casefold()
        for item in self.student_tree.get_children():
            self.student_tree.delete(item)
        for student in read_students():
            if query and query not in " ".join(student.values()).casefold():
                continue
            self.student_tree.insert("", "end", values=[student[header] for header in STUDENT_HEADERS])

    def delete_student(self) -> None:
        selection = self.student_tree.selection()
        if not selection:
            messagebox.showinfo("Select a student", "Select a student row before deleting.", parent=self)
            return
        student_id = str(self.student_tree.item(selection[0], "values")[0])
        if not messagebox.askyesno("Confirm deletion", f"Delete student {student_id} from the student register? Attendance history will be kept.", parent=self):
            return
        remaining = [student for student in read_students() if student["Student ID"] != student_id]
        try:
            write_students(remaining)
        except PermissionError:
            messagebox.showerror("Workbook is open", "Close attendance_records.xlsx in Excel, then try again.", parent=self)
            return
        self.refresh_students()
        messagebox.showinfo("Deleted", "Student removed from the student register.", parent=self)

    def _build_summary(self, parent: tk.Widget) -> None:
        row = tk.Frame(parent, bg=COLORS["light"])
        row.pack(fill="x")
        self.total_label = self._summary_card(row, "TOTAL RECORDS", "0")
        self.present_label = self._summary_card(row, "PRESENT TODAY", "0")
        self.absent_label = self._summary_card(row, "ABSENT TODAY", "0")

    def _summary_card(self, parent: tk.Widget, heading: str, value: str) -> tk.Label:
        card = tk.Frame(parent, bg=COLORS["white"], padx=17, pady=12,
                        highlightbackground=COLORS["line"], highlightthickness=1)
        card.pack(side="left", fill="x", expand=True, padx=(0, 12))
        tk.Label(card, text=heading, bg=COLORS["white"], fg=COLORS["muted"],
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        label = tk.Label(card, text=value, bg=COLORS["white"], fg=COLORS["navy"],
                         font=("Segoe UI", 20, "bold"))
        label.pack(anchor="w", pady=(3, 0))
        return label

    def _build_form(self, parent: tk.Widget) -> None:
        panel = tk.Frame(parent, bg=COLORS["white"], padx=18, pady=17,
                         highlightbackground=COLORS["line"], highlightthickness=1)
        panel.pack(side="left", fill="y", padx=(0, 16))
        panel.configure(width=315)
        panel.pack_propagate(False)
        tk.Label(panel, text="Attendance details", bg="white", fg=COLORS["navy"],
                 font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(panel, text="Add a new record or select one to edit.", bg="white", fg=COLORS["muted"],
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(3, 15))

        self.form_vars = {key: tk.StringVar() for key in
                          ["student_id", "student_name", "class_section", "subject", "record_date", "status", "remarks"]}
        fields = [
            ("Student ID *", "student_id", "entry"),
            ("Student name *", "student_name", "entry"),
            ("Class / Section *", "class_section", "entry"),
            ("Subject *", "subject", "entry"),
            ("Date (YYYY-MM-DD) *", "record_date", "entry"),
            ("Status *", "status", "combo"),
            ("Remarks", "remarks", "entry"),
        ]
        for label, key, kind in fields:
            tk.Label(panel, text=label, bg="white", fg=COLORS["text"],
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
            if kind == "combo":
                widget = ttk.Combobox(panel, textvariable=self.form_vars[key], values=["Present", "Absent", "Late", "Excused", "Other"],
                                      state="readonly", font=("Segoe UI", 10))
                widget.pack(fill="x", pady=(0, 10), ipady=3)
            else:
                widget = ttk.Entry(panel, textvariable=self.form_vars[key], font=("Segoe UI", 10))
                widget.pack(fill="x", pady=(0, 10), ipady=4)
        self.form_vars["record_date"].set(date.today().isoformat())
        self.form_vars["status"].set("Present")
        actions = tk.Frame(panel, bg="white")
        actions.pack(fill="x", pady=(4, 0))
        tk.Button(actions, text="Save record", command=self.save_record, bg=COLORS["blue"], fg="white",
                  relief="flat", padx=10, pady=8, cursor="hand2").pack(side="left", fill="x", expand=True, padx=(0, 5))
        tk.Button(actions, text="Update", command=self.update_record, bg=COLORS["green"], fg="white",
                  relief="flat", padx=10, pady=8, cursor="hand2").pack(side="left", fill="x", expand=True, padx=5)
        tk.Button(actions, text="Clear", command=self.clear_form, bg="#E8EEF4", fg=COLORS["text"],
                  relief="flat", padx=10, pady=8, cursor="hand2").pack(side="left", fill="x", expand=True, padx=(5, 0))
        tk.Button(panel, text="Delete selected record", command=self.delete_record, bg="#FCEAE8", fg=COLORS["red"],
                  relief="flat", padx=10, pady=8, cursor="hand2").pack(fill="x", pady=(9, 0))
        tk.Label(panel, text="* Required fields", bg="white", fg=COLORS["muted"],
                 font=("Segoe UI", 8)).pack(anchor="w", side="bottom", pady=(8, 0))

    def _build_records(self, parent: tk.Widget) -> None:
        panel = tk.Frame(parent, bg=COLORS["white"], padx=16, pady=16,
                         highlightbackground=COLORS["line"], highlightthickness=1)
        panel.pack(side="left", fill="both", expand=True)
        toolbar = tk.Frame(panel, bg="white")
        toolbar.pack(fill="x", pady=(0, 12))
        tk.Label(toolbar, text="Attendance records", bg="white", fg=COLORS["navy"],
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        tk.Button(toolbar, text="Refresh", command=self.refresh_records, bg="#E8EEF4", fg=COLORS["text"],
                  relief="flat", padx=10, pady=5, cursor="hand2").pack(side="right")

        filters = tk.Frame(panel, bg="white")
        filters.pack(fill="x", pady=(0, 10))
        search = ttk.Entry(filters, textvariable=self.search_var, font=("Segoe UI", 10))
        search.pack(side="left", fill="x", expand=True, ipady=5)
        search.insert(0, "")
        search.bind("<KeyRelease>", lambda _e: self.refresh_records())
        status_box = ttk.Combobox(filters, textvariable=self.status_filter,
                                  values=["All statuses", "Present", "Absent", "Late", "Excused", "Other"],
                                  state="readonly", width=15)
        status_box.pack(side="left", padx=(8, 0), ipady=4)
        status_box.bind("<<ComboboxSelected>>", lambda _e: self.refresh_records())

        columns = HEADERS
        self.tree = ttk.Treeview(panel, columns=columns, show="headings", selectmode="browse")
        widths = {"Record ID": 75, "Student ID": 90, "Student Name": 145, "Class / Section": 105,
                  "Subject": 110, "Date": 90, "Status": 80, "Remarks": 130}
        for column in columns:
            self.tree.heading(column, text=column)
            self.tree.column(column, width=widths[column], minwidth=55, anchor="w")
        yscroll = ttk.Scrollbar(panel, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(panel, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(fill="both", expand=True, side="top")
        yscroll.place(relx=1.0, rely=0.22, relheight=0.72, anchor="ne")
        xscroll.pack(fill="x")
        self.tree.bind("<<TreeviewSelect>>", self.load_selected_record)
        tk.Label(panel, text="Tip: select a row to load it into the form for editing.", bg="white",
                 fg=COLORS["muted"], font=("Segoe UI", 8)).pack(anchor="w", pady=(8, 0))

    def new_record(self) -> None:
        """Reset the form so a new attendance entry can be added."""
        self.clear_form()

    def _form_record(self) -> dict[str, str] | None:
        student_id = self.form_vars["student_id"].get().strip()
        student_name = self.form_vars["student_name"].get().strip()
        class_section = self.form_vars["class_section"].get().strip()
        subject = self.form_vars["subject"].get().strip()
        record_date = self.form_vars["record_date"].get().strip()
        status = self.form_vars["status"].get().strip()
        if not all([student_id, student_name, class_section, subject, record_date, status]):
            messagebox.showwarning("Missing information", "Fill in all required fields.", parent=self)
            return None
        try:
            datetime.strptime(record_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid date", "Enter the date in YYYY-MM-DD format, for example 2026-09-27.", parent=self)
            return None
        return {
            "Record ID": self.selected_record_id or "",
            "Student ID": student_id,
            "Student Name": student_name,
            "Class / Section": class_section,
            "Subject": subject,
            "Date": record_date,
            "Status": status,
            "Remarks": self.form_vars["remarks"].get().strip(),
        }

    def save_record(self) -> None:
        record = self._form_record()
        if record is None:
            return
        records = read_records()
        # One attendance entry per student per day prevents accidental duplicates.
        if any(r["Student ID"].casefold() == record["Student ID"].casefold() and
               r["Subject"].casefold() == record["Subject"].casefold() and r["Date"] == record["Date"] for r in records):
            messagebox.showwarning("Duplicate attendance", "This student already has an attendance record for that subject and date.", parent=self)
            return
        ids = [int(r["Record ID"]) for r in records if r["Record ID"].isdigit()]
        record["Record ID"] = str(max(ids, default=0) + 1)
        records.append(record)
        self._save(records, "Attendance record added.")

    def update_record(self) -> None:
        if not self.selected_record_id:
            messagebox.showinfo("Select a record", "Select a row in the table before updating.", parent=self)
            return
        record = self._form_record()
        if record is None:
            return
        records = read_records()
        if any(r["Record ID"] != self.selected_record_id and
               r["Student ID"].casefold() == record["Student ID"].casefold() and
               r["Subject"].casefold() == record["Subject"].casefold() and r["Date"] == record["Date"]
               for r in records):
            messagebox.showwarning("Duplicate attendance", "This student already has an attendance record for that subject and date.", parent=self)
            return
        for index, current in enumerate(records):
            if current["Record ID"] == self.selected_record_id:
                records[index] = record
                break
        else:
            messagebox.showerror("Record not found", "The selected record no longer exists in the workbook.", parent=self)
            self.refresh_records()
            return
        self._save(records, "Attendance record updated.")

    def delete_record(self) -> None:
        if not self.selected_record_id:
            messagebox.showinfo("Select a record", "Select a row in the table before deleting.", parent=self)
            return
        if not messagebox.askyesno("Confirm deletion", "Delete the selected attendance record from Excel?", parent=self):
            return
        records = [r for r in read_records() if r["Record ID"] != self.selected_record_id]
        self._save(records, "Attendance record deleted.")

    def _save(self, records: list[dict[str, str]], success_message: str) -> None:
        try:
            write_records(records)
        except PermissionError:
            messagebox.showerror("Workbook is open", "Close attendance_records.xlsx in Excel, then try again.", parent=self)
            return
        except OSError as error:
            messagebox.showerror("Could not save", f"The Excel file could not be saved.\n{error}", parent=self)
            return
        self.clear_form()
        self.refresh_records()
        messagebox.showinfo("Saved", success_message, parent=self)

    def clear_form(self) -> None:
        for key, variable in self.form_vars.items():
            variable.set("" if key != "record_date" else date.today().isoformat())
        self.form_vars["status"].set("Present")
        self.selected_record_id = None
        if hasattr(self, "tree"):
            for item in self.tree.selection():
                self.tree.selection_remove(item)

    def load_selected_record(self, _event: object = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        values = self.tree.item(selection[0], "values")
        if not values:
            return
        self.selected_record_id = str(values[0])
        self.form_vars["student_id"].set(values[1])
        self.form_vars["student_name"].set(values[2])
        self.form_vars["class_section"].set(values[3])
        self.form_vars["subject"].set(values[4])
        self.form_vars["record_date"].set(values[5])
        self.form_vars["status"].set(values[6])
        self.form_vars["remarks"].set(values[7])

    def refresh_records(self) -> None:
        if not hasattr(self, "tree"):
            return
        records = read_records()
        query = self.search_var.get().strip().casefold()
        status = self.status_filter.get()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for record in records:
            searchable = " ".join(record.values()).casefold()
            if query and query not in searchable:
                continue
            if status != "All statuses" and record["Status"] != status:
                continue
            self.tree.insert("", "end", values=[record[header] for header in HEADERS])
        today = date.today().isoformat()
        todays = [record for record in records if record["Date"] == today]
        self.total_label.configure(text=str(len(records)))
        self.present_label.configure(text=str(sum(1 for record in todays if record["Status"] == "Present")))
        self.absent_label.configure(text=str(sum(1 for record in todays if record["Status"] == "Absent")))

    def logout(self) -> None:
        if messagebox.askyesno("Log out", "Are you sure you want to log out?", parent=self):
            self.show_login()


if __name__ == "__main__":
    AttendanceApp().mainloop()



