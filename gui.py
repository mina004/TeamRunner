import json
import os
import subprocess
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional workbook readability check
try:
    import openpyxl  # type: ignore
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False


# ---------------- Paths ----------------
PROJ_ROOT = Path(__file__).resolve().parent
EXEC_DIR = PROJ_ROOT / "executables"
REPORT_DIR = PROJ_ROOT / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)


# ---------------- Task model ----------------
@dataclass
class Task:
    label: str
    exe: str
    json_name: str
    arg_style: str  # workbook_output | excel_output | workbook_report | workbook_only | workbook_defaultjson
    extra_args: list[str] | None = None


# ---------------- Your 10 EXEs ----------------
# This list matches the 10 teams (adjust only if your professor says otherwise).
TASKS: list[Task] = [
    Task("chart of accounts", "Chart_of_accounts.exe", "chart_of_accounts.json", "workbook_output"),
    Task("customers", "customers.exe", "customers.json", "workbook_output"),
    Task("vendors", "vendor_compare.exe", "vendors.json", "workbook_output"),
    Task("invoices", "qb-invoice-sync.exe", "invoices.json", "excel_output"),  # uses --excel
    Task("receive payments", "receive_payments.exe", "receive_payments.json", "workbook_output"),
    Task("service bills", "Service_bill_cli.exe", "service_bills.json", "workbook_report"),  # uses --report
    Task("item bills", "item_bills.exe", "item_bills.json", "workbook_output"),
    Task("misc income", "misc_income_cli.exe", "misc_income.json", "workbook_output"),
    Task("pay bills", "pay_bills.exe", "pay_bills.json", "workbook_output"),
    Task("payment terms", "payment_terms_dummy.exe", "payment_terms.json", "workbook_output"),
]


# ---------------- Runner result ----------------
@dataclass
class RunResult:
    label: str
    exe_name: str
    json_path: Path
    started: bool
    exit_code: int | None
    stdout: str
    stderr: str
    json_data: dict | None
    error_note: str | None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Team Runner GUI")
        self.geometry("980x580")
        self.minsize(820, 480)

        self.workbook_path: Path | None = None
        self._thread: threading.Thread | None = None
        self._results: list[RunResult] = []

        # --- Controls ---
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        self.select_btn = ttk.Button(top, text="Select Company Workbook (.xlsx)", command=self.select_workbook)
        self.select_btn.pack(side="left")

        self.start_btn = ttk.Button(top, text="Start", command=self.start_run, state="disabled")
        self.start_btn.pack(side="left", padx=(8, 0))

        self.progress = ttk.Progressbar(top, orient="horizontal", mode="determinate", length=420)
        self.progress.pack(side="left", padx=12)
        self.progress["maximum"] = len(TASKS)
        self.progress["value"] = 0

        self.status_var = tk.StringVar(value="Select the official Example Company Excel file to begin.")
        ttk.Label(top, textvariable=self.status_var).pack(side="left", padx=10)

        # --- Log area ---
        frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        frame.pack(fill="both", expand=True)

        self.text = tk.Text(frame, wrap="word", font=("Consolas", 10))
        self.text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=scroll.set)

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        if not HAS_OPENPYXL:
            self._prepend_block(
                "Note",
                "openpyxl not installed. Workbook validation will be minimal.\n"
                "Install: pip install openpyxl"
            )

    # ---------------- UI actions ----------------
    def select_workbook(self):
        path = filedialog.askopenfilename(
            title="Select Company Workbook (.xlsx)",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
        )
        if not path:
            self._set_status("Workbook selection canceled.")
            return

        wb = Path(path)
        ok, err = self._validate_workbook(wb)
        if not ok:
            messagebox.showerror("Invalid Workbook", err or "Invalid workbook.")
            self.workbook_path = None
            self.start_btn.config(state="disabled")
            return

        self.workbook_path = wb
        self.start_btn.config(state="normal")
        self._set_status(f"Workbook selected: {wb.name}")

    def start_run(self):
        if self._thread and self._thread.is_alive():
            messagebox.showinfo("Running", "Already running.")
            return
        if not self.workbook_path:
            messagebox.showwarning("Missing Workbook", "Please select a workbook first.")
            return

        self.text.delete("1.0", "end")
        self._results.clear()
        self.progress["value"] = 0

        self.select_btn.config(state="disabled")
        self.start_btn.config(state="disabled")
        self._set_status("Running tasks...")

        self._thread = threading.Thread(target=self._run_all, daemon=True)
        self._thread.start()

    # ---------------- Validation ----------------
    def _validate_workbook(self, wb: Path) -> tuple[bool, str | None]:
        if not wb.exists():
            return False, "Selected file does not exist."
        if wb.suffix.lower() != ".xlsx":
            return False, "Please select a .xlsx file."
        if wb.stat().st_size == 0:
            return False, "Workbook is empty (0 bytes). Please select the official Example Company Excel file."

        if HAS_OPENPYXL:
            try:
                book = openpyxl.load_workbook(filename=str(wb), read_only=True, data_only=True)
                _ = book.sheetnames
                book.close()
            except Exception as e:
                return False, f"Could not open workbook: {e}"

        return True, None

    # ---------------- Command building ----------------
    def _build_cmd(self, exe_path: Path, task: Task, workbook: str, out_json: Path) -> list[str]:
        style = task.arg_style

        if style == "workbook_output":
            cmd = [str(exe_path), "--workbook", workbook, "--output", str(out_json)]
        elif style == "excel_output":
            cmd = [str(exe_path), "--excel", workbook, "--output", str(out_json)]
        elif style == "workbook_report":
            cmd = [str(exe_path), "--workbook", workbook, "--report", str(out_json)]
        elif style == "workbook_only":
            cmd = [str(exe_path), "--workbook", workbook]
        else:
            # default safe behavior
            cmd = [str(exe_path), "--workbook", workbook, "--output", str(out_json)]

        if task.extra_args:
            cmd.extend(task.extra_args)

        return cmd

    # ---------------- Runner ----------------
    def _run_all(self):
        wb = str(self.workbook_path)
        total = len(TASKS)

        for idx, task in enumerate(TASKS, start=1):
            exe_path = EXEC_DIR / task.exe
            json_path = REPORT_DIR / task.json_name

            self._set_status(f"Running: {task.label}")

            # If exe missing, record an error result
            if not exe_path.exists():
                res = RunResult(
                    label=task.label,
                    exe_name=task.exe,
                    json_path=json_path,
                    started=False,
                    exit_code=None,
                    stdout="",
                    stderr="",
                    json_data=None,
                    error_note=f"EXE not found: {exe_path}"
                )
                self._results.append(res)
                self._prepend_block(f"[{task.label}] ERROR", res.error_note)
                self._update_progress(idx, total)
                continue

            cmd = self._build_cmd(exe_path, task, wb, json_path)

            try:
                proc = subprocess.run(
                    cmd,
                    cwd=str(EXEC_DIR),
                    capture_output=True,
                    text=True,
                    check=False
                )
            except Exception as e:
                res = RunResult(
                    label=task.label,
                    exe_name=task.exe,
                    json_path=json_path,
                    started=False,
                    exit_code=None,
                    stdout="",
                    stderr="",
                    json_data=None,
                    error_note=f"Error starting process: {e}"
                )
                self._results.append(res)
                self._prepend_block(f"[{task.label}] ERROR", res.error_note)
                self._update_progress(idx, total)
                continue

            stdout = (proc.stdout or "").strip()
            stderr = (proc.stderr or "").strip()

            json_data = None
            error_note = None

            if json_path.exists():
                try:
                    with open(json_path, "r", encoding="utf-8") as f:
                        json_data = json.load(f)
                except Exception as e:
                    error_note = f"JSON read error: {e}"
            else:
                error_note = "No JSON output produced."

            res = RunResult(
                label=task.label,
                exe_name=task.exe,
                json_path=json_path,
                started=True,
                exit_code=proc.returncode,
                stdout=stdout,
                stderr=stderr,
                json_data=json_data,
                error_note=error_note
            )
            self._results.append(res)

            # Show output at top of GUI
            self._prepend_block(f"[{task.label}] Completed", self._format_result_for_ui(res))

            self._update_progress(idx, total)

        self._write_final_report_and_cleanup()
        self._set_status("Done.")
        self.after(0, lambda: self.select_btn.config(state="normal"))
        self.after(0, lambda: self.start_btn.config(state="normal" if self.workbook_path else "disabled"))

    # ---------------- Formatting ----------------
    def _summarize_json(self, data: dict) -> str:
        status = data.get("status", "(no status)")
        err = data.get("error", None)

        same_keys = {k: v for k, v in data.items() if k.startswith("same_")}
        added_keys = {k: v for k, v in data.items() if k.startswith("added_")}
        conflicts = data.get("conflicts", None)

        def _len(v):
            return len(v) if isinstance(v, list) else v

        lines = [f"status: {status}"]
        if err:
            lines.append(f"error: {err}")

        if same_keys:
            lines.append("same_*:")
            for k, v in same_keys.items():
                lines.append(f"  - {k}: {_len(v)}")

        if added_keys:
            lines.append("added_*:")
            for k, v in added_keys.items():
                lines.append(f"  - {k}: {_len(v)}")

        if isinstance(conflicts, list):
            lines.append(f"conflicts: {len(conflicts)}")

        return "\n".join(lines)

    def _format_result_for_ui(self, res: RunResult) -> str:
        parts = []
        parts.append(f"EXE: {res.exe_name}")
        if res.exit_code is not None:
            parts.append(f"exit_code: {res.exit_code}")

        if res.json_data is not None:
            parts.append("\nSummary:")
            parts.append(self._summarize_json(res.json_data))
            parts.append("\nFull JSON:")
            parts.append(json.dumps(res.json_data, indent=2, ensure_ascii=False))
        else:
            parts.append(f"\n[ERROR] {res.error_note or 'Unknown error'}")
            if res.stdout:
                parts.append("\nSTDOUT:\n" + res.stdout)
            if res.stderr:
                parts.append("\nSTDERR:\n" + res.stderr)

        return "\n".join(parts)

    # ---------------- Final report + cleanup ----------------
    def _write_final_report_and_cleanup(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_report = REPORT_DIR / f"report_{ts}.txt"

        with open(final_report, "w", encoding="utf-8") as out:
            out.write(f"FINAL REPORT — {datetime.now().isoformat(timespec='seconds')}\n")
            out.write("=" * 80 + "\n\n")

            # IMPORTANT: write ALL 10 sections (even if json missing)
            for res in self._results:
                out.write(f"--- {res.label} ---\n")
                if res.json_data is not None:
                    out.write(json.dumps(res.json_data, indent=2, ensure_ascii=False))
                    out.write("\n\n")
                else:
                    out.write("[ERROR: No JSON output produced for this team. Details below:]\n")
                    if res.error_note:
                        out.write(res.error_note + "\n")
                    if res.exit_code is not None:
                        out.write(f"exit_code: {res.exit_code}\n")
                    if res.stdout:
                        out.write("\nSTDOUT:\n" + res.stdout + "\n")
                    if res.stderr:
                        out.write("\nSTDERR:\n" + res.stderr + "\n")
                    out.write("\n\n")

        # Cleanup: delete JSON files that exist
        deleted = 0
        for res in self._results:
            try:
                if res.json_path.exists():
                    os.remove(res.json_path)
                    deleted += 1
            except Exception:
                pass

        self._prepend_block(
            "All Tasks Finished",
            f"Final report written to:\n{final_report}\n\n"
            f"Temporary JSON files deleted: {deleted}"
        )

    # ---------------- UI helpers ----------------
    def _update_progress(self, completed: int, total: int):
        def _ui():
            self.progress["value"] = completed
            pct = int((completed / total) * 100)
            self.status_var.set(f"{completed}/{total} completed — {pct}%")
        self.after(0, _ui)

    def _prepend_block(self, header: str, body: str):
        def _ui():
            block = f"{header}\n{body}\n" + "-" * 80 + "\n"
            self.text.insert("1.0", block)
        self.after(0, _ui)

    def _set_status(self, msg: str):
        self.after(0, lambda: self.status_var.set(msg))


if __name__ == "__main__":
    App().mainloop()
