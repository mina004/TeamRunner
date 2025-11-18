import json
import os
import subprocess
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Try to import openpyxl for real workbook validation
try:
    import openpyxl  # type: ignore
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

# ---------- Config ----------
PROJ_ROOT = Path(__file__).resolve().parent
EXEC_DIR = PROJ_ROOT / "executables"
REPORT_DIR = PROJ_ROOT / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

# Validation rule: require this sheet name in the workbook
REQUIRED_SHEETS = {"PaymentTerms"}

# Define tasks. First is real dummy exe with args; others simulate until exes are provided.
TASKS = [
    {
        "label": "payment terms (dummy)",
        "exe": "payment_terms_dummy.exe",
        "json": "terms.json",
        "args": ["--workbook", None, "--output", None],  # we inject workbook and output path at runtime
    },
    {"label": "customers",          "exe": "customers.exe",          "json": "customers.json",          "args": None},
    {"label": "vendors",            "exe": "vendors.exe",            "json": "vendors.json",            "args": None},
    {"label": "items",              "exe": "items.exe",              "json": "items.json",              "args": None},
    {"label": "invoices",           "exe": "invoices.exe",           "json": "invoices.json",           "args": None},
    {"label": "payments",           "exe": "payments.exe",           "json": "payments.json",           "args": None},
    {"label": "purchase orders",    "exe": "purchase_orders.exe",    "json": "purchase_orders.json",    "args": None},
    {"label": "bills",              "exe": "bills.exe",              "json": "bills.json",              "args": None},
    {"label": "journal entries",    "exe": "journal_entries.exe",    "json": "journal_entries.json",    "args": None},
    {"label": "general ledger",     "exe": "general_ledger.exe",     "json": "general_ledger.json",     "args": None},
]

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Team Runner GUI")
        self.geometry("920x540")
        self.minsize(760, 420)

        # --- Controls (top bar) ---
        controls = ttk.Frame(self, padding=10)
        controls.pack(fill="x")

        self.select_btn = ttk.Button(controls, text="Select Company Workbook (.xlsx)", command=self.select_workbook)
        self.select_btn.pack(side="left")

        self.progress = ttk.Progressbar(controls, orient="horizontal", mode="determinate", length=420)
        self.progress.pack(side="left", padx=10)
        self.progress["maximum"] = len(TASKS)

        self.status_var = tk.StringVar(value="Select a company workbook to begin.")
        ttk.Label(controls, textvariable=self.status_var).pack(side="left", padx=10)

        # --- Log area ---
        text_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        text_frame.pack(fill="both", expand=True)

        self.text = tk.Text(text_frame, wrap="word", height=20, font=("Consolas", 10))
        self.text.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=yscroll.set)

        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        # State
        self._completed_jsons = []
        self._lock = threading.Lock()
        self._workbook_path = None  # set after user picks

        # Hint if openpyxl missing
        if not HAS_OPENPYXL:
            self._prepend_text_block(
                "Note",
                "Optional dependency 'openpyxl' not found. Install with:\n    pip install openpyxl\n"
                "Falling back to basic validation (extension + readability)."
            )

    # ------------------ Workbook selection flow ------------------
    def select_workbook(self):
        # loop until valid or user cancels
        while True:
            file_path = filedialog.askopenfilename(
                title="Select Company Workbook",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All Files", "*.*")]
            )
            if not file_path:
                # user canceled
                self.status_var.set("Selection canceled. Please choose a workbook to begin.")
                return

            is_valid, err = self._validate_workbook(Path(file_path))
            if is_valid:
                self._workbook_path = Path(file_path)
                self._start_run()
                return
            else:
                messagebox.showerror("Invalid Workbook", err)
                # Loop continues and file dialog opens again

    def _validate_workbook(self, path: Path):
        # Basic checks
        if not path.exists():
            return False, "The selected file does not exist."
        if path.suffix.lower() != ".xlsx":
            return False, "Please select a .xlsx Excel workbook."
        if path.stat().st_size == 0:
            return False, "File does not match Beulah Inc. format. Please reach out to David Nevill dnevill@beulahinc.com"

        if HAS_OPENPYXL:
            try:
                wb = openpyxl.load_workbook(filename=str(path), read_only=True, data_only=True)
                sheets = set(wb.sheetnames)
                wb.close()
                missing = REQUIRED_SHEETS - sheets
                if missing:
                    return False, (
                        "Selected workbook is not in the expected format.\n"
                        f"Missing required sheet(s): {', '.join(sorted(missing))}\n"
                        "Please choose a valid file."
                    )
            except Exception as e:
                return False, f"Could not read workbook: {e}\nPlease choose a valid .xlsx file."
        else:
            # Minimal check when openpyxl not installed
            # (We can't inspect sheet names; encourage installing openpyxl)
            pass

        return True, None

    # ------------------ Runner ------------------
    def _start_run(self):
        if hasattr(self, "_thread") and self._thread.is_alive():
            messagebox.showinfo("Running", "Process already running.")
            return

        self.progress["value"] = 0
        self.status_var.set("Running tasks...")
        self.text.delete("1.0", "end")
        self._completed_jsons.clear()

        self._thread = threading.Thread(target=self._run_all_tasks, daemon=True)
        self._thread.start()

    def _run_all_tasks(self):
        total = len(TASKS)
        workbook = str(self._workbook_path) if self._workbook_path else None

        for i, task in enumerate(TASKS, start=1):
            label = task["label"]
            exe_name = task["exe"]
            json_name = task["json"]
            args = task.get("args")

            exe_path = (EXEC_DIR / exe_name)
            json_path = (REPORT_DIR / json_name)
            json_path.parent.mkdir(parents=True, exist_ok=True)

            self._append_status(f"Starting: {label}")

            if exe_path.exists():
                # Build command with workbook + absolute output path
                cmd = [str(exe_path)]
                if args:
                    final_args = []
                    it = iter(args)
                    for a in it:
                        if a == "--workbook":
                            final_args.extend(["--workbook", workbook])
                        elif a == "--output":
                            final_args.extend(["--output", str(json_path)])
                        else:
                            final_args.append(a)
                    cmd += final_args

                try:
                    result = subprocess.run(
                        cmd,
                        cwd=str(EXEC_DIR),  # run from executables folder
                        capture_output=True,
                        text=True,
                        check=False
                    )
                    if result.returncode != 0:
                        body = []
                        if result.stdout:
                            body.append("STDOUT:\n" + result.stdout.strip())
                        if result.stderr:
                            body.append("STDERR:\n" + result.stderr.strip())
                        self._prepend_text_block(
                            f"[{self._now()}] {label} — FAILED (exit {result.returncode})",
                            "\n\n".join(body) or "(no output)"
                        )
                    else:
                        self._handle_json_output(label, json_path)
                except Exception as e:
                    self._prepend_text_block(f"[{self._now()}] {label} — ERROR STARTING", str(e))
            else:
                # Simulate task if exe not present
                time.sleep(0.5)
                dummy_data = {
                    "team": label,
                    "status": "simulated",
                    "workbook": workbook,
                    "records_processed": i * 10,
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                }
                with open(json_path, "w", encoding="utf-8") as f:
                    json.dump(dummy_data, f, indent=2)
                self._handle_json_output(label, json_path)

            self._update_progress(i, total)

        self._finish_and_cleanup()

    def _handle_json_output(self, label, json_path: Path):
        if json_path.exists():
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                parsed = json.dumps(data, indent=2, ensure_ascii=False)
                self._prepend_text_block(f"[{label}] Completed", parsed)
                with self._lock:
                    self._completed_jsons.append((label, json_path))
            except Exception as e:
                self._prepend_text_block(f"[{self._now()}] {label} — JSON READ ERROR", str(e))
        else:
            self._prepend_text_block(f"[{self._now()}] {label} — WARNING", f"Expected JSON not found at: {json_path}")

    def _finish_and_cleanup(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_report = REPORT_DIR / f"report_{ts}.txt"
        try:
            with open(final_report, "w", encoding="utf-8") as out:
                out.write(f"FINAL REPORT — {datetime.now().isoformat(timespec='seconds')}\n")
                out.write("=" * 80 + "\n")
                for label, json_path in self._completed_jsons:
                    out.write(f"\n--- {label} ---\n")
                    try:
                        with open(json_path, "r", encoding="utf-8") as f:
                            data = json.load(f)
                        out.write(json.dumps(data, indent=2, ensure_ascii=False))
                        out.write("\n")
                    except Exception as e:
                        out.write(f"[ERROR reading {json_path}: {e}]\n")

            # Delete JSONs
            for _, json_path in self._completed_jsons:
                try:
                    os.remove(json_path)
                except Exception as e:
                    self._prepend_text_block("CLEANUP WARNING", f"Could not delete {json_path}: {e}")

            self._prepend_text_block("All Tasks Finished", f"Final report written to: {final_report}\nTemporary JSON files deleted.")
        finally:
            self.after(0, lambda: self.select_btn.config(state="normal"))
            self.status_var.set("Done.")

    # ---- UI helpers ----
    def _update_progress(self, completed, total):
        def _ui():
            self.progress["value"] = completed
            pct = int((completed / total) * 100)
            self.status_var.set(f"{completed}/{total} completed — {pct}%")
        self.after(0, _ui)

    def _prepend_text_block(self, header, body):
        def _ui():
            block = f"{header}\n{body}\n{'-'*70}\n"
            self.text.insert("1.0", block)
        self.after(0, _ui)

    def _append_status(self, msg):
        self.after(0, lambda: self.status_var.set(msg))

    @staticmethod
    def _now():
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


if __name__ == "__main__":
    app = App()
    app.mainloop()
