from __future__ import annotations

import argparse
import queue
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path

import customtkinter as ctk


@dataclass(frozen=True)
class ScriptEntry:
    key: str
    label: str
    filename: str


SCRIPT_ORDER = [
    ScriptEntry("cash_ups", "Cash Ups", "cash_ups_rollover.py"),
    ScriptEntry("cashiers_shorts_and_overs", "Cashiers Shorts and Overs", "cashiers_shorts_and_overs_rollover.py"),
    ScriptEntry("daily_bill_payment_recon", "Daily Bill Payment Recon", "daily_bill_payment_recon_rollover.py"),
    ScriptEntry("daily_drinking_water", "Daily Drinking Water", "daily_drinking_water_rollover.py"),
    ScriptEntry("daily_ice_stock", "Daily Ice Stock", "daily_ice_stock_rollover.py"),
    ScriptEntry("daily_sales", "Daily Sales", "daily_sales_rollover.py"),
    ScriptEntry("gas_bottles", "Gas Bottles", "gas_bottles_rollover.py"),
    ScriptEntry("speed_point_control", "Speed Point Control", "speed_point_control_rollover.py"),
    ScriptEntry("fms_recon", "FMS Recon", "fms_recon_rollover.py"),
]


def resolve_root(root: Path | None) -> Path:
    base = root if root is not None else Path(__file__).resolve().parent
    return base


def run_script(script_path: Path, root: Path, dry_run: bool) -> subprocess.CompletedProcess[str]:
    command = [sys.executable, str(script_path), "--root", str(root)]
    if dry_run:
        command.append("--dry-run")
    return subprocess.run(command, check=False, capture_output=True, text=True)


def run_scripts(selected: list[ScriptEntry], workspace_root: Path, dry_run: bool, log) -> bool:
    failures: list[tuple[str, int]] = []
    for entry in selected:
        script_path = workspace_root / entry.filename
        log(f"[{entry.key}] Running {script_path.name}")
        if not script_path.exists():
            log(f"[{entry.key}] Missing script: {script_path}")
            failures.append((entry.key, 1))
            continue

        completed = run_script(script_path, workspace_root, dry_run)
        if completed.stdout:
            log(completed.stdout.rstrip())
        if completed.stderr:
            log(completed.stderr.rstrip())
        if completed.returncode != 0:
            failures.append((entry.key, completed.returncode))

    if failures:
        log("\nSome rollover scripts failed:")
        for name, code in failures:
            log(f"  - {name}: exit code {code}")
        return False

    log("\nAll selected rollover scripts completed successfully.")
    return True


def cli_main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Run one or more monthly rollover scripts.")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path(__file__).resolve().parent,
        help="Project root that contains the workbook folders and rollover scripts.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Ask each script to show what it would do without writing changes.",
    )
    parser.add_argument(
        "--only",
        nargs="+",
        choices=[entry.key for entry in SCRIPT_ORDER],
        help="Run only the selected script keys instead of the full sequence.",
    )
    args = parser.parse_args(argv)

    root = resolve_root(args.root)
    selected = SCRIPT_ORDER if not args.only else [entry for entry in SCRIPT_ORDER if entry.key in set(args.only)]
    if not selected:
        print("No scripts selected.")
        return 1

    ok = run_scripts(selected, root, args.dry_run, print)
    return 0 if ok else 1


class LauncherApp:
    def __init__(self, root_dir: Path) -> None:
        self.root_dir = root_dir
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")
        self.window = ctk.CTk()
        self.window.title("Monthly Rollover Launcher")

		# --- Center the window on the screen ---
        window_width = 980
        window_height = 750

		# Ensure Tkinter has registered the screen dimensions
        self.window.update_idletasks()
        
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

		 # Calculate the X and Y coordinates to center the window
        x = int((screen_width / 2) - (window_width / 2))
        y = int((screen_height / 2) - (window_height / 2))

		# Apply the geometry with the +x+y offset
        self.window.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.window.minsize(900, 600)

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.dry_run_var = ctk.BooleanVar(value=False)
        self.script_vars: dict[str, ctk.BooleanVar] = {}

        self._build_ui()

    def _modal(self, title: str, message: str) -> None:
        dialog = ctk.CTkToplevel(self.window)
        dialog.title(title)
        dialog.geometry("420x180")
        dialog.resizable(False, False)
        dialog.transient(self.window)
        dialog.grab_set()

        frame = ctk.CTkFrame(dialog, corner_radius=14)
        frame.pack(fill="both", expand=True, padx=16, pady=16)
        frame.grid_columnconfigure(0, weight=1)

        label = ctk.CTkLabel(frame, text=message, wraplength=360, justify="left")
        label.grid(row=0, column=0, sticky="w", padx=14, pady=(16, 10))
        button = ctk.CTkButton(frame, text="OK", command=dialog.destroy)
        button.grid(row=1, column=0, sticky="e", padx=14, pady=(8, 14))
        dialog.focus()

    def _build_ui(self) -> None:
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_rowconfigure(0, weight=1)        

        container = ctk.CTkFrame(self.window, corner_radius=18)
        container.grid(row=0, column=0, padx=18, pady=18, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(2, weight=1)

        header = ctk.CTkFrame(container, corner_radius=14)
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 10))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="Monthly Rollover Launcher", font=("Segoe UI", 24, "bold")).grid(
            row=0, column=0, sticky="w", padx=16, pady=(14, 2)
        )
        ctk.CTkLabel(
            header,
            text="Select one or more rollover scripts, then run them in a single pass.",
            wraplength=860,
            justify="left",
        ).grid(row=1, column=0, sticky="w", padx=16, pady=(0, 14))

        body = ctk.CTkFrame(container, corner_radius=14)
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        scripts_panel = ctk.CTkFrame(body, corner_radius=12)
        scripts_panel.grid(row=0, column=0, sticky="nsew", padx=(12, 6), pady=12)
        scripts_panel.grid_columnconfigure(0, weight=1)
        scripts_panel.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(scripts_panel, text="Scripts", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 8)
        )

        scroll = ctk.CTkScrollableFrame(scripts_panel, corner_radius=10)
        scroll.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        scroll.grid_columnconfigure(0, weight=1)

        for entry in SCRIPT_ORDER:
            var = ctk.BooleanVar(value=True)
            self.script_vars[entry.key] = var
            ctk.CTkCheckBox(scroll, text=entry.label, variable=var).pack(anchor="w", padx=8, pady=5)

        right_panel = ctk.CTkFrame(body, corner_radius=12)
        right_panel.grid(row=0, column=1, sticky="nsew", padx=(6, 12), pady=12)
        right_panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(right_panel, text="Options", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 8)
        )

        options_inner = ctk.CTkFrame(right_panel, corner_radius=10)
        options_inner.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        options_inner.grid_columnconfigure(0, weight=1)

        ctk.CTkCheckBox(options_inner, text="Dry run", variable=self.dry_run_var).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 8)
        )
        ctk.CTkButton(options_inner, text="Select All", command=self.select_all).grid(
            row=1, column=0, sticky="ew", padx=14, pady=(8, 6)
        )
        ctk.CTkButton(options_inner, text="Clear All", command=self.clear_all).grid(
            row=2, column=0, sticky="ew", padx=14, pady=6
        )
        ctk.CTkButton(options_inner, text="Run Selected", command=self.start_run).grid(
            row=3, column=0, sticky="ew", padx=14, pady=(12, 6)
        )
        ctk.CTkButton(options_inner, text="Run All", command=self.start_run_all).grid(
            row=4, column=0, sticky="ew", padx=14, pady=6
        )
        ctk.CTkButton(options_inner, text="Close", command=self.window.destroy).grid(
            row=5, column=0, sticky="ew", padx=14, pady=(12, 14)
        )

        log_frame = ctk.CTkFrame(container, corner_radius=14)
        log_frame.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 14))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(log_frame, text="Output", font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w", padx=14, pady=(14, 8)
        )

        self.log_widget = ctk.CTkTextbox(log_frame, wrap="word")
        self.log_widget.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))

        self._append_log("Ready.")
        self.window.after(100, self._drain_log_queue)

    def select_all(self) -> None:
        for var in self.script_vars.values():
            var.set(True)

    def clear_all(self) -> None:
        for var in self.script_vars.values():
            var.set(False)

    def _selected_entries(self) -> list[ScriptEntry]:
        return [entry for entry in SCRIPT_ORDER if bool(self.script_vars[entry.key].get())]

    def _append_log(self, text: str) -> None:
        self.log_widget.insert("end", text + "\n")
        self.log_widget.see("end")

    def _log(self, text: str) -> None:
        self.log_queue.put(text)

    def _drain_log_queue(self) -> None:
        try:
            while True:
                message = self.log_queue.get_nowait()
                self._append_log(message)
        except queue.Empty:
            pass

        if self.running:
            self.window.after(100, self._drain_log_queue)
        else:
            self.window.after(200, self._drain_log_queue)

    def _set_controls_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for child in self.window.winfo_children():
            try:
                child.configure(state=state)
            except Exception:
                pass

    def _start_run(self, selected: list[ScriptEntry]) -> None:
        if self.running:
            return
        if not selected:
            self._modal("Monthly Rollover Launcher", "Select at least one script to run.")
            return

        self.running = True
        self._append_log("")
        self._append_log("Starting run...")

        worker = threading.Thread(target=self._run_worker, args=(selected,), daemon=True)
        worker.start()

    def start_run(self) -> None:
        self._start_run(self._selected_entries())

    def start_run_all(self) -> None:
        self._start_run(list(SCRIPT_ORDER))

    def _run_worker(self, selected: list[ScriptEntry]) -> None:
        ok = run_scripts(selected, self.root_dir, bool(self.dry_run_var.get()), self._log)
        self.window.after(0, self._finish_run, ok)

    def _finish_run(self, ok: bool) -> None:
        self.running = False
        if ok:
            self._modal("Monthly Rollover Launcher", "All selected rollover scripts completed successfully.")
        else:
            self._modal("Monthly Rollover Launcher", "One or more rollover scripts failed. See the log for details.")

    def run(self) -> None:
        self.window.mainloop()


def main(argv: list[str] | None = None) -> int:
    if argv is None:
        argv = sys.argv[1:]

    if argv:
        return cli_main(argv)

    app = LauncherApp(resolve_root(None))
    app.run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
