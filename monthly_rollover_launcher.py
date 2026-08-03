from __future__ import annotations

import argparse
import queue
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, scrolledtext, ttk

try:
    import customtkinter as ctk
except ModuleNotFoundError:
    ctk = None


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
    main_folder = base / "MAIN FOLDER"
    if main_folder.exists():
        return main_folder
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
        self.window = self._build_window()
        self.log_queue: queue.Queue[str] = queue.Queue()
        self.running = False
        self.dry_run_var = self._make_boolean_var(False)
        self.script_vars: dict[str, object] = {}
        self._build_ui()

    def _build_window(self):
        if ctk is not None:
            ctk.set_appearance_mode("system")
            ctk.set_default_color_theme("blue")
            window = ctk.CTk()
            window.title("Monthly Rollover Launcher")
            window.geometry("900x640")
            window.minsize(820, 560)
            return window

        window = tk.Tk()
        window.title("Monthly Rollover Launcher")
        window.geometry("900x640")
        window.minsize(820, 560)
        return window

    def _make_boolean_var(self, value: bool):
        if ctk is not None:
            return ctk.BooleanVar(value=value)
        return tk.BooleanVar(value=value)

    def _make_string_var(self, value: str):
        if ctk is not None:
            return ctk.StringVar(value=value)
        return tk.StringVar(value=value)

    def _frame(self, parent, **kwargs):
        if ctk is not None:
            return ctk.CTkFrame(parent, **kwargs)
        return ttk.Frame(parent, **kwargs)

    def _label(self, parent, **kwargs):
        if ctk is not None:
            return ctk.CTkLabel(parent, **kwargs)
        return ttk.Label(parent, **kwargs)

    def _button(self, parent, **kwargs):
        if ctk is not None:
            return ctk.CTkButton(parent, **kwargs)
        return ttk.Button(parent, **kwargs)

    def _checkbox(self, parent, **kwargs):
        if ctk is not None:
            return ctk.CTkCheckBox(parent, **kwargs)
        return ttk.Checkbutton(parent, **kwargs)

    def _textbox(self, parent, **kwargs):
        if ctk is not None:
            return ctk.CTkTextbox(parent, **kwargs)
        return scrolledtext.ScrolledText(parent, **kwargs)

    def _build_ui(self) -> None:
        if ctk is not None:
            self.window.grid_columnconfigure(0, weight=1)
            self.window.grid_rowconfigure(3, weight=1)
            container = ctk.CTkFrame(self.window, corner_radius=16)
            container.grid(row=0, column=0, padx=18, pady=18, sticky="nsew")
            container.grid_columnconfigure(0, weight=1)
            container.grid_rowconfigure(3, weight=1)
        else:
            self.window.columnconfigure(0, weight=1)
            self.window.rowconfigure(3, weight=1)
            container = ttk.Frame(self.window, padding=18)
            container.grid(row=0, column=0, sticky="nsew")
            container.columnconfigure(0, weight=1)
            container.rowconfigure(3, weight=1)

        title = self._label(
            container,
            text="Monthly Rollover Launcher",
            font=("Segoe UI", 20, "bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=4, pady=(4, 0))

        subtitle = self._label(
            container,
            text="Pick one script or select multiple. The launcher automatically uses MAIN FOLDER when it exists.",
            wraplength=820,
            justify="left",
        )
        subtitle.grid(row=1, column=0, sticky="w", padx=4, pady=(6, 14))

        controls = self._frame(container)
        controls.grid(row=2, column=0, sticky="ew", padx=4)
        if ctk is None:
            controls.columnconfigure(0, weight=1)
            controls.columnconfigure(1, weight=1)

        left = self._frame(controls)
        right = self._frame(controls)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        if ctk is None:
            left.columnconfigure(0, weight=1)
            right.columnconfigure(0, weight=1)

        script_box = self._frame(left)
        script_box.grid(row=0, column=0, sticky="ew")
        script_title = self._label(script_box, text="Scripts", font=("Segoe UI", 12, "bold"))
        script_title.grid(row=0, column=0, sticky="w", pady=(0, 8))

        if ctk is not None:
            scroll = ctk.CTkScrollableFrame(script_box, height=250, corner_radius=12)
            scroll.grid(row=1, column=0, sticky="nsew")
            scroll.grid_columnconfigure(0, weight=1)
        else:
            scroll = ttk.Frame(script_box)
            scroll.grid(row=1, column=0, sticky="nsew")

        for entry in SCRIPT_ORDER:
            var = self._make_boolean_var(True)
            self.script_vars[entry.key] = var
            label = f"{entry.label}"
            if ctk is not None:
                checkbox = ctk.CTkCheckBox(scroll, text=label, variable=var)
                checkbox.pack(anchor="w", padx=6, pady=3)
            else:
                checkbox = ttk.Checkbutton(scroll, text=label, variable=var)
                checkbox.pack(anchor="w", padx=6, pady=3)

        button_row = self._frame(left)
        button_row.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        self._button(button_row, text="Select All", command=self.select_all).pack(side="left", padx=(0, 8))
        self._button(button_row, text="Clear All", command=self.clear_all).pack(side="left")

        options_box = self._frame(right)
        options_box.grid(row=0, column=0, sticky="nsew")
        options_title = self._label(options_box, text="Options", font=("Segoe UI", 12, "bold"))
        options_title.grid(row=0, column=0, sticky="w", pady=(0, 8))

        dry_run_cb = self._checkbox(options_box, text="Dry run", variable=self.dry_run_var)
        if ctk is not None:
            dry_run_cb.grid(row=1, column=0, sticky="w", pady=(0, 10))
        else:
            dry_run_cb.grid(row=1, column=0, sticky="w")

        self.run_all_button = self._button(options_box, text="Run Selected", command=self.start_run)
        self.run_all_button.grid(row=2, column=0, sticky="ew", pady=(8, 8))

        self.run_all_only_button = self._button(options_box, text="Run All", command=self.start_run_all)
        self.run_all_only_button.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        self.close_button = self._button(options_box, text="Close", command=self.window.destroy)
        self.close_button.grid(row=4, column=0, sticky="ew")

        log_label = self._label(container, text="Output", font=("Segoe UI", 12, "bold"))
        log_label.grid(row=3, column=0, sticky="w", padx=4, pady=(14, 6))

        log_frame = self._frame(container)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=4, pady=(0, 4))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        if ctk is not None:
            self.log_widget = ctk.CTkTextbox(log_frame, wrap="word")
            self.log_widget.grid(row=0, column=0, sticky="nsew")
        else:
            self.log_widget = scrolledtext.ScrolledText(log_frame, wrap="word", state="disabled")
            self.log_widget.grid(row=0, column=0, sticky="nsew")

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
        if ctk is not None:
            self.log_widget.insert("end", text + "\n")
            self.log_widget.see("end")
        else:
            self.log_widget.configure(state="normal")
            self.log_widget.insert("end", text + "\n")
            self.log_widget.see("end")
            self.log_widget.configure(state="disabled")

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

    def _disable_buttons(self, disabled: bool) -> None:
        state = "disabled" if disabled else "normal"
        for widget in (self.run_all_button, self.run_all_only_button, self.close_button):
            try:
                widget.configure(state=state)
            except Exception:
                pass

    def start_run(self) -> None:
        self._start_run(self._selected_entries())

    def start_run_all(self) -> None:
        self._start_run(list(SCRIPT_ORDER))

    def _start_run(self, selected: list[ScriptEntry]) -> None:
        if self.running:
            return
        if not selected:
            messagebox.showwarning("Monthly Rollover Launcher", "Select at least one script to run.")
            return

        self.running = True
        self._disable_buttons(True)
        self._append_log("")
        self._append_log("Starting run...")

        worker = threading.Thread(target=self._run_worker, args=(selected,), daemon=True)
        worker.start()

    def _run_worker(self, selected: list[ScriptEntry]) -> None:
        ok = run_scripts(selected, self.root_dir, bool(self.dry_run_var.get()), self._log)
        self.window.after(0, self._finish_run, ok)

    def _finish_run(self, ok: bool) -> None:
        self.running = False
        self._disable_buttons(False)
        if ok:
            messagebox.showinfo("Monthly Rollover Launcher", "All selected rollover scripts completed successfully.")
        else:
            messagebox.showerror("Monthly Rollover Launcher", "One or more rollover scripts failed. See the log for details.")

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
