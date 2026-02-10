#!/usr/bin/env python3
"""
DBC Compare Tool - GUI Application
Cross-platform visual tool to compare CAN DBC files between two versions.
Works on Windows, macOS, and Linux.

                  _                                        _
  _ __ ___   __ _| | _____   _ __   ___  __ _  ___ ___   | |
 | '_ ` _ \ / _` | |/ / _ \ | '_ \ / _ \/ _` |/ __/ _ \  | |
 | | | | | | (_| |   <  __/ | |_) |  __/ (_| | (_|  __/  |_|
 |_| |_| |_|\__,_|_|\_\___| | .__/ \___|\__,_|\___\___|  (_)
                             |_|
              not war  --  5c1c30200c080e3ff581251ea1dfce8b9e0b12db7b194f8c8a6beb687939f492
   Find yourself in peace: https://www.youtube.com/watch?v=4-079YIasck
"""

import os
import sys
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from dbc_compare import (
    find_dbc_files, extract_bus_name, parse_dbc, compare_dbc_files,
    write_comparison_xlsx, write_html_report, write_pdf_report,
    categorize_changes, NUM_MSG_COLS, NUM_SIG_COLS, NUM_COLS_PER_SIDE,
    MSG_HEADERS, SIG_HEADERS
)


class DBCCompareApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DBC Compare Tool")
        self.root.minsize(750, 620)
        self.root.resizable(True, True)

        # Center window on screen
        self.root.update_idletasks()
        w, h = 750, 620
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        # Variables
        self.old_folder = tk.StringVar()
        self.new_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.is_running = False

        self._build_ui()

    def _build_ui(self):
        # Main frame with padding
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        # ── Title ──
        title_frame = ttk.Frame(main)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        title_label = ttk.Label(title_frame, text="DBC Compare Tool",
                  font=("Helvetica", 18, "bold"))
        title_label.pack(side=tk.LEFT)
        title_label.bind("<Triple-Button-1>", lambda e: self._show_about())
        ttk.Label(title_frame, text="Compare CAN DBC files between two versions",
                  font=("Helvetica", 10)).pack(side=tk.LEFT, padx=(15, 0), pady=(5, 0))

        # ── Separator ──
        ttk.Separator(main, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)

        # ── Folder Selection ──
        folder_frame = ttk.LabelFrame(main, text="  Folders  ", padding=10)
        folder_frame.pack(fill=tk.X, pady=(5, 5))

        # Old folder
        row1 = ttk.Frame(folder_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="Old Version:", width=14, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.old_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(row1, text="Browse...", width=10,
                   command=lambda: self._browse_folder(self.old_folder, "Select OLD version folder")).pack(side=tk.RIGHT)

        # New folder
        row2 = ttk.Frame(folder_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="New Version:", width=14, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self.new_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(row2, text="Browse...", width=10,
                   command=lambda: self._browse_folder(self.new_folder, "Select NEW version folder")).pack(side=tk.RIGHT)

        # Output folder
        row3 = ttk.Frame(folder_frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="Output Folder:", width=14, anchor=tk.W).pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self.output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(row3, text="Browse...", width=10,
                   command=lambda: self._browse_folder(self.output_folder, "Select OUTPUT folder")).pack(side=tk.RIGHT)

        ttk.Label(folder_frame, text="(Leave output empty for auto-generated name)",
                  font=("Helvetica", 9), foreground="gray").pack(anchor=tk.W, pady=(2, 0))

        # ── Compare Button ──
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        self.compare_btn = ttk.Button(btn_frame, text="  Compare  ",
                                       command=self._start_compare)
        self.compare_btn.pack(side=tk.LEFT)
        self.open_btn = ttk.Button(btn_frame, text="  Open Output Folder  ",
                                    command=self._open_output, state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT, padx=(10, 0))

        # ── Progress ──
        progress_frame = ttk.Frame(main)
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        self.progress_label = ttk.Label(progress_frame, text="Ready", font=("Helvetica", 10))
        self.progress_label.pack(side=tk.LEFT)
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=200)
        self.progress_bar.pack(side=tk.RIGHT)

        # ── Results ──
        results_frame = ttk.LabelFrame(main, text="  Results  ", padding=5)
        results_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        # Treeview for summary table
        columns = ("bus", "total", "differences", "status")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings", height=12)
        self.tree.heading("bus", text="Bus")
        self.tree.heading("total", text="Total Rows")
        self.tree.heading("differences", text="Differences")
        self.tree.heading("status", text="Status")
        self.tree.column("bus", width=120, anchor=tk.W)
        self.tree.column("total", width=100, anchor=tk.CENTER)
        self.tree.column("differences", width=100, anchor=tk.CENTER)
        self.tree.column("status", width=300, anchor=tk.W)

        scrollbar = ttk.Scrollbar(results_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # ── Status bar ──
        self.status_var = tk.StringVar(value="Select old and new DBC folders, then click Compare.")
        status_bar = ttk.Label(main, textvariable=self.status_var, font=("Helvetica", 9),
                               foreground="gray", anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(5, 0))

    def _show_about(self):
        """Hidden about dialog."""
        about = tk.Toplevel(self.root)
        about.title("About")
        about.resizable(False, False)
        about.transient(self.root)
        about.grab_set()
        w, h = 420, 240
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (w // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (h // 2)
        about.geometry(f"{w}x{h}+{x}+{y}")

        f = ttk.Frame(about, padding=20)
        f.pack(fill=tk.BOTH, expand=True)
        ttk.Label(f, text="DBC Compare Tool", font=("Helvetica", 16, "bold")).pack(pady=(0, 5))
        ttk.Label(f, text="Make Peace, Not War", font=("Helvetica", 12, "italic")).pack(pady=(0, 10))
        ttk.Separator(f, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        ttk.Label(f, text="See CHECKSUM.sha256 to verify author",
                  font=("Helvetica", 9), foreground="gray").pack(pady=(5, 5))

        link = ttk.Label(f, text="Find yourself in peace (TED Talk)",
                         font=("Helvetica", 10, "underline"), foreground="blue", cursor="hand2")
        link.pack(pady=(0, 10))
        link.bind("<Button-1>", lambda e: webbrowser.open("https://www.youtube.com/watch?v=4-079YIasck"))

        ttk.Button(f, text="Close", command=about.destroy).pack()

    def _browse_folder(self, var, title):
        path = filedialog.askdirectory(title=title)
        if path:
            var.set(path)

    def _open_output(self):
        out = self.output_folder.get()
        if out and os.path.isdir(out):
            if sys.platform == "darwin":
                os.system(f'open "{out}"')
            elif sys.platform == "win32":
                os.startfile(out)
            else:
                os.system(f'xdg-open "{out}"')

    def _start_compare(self):
        old = self.old_folder.get().strip()
        new = self.new_folder.get().strip()

        if not old or not os.path.isdir(old):
            messagebox.showerror("Error", "Please select a valid OLD version folder.")
            return
        if not new or not os.path.isdir(new):
            messagebox.showerror("Error", "Please select a valid NEW version folder.")
            return

        # Determine output
        out = self.output_folder.get().strip()
        if not out:
            old_name = os.path.basename(old.rstrip('/\\'))
            new_name = os.path.basename(new.rstrip('/\\'))
            out = os.path.join(
                os.path.dirname(old.rstrip('/\\')),
                f"DBC_Compare_{old_name}_vs_{new_name}"
            )
            self.output_folder.set(out)

        self.is_running = True
        self.compare_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self.progress_bar["value"] = 0

        # Run comparison in background thread
        thread = threading.Thread(target=self._run_compare, args=(old, new, out), daemon=True)
        thread.start()

    def _run_compare(self, old_folder, new_folder, output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
            old_name = os.path.basename(old_folder.rstrip('/\\'))
            new_name = os.path.basename(new_folder.rstrip('/\\'))

            old_dbc = find_dbc_files(old_folder)
            new_dbc = find_dbc_files(new_folder)

            if not old_dbc and not new_dbc:
                self._ui_update(lambda: messagebox.showerror("Error", "No .dbc files found in either folder."))
                self._finish()
                return

            all_buses = sorted(set(list(old_dbc.keys()) + list(new_dbc.keys())))
            total_buses = len(all_buses)
            total_rows_all = 0
            total_diffs_all = 0

            for idx, bus_prefix in enumerate(all_buses):
                bus_name = extract_bus_name(bus_prefix)
                old_path = old_dbc.get(bus_prefix)
                new_path = new_dbc.get(bus_prefix)

                self._ui_update(lambda bn=bus_name, i=idx: (
                    self.progress_label.config(text=f"Comparing {bn}..."),
                    self._set_progress((i / total_buses) * 100)
                ))

                if not old_path or not new_path:
                    status = "Only in OLD" if old_path else "Only in NEW"
                    self._ui_update(lambda bn=bus_name, s=status: self.tree.insert(
                        "", tk.END, values=(bn, "-", "-", s)
                    ))
                    continue

                # Parse
                self._ui_update(lambda bn=bus_name: self.progress_label.config(text=f"Parsing {bn}..."))
                old_db = parse_dbc(old_path)
                new_db = parse_dbc(new_path)

                # Compare
                self._ui_update(lambda bn=bus_name: self.progress_label.config(text=f"Comparing {bn}..."))
                comparison_rows = compare_dbc_files(old_db, new_db)
                diff_count = sum(1 for _, _, has_diff, _ in comparison_rows if has_diff)
                total = len(comparison_rows)
                total_rows_all += total
                total_diffs_all += diff_count

                # Write xlsx
                self._ui_update(lambda bn=bus_name: self.progress_label.config(text=f"Writing {bn}.xlsx..."))
                old_rel = f"{old_name}\\{os.path.basename(old_path)}"
                new_rel = f"{new_name}\\{os.path.basename(new_path)}"
                base_name = f"DBC_Compare_{old_name}_vs_{new_name}_{bus_name}"
                out_file = os.path.join(output_dir, f"{base_name}.xlsx")
                write_comparison_xlsx(out_file, old_rel, new_rel, comparison_rows,
                                      old_db=old_db, new_db=new_db,
                                      old_label=old_name, new_label=new_name)

                # Generate HTML and PDF reports
                cats = categorize_changes(old_db, new_db)
                write_html_report(
                    os.path.join(output_dir, f"{base_name}.html"),
                    old_name, new_name, cats, old_db, new_db)
                write_pdf_report(
                    os.path.join(output_dir, f"{base_name}.pdf"),
                    old_name, new_name, cats, old_db, new_db)

                # Update tree
                status = "Identical" if diff_count == 0 else f"{diff_count} differences found"
                self._ui_update(lambda bn=bus_name, t=total, d=diff_count, s=status: self.tree.insert(
                    "", tk.END, values=(bn, f"{t:,}", f"{d:,}", s)
                ))

            # Add total row
            self._ui_update(lambda tr=total_rows_all, td=total_diffs_all: (
                self.tree.insert("", tk.END, values=("", "", "", "")),
                self.tree.insert("", tk.END, values=(
                    "TOTAL", f"{tr:,}", f"{td:,}",
                    f"Saved to: {os.path.basename(output_dir)}"
                ))
            ))

            self._ui_update(lambda: (
                self.progress_label.config(text="Done!"),
                self._set_progress(100),
                self.status_var.set(f"Comparison complete. Output: {output_dir}")
            ))

        except Exception as e:
            self._ui_update(lambda err=str(e): (
                messagebox.showerror("Error", f"Comparison failed:\n{err}"),
                self.progress_label.config(text="Error"),
            ))
        finally:
            self._finish()

    def _finish(self):
        self._ui_update(lambda: (
            self.compare_btn.config(state=tk.NORMAL),
            self.open_btn.config(state=tk.NORMAL),
        ))
        self.is_running = False

    def _set_progress(self, value):
        self.progress_bar["value"] = value

    def _ui_update(self, func):
        """Schedule a function to run on the main UI thread."""
        self.root.after(0, func)


def main():
    root = tk.Tk()

    # Set theme
    style = ttk.Style()
    available = style.theme_names()
    for preferred in ("clam", "aqua", "vista", "default"):
        if preferred in available:
            style.theme_use(preferred)
            break

    app = DBCCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
