"""Windows GUI entry point for mfin PDF table extraction.

Usage:
    mfin.exe <path.pdf>       -- process PDF with progress window
    mfin.exe --install        -- register right-click context menu
    mfin.exe --uninstall      -- remove right-click context menu
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import scrolledtext

from extract_tables import process_pdf


class ProgressWindow:
    """A small tkinter window that shows extraction progress."""

    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.root = tk.Tk()
        self.root.title("mfin — Extracting Tables")
        self.root.geometry("500x350")
        self.root.resizable(True, True)

        # Filename label
        label = tk.Label(
            self.root,
            text=f"Processing: {os.path.basename(pdf_path)}",
            anchor="w",
            padx=10,
            pady=5,
        )
        label.pack(fill="x")

        # Scrolling log area
        self.log_area = scrolledtext.ScrolledText(
            self.root, wrap="word", state="disabled", height=15
        )
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        # Close button (disabled until done)
        self.close_btn = tk.Button(
            self.root, text="Close", state="disabled", command=self.root.destroy
        )
        self.close_btn.pack(pady=(0, 10))

        # Handle window close via X button
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._done = False

    def log(self, message):
        """Thread-safe logging to the text area."""
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.config(state="disabled")

    def _on_close(self):
        if self._done:
            self.root.destroy()

    def _mark_done(self, success):
        self._done = True
        self.close_btn.config(state="normal")
        if success:
            self.log("\n--- Done! ---")
        else:
            self.log("\n--- Failed (see errors above) ---")

    def run(self):
        """Start processing in a background thread, run the GUI main loop."""
        thread = threading.Thread(target=self._process, daemon=True)
        thread.start()
        self.root.mainloop()

    def _process(self):
        try:
            output_dir = os.path.join(os.path.dirname(self.pdf_path), "tabele")
            process_pdf(self.pdf_path, output_dir, log=self.log)
            self.root.after(0, self._mark_done, True)
        except Exception as e:
            self.log(f"\nERROR: {e}")
            self.root.after(0, self._mark_done, False)


def run_gui(pdf_path):
    """Launch the progress window for a single PDF."""
    window = ProgressWindow(pdf_path)
    window.run()
