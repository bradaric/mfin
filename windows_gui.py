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


REGISTRY_PATH = r"Software\Classes\SystemFileAssociations\.pdf\shell\MfinExtract"


def install_context_menu():
    """Register the right-click context menu entry for PDF files (per-user).

    Returns a status message string.
    """
    if sys.platform != "win32":
        return "Context menu installation is only supported on Windows."

    import winreg

    exe_path = os.path.abspath(sys.executable)
    # When frozen by PyInstaller, sys.executable is the .exe itself
    # When running as script, point to this script via pythonw
    if not getattr(sys, 'frozen', False):
        exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
    else:
        exe_path = f'"{exe_path}"'

    try:
        # Create shell key with display name
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Extract Tables (mfin)")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)

        # Create command subkey
        cmd_key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command"
        )
        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, f'{exe_path} "%1"')
        winreg.CloseKey(cmd_key)

        return "Context menu installed successfully.\nRight-click any PDF to see 'Extract Tables (mfin)'."
    except OSError as e:
        return f"Failed to install context menu: {e}"


def uninstall_context_menu():
    """Remove the right-click context menu entry.

    Returns a status message string.
    """
    if sys.platform != "win32":
        return "Context menu removal is only supported on Windows."

    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        return "Context menu removed successfully."
    except FileNotFoundError:
        return "Context menu entry not found (already removed?)."
    except OSError as e:
        return f"Failed to remove context menu: {e}"


def show_setup_dialog():
    """Show a simple GUI for installing/uninstalling the right-click menu."""
    root = tk.Tk()
    root.title("mfin — Setup")
    root.geometry("400x200")
    root.resizable(False, False)

    tk.Label(
        root,
        text="mfin — PDF Table Extractor",
        font=("", 14, "bold"),
        pady=10,
    ).pack()

    tk.Label(
        root,
        text="Add or remove the right-click menu entry for PDF files.",
        pady=5,
    ).pack()

    status_var = tk.StringVar()
    status_label = tk.Label(root, textvariable=status_var, wraplength=360, pady=10)
    status_label.pack()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=5)

    def on_install():
        msg = install_context_menu()
        status_var.set(msg)

    def on_uninstall():
        msg = uninstall_context_menu()
        status_var.set(msg)

    tk.Button(btn_frame, text="Install", width=12, command=on_install).pack(
        side="left", padx=10
    )
    tk.Button(btn_frame, text="Uninstall", width=12, command=on_uninstall).pack(
        side="left", padx=10
    )

    root.mainloop()


def main():
    if len(sys.argv) < 2:
        show_setup_dialog()
        return

    arg = sys.argv[1]

    if arg == "--install":
        print(install_context_menu())
    elif arg == "--uninstall":
        print(uninstall_context_menu())
    elif os.path.isfile(arg) and arg.lower().endswith(".pdf"):
        run_gui(arg)
    else:
        print(f"Unknown argument or file not found: {arg}")
        sys.exit(1)


if __name__ == "__main__":
    main()
