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
            from extract_tables import process_pdf

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

# Windows 11 hides legacy context menu items behind "Show more options".
# This well-known per-user registry key restores the full classic context menu.
WIN11_CLASSIC_MENU_PATH = (
    r"Software\Classes\CLSID"
    r"\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
)


def _is_windows_11():
    """Check if running on Windows 11 (build >= 22000)."""
    try:
        ver = sys.getwindowsversion()
        return ver.build >= 22000
    except AttributeError:
        return False


def _has_classic_context_menu():
    """Check if the classic context menu tweak is already applied."""
    import winreg

    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, WIN11_CLASSIC_MENU_PATH)
        winreg.CloseKey(key)
        return True
    except OSError:
        return False


def _enable_classic_context_menu():
    """Apply registry tweak to show the full context menu on Windows 11.

    Creates an empty InprocServer32 key that tells Explorer to use the
    classic context menu instead of the simplified Windows 11 one.
    """
    import winreg

    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, WIN11_CLASSIC_MENU_PATH)
    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "")
    winreg.CloseKey(key)


def _disable_classic_context_menu():
    """Remove the classic context menu registry tweak."""
    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, WIN11_CLASSIC_MENU_PATH)
        parent = r"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, parent)
    except OSError:
        pass


def _restart_explorer():
    """Restart Windows Explorer so context menu changes take effect."""
    import subprocess

    subprocess.run(["taskkill", "/f", "/im", "explorer.exe"], capture_output=True)
    subprocess.Popen(["explorer.exe"])


def install_context_menu(fix_win11=False):
    """Register the right-click context menu entry for PDF files (per-user).

    Args:
        fix_win11: If True, also apply the Windows 11 classic context menu
            tweak so the entry is visible without clicking "Show more options".

    Returns a status message string.
    """
    if sys.platform != "win32":
        return "Context menu installation is only supported on Windows."

    import winreg

    exe_path = os.path.abspath(sys.executable)
    if not getattr(sys, "frozen", False):
        exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
    else:
        exe_path = f'"{exe_path}"'

    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Extract Tables (mfin)")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)

        cmd_key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command"
        )
        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, f'{exe_path} "%1"')
        winreg.CloseKey(cmd_key)

        msg = "Context menu installed successfully.\nRight-click any PDF to see 'Extract Tables (mfin)'."

        if fix_win11:
            _enable_classic_context_menu()
            _restart_explorer()
            msg += (
                "\n\nWindows 11 fix applied: full context menu restored."
                "\nExplorer has been restarted."
            )

        return msg
    except OSError as e:
        return f"Failed to install context menu: {e}"


def uninstall_context_menu(undo_win11_fix=False):
    """Remove the right-click context menu entry.

    Args:
        undo_win11_fix: If True, also remove the classic context menu tweak.

    Returns a status message string.
    """
    if sys.platform != "win32":
        return "Context menu removal is only supported on Windows."

    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        msg = "Context menu removed successfully."
    except FileNotFoundError:
        msg = "Context menu entry not found (already removed?)."
    except OSError as e:
        return f"Failed to remove context menu: {e}"

    if undo_win11_fix and _has_classic_context_menu():
        _disable_classic_context_menu()
        _restart_explorer()
        msg += (
            "\nWindows 11 classic menu fix removed."
            "\nExplorer has been restarted."
        )

    return msg


def show_setup_dialog():
    """Show a simple GUI for installing/uninstalling the right-click menu."""
    root = tk.Tk()
    root.title("mfin — Setup")

    is_win11 = _is_windows_11()
    root.geometry("420x250" if is_win11 else "400x200")
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

    # Windows 11: offer to restore the classic full context menu
    win11_var = tk.BooleanVar(value=is_win11)
    if is_win11:
        tk.Checkbutton(
            root,
            text="Show in top-level menu on Windows 11\n"
                 "(restores classic right-click menu, restarts Explorer)",
            variable=win11_var,
            justify="left",
        ).pack(pady=(0, 5))

    status_var = tk.StringVar()
    tk.Label(root, textvariable=status_var, wraplength=380, pady=10).pack()

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=5)

    def on_install():
        msg = install_context_menu(fix_win11=win11_var.get())
        status_var.set(msg)

    def on_uninstall():
        msg = uninstall_context_menu(undo_win11_fix=win11_var.get())
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
