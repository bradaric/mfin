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

# CLSID for our IExplorerCommand shell extension (stable, generated once).
SHELL_EXT_CLSID = "{e7a20a14-3b78-4d9a-9c12-5f4b8a3e6d01}"

# Per-user COM class registration path.
CLSID_REG_PATH = rf"Software\Classes\CLSID\{SHELL_EXT_CLSID}"


def _get_shell_ext_dll_path():
    """Return expected path of mfin_shell.dll next to the executable."""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "mfin_shell.dll")


def _is_windows_11():
    """Check if running on Windows 11 (build >= 22000)."""
    try:
        ver = sys.getwindowsversion()
        return ver.build >= 22000
    except AttributeError:
        return False


def _shell_ext_available():
    """Check if the shell extension DLL is present."""
    return os.path.isfile(_get_shell_ext_dll_path())


def _register_shell_extension():
    """Register the IExplorerCommand COM DLL per-user for .pdf files.

    Sets ExplorerCommandHandler on the shell verb so Windows 11 uses the
    COM DLL for the modern context menu, while Win10 ignores it and uses
    the command subkey as a fallback.
    """
    import winreg

    dll_path = _get_shell_ext_dll_path()

    # Register CLSID -> InprocServer32 pointing to our DLL
    inproc_path = CLSID_REG_PATH + r"\InprocServer32"
    key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, inproc_path)
    winreg.SetValueEx(key, "", 0, winreg.REG_SZ, dll_path)
    winreg.SetValueEx(key, "ThreadingModel", 0, winreg.REG_SZ, "Apartment")
    winreg.CloseKey(key)

    # Point the shell verb to our COM handler via ExplorerCommandHandler.
    # Windows 11 uses this for the modern context menu; Win10 ignores it
    # and falls through to the command subkey.
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_SET_VALUE
    )
    winreg.SetValueEx(
        key, "ExplorerCommandHandler", 0, winreg.REG_SZ, SHELL_EXT_CLSID
    )
    winreg.CloseKey(key)


def _unregister_shell_extension():
    """Remove per-user COM registration for the shell extension."""
    import winreg

    # Remove ExplorerCommandHandler from the shell verb
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH, 0, winreg.KEY_SET_VALUE
        )
        winreg.DeleteValue(key, "ExplorerCommandHandler")
        winreg.CloseKey(key)
    except OSError:
        pass

    # Remove CLSID registration
    for path in [
        CLSID_REG_PATH + r"\InprocServer32",
        CLSID_REG_PATH,
    ]:
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
        except OSError:
            pass


def install_context_menu():
    """Register the right-click context menu entry for PDF files (per-user).

    On Windows 11 with the shell extension DLL present, registers the COM
    IExplorerCommand so the item appears in the modern context menu.
    Otherwise falls back to the classic registry approach.

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
        # Always register the classic shell entry (works on Win10, fallback on Win11)
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

        # On Windows 11, also register the shell extension for the modern menu
        if _is_windows_11() and _shell_ext_available():
            _register_shell_extension()
            msg += "\n\nWindows 11: registered shell extension for top-level menu."
        elif _is_windows_11():
            msg += (
                "\n\nNote: mfin_shell.dll not found — on Windows 11 the menu item"
                "\nwill appear under 'Show more options'."
            )

        return msg
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
        msg = "Context menu removed successfully."
    except FileNotFoundError:
        msg = "Context menu entry not found (already removed?)."
    except OSError as e:
        return f"Failed to remove context menu: {e}"

    _unregister_shell_extension()
    return msg


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
