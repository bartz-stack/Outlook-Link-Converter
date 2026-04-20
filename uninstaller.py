"""
Outlook Link Converter - Silent Uninstaller

Removes the "Copy as Outlook Link" context menu entries and deletes
the local installation folder. No admin rights required.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def show_popup(title, message, duration_ms=3000):
    """Show a friendly popup notification that auto-closes."""
    try:
        import tkinter as tk

        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes("-topmost", True)
        root.configure(bg="#2d2d2d")

        window_width = 380
        window_height = 90
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        container = tk.Frame(root, bg="#2d2d2d")
        container.pack(fill="both", expand=True, padx=12, pady=10)

        text_frame = tk.Frame(container, bg="#2d2d2d")
        text_frame.pack(side="left", fill="both", expand=True)

        tk.Label(
            text_frame, text=title, font=("Segoe UI", 11, "bold"),
            fg="white", bg="#2d2d2d", anchor="w",
        ).pack(fill="x")

        tk.Label(
            text_frame, text=message, font=("Segoe UI", 9),
            fg="#cccccc", bg="#2d2d2d", anchor="w",
        ).pack(fill="x", pady=(2, 0))

        root.after(duration_ms, root.destroy)
        root.mainloop()
    except Exception:
        pass


def main():
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0

    # Remove registry entries
    for scope in ["*", "Directory"]:
        key = f"HKCU\\Software\\Classes\\{scope}\\shell\\OutlookLinkConverter"
        subprocess.run(
            f'reg delete "{key}" /f',
            shell=True, startupinfo=startupinfo,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )

    # Remove installed files
    install_dir = (
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "OutlookLinkConverter"
    )
    if install_dir.exists():
        # Can't delete our own exe while running, so schedule deletion
        # via a fire-and-forget cmd that waits a moment then removes the folder
        cmd = f'ping -n 2 127.0.0.1 >nul & rmdir /s /q "{install_dir}"'
        subprocess.Popen(
            cmd, shell=True, startupinfo=startupinfo,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )

    show_popup(
        "Outlook Link Converter Removed",
        "Context menu entry and local files have been removed.",
        duration_ms=3000,
    )


if __name__ == "__main__":
    main()
