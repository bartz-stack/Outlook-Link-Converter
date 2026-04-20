"""
Outlook Link Converter - Silent Installer

Copies the converter application and support files into a local directory
(%LOCALAPPDATA%\OutlookLinkConverter) and registers the "Copy as Outlook Link"
right-click context menu for files and folders.

No admin rights required - uses HKCU registry keys.
"""

import os
import shutil
import sys
import subprocess
from pathlib import Path


def bundled_path(name: str) -> Path:
    """Locate a file bundled inside the PyInstaller package."""
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", "."))
    else:
        base = Path(__file__).parent
    return base / name


def show_popup(title, message, icon_path=None, duration_ms=3000):
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

        if icon_path and Path(icon_path).exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(icon_path).resize((40, 40), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                icon_label = tk.Label(container, image=photo, bg="#2d2d2d")
                icon_label.image = photo
                icon_label.pack(side="left", padx=(0, 12))
            except Exception:
                pass

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


def register_context_menu(exe_path: Path):
    """Create HKCU registry keys for files and directories."""
    icon = exe_path.parent / "Outlook365_icon.ico"
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0

    for scope in ["*", "Directory"]:
        key = f"HKCU\\Software\\Classes\\{scope}\\shell\\OutlookLinkConverter"
        subprocess.run(
            f'reg add "{key}" /ve /d "Copy as Outlook Link" /f',
            shell=True, startupinfo=startupinfo,
        )
        subprocess.run(
            f'reg add "{key}" /v "Icon" /d "{icon}" /f',
            shell=True, startupinfo=startupinfo,
        )
        subprocess.run(
            f'reg add "{key}\\command" /ve /d "\\"{exe_path}\\" \\"%1\\"" /f',
            shell=True, startupinfo=startupinfo,
        )


def main():
    target_dir = (
        Path(os.environ.get("LOCALAPPDATA", "C:\\OutlookLinkConverter"))
        / "OutlookLinkConverter"
    )
    target_dir.mkdir(parents=True, exist_ok=True)

    files_to_copy = [
        "OutlookLinkConverter.exe",
        "gui_config.json",
        "Outlook365_icon.ico",
        "Uninstall.exe",
    ]
    for name in files_to_copy:
        src = bundled_path(name)
        if src.exists():
            shutil.copy2(src, target_dir / name)

    exe_path = target_dir / "OutlookLinkConverter.exe"
    if not exe_path.exists():
        show_popup("Installation Failed", "Converter executable missing from installer bundle.")
        sys.exit(1)

    register_context_menu(exe_path)

    icon = target_dir / "Outlook365_icon.ico"
    show_popup(
        "Outlook Link Converter Installed",
        "Right-click any file or folder to copy an Outlook link.",
        icon_path=str(icon),
        duration_ms=3000,
    )


if __name__ == "__main__":
    main()
