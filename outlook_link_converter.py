"""
Outlook Link Converter - Right-Click Context Menu Version
Converts network file paths to Outlook 365-compatible links
No GUI - runs silently and copies link to clipboard
"""

import sys
import subprocess
import json
from pathlib import Path
from urllib.parse import quote


def load_config():
    """Load configuration from JSON file"""
    try:
        if getattr(sys, 'frozen', False):
            script_dir = Path(sys.executable).parent
        else:
            script_dir = Path(__file__).parent
        config_path = script_dir / 'gui_config.json'
        
        if config_path.exists():
            with open(config_path, 'r') as f:
                return json.load(f)
    except Exception:
        pass
    
    return {
        "unc_to_drive_mappings": {
            "\\\\YOUR-SERVER\\share1": "P:",
            "\\\\YOUR-SERVER\\share2": "G:"
        }
    }


def copy_to_clipboard(text):
    """Copy text to clipboard using Windows clip.exe"""
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0  # SW_HIDE
    
    process = subprocess.Popen(
        ['clip'], 
        stdin=subprocess.PIPE,
        startupinfo=startupinfo
    )
    process.communicate(text.encode('utf-8'))


def show_notification(title, message, icon_path=None, duration_ms=2000):
    """Show a quick popup notification that auto-closes"""
    try:
        import tkinter as tk
        
        root = tk.Tk()
        root.overrideredirect(True)
        root.attributes('-topmost', True)
        root.configure(bg='#2d2d2d')
        
        # Position bottom-right of screen
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = 320
        window_height = 80
        x = screen_width - window_width - 20
        y = screen_height - window_height - 60
        root.geometry(f'{window_width}x{window_height}+{x}+{y}')
        
        # Main container
        container = tk.Frame(root, bg='#2d2d2d')
        container.pack(fill='both', expand=True, padx=12, pady=10)
        
        # Try to load icon
        if icon_path and Path(icon_path).exists():
            try:
                from PIL import Image, ImageTk
                img = Image.open(icon_path)
                img = img.resize((40, 40), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                icon_label = tk.Label(container, image=photo, bg='#2d2d2d')
                icon_label.image = photo  # Keep reference
                icon_label.pack(side='left', padx=(0, 12))
            except:
                pass
        
        # Text container
        text_frame = tk.Frame(container, bg='#2d2d2d')
        text_frame.pack(side='left', fill='both', expand=True)
        
        # Title
        title_label = tk.Label(
            text_frame, 
            text=title, 
            font=('Segoe UI', 11, 'bold'),
            fg='white',
            bg='#2d2d2d',
            anchor='w'
        )
        title_label.pack(fill='x')
        
        # Message
        msg_label = tk.Label(
            text_frame,
            text=message,
            font=('Segoe UI', 9),
            fg='#cccccc',
            bg='#2d2d2d',
            anchor='w'
        )
        msg_label.pack(fill='x', pady=(2, 0))
        
        # Auto-close
        root.after(duration_ms, root.destroy)
        root.mainloop()
        
    except:
        pass


def convert_path_to_outlook_link(file_path, config):
    """Convert a file path to Outlook 365-compatible link format"""
    file_path = file_path.strip().strip('"').strip("'")
    file_path = file_path.replace('/', '\\')
    
    unc_mappings = config.get("unc_to_drive_mappings", {})
    sorted_mappings = sorted(unc_mappings.items(), key=lambda x: len(x[0]), reverse=True)
    
    for unc_path, drive_letter in sorted_mappings:
        unc_path_normalized = unc_path.replace('/', '\\')
        if file_path.upper().startswith(unc_path_normalized.upper()):
            remainder = file_path[len(unc_path_normalized):].lstrip('\\')
            if not drive_letter.endswith(':'):
                drive_letter = drive_letter + ':'
            file_path = drive_letter + '\\' + remainder
            break
    
    # Encode each path segment, preserving backslashes and the drive colon
    parts = file_path.split('\\')
    encoded_parts = []
    for part in parts:
        # Keep drive letter (e.g. "P:") as-is, encode everything else
        if len(part) == 2 and part[1] == ':':
            encoded_parts.append(part)
        else:
            encoded_parts.append(quote(part, safe=''))
    encoded_path = '\\'.join(encoded_parts)
    outlook_link = f'https://outlook.office.com/local/path/file://{encoded_path}'
    
    return outlook_link


def main():
    if len(sys.argv) < 2:
        show_notification("Outlook Link Converter", "No file path provided.")
        sys.exit(1)
    
    file_path = sys.argv[1]
    config = load_config()
    outlook_link = convert_path_to_outlook_link(file_path, config)
    copy_to_clipboard(outlook_link)
    
    filename = Path(file_path).name
    
    # Get icon path
    if getattr(sys, 'frozen', False):
        script_dir = Path(sys.executable).parent
    else:
        script_dir = Path(__file__).parent
    icon_path = script_dir / 'Outlook365_icon.ico'
    
    show_notification("Link Copied!", filename, icon_path=str(icon_path), duration_ms=2000)


if __name__ == '__main__':
    main()