Outlook Link Converter - Setup Instructions
=============================================

What does this do?
------------------
Adds a "Copy as Outlook Link" option to your right-click menu.
When you right-click a file or folder and choose it, a clickable
link is copied to your clipboard that you can paste into Outlook 365
emails. Recipients can click the link to open the file directly.


How to install
--------------
1. Double-click  OutlookLinkConverterInstaller.exe
2. A small popup will confirm it installed. That's it.

No admin password is needed. Nothing else to do.
Right-click any file or folder and you'll see "Copy as Outlook Link".


How to use
----------
1. Right-click any file or folder in File Explorer.
2. Click "Copy as Outlook Link".
3. A small notification appears confirming the link was copied.
4. Paste (Ctrl+V) into your Outlook email.
5. The recipient clicks the link to open the file.


How to uninstall
----------------
1. Open File Explorer and type this in the address bar:
      %LOCALAPPDATA%\OutlookLinkConverter
2. Double-click  Uninstall.exe
3. A popup will confirm it was removed.


Troubleshooting
---------------
- "Copy as Outlook Link" doesn't appear in the right-click menu:
    Re-run the installer. If it still doesn't show, try signing out
    and back in to Windows (or restart Explorer).

- The link doesn't work for the recipient:
    The recipient needs the same network drive mapped. For example if
    the link uses P:\ they need P:\ mapped to the same network share.

- Windows Defender blocks the installer:
    This is a common false positive with self-contained Python apps.
    Click "More info" then "Run anyway". The tool is safe.

- Right-clicking an .exe or .bat on a network drive doesn't work:
    This is a Windows limitation. It works for all other file types.


Need help?
----------
Contact IT or the person who shared this installer with you.
