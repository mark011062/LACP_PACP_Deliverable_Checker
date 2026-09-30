Patched version notes
=====================

Updated pacp_lacp_checker.py so media folder paths are resolved more safely.

What changed:
- Database values like \Pictures are treated as Pictures, not as an absolute root path.
- If you select the Video folder, the checker can also look one level up for sibling folders like Pictures.
- This supports layouts like:

  Project Folder/
    Pictures/
    Video/

Important:
- The included dist/pacp_lacp_checker.exe from the original zip was NOT rebuilt in this sandbox.
- To use the fix immediately, run the patched pacp_lacp_checker.py with Python.
- To use it as an .exe, rebuild the executable on your Windows machine with PyInstaller.

Suggested rebuild command from inside the project folder:

  pyinstaller --onefile --windowed pacp_lacp_checker.py

