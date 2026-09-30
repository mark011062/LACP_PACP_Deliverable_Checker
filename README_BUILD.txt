ACME PACP/LACP CHECKER - BUILD INSTRUCTIONS

1. Extract this ZIP to a normal folder such as C:\Acme_Checker.
2. Open the Acme_Checker folder.
3. Double-click BUILD.bat.
4. Wait for the build to finish.
5. The finished program will be:
   dist\pacp_lacp_checker.exe

BUILD.bat prefers Python 3.14, then 3.13, then 3.12 when installed.
It recreates the local .venv each time, so failed/old environments do not interfere.

Required packages are installed automatically:
- pyodbc 5.3.0
- PyInstaller 6.22.2
