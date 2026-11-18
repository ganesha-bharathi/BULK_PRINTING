WORD BULK PRINTER - DESKTOP APPLICATION BUILD KIT

This package builds a Windows .exe so you can open the printer app by double-clicking it. You do not need to run the .py file each time.

WHAT YOU NEED (one time on the PC used to build the EXE)
- Windows 10 or Windows 11
- Python 3.10 or newer
- Microsoft Word desktop installed
- A configured Windows default printer
- Internet connection for the first-time package installation

BUILD THE APPLICATION
1. Extract this ZIP folder.
2. Double-click Build_Windows_EXE.bat.
3. Wait for the build to finish. It installs openpyxl, pywin32 and PyInstaller, then packages the program.
4. The EXE will be created at: dist\WordBulkPrinter.exe
5. Copy WordBulkPrinter.exe to your Desktop. Double-click it whenever you need to print.

USING THE APP
1. Open WordBulkPrinter.exe.
2. Browse to your Excel .xlsx workbook.
3. Browse to the folder containing the Word documents.
4. Click Check files & build preview.
5. Review all IDs, quantities, missing files and ambiguous filename matches.
6. Ensure the intended printer is the Windows default printer.
7. Click Confirm & print.

IMPORTANT
- The Excel sheet must contain headers External Id and ASN quantity.
- The program uses the active worksheet in the workbook.
- Word files can be .docx, .doc or .docm. The filename (without extension) must match the External ID.
- Matching is case-insensitive and searches subfolders.
- Duplicate External IDs are summed across the worksheet.
- IDs with missing files or more than one matching file are skipped and shown in the preview.
- The app sends one Word print command per requested copy, using the working copy-count fix.
- Microsoft Word desktop must be installed. The EXE does not include Microsoft Word or a printer driver.
- Windows may show a SmartScreen warning for an EXE built locally because it is not code-signed. Only run the EXE you built yourself from this source package.
- Test a small print batch first. A message that jobs were submitted does not guarantee the printer physically completed them.
