# LJexm 1.0.5

## New Features
- Added support for converting **HTM and HTML** files to PDF using `webshot` and `PhantomJS`.
- File extension handling is now case-insensitive, ensuring that .pdf, .htm, and .html files are processed correctly regardless of letter casing (e.g., .PDF, .HTM, .HTML).

## Enhancements
- **Excel Print Area Fix**: Resolved an issue where only the print area was being converted to PDF. Now, the entire sheet is properly included in the output.
- Improved handling of missing PhantomJS installation—now automatically installs if not present.

## Bug Fixes
- Fixed minor inconsistencies in file handling during extraction and conversion.

---

# LJexm 1.0.4

- Addressed note regarding possibly misspelled words in the **`DESCRIPTION`** file (i.e., **'PDFs'** and **'VBScript'**).
- Updated the **`DESCRIPTION`** file to use single quotes for software names and clarified platform support (i.e., **Windows-only**).
- Fixed testing issues by ensuring the **`run_app()`** function gracefully handles non-Windows platforms.
- Used a mock environment during tests to avoid platform-specific dependency issues with **'VBScript'** and **'wscript.exe'**.
- **Updated error handling**: Errors are now **logged to `conversion_log.txt`**, printed in R, and **deleted automatically** after execution.
- Updated DESCRIPTION file to correctly format software names (e.g., 'pdf', 'zip') as per CRAN's request.
- Fixed CRAN test failure by skipping `run_app()` test on non-Windows systems.  
- Improved test handling for non-interactive environments.  
- Ensured package compliance with CRAN’s testing policies.
- Replaced all cat() statements with message() or warning() as per CRAN guidelines.
- Ensured that status messages are now printed using message(), allowing suppressibility via suppressMessages().
- Ensured warnings are handled appropriately with warning() where necessary.
- Maintained compatibility with CRAN's policies regarding console message handling and logging behavior.

