# Workload Merge Console

## 1. Purpose

Workload Merge Console is a browser-based Excel processing system for
combining workload-related Excel files and verifying subject codes.

The system: - Processes the required Excel input files. - Merges
workload information. - Verifies subject codes. - Auto-corrects certain
unique one-character/code-format mistakes. - Replaces valid subject
names with the official name in **UPPERCASE**. - Keeps unmatched/invalid
Excel rows unchanged. - Generates the final Excel file for download.

------------------------------------------------------------------------

## 2. System Files

Keep all files in the same folder:

``` text
Tested system/
├── Workload_Merge_Console.bat
├── server.js
├── workload_merge_console_corrected.html
└── README.md
```

### `Workload_Merge_Console.bat`

The normal one-click launcher.

Double-click this file. It starts the local server and opens the
application automatically.

### `server.js`

Runs the local Node.js server. It serves the HTML application and
retrieves official UVPCE pages server-side so browser CORS restrictions
do not prevent B.Tech verification.

### `workload_merge_console_corrected.html`

The main application. It contains the interface, Excel processing,
subject verification, auto-correction, result display, and final Excel
generation.

### `README.md`

This instruction file.

------------------------------------------------------------------------

## 3. Requirements

The computer needs: - Windows - Node.js - A modern browser such as
Chrome, Edge, or Brave - Internet access for official B.Tech
verification

To check Node.js:

``` cmd
node --version
```

A version such as `v24.16.0` confirms Node.js is installed.

------------------------------------------------------------------------

## 4. How to Start the System

### Recommended method

Do **not** manually run CMD or `node server.js`.

1.  Open the system folder.
2.  Double-click:

``` text
Workload_Merge_Console.bat
```

3.  The browser should open automatically at:

``` text
http://127.0.0.1:8787
```

4.  Upload the required Excel files.
5.  Click **Run Pipeline**.
6.  Download the generated final Excel file.

Keep the black server window open while using the application.

------------------------------------------------------------------------

## 5. Required Excel Inputs

The application uses the input files requested by the upload fields in
the interface. The normal inputs are:

1.  Coordinator
2.  Subject Faculty
3.  Workload
4.  M.Tech Scheme

Use the correct workbook for each field.

The **M.Tech Scheme** workbook is the source used for M.Tech
subject-code verification.

------------------------------------------------------------------------

# 6. Subject-Code Verification Rules

These rules must be preserved.

## M.Tech

M.Tech subject codes are verified **ONLY against the provided M.Tech
Excel**.

The B.Tech website is not used for M.Tech verification.

### Valid M.Tech code

The subject name is replaced with the official name from the M.Tech
Excel and converted to **UPPERCASE**.

### Correctable M.Tech code

If there is exactly one valid correction: - Correct the subject code. -
Replace the subject name with the official M.Tech name. - Convert the
official name to uppercase.

### Unmatched M.Tech code

The original Excel row remains unchanged.

For example:

``` text
Code: 3ABC999
Name: Existing Subject Name
```

remains exactly:

``` text
3ABC999
Existing Subject Name
```

The system must not replace the name with `INVALID CODE`.

------------------------------------------------------------------------

## B.Tech

B.Tech subject codes are verified **ONLY against the official UVPCE
website**.

The system uses the official syllabus source for the relevant B.Tech
stream, including the configured CE, CE-AI, and CSBS sources.

CE/IT uses the configured CE syllabus source.

### Valid B.Tech code

The system: 1. Finds the official code. 2. Gets the official subject
name. 3. Replaces the Excel subject name. 4. Converts the official name
to **UPPERCASE**.

Example:

``` text
Official name:
Database Management System
```

becomes:

``` text
DATABASE MANAGEMENT SYSTEM
```

### Correctable B.Tech code

If there is exactly one valid correction: - Correct the subject code. -
Replace the subject name with the official website name. - Convert the
official name to uppercase.

### Unmatched B.Tech code

The original Excel row remains unchanged.

Do not write:

``` text
INVALID CODE
```

into the Excel subject-name cell.

------------------------------------------------------------------------

# 7. Auto-Correction

The system can correct certain obvious subject-code mistakes.

Examples include: - `I` versus `1` - A single-character difference when
exactly one official candidate exists

A correction should only be made when there is **one unique valid
match**.

The intended flow is:

``` text
Input Code
   ↓
Normalize
   ↓
Find exact match
   ↓
If no match, try allowed correction
   ↓
Exactly one valid candidate?
   ├── YES → Correct code + official name
   └── NO  → Keep original row
```

------------------------------------------------------------------------

# 8. Unmatched Codes

An unmatched code means that no valid matching code was found in the
required source.

The original: - Subject Code - Subject Name

must remain unchanged.

The application can still count/display the code as unmatched for
reporting, but it must not overwrite the original Excel data.

------------------------------------------------------------------------

# 9. Official Website Verification

For B.Tech, the official UVPCE website is the source of truth.

The system must not silently use an old bundled or hard-coded
subject-code cache instead of the official website.

The intended flow is:

``` text
B.Tech Excel
     ↓
Subject Code
     ↓
Local server
     ↓
Official UVPCE website
     ↓
Official Subject Code + Subject Name
     ↓
Verification
     ↓
Valid / Corrected / Unmatched
```

`server.js` is only a local transport layer for accessing the official
website. It is not a replacement source of subject data.

------------------------------------------------------------------------

# 10. Why the HTML Should Not Be Double-Clicked

Do not normally open:

``` text
workload_merge_console_corrected.html
```

directly.

Opening it directly creates a `file://` browser page. Browser
security/CORS restrictions can prevent direct requests to the UVPCE
website.

That can produce:

``` text
Failed to fetch
```

or many:

``` text
skipped
```

rows.

Always start the application with:

``` text
Workload_Merge_Console.bat
```

and use:

``` text
http://127.0.0.1:8787
```

------------------------------------------------------------------------

# 11. Normal Workflow

``` text
1. Open system folder
        ↓
2. Double-click Workload_Merge_Console.bat
        ↓
3. Browser opens
        ↓
4. Upload required Excel files
        ↓
5. Click Run Pipeline
        ↓
6. System processes and merges data
        ↓
7. M.Tech → M.Tech Excel verification
        ↓
8. B.Tech → Official UVPCE verification
        ↓
9. Valid → Official uppercase name
        ↓
10. Correctable → Correct code + official uppercase name
        ↓
11. Unmatched → Original Excel values retained
        ↓
12. Final Excel generated
        ↓
13. Download final_input.xlsx
```

------------------------------------------------------------------------

# 12. Understanding Verification Results

### Valid

The subject code exactly matches an official code.

### Auto-corrected

The input contained a supported typo and was changed to a unique valid
code.

### Unmatched

No valid code was found. The original Excel row is retained.

### Skipped

Verification could not be completed for that source. Do not treat
skipped rows as verified.

If a large number of B.Tech rows are skipped, check the server, internet
connection, and official UVPCE website access before trusting the
result.

------------------------------------------------------------------------

# 13. Troubleshooting

## `node is not recognized`

Run:

``` cmd
node --version
```

If Windows says Node is not recognized, install Node.js and reopen the
application.

------------------------------------------------------------------------

## `Cannot find module ... server.js`

The server was started from the wrong folder.

Make sure `server.js` is in the same folder as the launcher and HTML
file.

Normally this is avoided by using:

``` text
Workload_Merge_Console.bat
```

------------------------------------------------------------------------

## `ENOENT ... workload_merge_console_corrected.html`

The server cannot find the HTML file.

Check that these are in the same folder:

``` text
server.js
workload_merge_console_corrected.html
```

Also make sure the HTML filename has not accidentally become:

``` text
workload_merge_console_corrected.html.html
```

------------------------------------------------------------------------

## `Failed to fetch`

First confirm the application was started with:

``` text
Workload_Merge_Console.bat
```

and not by opening the HTML directly.

Then confirm: - Internet is working. - The official UVPCE website is
reachable. - The black server window is still open.

------------------------------------------------------------------------

## Browser does not open automatically

Open this manually while the server is running:

``` text
http://127.0.0.1:8787
```

------------------------------------------------------------------------

## Server window was closed

Double-click:

``` text
Workload_Merge_Console.bat
```

again.

------------------------------------------------------------------------

# 14. Sharing the System

When giving this system to another person, send the complete folder:

``` text
Workload_Merge_Console.bat
server.js
workload_merge_console_corrected.html
README.md
```

The recipient should install Node.js once.

After that, their normal procedure is only:

``` text
Double-click Workload_Merge_Console.bat
        ↓
Upload Excel files
        ↓
Run Pipeline
        ↓
Download final Excel
```

They do not need to manually open CMD or understand Node.js for normal
use.

------------------------------------------------------------------------

# 15. Important Rules for Future Developers

Do not change these source rules unless the project requirements are
intentionally changed:

``` text
M.Tech → M.Tech Excel ONLY
B.Tech  → Official UVPCE website ONLY
```

Do not replace official B.Tech verification with a cached/hard-coded
code list.

Name rules:

``` text
Valid code     → Official name → UPPERCASE
Corrected code → Official name → UPPERCASE
```

Unmatched rule:

``` text
Unmatched code → Keep original Excel row unchanged
```

Do not write `INVALID CODE` into the original subject-name cell.

The local server is required for reliable browser access to the official
B.Tech website.

------------------------------------------------------------------------


---

# 17. Port Number Conflict

The default local server port is:

```text
8787
```

If another application/system is already using port `8787`, the Workload Merge Console may not start correctly.

In that case, you can change the port number in **`server.js`**.

Find this line near the top:

```javascript
const PORT = 8787;
```

Change `8787` to another available port, for example:

```javascript
const PORT = 8080;
```

or:

```javascript
const PORT = 9000;
```

### Important: Update the URL if the port is changed

If you change the port from:

```text
8787
```

to:

```text
8080
```

the application URL becomes:

```text
http://127.0.0.1:8080
```

The `Workload_Merge_Console.bat` launcher normally opens the configured application URL. Therefore, when changing the port, make sure the port used by the launcher matches the port in `server.js`.

For example:

```text
server.js
PORT = 8080
```

and the launcher should open:

```text
http://127.0.0.1:8080
```

### Choosing a port

Use a port that is not already being used by another application.

Common alternatives include:

```text
8080
8081
8888
9000
9090
```

If a selected port is also occupied, choose another available port.

### Quick troubleshooting

If the server reports an error similar to:

```text
EADDRINUSE
```

it usually means the selected port is already being used.

Change:

```javascript
const PORT = 8787;
```

to another available port and make sure the launcher uses the same port.


# 17. Quick Start

For a normal user:

1.  Keep all system files in one folder.
2.  Install Node.js once.
3.  Double-click `Workload_Merge_Console.bat`.
4.  Wait for the browser to open.
5.  Upload the required Excel files.
6.  Click **Run Pipeline**.
7.  Check the verification results.
8.  Download the final Excel.

**No manual CMD commands are required for normal use.**
