# Faculty Role Assigner

## 1. What is this project?

**Faculty Role Assigner** is a Windows-based web application that automatically assigns faculty members to examination roles using faculty workload data from an Excel (`.xlsx`) file.

The application:
- Detects each subject as **Theory**, **Lab**, or **Lab+Theory**.
- Assigns **Paper Setter I** and **Paper Setter II**.
- Generates **Practical Only**, **Examiner**, and **Reserve** assignments where applicable.
- Creates separate Excel sheets for each semester.
- Creates a **Faculty Counter** sheet for setter workload/repetition.
- Uses colors to highlight unavailable assignments and repeated setters.

The main algorithm is in `app_prot.py`. `launcher.py` starts the local web application and opens the browser automatically.

---

## 2. Project files

Keep these files together:

| File | Purpose |
|---|---|
| `START.bat` | Recommended Windows startup file |
| `launcher.py` | Finds a free port, starts Flask, and opens the browser |
| `app_prot.py` | Main application, assignment algorithm, and Excel generation |
| `build_exe.spec` | PyInstaller configuration for creating a standalone `.exe` |
| `.deps_ok` | Marker used by `START.bat` after dependencies are installed |
| `README.md` | Project instructions |

---

# 3. Quick start — Python version

### Requirements

- Windows
- Python **3.9 or later**
- Internet connection on the first run

During Python installation, enable **"Add Python to PATH"**.

### Steps

1. Copy the complete project folder to the Windows computer.
2. Keep all project files in the same folder.
3. Double-click `START.bat`.
4. On the first run, it installs:
   - Flask
   - Pandas
   - OpenPyXL
5. The browser opens automatically.
6. Upload the required `.xlsx` workload file.
7. Click **Generate assignment file**.
8. The generated file is downloaded as `faculty_assignments.xlsx`.
9. Close the Command Prompt window to stop the application.

After a successful first setup, `.deps_ok` is created, so the dependencies are not installed again by `START.bat`.

---

# 4. Input Excel file

The uploaded file must be an Excel `.xlsx` file.

The following columns are required:

- `Stream`
- `Subject code`
- `Subject Name`
- `Is Coordinator`
- `Theory Workload`
- `Lab Workload`
- `Total Workload`
- `Faculty name`

The application groups records by **Stream** and **Subject code**.

If any required column is missing, the application displays an error instead of generating the output.

---



# 5. Output Excel file

The generated file is:

`faculty_assignments.xlsx`

## Semester sheets

A separate sheet is created for each detected semester, such as `Sem 2`, `Sem 4`, or `Sem 6`.

Each sheet contains:

| Column | Description |
|---|---|
| Sr. No. | Serial number |
| Stream | Academic stream |
| Subject code | Subject code |
| Subject Name | Subject name |
| Course Type | B. Tech., M. Tech., or Unknown |
| Subject Type | Theory, Lab, or Lab+Theory |
| Paper Setter I | First paper setter |
| Paper Setter II | Second paper setter |
| Practical Only | Practical-only faculty where applicable |
| Examiner | Examiner faculty list |
| Reserve | Reserve faculty |

## Faculty Counter

The workbook also contains `Faculty Counter`.

It shows:
- Faculty Name
- Repetition
- Semester-wise setter assignments

The repetition count is based on unique subject codes assigned to the faculty as a setter.

---

# 6. Color meanings

### Red

Red indicates that no suitable faculty/assignment was found for that role.

Examples:
- No setter available
- No examiner available
- No reserve available

### Yellow

Yellow indicates a faculty member has been repeated as a setter across different subject codes.

### `not applicable`

The role does not apply to that subject type.


# 7. Building a standalone `.exe`

The project includes `build_exe.spec` for creating a standalone Windows executable with PyInstaller.

This is useful when the recipient does not have Python installed.

### Build steps

On a Windows computer with Python installed:

1. Open Command Prompt in the project folder.
2. Run:

```text
pip install pyinstaller flask pandas openpyxl
```

3. Build the executable:

```text
pyinstaller build_exe.spec
```

4. The executable will be created as:

```text
dist\FacultyRoleAssigner.exe
```

5. Double-click the `.exe` to run it.

The supplied spec file includes `app_prot.py` and the required imports for Flask, Pandas, OpenPyXL, and the launcher.

---

# 8. Recommended distribution

## If the recipient has Python

Send the complete project folder:

```text
START.bat
launcher.py
app_prot.py
build_exe.spec
.deps_ok
README.md
```

They only need to double-click `START.bat`.

## If the recipient does not have Python

Build `FacultyRoleAssigner.exe` using `build_exe.spec`, then distribute the executable.


# 9. Technical overview

The project uses:

- **Python** — application and assignment logic
- **Flask** — local web interface
- **Pandas** — Excel input/data processing
- **OpenPyXL** — Excel output generation and formatting
- **PyInstaller** — optional standalone Windows executable

The web interface is embedded directly inside `app_prot.py`, so no separate HTML/CSS/JavaScript files are required.

The application runs locally on:

```text
127.0.0.1
```

and is intended to be used from the same Windows computer.

---

# 10. Quick reference

For normal use:

1. Keep all project files together.
2. Double-click `START.bat`.
3. Wait for the browser to open.
4. Upload the `.xlsx` workload file.
5. Click **Generate assignment file**.
6. Use the downloaded `faculty_assignments.xlsx`.
7. Close the Command Prompt window when finished.

If port `5000` is busy, the launcher automatically attempts the next available port.



-------------------------------------------------------------------------------------------------------------------------------------------

# Below is the logic of the algorithm created in the app_prot.py file, which is detecting the subject and assinging the roles to facuty:


# Subject type detection

The application automatically determines the subject type from workload values.

### Lab

If all records for the subject have:

`Theory Workload = 0`

the subject is classified as **Lab**.

### Theory

If all records for the subject have:

`Lab Workload = 0`

the subject is classified as **Theory**.

### Lab+Theory

If both theory and lab workload are present, the subject is classified as **Lab+Theory**.

---

#  Assignment logic

## Paper Setter I

For applicable subjects, the selection priority is:

1. Subject Coordinator
2. Highest `Total Workload` among all faculty
3. Highest `Total Workload` among theory faculty
4. Highest `Lab Workload` among lab faculty

The first priority level with an available candidate is used.

## Paper Setter II

Paper Setter II follows the same priority system, but the faculty selected as Paper Setter I for that subject is excluded.

The application also maintains a global setter counter. If a faculty member is assigned as a setter for different subject codes, the repeated assignment is highlighted.

## Practical Only, Examiner and Reserve

These roles are handled according to subject type.

- **Theory:** Practical Only is `not applicable`.
- **Lab:** Paper Setter I, Paper Setter II and Examiner are not applicable.
- **Lab+Theory:** setter faculty are excluded from Practical Only, Examiner and Reserve.
- If no suitable faculty can be found, the result is marked `NOT AVAILABLE`.

---