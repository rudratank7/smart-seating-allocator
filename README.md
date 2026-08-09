# Smart Seating Allocator

This file describes how to run the project, expected Excel input format, and sample files in the repository.

## 1. Project overview
- Django application with a single `seating` app.
- Upload two Excel files (class + student) via web form.
- Applies a custom block-based seating allocation algorithm.
- Returns a generated output Excel (`output_data_final_seating.xlsx`).

## 2. Folder layout
- `manage.py`
- `requirements.txt`
- `run_project.bat` (Windows cmd helper)
- `run_project.ps1` (PowerShell helper)
- `seating/` (Django app)
- `seating_project/` (Django settings/urls)
- `sample/` (example input and output Excel files)

## 3. Sample folder content
`sample/` currently contains:
- `class_data.xlsx`
- `students_data.xlsx`
- `output_data_final_seating.xlsx`

These are real data files you can use immediately for testing.

## 4. Required Python packages
Install dependencies:
```powershell
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

`requirements.txt` should include at least:
- Django
- pandas
- openpyxl

## 5. Run instructions (recommended)
### Option A: PowerShell helper
```powershell
cd "D:\Smart Seating Allocator"
.\run_project.ps1
```
This script:
- Creates/uses `venv` (in project root) if missing
- Activates virtual environment in current shell
- Installs requirements
- Runs `python manage.py migrate`
- Starts server on `0.0.0.0:8000`

### Option B: CMD helper
```cmd
cd "D:\Smart Seating Allocator"
run_project.bat
```
### Option C: manual
```powershell
cd "D:\Smart Seating Allocator"
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
python manage.py migrate
python manage.py runserver 0.0.0.0:8000
```

## 6. Web usage
- Open `http://127.0.0.1:8000/`
- Upload both files:
  - Class file: `class_file`
  - Students file: `student_file`
- Optional text field `user_data` appears in output Excel row 3.
- Response is downloadable `output_data_final_seating.xlsx`.

## 7. Excel input format
### Class file (first row header):
- `Class Room` (string)
- `Class Capacity` (integer)

Example:
| Class Room | Class Capacity |
| ---------- | -------------- |
| A1         | 30             |
| B2         | 25             |

### Students file (first row header):
- `Branch` (string)
- `No. of Students` (integer)

Example:
| Branch  | No. of Students |
| ------- | --------------- |
| CSE     | 58              |
| ECE     | 42              |

## 8. Output format
`output_data_final_seating.xlsx` includes:
- metadata rows (P P SAVANI UNIVERSITY, School of Engineering, user_data, Seating Arrangement)
- table with:
  - Sr No, Class Room, Class Capacity, Available
  - Block 1/2/3/4 branch and count columns
  - Total students, Total Block
- right summary table:
  - Branch, Original Students, Remaining Students, Class Rooms Allocated

## 9. Notes
- Empty rows fully blank are dropped from input.
- Capacity under 3 rooms are ignored by allocator.
- All numbers are coerced to numeric; invalid values become 0.
- If one file is missing or invalid in upload, HTTP 400 returned.

## 10. Troubleshooting
- On PowerShell script blocked, run as admin and execute:
  `Set-ExecutionPolicy RemoteSigned -Scope LocalMachine`
- If pandas/openpyxl install fails for binary package issues:
  `pip install numpy pandas openpyxl`

---

This file is created as the primary README with run details and sample/format references.
