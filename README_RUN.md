Running this Django project (Windows)

Files added:
- run_project.bat  -> Helper for cmd
- run_project.ps1  -> Helper for PowerShell

Project layout (relevant):
- [project root] (contains manage.py and requirements.txt)
  - manage.py
  - requirements.txt
  - seating/ (app)
  - seating_project/ (Django project)

Quick instructions (recommended):

1) Open PowerShell, change to project root (where this README is):

   cd D:\project\Final_project

2) Run the PowerShell helper (recommended):

   .\run_project.ps1

   - The script will create a virtual environment at D:\project\Final_project\venv if missing,
     activate it in the current PowerShell session, install packages from requirements.txt,
     apply migrations and start the development server on 0.0.0.0:8000.

3) Or use the Batch helper from Command Prompt (cmd):

   cd D:\project\Final_project
   run_project.bat

Notes and troubleshooting:
- If PowerShell prevents scripts from running, temporarily allow execution for the current session:
  Open PowerShell as Administrator and run:
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine

- If installation fails for binary packages like NumPy, try installing wheels matching your Python version:
    python -m pip install --upgrade pip
    python -m pip install numpy pandas openpyxl

- If you created the venv in a different folder (e.g., seating_project\venv), run the commands from the project root but activate that venv instead.

Testing upload endpoint:
- Open http://127.0.0.1:8000/ in your browser. The upload page is served by the `seating` app.
- Or use curl to POST files (adjust paths):
    curl -F "class_file=@C:\path\to\classes.xlsx" -F "student_file=@C:\path\to\students.xlsx" http://127.0.0.1:8000/ -o result.xlsx
