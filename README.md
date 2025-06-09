# Python Project Launcher GUI

A clean, modular project scaffolding tool with Excel-integrated logging, Tkinter + `ttkbootstrap` GUI, and macro/VBA automation support.

* * *

## Features

-  Create new project folders with predefined structure
- Auto-generate project metadata (`.projectmeta.json`)
- Log all actions to `project_dashboard.xlsm`
- Dropdowns and validation sourced from `DataValidation` sheet
- Launch from Excel via VBA macro or standalone `.bat` file
- Built using `ttkbootstrap`, `xlwings`, and `argparse`

* * *

## Project Structure

```plaintext
project_launcher/
├── projectgen.py              # Core CLI project generator
├── project_launcher_gui.py    # ttkbootstrap-based UI
├── project_launcher_core.py   # Modular core logic for loading, logging, tagging
├── launch_project_gui.bat     # Fast launcher (no Excel required)
├── .venv/                     # Virtual environment
└── project_dashboard.xlsm     # Excel log + dropdown definitions
```

## Requirements

- Python 3.10+
    
- Packages:
    
    `pip install ttkbootstrap openpyxl xlwings`
    
- Excel (with macros enabled)
    
- Optional: `.venv` with `pythonw.exe` for silent GUI launching
    

* * *

## Usage

### 1. From Console

`python project_launcher_gui.py`

### 2. Batch file (.bat):

`launch_project_gui.bat`

### 3. From Excel (VBA)

Assign a form button in `project_dashboard.xlsm` to this VBA macro:

```vba
Sub LaunchProjectGUI()
    Shell "C:\Users\damian\projects\newproject\.venv\Scripts\pythonw.exe C:\Users\damian\projects\newproject\project_launcher_gui.py", vbHide
End Sub

```

* * *

## Excel Setup

- Sheet `DataValidation`: Defines dropdown values (Type, Language, Status, etc.)
    
- Sheet `Dashboard`: Logs:
    
    - Timestamp
        
    - Project name
        
    - Type, Language, Status, Tags
        
    - Dev path
        
    - Output result / error
        

* * *

## Project Metadata Output

Each project contains a `.projectmeta.json` like:

```json
{
  "project_name": "test-fastapi-cli",
  "created": "2025-06-08T11:52:03",
  "type": "backend",
  "lang": "python",
  "status": "Planning",
  "tags": ["fastapi", "cli", "python"],
  ...
}

```

* * *

## Future Enhancements (Ideas)

- Support multiple frameworks
    
- Auto-open VS Code after creation
    
- Git init + remote hook
    
- `.csv` fallback if Excel is locked
    
- Packaged `.exe` with `pyinstaller`
    

* * *

## Author

Built by Damian Damjanovic for personal project scaffolding, automation and centralized tracking.

* * *

> Designed to launch fast, log smart, and grow with your workflow.

* * *
