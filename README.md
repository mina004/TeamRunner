# TeamRunner – Beulah Inc. Workbook Integration GUI

A Python/Tkinter-based GUI that:

- Lets the user **select a Beulah Inc. company workbook (.xlsx)** from any folder  
- **Validates** that the workbook matches the expected format  
- Automatically starts running configured **.exe connector programs** (dummy or real)  
- Tracks progress with a **progress bar**  
- Reads each executable’s **JSON output**, displays it in a log window (newest first)  
- Generates a combined, human-readable **final report** and cleans up temporary JSON files.

> This project is part of the CS Project — Fall 2025 course and is designed to integrate multiple QuickBooks connector teams into one easy-to-use GUI.

## ✨ Features

- **GUI built with Tkinter**
  - Single main window  
  - Button to **Select Company Workbook**  
  - **Progress bar** showing task completion  
  - Scrollable **log area** with newest entries at the top  

- **Workbook selection & auto-start**
  - User selects a `.xlsx` file from any folder  
  - As soon as a valid workbook is selected, the GUI **automatically begins** executing tasks  

- **Workbook validation**
  - Ensures:
    - File ends with `.xlsx`  
    - File is readable  
    - File is not empty  
  - Special Beulah Inc. validation message:
    > `File does not match Beulah Inc. format. Please reach out to David Nevill dnevill@beulahinc.com`

- **Executable integration**
  - Looks for `.exe` files inside the `executables/` folder  
  - Passes the workbook path using:  
    ```
    --workbook <path_to_workbook>
    ```
  - Passes output path for JSON:  
    ```
    --output <path_to_json>
    ```
  - If an exe is missing, the GUI **simulates** the task for demo purposes

- **JSON parsing & log display**
  - JSON results appear at the **top** of the log  
  - Log includes headers, separators, timestamps

- **Final Report**
  - After all tasks complete:
    - A final report is generated at:  
      `reports/report_YYYYmmdd_HHMMSS.txt`  
    - Temporary JSON files are deleted  

## 📂 Project Structure

```
TeamRunner/
├── gui.py
├── README.md
├── executables/
│   └── payment_terms_dummy.exe
└── reports/
```

## 🚀 Running the Application

### Install optional dependency
```
pip install openpyxl
```

### Run directly
```
python gui.py
```

### Build a Windows EXE
```
pyinstaller --onefile --windowed gui.py
```

## 👤 Author
- **Student:** Minahil Rao  
- **Course:** CS Project — Fall 2025  
- **Organization:** Beulah Inc.
