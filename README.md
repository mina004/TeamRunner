TeamRunner – QuickBooks Integration GUI

CS 480 – Internship / Final Project
Purdue University Northwest (Hammond Campus)

Project Overview

This project provides a Tkinter-based GUI that integrates 10 independent QuickBooks connector executables developed by different teams.

The GUI:

Runs each team’s .exe as a subprocess

Tracks progress with a progress bar

Displays parsed JSON results in real time

Produces a final human-readable report

Cleans up temporary JSON files automatically

This project is designed to be executed inside a QuickBooks Desktop Virtual Machine (VM) using 32-bit Python, as required by the QuickBooks SDK.

⚠️ IMPORTANT: Environment Requirements
✅ Required Environment

Windows OS

QuickBooks Desktop installed

QuickBooks SDK available

32-bit Python only

Run inside QuickBooks VM

QuickBooks Desktop uses 32-bit COM objects.
Running this project with 64-bit Python or a 64-bit GUI EXE will cause COM errors such as:

Invalid class string

Class not registered

Verifying Python Bit-Version

Inside the QuickBooks VM, open Command Prompt and run:

python -c "import struct; print(struct.calcsize('P') * 8)"


Expected output:

32


If output is 64, Python must be reinstalled as 32-bit.

Folder Structure
TeamRunner/
│
├── gui.py                    # Tkinter GUI
├── Example Company Excel.xlsx # Official workbook (provided by instructor)
│
├── executables/               # All team .exe files
│   ├── Chart_of_accounts.exe
│   ├── customer_compare.exe
│   ├── vendor_compare.exe
│   ├── qb-invoice-sync.exe
│   ├── receive_payments.exe
│   ├── service_bill_cli.exe
│   ├── item_bills.exe
│   ├── misc_income_cli.exe
│   ├── pay_bills.exe
│   ├── payment_terms_dummy.exe
│
├── reports/                   # Auto-generated reports
│   └── report_YYYYMMDD_HHMMSS.txt

Integrated Team Executables (10 Total)
Team	Executable
Chart of Accounts	Chart_of_accounts.exe
Customers	customer_compare.exe
Vendors	vendor_compare.exe
Invoices	qb-invoice-sync.exe
Receive Payments	receive_payments.exe
Service Bills	service_bill_cli.exe
Item Bills	item_bills.exe
Misc Income	misc_income_cli.exe
Pay Bills	pay_bills.exe
Payment Terms	payment_terms_dummy.exe
Running the GUI
Step 1: Open QuickBooks VM

Log into the VM using Windows Remote Desktop with instructor credentials.

Step 2: Open QuickBooks

Ensure the Example Company file is open in QuickBooks Desktop.

Step 3: Run the GUI

From inside the VM:

cd TeamRunner
python gui.py

GUI Behavior

Click Start

Progress bar updates in 10% increments

Each .exe runs sequentially

JSON output is parsed and displayed

Latest results appear at the top

After completion:

A final report is written to /reports

Temporary JSON files are deleted automatically

Example message shown in GUI:

Temporary JSON files deleted: 6

JSON Handling & Schema Differences

Each team produces JSON using slightly different schemas, for example:

same_accounts

same_customers

same_items

same_payments

The GUI dynamically detects keys and:

Displays relevant fields

Handles missing keys gracefully

Logs errors without crashing

Known & Expected Errors (Outside VM)

If the GUI is run outside the QuickBooks VM, the following errors are expected:

(-2147221005, 'Invalid class string', None, None)


These errors do not indicate a bug in the GUI and occur because:

QuickBooks COM objects are unavailable

QuickBooks Desktop is not running

Final Notes for Grading

The project fully satisfies the CS 480 final project description

All required executables are integrated

GUI is compatible with QuickBooks Desktop VM

Errors outside the VM are expected and documented

Tested with the official Example Company Excel file

Author

Minahil Rao
CS 480 – Internship
Purdue University Northwest
Fall 2025
