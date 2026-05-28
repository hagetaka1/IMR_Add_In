# OPEX IMR Chart  -  Python / xlwings version

## Quick start (2 minutes)

### 1  Install dependencies
```
pip install xlwings
```

### 2  Copy opex_imr.py somewhere permanent
```
C:\AddIns\opex_imr.py
```

### 3  Run it
Double-click `opex_imr.py`, or from a terminal:
```
python opex_imr.py
```

---

## How to use

1. Open Excel and enter your data in a single column (or row)
2. Select the data range (no headers)
3. Run `python opex_imr.py`
4. Fill in the settings dialog – all fields are visible at once
5. Click **Create Charts**

Two charts appear to the right of the data on a new sheet:
- **Individuals chart** – Y values with Mean (green), LCL and UCL (red)
- **Moving Range chart** – MR values with MRBar (green) and MRUCL (red)

Labels on the right of each control line show the rounded value only (e.g. `52.55`).

---

## Optional: Add a ribbon button

If you want a ribbon button instead of running from the terminal:

### 1  Install the xlwings Excel add-in
```
xlwings addin install
```

### 2  Create a companion workbook
- Open Excel, create a new workbook, save as `OPEX.xlsm`
- Open the VBA editor (Alt+F11)
- Add this one sub to Module1:

```vb
Sub IMR_Chart()
    RunPython "import sys; sys.path.insert(0, r'C:\AddIns'); import opex_imr; opex_imr.create_imr_chart()"
End Sub
```
*(Change `C:\AddIns` to wherever you saved `opex_imr.py`)*

### 3  Add a Quick Access Toolbar button
- File → Options → Quick Access Toolbar
- Choose commands from: Macros
- Select `IMR_Chart` → Add → OK

Or assign it to a keyboard shortcut:
- Developer tab → Macros → IMR_Chart → Options → set a shortcut key

---

## Why Python instead of VBA?

| | VBA (.xla) | Python (xlwings) |
|---|---|---|
| Debug errors | No debug button in add-in mode | Full stack traces in terminal |
| Chart API | Deprecated methods in Excel 365 | Direct COM access, always current |
| Dialog | Multi-step InputBox or worksheet hack | Native tkinter window, all fields at once |
| Code editing | VBA editor | Any editor (VS Code, etc.) |
| Version control | Binary .xla file | Plain text .py file |
| Testing | Manual only | Unit testable |

---

## IMR constants used

| Constant | Value | Purpose |
|---|---|---|
| D2 | 2.659 | I-chart limit width: Mean ± D2 × RBar |
| D4 | 3.267 | MRUCL = RBar × D4 |
