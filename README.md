# OPEX IMR Chart Add-in

Individuals and Moving Range (IMR) control charts for Excel, with a dedicated **OPEX** ribbon tab.

---

## Files

| File | Purpose |
|---|---|
| `OPEX.xlam` | Excel add-in with OPEX ribbon tab and IMR Chart button |
| `opex_imr.py` | Python chart engine (xlwings) |

---

## Setup (one time)

### 1  Install Python dependencies
```
pip install xlwings
```

### 2  Install the xlwings Excel add-in
```
xlwings addin install
```
This gives Excel the `RunPython` function that the ribbon button uses.

### 3  Place opex_imr.py in a permanent folder
```
C:\AddIns\opex_imr.py
```
(Any folder works — just remember the path for step 5.)

### 4  Install OPEX.xlam
- Right-click `OPEX.xlam` -> Properties -> check **Unblock** -> OK
- Excel -> File -> Options -> Add-Ins -> Manage: Excel Add-Ins -> Go
- Browse to `OPEX.xlam` -> OK

You will now see an **OPEX** tab in the Excel ribbon.

### 5  Tell xlwings where opex_imr.py lives
- Click the **xlwings** ribbon tab
- Click **Settings** (or **Edit Config**)
- Set `PYTHONPATH` to the folder containing `opex_imr.py`, e.g.:
  ```
  PYTHONPATH = C:\AddIns
  ```
- Save and close

---

## Usage

1. Open Excel and enter your data in a single column (no header)
2. Select the data range
3. Click **OPEX** tab -> **IMR Chart**
4. Fill in the settings dialog (all fields visible at once)
5. Click **Create Charts**

Both charts appear to the right of the data on a new sheet.

---

## Chart design

| Series | Colour | Markers |
|---|---|---|
| Y / MR data | Navy blue | Circle |
| Mean / MRBar | Green | None |
| UCL / LCL / MRUCL | Red | None |

End labels show rounded values only (e.g. `52.55`). Label count and precision are set in the dialog.

---

## IMR constants

| Constant | Value | Use |
|---|---|---|
| D2 | 2.659 | I-chart limits: Mean ± D2 × RBar |
| D4 | 3.267 | MRUCL = RBar × D4 |

---

## Running without the ribbon button

```
python opex_imr.py
```

Select your data in Excel first, then run the script.
