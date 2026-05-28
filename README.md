# OPEX IMR Chart Add-in

Individuals and Moving Range (IMR) control charts for Excel, with a dedicated **OPEX** ribbon tab.

---

## Files

| File | Purpose |
|---|---|
| `OPEX.xlam` | Excel add-in — adds the OPEX ribbon tab with IMR Chart button |
| `opex_imr.py` | Python chart engine |

---

## Setup (one time, ~2 minutes)

### 1  Install Python dependencies
```
pip install xlwings
```

### 2  Place both files in the same folder
```
C:\AddIns\OPEX.xlam
C:\AddIns\opex_imr.py
```

### 3  Install OPEX.xlam
- Right-click `OPEX.xlam` -> Properties -> check **Unblock** -> OK
- Excel -> File -> Options -> Add-Ins -> Manage: Excel Add-Ins -> Go
- Browse to `C:\AddIns\OPEX.xlam` -> OK

You will now see an **OPEX** tab in the Excel ribbon.

> **Note:** The xlwings Excel add-in (`xlwings addin install`) is **not** required.
> The ribbon button launches Python directly — no intermediate add-in needed.

---

## Usage

1. Open Excel and enter your data in a single column (no header)
2. Select the data range
3. Click **OPEX** tab -> **IMR Chart**
4. Fill in the settings dialog — all fields visible at once:
   - Individuals chart title
   - Moving Range chart title
   - Y-axis label
   - X-axis label
   - Decimal places (default: 2)
5. Click **Create Charts**

Both charts appear to the right of the data on a new sheet.

---

## Chart design

| Series | Colour | Markers |
|---|---|---|
| Y / MR data | Navy blue | Circle |
| Mean / MRBar | Green | None |
| UCL / LCL / MRUCL | Red | None |

End labels show rounded values only (e.g. `52.55`).
The X axis is pinned to the bottom of the chart.
Charts are positioned to the right of the data — they never overlap it.

---

## Running without the ribbon button

Select your data in Excel first, then:
```
python opex_imr.py
```

---

## IMR constants

| Constant | Value | Use |
|---|---|---|
| D2 | 2.659 | I-chart limits: Mean +/- D2 x RBar |
| D4 | 3.267 | MRUCL = RBar x D4 |
