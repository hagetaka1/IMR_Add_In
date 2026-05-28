"""
OPEX IMR Chart  -  xlwings version
Calculates Individuals and Moving Range control charts from selected Excel data.

Usage:
    python opex_imr.py                  # Standalone - prompts for open workbook
    Called via xlwings RunPython button  # Add-in mode

Requirements:
    pip install xlwings
    xlwings addin install               # Only needed for ribbon button mode
"""

import math
import tkinter as tk
from tkinter import ttk, messagebox
import xlwings as xw


# ── Control chart constants (n=2 subgroup) ────────────────────────────────────
D4 = 3.267   # MRUCL = RBar * D4
D2 = 2.659   # I-chart limits = Mean +/- D2 * RBar

# ── Colours ───────────────────────────────────────────────────────────────────
BLUE_DARK  = (31,  73,  125)   # data line
GREEN      = (0,   153, 0)     # mean / MRBar
RED        = (192, 0,   0)     # control limits
WHITE      = (255, 255, 255)
NAVY       = (31,  73,  125)
GREY_TEXT  = (89,  89,  89)

# ── Chart dimensions (points) ─────────────────────────────────────────────────
CHART_W = 560
CHART_H = 300
CHART_GAP = 20


# ══════════════════════════════════════════════════════════════════════════════
#  Settings dialog
# ══════════════════════════════════════════════════════════════════════════════

class SettingsDialog:
    """Single-window dialog – all fields visible at once."""

    def __init__(self):
        self.result = None

        root = tk.Tk()
        root.title("OPEX IMR Chart – Settings")
        root.resizable(False, False)
        root.configure(bg="#f0f0f0")

        # ── Header ────────────────────────────────────────────────────────────
        header = tk.Frame(root, bg="#0070C0", pady=10)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        tk.Label(header, text="OPEX IMR Chart Settings",
                 font=("Calibri", 14, "bold"), fg="white", bg="#0070C0"
                 ).pack()

        # ── Form fields ───────────────────────────────────────────────────────
        fields = [
            ("Individuals Chart Title",  "Individuals"),
            ("Moving Range Chart Title", "Moving Range"),
            ("Y-Axis Label",             "Value"),
            ("X-Axis Label",             "Observation"),
            ("Decimal Places (0-6)",     "2"),
        ]

        self.vars = []
        form = tk.Frame(root, bg="#f0f0f0", padx=20, pady=12)
        form.grid(row=1, column=0, columnspan=2, sticky="ew")

        for i, (label, default) in enumerate(fields):
            tk.Label(form, text=label + ":", font=("Calibri", 10),
                     bg="#f0f0f0", anchor="e", width=26
                     ).grid(row=i, column=0, padx=(0, 8), pady=5, sticky="e")
            var = tk.StringVar(value=default)
            entry = ttk.Entry(form, textvariable=var, width=28,
                              font=("Calibri", 10))
            entry.grid(row=i, column=1, pady=5, sticky="w")
            self.vars.append(var)
            if i == 0:
                entry.focus_set()

        # ── Buttons ───────────────────────────────────────────────────────────
        btn_frame = tk.Frame(root, bg="#f0f0f0", pady=10)
        btn_frame.grid(row=2, column=0, columnspan=2)

        ttk.Button(btn_frame, text="Create Charts",
                   command=self._ok).pack(side="left", padx=8, ipadx=12, ipady=4)
        ttk.Button(btn_frame, text="Cancel",
                   command=root.destroy).pack(side="left", padx=8, ipadx=12, ipady=4)

        root.bind("<Return>", lambda e: self._ok())
        root.bind("<Escape>", lambda e: root.destroy())

        # Centre on screen
        root.update_idletasks()
        w, h = root.winfo_width(), root.winfo_height()
        x = (root.winfo_screenwidth()  - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry(f"+{x}+{y}")

        self._root = root
        root.mainloop()

    def _ok(self):
        try:
            decimals = int(self.vars[4].get())
            if not 0 <= decimals <= 6:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid input",
                                 "Decimal places must be a whole number 0-6.")
            return

        self.result = {
            "indiv_title": self.vars[0].get().strip() or "Individuals",
            "mr_title":    self.vars[1].get().strip() or "Moving Range",
            "y_axis":      self.vars[2].get().strip(),
            "x_axis":      self.vars[3].get().strip(),
            "decimals":    decimals,
        }
        self._root.destroy()


# ══════════════════════════════════════════════════════════════════════════════
#  IMR statistics
# ══════════════════════════════════════════════════════════════════════════════

def calc_imr(values: list) -> dict:
    n    = len(values)
    mean = sum(values) / n
    mrs  = [abs(values[i] - values[i-1]) for i in range(1, n)]
    rbar = sum(mrs) / len(mrs)
    return {
        "mean":   mean,
        "rbar":   rbar,
        "lcl":    mean - D2 * rbar,
        "ucl":    mean + D2 * rbar,
        "mrucl":  rbar * D4,
        "mrs":    mrs,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  Excel helpers
# ══════════════════════════════════════════════════════════════════════════════

def rgb(r, g, b) -> int:
    """Convert RGB tuple to Excel BGR integer."""
    return r + g * 256 + b * 65536


def unique_sheet_name(wb, prefix: str) -> str:
    names = {s.name for s in wb.sheets}
    i = 1
    while f"{prefix}{i}" in names:
        i += 1
    return f"{prefix}{i}"


def build_worksheet(wb, values: list, stats: dict) -> xw.Sheet:
    """Write data and computed columns to a new worksheet."""
    name = unique_sheet_name(wb, "IMR")
    ws   = wb.sheets.add(name)

    headers = ["Y", "Mean", "LCL", "UCL", "MR", "MRBar", "MRUCL"]
    ws.range("A1").value = headers
    ws.range("A1:G1").api.Font.Bold = True

    n = len(values)
    for i, v in enumerate(values):
        row = i + 2
        ws.cells(row, 1).value = v
        ws.cells(row, 2).value = stats["mean"]
        ws.cells(row, 3).value = stats["lcl"]
        ws.cells(row, 4).value = stats["ucl"]
        ws.cells(row, 5).value = stats["mrs"][i-1] if i > 0 else None
        ws.cells(row, 6).value = stats["rbar"]
        ws.cells(row, 7).value = stats["mrucl"]

    # Auto-fit data columns
    ws.range("A:G").columns.autofit()
    return ws


# ══════════════════════════════════════════════════════════════════════════════
#  Chart helpers
# ══════════════════════════════════════════════════════════════════════════════

def _fmt(value: float, decimals: int) -> str:
    return f"{value:.{decimals}f}"


def apply_presentation_style(cht_api, y_label: str, x_label: str):
    """Apply clean presentation style via the Excel COM API."""

    # Chart area
    cht_api.ChartArea.Border.LineStyle = -4142          # xlNone
    cht_api.ChartArea.Interior.Color   = rgb(*WHITE)
    cht_api.ChartArea.Format.TextFrame2.TextRange.Font.Name = "Calibri"

    # Plot area
    cht_api.PlotArea.Border.LineStyle  = -4142
    cht_api.PlotArea.Interior.Color    = rgb(*WHITE)

    # Title
    cht_api.ChartTitle.Font.Name  = "Calibri"
    cht_api.ChartTitle.Font.Size  = 13
    cht_api.ChartTitle.Font.Bold  = True
    cht_api.ChartTitle.Font.Color = rgb(*NAVY)

    # Value axis (Y) – pin to bottom
    val_axis = cht_api.Axes(2)          # xlValue = 2
    val_axis.HasMajorGridlines = False
    val_axis.Crosses           = 4      # xlMinimum = 4
    val_axis.TickLabels.Font.Name = "Calibri"
    val_axis.TickLabels.Font.Size = 9
    if y_label:
        val_axis.HasTitle = True
        val_axis.AxisTitle.Text       = y_label
        val_axis.AxisTitle.Font.Name  = "Calibri"
        val_axis.AxisTitle.Font.Size  = 10
        val_axis.AxisTitle.Font.Color = rgb(*GREY_TEXT)

    # Category axis (X)
    cat_axis = cht_api.Axes(1)          # xlCategory = 1
    cat_axis.HasMajorGridlines = False
    cat_axis.TickLabels.Font.Name = "Calibri"
    cat_axis.TickLabels.Font.Size = 9
    if x_label:
        cat_axis.HasTitle = True
        cat_axis.AxisTitle.Text       = x_label
        cat_axis.AxisTitle.Font.Name  = "Calibri"
        cat_axis.AxisTitle.Font.Size  = 10
        cat_axis.AxisTitle.Font.Color = rgb(*GREY_TEXT)


def style_series(ser_api, color: tuple, weight: float,
                 marker_style: int, marker_color: tuple):
    """Style a chart series via COM."""
    ser_api.Format.Line.ForeColor.RGB = rgb(*color)
    ser_api.Format.Line.Weight        = weight
    ser_api.MarkerStyle = marker_style       # -4142=none, 8=circle
    if marker_style != -4142:
        ser_api.MarkerSize            = 5
        ser_api.MarkerForegroundColor = rgb(*marker_color)
        ser_api.MarkerBackgroundColor = rgb(*WHITE)


def add_end_label(ser_api, value: float, decimals: int, color: tuple):
    """Add a value-only label to the last point of a series."""
    last_pt = ser_api.Points(ser_api.Points().Count)
    last_pt.ApplyDataLabels(Type=2)              # xlDataLabelsShowValue = 2
    dl = last_pt.DataLabel
    dl.Text           = _fmt(value, decimals)
    dl.Position       = -4152                    # xlLabelPositionRight = -4152
    dl.Font.Name      = "Calibri"
    dl.Font.Size      = 9
    dl.Font.Bold      = True
    dl.Font.Color     = rgb(*color)
    dl.Format.Line.Visible = 1                   # msoFalse = 1... actually use 0
    try:
        dl.Format.Line.Visible = False
        dl.Format.Fill.Visible = False
    except Exception:
        pass


def build_chart(ws: xw.Sheet, src_range: xw.Range,
                title: str, left: float, top: float,
                series_styles: list, end_labels: list,
                y_label: str, x_label: str) -> None:
    """
    Generic chart builder.
    series_styles: list of (color, weight, marker_style, marker_color)
    end_labels:    list of (series_index_1based, value, decimals, color) or None
    """
    co  = ws.charts.add(left=left, top=top, width=CHART_W, height=CHART_H)
    api = co.api[1]      # xlwings Chart COM object

    api.ChartType = 65   # xlLine = 65
    api.SetSourceData(Source=src_range.api, PlotBy=2)  # xlColumns = 2
    api.HasTitle       = True
    api.ChartTitle.Text = title
    api.HasLegend      = False

    apply_presentation_style(api, y_label, x_label)

    for i, (color, weight, marker, mcolor) in enumerate(series_styles, start=1):
        style_series(api.SeriesCollection(i), color, weight, marker, mcolor)

    for label in end_labels:
        if label:
            idx, value, decimals, color = label
            add_end_label(api.SeriesCollection(idx), value, decimals, color)


# ══════════════════════════════════════════════════════════════════════════════
#  Main entry points
# ══════════════════════════════════════════════════════════════════════════════

def create_imr_chart():
    """
    Entry point called from:
      - OPEX ribbon button (via xlwings RunPython)
      - Command line: python opex_imr.py
    """
    try:
        # RunPython mode: xlwings injects the calling workbook
        wb = xw.Book.caller()
    except Exception:
        # Standalone mode: attach to the active Excel instance
        try:
            app = xw.apps.active
            if app is None:
                raise RuntimeError("No Excel instance found.")
            wb = app.books.active
        except Exception:
            messagebox.showerror("OPEX IMR",
                "No Excel workbook found.\n\n"
                "Please open Excel with your data selected, then run this script.")
            return

    # ── Get data selection ────────────────────────────────────────────────────
    sel = wb.app.selection
    if sel is None:
        messagebox.showerror("OPEX IMR", "Please select a range of data in Excel first.")
        return

    raw = sel.value
    if raw is None:
        messagebox.showerror("OPEX IMR", "Selected range is empty.")
        return

    # Flatten rows or columns into a 1-D list of numbers
    if isinstance(raw, (int, float)):
        raw = [raw]
    elif isinstance(raw[0], list):
        # 2-D: flatten column or row
        if len(raw[0]) == 1:
            raw = [r[0] for r in raw]      # single column
        else:
            raw = raw[0]                   # single row

    try:
        values = [float(v) for v in raw if v is not None]
    except (TypeError, ValueError):
        messagebox.showerror("OPEX IMR",
            "Selection contains non-numeric values. "
            "Please select a single row or column of numbers without headers.")
        return

    if len(values) < 3:
        messagebox.showerror("OPEX IMR", "At least 3 data points are required.")
        return

    # ── Settings dialog ───────────────────────────────────────────────────────
    dlg = SettingsDialog()
    if dlg.result is None:
        return      # user cancelled

    s        = dlg.result
    decimals = s["decimals"]
    stats    = calc_imr(values)
    n        = len(values)

    # ── Build worksheet ───────────────────────────────────────────────────────
    ws = build_worksheet(wb, values, stats)

    # Chart left edge = just right of column G (col 7) with a small gap
    chart_left = ws.range("I1").left
    chart_top  = ws.range("A1").top

    # ── Individuals chart  (columns A:D = Y, Mean, LCL, UCL) ─────────────────
    indiv_src = ws.range(f"A1:D{n+1}")
    build_chart(
        ws       = ws,
        src_range = indiv_src,
        title    = s["indiv_title"],
        left     = chart_left,
        top      = chart_top,
        series_styles = [
            (BLUE_DARK, 2.0,  8,    BLUE_DARK),   # Y     – circle markers
            (GREEN,     1.75, -4142, GREEN),       # Mean  – no markers
            (RED,       1.75, -4142, RED),         # LCL   – no markers
            (RED,       1.75, -4142, RED),         # UCL   – no markers
        ],
        end_labels = [
            None,
            (2, stats["mean"], decimals, GREEN),
            (3, stats["lcl"],  decimals, RED),
            (4, stats["ucl"],  decimals, RED),
        ],
        y_label = s["y_axis"],
        x_label = s["x_axis"],
    )

    # ── Moving Range chart  (columns E:G = MR, MRBar, MRUCL, rows 2 onwards) ─
    mr_src = ws.range(f"E1:G{n+1}")   # include header row so Excel sees 3 named series
    try:
        build_chart(
            ws        = ws,
            src_range = mr_src,
            title     = s["mr_title"],
            left      = chart_left + CHART_W + CHART_GAP,
            top       = chart_top,
            series_styles = [
                (BLUE_DARK, 2.0,  8,    BLUE_DARK),   # MR    – circle markers
                (GREEN,     1.75, -4142, GREEN),       # MRBar – no markers
                (RED,       1.75, -4142, RED),         # MRUCL – no markers
            ],
            end_labels = [
                None,
                (2, stats["rbar"],  decimals, GREEN),
                (3, stats["mrucl"], decimals, RED),
            ],
            y_label = s["y_axis"],
            x_label = s["x_axis"],
        )
    except Exception as e:
        import traceback
        messagebox.showerror("OPEX IMR - MR Chart Error",
            "The Moving Range chart failed:\n\n" + traceback.format_exc())

    ws.activate()
    ws.range("A1").select()
    wb.app.screen_updating = True


# ── Standalone launcher ───────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        create_imr_chart()
    except Exception as e:
        import traceback
        messagebox.showerror("OPEX IMR Error", traceback.format_exc())
