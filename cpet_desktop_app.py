import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import csv
import re
import os
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# --- CORE LOGIC FROM cpet_generate_analysed.py ---

SUMMARY_ROWS = 15
ROLLING_WINDOW_ROWS = 6  # 30 seconds in the 5-sec averaged file

COL_TIME = 1
COL_LOAD = 2
COL_HR = 3
COL_VE = 4
COL_VO2 = 6
COL_VO2_ROLLING = 7
COL_VCO2 = 8
COL_RER = 9
COL_VO2KG = 10
COL_VO2KG_ROLLING = 11
COL_O2PULSE = 12

class CPETAnalysisError(RuntimeError):
    pass

def clean_text(value: Any) -> str:
    return str(value).strip().lower() if value is not None else ""

def to_number(value: Any) -> Optional[float]:
    if value is None or value == "": return None
    if isinstance(value, (int, float)) and not isinstance(value, bool): return float(value)
    try: return float(str(value).strip())
    except: return None

def parse_time_to_seconds(value: Any) -> Optional[int]:
    if value is None: return None
    if isinstance(value, (time, datetime)): return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, timedelta): return int(round(value.total_seconds()))
    if isinstance(value, (int, float)):
        if 0 <= value < 1: return int(round(value * 86400))
        return None
    text = str(value).strip()
    match = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", text)
    if not match: return None
    h_or_m = int(match.group(1))
    m_or_s = int(match.group(2))
    s = int(match.group(3)) if match.group(3) else None
    return h_or_m * 60 + m_or_s if s is None else h_or_m * 3600 + m_or_s * 60 + s

def coerce_cell(value: str) -> Any:
    if value is None: return None
    text = str(value).rstrip("\n\r")
    if text == "": return None
    if re.fullmatch(r"\s*\d{1,2}:\d{2}(?::\d{2})?\s*", text): return text
    try:
        num = float(text)
        return int(num) if num.is_integer() else num
    except: return text

def load_raw_file(path: Path) -> Workbook:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}: return load_workbook(path)
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    wb = Workbook()
    ws = wb.active
    ws.title = path.stem[:31]
    reader = csv.reader(raw.splitlines(), delimiter="\t")
    for row in reader: ws.append([coerce_cell(cell) for cell in row])
    return wb

def find_row(ws, label: str) -> int:
    target = clean_text(label)
    for row in range(1, ws.max_row + 1):
        if clean_text(ws.cell(row, 1).value) == target: return row
    raise CPETAnalysisError(f"Could not find block label: {label}")

def block_rows(ws, label: str, next_label: Optional[str] = None) -> Dict[str, int]:
    label_row = find_row(ws, label)
    header_row, unit_row, data_start = label_row + 1, label_row + 2, label_row + 3
    data_end = find_row(ws, next_label) - 1 if next_label else ws.max_row
    while data_end >= data_start:
        if any(ws.cell(data_end, c).value is not None for c in range(1, 15)): break
        data_end -= 1
    return {"label": label_row, "header": header_row, "unit": unit_row, "start": data_start, "end": data_end}

def remove_irregular_5s_rows(ws, start: int, end: int) -> int:
    deleted = 0
    for row in range(end, start - 1, -1):
        secs = parse_time_to_seconds(ws.cell(row, COL_TIME).value)
        if secs is not None and secs % 5 != 0:
            ws.delete_rows(row, 1)
            deleted += 1
    return deleted

def max_load_row(ws, start: int, end: int) -> int:
    m_val, m_row = -1.0, -1
    for row in range(start, end + 1):
        v = to_number(ws.cell(row, COL_LOAD).value)
        if v is not None and v >= m_val: m_val, m_row = v, row
    if m_row == -1: raise CPETAnalysisError("No numeric load values found.")
    return m_row

def process_file(input_path: Path):
    output_path = input_path.with_name(f"{input_path.stem}_ANALYSED.xlsx")
    wb = load_raw_file(input_path)
    ws = wb.active
    ws.insert_rows(1, SUMMARY_ROWS)
    ws.insert_cols(COL_VO2_ROLLING, 1)
    ws.insert_cols(COL_VO2KG_ROLLING, 1)

    five = block_rows(ws, "5s AVERAGED DATA", "BP DATA")
    removed = remove_irregular_5s_rows(ws, five["start"], five["end"])
    five = block_rows(ws, "5s AVERAGED DATA", "BP DATA")
    bp = block_rows(ws, "BP DATA", "BREATH x BREATH DATA")
    bxb = block_rows(ws, "BREATH x BREATH DATA", None)

    for row in range(five["start"], five["end"] + 1):
        s = max(five["start"], row - 5)
        ws.cell(row, COL_VO2_ROLLING).value = f"=AVERAGE(F{s}:F{row})"
        ws.cell(row, COL_VO2KG_ROLLING).value = f"=AVERAGE(J{s}:J{row})"
        ws.cell(row, COL_VO2_ROLLING).number_format = "0.0"
        ws.cell(row, COL_VO2KG_ROLLING).number_format = "0.0"

    p5, pbp, pbx = max_load_row(ws, five["start"], five["end"]), max_load_row(ws, bp["start"], bp["end"]), max_load_row(ws, bxb["start"], bxb["end"])
    
    # Simple pre-exercise HR detection (average of 0 load rows at start)
    pre_hr_end = five["start"]
    for r in range(five["start"], five["end"]+1):
        if to_number(ws.cell(r, COL_LOAD).value) == 0: pre_hr_end = r
        else: break
    
    pre_bp_row = bp["start"]
    for r in range(bp["start"], bp["end"]+1):
        if to_number(ws.cell(r, COL_LOAD).value) == 0: pre_bp_row = r
        else: break

    l30 = max(five["start"], p5 - 5)
    summary = [
        (1, "Pre-exercise", None), (2, "HR", f"=AVERAGE(C{five['start']}:C{pre_hr_end})"),
        (3, "Sys", f"=E{pre_bp_row}"), (4, "Dia", f"=F{pre_bp_row}"),
        (5, "Peak Values", None), (6, "HR", f"=MAX(C{bxb['start']}:C{pbx})"),
        (7, "Load (W)", f"=MAX(B{bxb['start']}:B{bxb['end']})"),
        (8, "VO2 (mL/min)", f"=MAX(G{five['start']}:G{p5})"),
        (9, "VO2 (mL/min)/kg", f"=MAX(K{five['start']}:K{p5})"),
        (10, "RER", f"=AVERAGE(I{l30}:I{p5})"), (11, "V'E", f"=AVERAGE(D{l30}:D{p5})"),
        (12, "O2 pulse", f"=AVERAGE(L{l30}:L{p5})"),
        (13, "Sys", f"=E{pbp}"), (14, "Dia", f"=F{pbp}")
    ]
    for r, lab, form in summary:
        ws.cell(r, 1).value = lab
        if form: ws.cell(r, 2).value = form
    
    # Styling
    bold, fill = Font(bold=True), PatternFill("solid", fgColor="D9EAF7")
    for r in [1, 5]: ws.cell(r, 1).font, ws.cell(r, 1).fill = bold, fill
    for r in range(1, 15): ws.cell(r, 2).number_format = "0.0"
    ws.cell(10, 2).number_format = "0.00"
    
    wb.save(output_path)
    return output_path, removed

# --- GUI ---

class App:
    def __init__(self, root):
        self.root = root
        root.title("CPET Analyser")
        root.geometry("400x250")
        
        frame = ttk.Frame(root, padding="20")
        frame.pack(expand=True, fill="both")
        
        ttk.Label(frame, text="🫁 CPET Analysis Tool", font=("Helvetica", 16, "bold")).pack(pady=10)
        ttk.Label(frame, text="Select a raw CPET file (.xlsx or .xls)").pack()
        
        self.btn = ttk.Button(frame, text="Browse and Process", command=self.run)
        self.btn.pack(pady=20)
        
        self.status = ttk.Label(frame, text="Ready", foreground="gray")
        self.status.pack()

    def run(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if not file_path: return
        
        try:
            self.status.config(text="Processing...", foreground="blue")
            self.root.update()
            out, rem = process_file(Path(file_path))
            self.status.config(text="Done!", foreground="green")
            messagebox.showinfo("Success", f"File saved:\n{out.name}\n\nRemoved {rem} irregular rows.")
        except Exception as e:
            self.status.config(text="Error", foreground="red")
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
