import os
import sys
import glob
import shutil
import logging
import traceback
import pandas as pd
from pyhwpx import Hwp
from datetime import datetime
import constants as c
from tkinter import messagebox, Tk

#################################################################### debug
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
stream_handler = logging.StreamHandler()
stream_handler.setFormatter(logging.Formatter('%(message)s'))
logger.addHandler(stream_handler)

def save_error_log():
  file_handler = logging.FileHandler("debug_log.txt", encoding='utf-8')
  formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
  file_handler.setFormatter(formatter)
  logger.addHandler(file_handler)
  
  logger.error("Detailed traceback information:")
  logger.error(traceback.format_exc())
  
  file_handler.flush()
  logger.removeHandler(file_handler)


#################################################################### utils
def show_alert(title, message):
  root = Tk()
  root.withdraw()
  messagebox.showinfo(title, message)
  root.destroy()

def get_formatted_date():
  now = datetime.now()
  date_text = f"{now.strftime('%y')}. {now.month}."
  print(f"[WITTEN] {date_text}")
  return date_text

def get_paths():
  if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(os.path.abspath(sys.executable))
  else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

  excel_files = glob.glob(os.path.join(current_dir, "*.xlsx"))
  excel_files = [f for f in excel_files if not os.path.basename(f).startswith("~$")]

  if not excel_files:
    show_alert("Error", f"No Excel (.xlsx) file was found in the specified folder.")
  selected_excel = excel_files[0]

  template_path = os.path.join(current_dir, c.TEMPLATE_FILE)
  if not os.path.exists(template_path):
    show_alert("Error", f"No template (.hwpx) file was found in the specified folder.")
  
  return {
    "excel": selected_excel,
    "template": os.path.join(current_dir, c.TEMPLATE_FILE),
    "output": os.path.join(current_dir, f"{datetime.now().strftime('%Y-%m-%d-%H%M%S')}.hwp")
  }

def to_ratio(val):
  if pd.isna(val) or val == 0: return "-"
  return f"{'△' if val < 0 else ''}{abs(val):.1f}"

def to_thousands(val):
  try:
    num = float(str(val).replace(',', '').strip())
    if num == 0: return "-"
    formatted = f"{int(abs(num) / 1000):,}"
    return f"△{formatted}" if num < 0 else formatted
  except (ValueError, TypeError):
    return "-"

def growth_rate(target, now):
  if not now or not target: return 0.0
  return round((target - now) / now * 100, 1)

def update_progress(writer, field_name=''): 
  bar_length=20
  percent = (writer.current_count / writer.total_tasks)
  filled_length = int(bar_length * percent)
  bar = "■" * filled_length + "□" * (bar_length - filled_length)
  clean_name = str(field_name).strip()
  display_name = (clean_name[:22] + '..') if len(clean_name) > 24 else clean_name
  print(f"\r{bar} {int(percent * 100):>3}% ({writer.current_count}/{writer.total_tasks}) {display_name:<30}", end="", flush=True)

#################################################################### hwp
class HwpWriter:
  def __init__(self, output_path):
    try:
      print(f"[INIT] Opening HWP file: {os.path.basename(output_path)}")
      self.hwp = Hwp()
      self.hwp.open(output_path)
      self.budget_storage = {}
      self.total_tasks = 0
      self.current_count = 0
    except Exception as e:
      print(f"[ERROR] Opening HWP file: {e}")
      save_error_log()
      raise

  def apply_cell_style(self, req_txt, fix_txt, req_num, fix_num, is_bold=False):
    # initialize
    self.hwp.SelectAll() 
    self.hwp.Delete()
    self.hwp.set_font(StrikeOutType=False, Bold=False, TextColor="Black")

    # single-line
    if fix_txt is None:
      self.hwp.insert_text(str(req_txt or "0"))
      self.hwp.MoveLineBegin(); self.hwp.MoveSelLineEnd()
      req_color = "Red" if (req_num and req_num < 0) else "Black"
      self.hwp.set_font(TextColor=req_color, Bold=is_bold)
    
    # multi-line
    else:
      self.hwp.insert_text(f"{req_txt or '-'}\r\n{fix_txt or '-'}")
      self.hwp.MoveLineBegin(); self.hwp.MoveUp(); self.hwp.MoveLineBegin(); self.hwp.MoveSelLineEnd()
      self.hwp.set_font(StrikeOutType=True, StrikeOutShape=0, Bold=is_bold) 
      self.hwp.MoveDown(); self.hwp.MoveLineBegin(); self.hwp.MoveSelLineEnd()
      self.hwp.set_font(TextColor="Red" if (fix_num and fix_num < 0) else "Blue", Bold=is_bold)
    self.hwp.Cancel()

  def write_budget_cell(self, field, data, is_bold=False, is_single_line=False):
    if not self.hwp.field_exist(field): 
      print(f"\n[WARN] Field not found: {field}")
      return

    try:
      self.budget_storage[field] = data['now']
      self.hwp.move_to_field(field)
      steps = [
        (data['now'], None),
        (data['req'], None if is_single_line else data['fix']),
        (data['req']-data['now'], None if is_single_line else data['fix']-data['now']),
        (growth_rate(data['req'], data['now']), None if is_single_line else growth_rate(data['fix'], data['now']))
      ]

      for i, (req_val, fix_val) in enumerate(steps):
        is_ratio = (i == 3)
        req_txt = to_ratio(req_val) if is_ratio else to_thousands(req_val)
        fix_txt = (to_ratio(fix_val) if is_ratio else to_thousands(fix_val)) if fix_val is not None else None
        self.apply_cell_style(req_txt, fix_txt, req_val, fix_val or 0, is_bold)
        if i < len(steps) - 1: self.hwp.TableRightCell()
    except Exception as e:
      print(f"\n[ERROR] Writing failed at [{field}]: {e}")
      save_error_log()
  

#################################################################### excel
def load_excel(file_path):
  df = pd.read_excel(file_path, header=None)
  df_detail = pd.read_excel(file_path, header=11)

  target_cols = [df_detail.columns[c.COL_NOW], df_detail.columns[c.COL_REQ], df_detail.columns[c.COL_FIX]]
  for col in target_cols:
    df_detail[col] = pd.to_numeric(df_detail[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

  has_code = df_detail.iloc[:, c.COL_PROG].astype(str).str.contains(r'^\[\d+\]') | \
               df_detail.iloc[:, c.COL_UNIT].astype(str).str.contains(r'^\[\d+\]') | \
               df_detail.iloc[:, c.COL_SUB].astype(str).str.contains(r'^\[\d+\]')
    
  is_header_row = df_detail.iloc[:, c.COL_DETAIL].isna() & has_code
  df_detail.iloc[:, [c.COL_PROG, c.COL_UNIT, c.COL_SUB]] = df_detail.iloc[:, [c.COL_PROG, c.COL_UNIT, c.COL_SUB]].ffill()
  df_sums = df_detail[is_header_row].copy()
  
  return df, df_detail, df_sums

#################################################################### boilerplate
def match_field(writer, df, mapping, col_indices, is_bold=False, is_single_line=False):
  
  for idx, (keys, field) in enumerate(mapping.items(), 1):
    writer.current_count += 1
    update_progress(writer, field)
  
    key_list = list(keys) if isinstance(keys, tuple) else [keys]
    cond = True
    for i, key in enumerate(key_list):
      if key is not None:
        cond &= (df.iloc[:, col_indices[i]] == key)
    
    rows = df[cond]
    if not rows.empty:
      now = rows.iloc[:, c.COL_NOW].sum()
      req = rows.iloc[:, c.COL_REQ].sum()
      fix = rows.iloc[:, c.COL_FIX].sum()
      metrics = {"now": now, "req": req, "fix": fix}
      writer.write_budget_cell(
        field, 
        metrics, 
        is_bold=is_bold, 
        is_single_line=is_single_line or (req == fix)
      )
    else:
      print(f"\n[MISS] No match found for: {keys} -> Field: {field}")
  

#################################################################### core
def fill_summary(writer, df, df_sums):
  # nature
  nature_vals = df.iloc[c.NATURE_ROWS_NR, c.NATURE_COL_NOW].tolist() + \
                df.iloc[c.NATURE_ROWS_NR, c.NATURE_COL_REQ].tolist() + \
                df.iloc[c.NATURE_ROWS_F, c.NATURE_COL_FIX].tolist()
  for field, val in zip(c.SUM_BY_NATURE, nature_vals):
    writer.current_count += 1
    if writer.hwp.field_exist(field):
      update_progress(writer, field)
      writer.hwp.move_to_field(field)
      writer.apply_cell_style(to_thousands(val), None, val, None, False)

  # organization
  # 1~5
  for field, target_fields in c.SUM_BY_ORGANIZATION_1.items():
    if not writer.hwp.field_exist(field): continue
    writer.current_count += 1
    update_progress(writer, field)
    total_fix = sum(writer.budget_storage.get(f, 0) for f in target_fields)
    writer.hwp.move_to_field(field)
    writer.apply_cell_style(to_thousands(total_fix), None, total_fix, None, is_bold=False)

  # 6~15
  org_vals = df.iloc[c.ORG_ROW, c.ORG_COLS_REQ].tolist() + df.iloc[c.ORG_ROW, c.ORG_COLS_FIX].tolist()
  for field, val in zip(c.SUM_BY_ORGANIZATION_2, org_vals):
    if writer.hwp.field_exist(field):
      writer.current_count += 1
      update_progress(writer, field)
      writer.hwp.move_to_field(field)
      writer.apply_cell_style(to_thousands(val), None, val, None, False)
  
  # project
  cols = [c.COL_NOW, c.COL_REQ, c.COL_FIX]
  for col_idx in cols:
    df_sums.iloc[:, col_idx] = pd.to_numeric(df_sums.iloc[:, col_idx], errors='coerce').fillna(0)
  
  actual_progs = df_sums.iloc[:, c.COL_PROG].astype(str).str.strip()
  total_metrics = {"now": 0, "req": 0, "fix": 0}

  for field, target_prog_names in c.SUM_PROJECT.items():
    now, req, fix = 0, 0, 0
    for prog_name in target_prog_names:
      target = str(prog_name).strip()
      prog_rows = df_sums[actual_progs == target]
      
      if not prog_rows.empty:
        now += prog_rows.iloc[:, c.COL_NOW].sum()
        req += prog_rows.iloc[:, c.COL_REQ].sum()
        fix += prog_rows.iloc[:, c.COL_FIX].sum()

    metrics = {"now": now, "req": req, "fix": fix}
    writer.current_count += 1
    update_progress(writer, field)
    writer.write_budget_cell(field, metrics, is_bold=False, is_single_line=True)

    total_metrics["now"] += now
    total_metrics["req"] += req
    total_metrics["fix"] += fix
  if c.SUM_PROJECT_TOTAL:
    writer.current_count += 1
    update_progress(writer, c.SUM_PROJECT_TOTAL)
    writer.write_budget_cell(c.SUM_PROJECT_TOTAL, total_metrics, is_bold=True, is_single_line=True)


def generate_hwp():
  try:
    print("[START] Beginning HWP generation process")
    paths = get_paths()

    print(f"[COPY] Creating output file: {os.path.basename(paths['output'])}")
    shutil.copy(paths['template'], paths['output'])

    # road
    print("[LOAD] Reading Excel data...")
    df, df_detail, df_sums = load_excel(paths['excel'])
    writer = HwpWriter(paths['output'])

    # written
    date_str = get_formatted_date()
    if writer.hwp.field_exist(c.WRITTEN): writer.hwp.put_field_text(c.WRITTEN, date_str)

    # percentage
    all_mappings = [
      c.ASSOCIATED_SUB_PROJECTS, c.UNITS, c.PROGRAMS,
      c.ASSOCIATED_DETAILS, c.ASSOCIATED_SUB_UNITS, 
      c.SUB_PROJECTS, c.TOTAL_UNITS
    ]
    summary_tasks_count = (
        len(c.SUM_BY_NATURE) + 
        len(c.SUM_BY_ORGANIZATION_1) + 
        len(c.SUM_BY_ORGANIZATION_2) + 
        len(c.SUM_PROJECT) + 
        (1 if c.SUM_PROJECT_TOTAL else 0)
    )
    writer.total_tasks = sum(len(m) for m in all_mappings) + summary_tasks_count

    print(f"[RUN] Matching fields and filling data...")
    match_field(writer, df_sums, c.ASSOCIATED_SUB_PROJECTS, [c.COL_PROG, c.COL_UNIT, c.COL_SUB], is_bold=True, is_single_line=True)
    match_field(writer, df_sums, c.UNITS, [c.COL_PROG, c.COL_UNIT], is_bold=True)
    match_field(writer, df_sums, c.PROGRAMS, [c.COL_PROG], is_bold=True)
    match_field(writer, df_detail, c.ASSOCIATED_DETAILS, [c.COL_PROG, c.COL_UNIT, c.COL_SUB, c.COL_DETAIL], is_single_line=True)
    match_field(writer, df_sums, c.ASSOCIATED_SUB_UNITS, [c.COL_PROG, c.COL_UNIT, c.COL_SUB], is_single_line=True)
    match_field(writer, df_sums, c.SUB_PROJECTS, [c.COL_PROG, c.COL_UNIT, c.COL_SUB])
    match_field(writer, df_sums, c.TOTAL_UNITS, [c.COL_PROG, c.COL_UNIT])
    fill_summary(writer, df, df_sums)
    
    # save
    writer.hwp.save()

  except Exception as e:
    print(f"\n[Error] Process halted due to error: {e}")
    save_error_log()
    show_alert("Error", f"Process halted due to error")

  finally:
    print("\n[EXIT] Program terminated.")

if __name__ == "__main__":
    generate_hwp()