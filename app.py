import streamlit as st
import pdfplumber
import re
from docxtpl import DocxTemplate
from datetime import datetime, timezone, timedelta
import io
import os
import sys
import pandas as pd
import openpyxl
import openpyxl.utils  
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.cell.cell import MergedCell
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage, ImageChops
import fitz  # PyMuPDF
import numpy as np
import gc
import math
import zipfile

# ==============================================================================
# [공통 유틸리티]
# ==============================================================================
st.set_page_config(page_title="통합 양식 변환기", layout="wide")

def get_resource_path(relative_path):
    """실행 파일(exe) 내부나 일반 환경에서 절대 경로를 찾습니다."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 세션 상태 초기화 (결과물 유지용)
SESSION_KEYS = [
    'spec_res', 'spec_fname',
    'allergy_res_83', 'allergy_res_26', 'allergy_fname_83', 'allergy_fname_26',
    'ifra_res', 'ifra_fname',
    'msds_res', # 리스트 형태: [{'fname': '...', 'data': ...}]
    'others_res', 'others_fname'
]
for k in SESSION_KEYS:
    if k not in st.session_state:
        st.session_state[k] = None if 'res' not in k or k == 'msds_res' else ""
if st.session_state['msds_res'] is None:
    st.session_state['msds_res'] = []

# ==============================================================================
# [1번: SPEC 변환기 로직]
# ==============================================================================
def process_spec(pdf_file, product_name, mode):
    pdf_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                pdf_text += extracted + "\n"
                
    context = {
        "PRODUCT": product_name,
        "COLOR": "PALE YELLOW TO YELLOW",
        "SG": "0.902 ~ 0.922",
        "RI": "1.466 ~ 1.476",
        "DATE": datetime.now().strftime("%d. %b. %Y").upper()
    }
    
    if mode == "CFF":
        color_match = re.search(r'COLOR\s*:(.*?)APPEARANCE\s*:', pdf_text, re.DOTALL | re.IGNORECASE)
        if color_match: context["COLOR"] = color_match.group(1).strip().upper()
        sg_match = re.search(r'SPECIFIC GRAVITY.*?\(\d+°C\)\s*:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
        if sg_match:
            sg_base = float(sg_match.group(1))
            context["SG"] = f"{sg_base - 0.01:.3f} ~ {sg_base + 0.01:.3f}"
        ri_match = re.search(r'REFRACTIVE INDEX.*?\(\d+°C\)\s*:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
        if ri_match:
            ri_base = float(ri_match.group(1))
            context["RI"] = f"{ri_base - 0.005:.3f} ~ {ri_base + 0.005:.3f}"
    elif mode == "HP":
        color_match = re.search(r'■\s*COLOR\s*:(.*?)■\s*APPEARANCE\s*:', pdf_text, re.DOTALL | re.IGNORECASE)
        if color_match: context["COLOR"] = color_match.group(1).strip().upper()
        sg_match = re.search(r'■\s*SPECIFIC GRAVITY.*?\:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
        if sg_match:
            sg_base = float(sg_match.group(1))
            context["SG"] = f"{sg_base - 0.01:.3f} ~ {sg_base + 0.01:.3f}"
        ri_match = re.search(r'■\s*REFRACTIVE INDEX.*?\:\s*([\d\.]+)\s*[±\+/-]\s*[\d\.]+', pdf_text, re.IGNORECASE)
        if ri_match:
            ri_base = float(ri_match.group(1))
            context["RI"] = f"{ri_base - 0.005:.3f} ~ {ri_base + 0.005:.3f}"

    doc_path = get_resource_path("SPEC templates/spec.docx")
    doc = DocxTemplate(doc_path)
    doc.render(context)
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio, f"{product_name} SPEC.docx"


# ==============================================================================
# [2번: ALLERGY 변환기 로직]
# ==============================================================================
def extract_cas(text):
    if pd.isna(text): return []
    clean_text = str(text).replace('/', ' ').replace('\n', ' ').replace('\r', ' ')
    return re.findall(r'\d{2,7}-\d{2}-\d', clean_text)

def logic_cff_83(input_df, template_path, customer_name, product_name):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    for row in ws.iter_rows(min_col=3, max_col=3, min_row=1):
        for cell in row:
            if str(cell.value).startswith('='): cell.value = None
    if "Sheet2" in wb.sheetnames: del wb["Sheet2"]
    source_data = {}
    for idx, row in input_df.iterrows():
        cas_text = row.iloc[5] if len(row) > 5 else None
        val = row.iloc[11] if len(row) > 11 else None
        for cas in extract_cas(cas_text): source_data[cas] = val
    for r in range(1, ws.max_row + 1):
        template_cas_text = ws.cell(row=r, column=2).value
        if template_cas_text:
            for t_cas in extract_cas(template_cas_text):
                if t_cas in source_data:
                    ws.cell(row=r, column=3).value = source_data[t_cas]
                    break 
    ws['B9'] = customer_name
    ws['B10'] = product_name
    ws['E10'] = datetime.now().strftime("%Y-%m-%d")
    return wb

def logic_cff_26(input_df, template_path, customer_name, product_name):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    for row in ws.iter_rows(min_col=3, max_col=3, min_row=18, max_row=43):
        for cell in row: cell.value = None
    source_data = {}
    for idx, row in input_df.iterrows():
        cas_text = row.iloc[5] if len(row) > 5 else None
        val = row.iloc[11] if len(row) > 11 else None
        for cas in extract_cas(cas_text): source_data[cas] = val
    for r in range(1, ws.max_row + 1):
        template_cas_text = ws.cell(row=r, column=2).value
        if template_cas_text:
            for t_cas in extract_cas(template_cas_text):
                if t_cas in source_data:
                    ws.cell(row=r, column=3).value = source_data[t_cas]
                    break
    ws['B11'] = customer_name; ws['B12'] = product_name; ws['E13'] = datetime.now().strftime("%Y-%m-%d")
    align_center = Alignment(horizontal='center', vertical='center')
    for row in ws.iter_rows(min_col=3, max_col=6, min_row=18, max_row=43):
        for cell in row: cell.alignment = align_center
    return wb

def logic_hp_83(input_df, template_path, customer_name, product_name):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    for row in ws.iter_rows(min_col=3, max_col=3, min_row=1):
        for cell in row:
            if str(cell.value).startswith('='): cell.value = None
    if "Sheet2" in wb.sheetnames: del wb["Sheet2"]
    source_data = {}
    for idx, row in input_df.iterrows():
        cas_text = row.iloc[1] if len(row) > 1 else None
        val = row.iloc[2] if len(row) > 2 else None
        for cas in extract_cas(cas_text): source_data[cas] = val
    for r in range(1, ws.max_row + 1):
        template_cas_text = ws.cell(row=r, column=2).value
        if template_cas_text:
            for t_cas in extract_cas(template_cas_text):
                if t_cas in source_data:
                    val_to_insert = source_data[t_cas]
                    if pd.notna(val_to_insert) and str(val_to_insert).strip() not in ['0', '0.0']:
                        ws.cell(row=r, column=3).value = val_to_insert
                    break 
    ws['B9'] = customer_name; ws['B10'] = product_name; ws['E10'] = datetime.now().strftime("%Y-%m-%d")
    return wb

def logic_hp_26(input_df, template_path, customer_name, product_name):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    for row in ws.iter_rows(min_col=3, max_col=3, min_row=18, max_row=43):
        for cell in row: cell.value = None
    source_data = {}
    for idx, row in input_df.iterrows():
        cas_text = row.iloc[1] if len(row) > 1 else None
        val = row.iloc[2] if len(row) > 2 else None
        for cas in extract_cas(cas_text): source_data[cas] = val
    for r in range(1, ws.max_row + 1):
        template_cas_text = ws.cell(row=r, column=2).value
        if template_cas_text:
            for t_cas in extract_cas(template_cas_text):
                if t_cas in source_data:
                    val_to_insert = source_data[t_cas]
                    if pd.notna(val_to_insert) and str(val_to_insert).strip() not in ['0', '0.0']:
                        ws.cell(row=r, column=3).value = val_to_insert
                    break
    ws['B11'] = customer_name; ws['B12'] = product_name; ws['E13'] = datetime.now().strftime("%Y-%m-%d")
    align_center = Alignment(horizontal='center', vertical='center')
    for row in ws.iter_rows(min_col=3, max_col=6, min_row=18, max_row=43):
        for cell in row: cell.alignment = align_center
    return wb

def to_excel(data):
    output = io.BytesIO()
    if isinstance(data, pd.DataFrame):
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            data.to_excel(writer, index=False, sheet_name='Sheet1')
    else:
        data.save(output)
    return output.getvalue()


# ==============================================================================
# [3번: IFRA 변환기 로직]
# ==============================================================================
def extract_text_between_ifra(text, start_keyword, end_keyword=None):
    def flexible_escape(kw):
        escaped = re.escape(kw).replace(r'\ ', r'\s+').replace(' ', r'\s+')
        escaped = escaped.replace(r'\.', r'\s*\.?\s*').replace(r'\*', r'\s*\*\s*')
        return escaped
    start_pattern = flexible_escape(start_keyword)
    if end_keyword:
        end_pattern = flexible_escape(end_keyword)
        match = re.search(f"{start_pattern}(.*?){end_pattern}", text, re.DOTALL | re.IGNORECASE)
    else:
        match = re.search(f"{start_pattern}(.*)", text, re.IGNORECASE)
    if match: return process_value_ifra(match.group(1).strip())
    return "Not Permitted"

def process_value_ifra(val_str):
    if isinstance(val_str, (int, float)): val_str = str(val_str)
    if not val_str: return "Not Permitted"
    val_lower = val_str.lower()
    if "not" in val_lower and "permitted" in val_lower: return "Not Permitted"
    if "not" in val_lower and "restricted" in val_lower: return "Not Restricted"
    num_match = re.search(r'\d+\.?\d*', val_str)
    if not num_match: return "Not Permitted"
    try:
        clean_str = num_match.group(0)
        val_float = float(clean_str)
        if val_float == 0.0: return "Not Permitted"
        s = str(val_float)
        if '.' in s:
            int_part, dec_part = s.split('.')
            dec_part = dec_part[:2].ljust(2, '0')
            return f"{int_part}.{dec_part}"
        return f"{s}.00"
    except ValueError: return "Not Permitted"

def process_ifra(pdf_file, customer_name, product_name, mode):
    full_text = ""
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages: full_text += page.extract_text() + "\n"
    if mode == "CFF":
        context = {
            "CUSTOMER": customer_name, "PRODUCT": product_name,
            "CATEGORY1": extract_text_between_ifra(full_text, "Category 1", "Category 2"),
            "CATEGORY2": extract_text_between_ifra(full_text, "Category 2", "Category 3"),
            "CATEGORY3": extract_text_between_ifra(full_text, "Category 3", "Category 4"),
            "CATEGORY4": extract_text_between_ifra(full_text, "Category 4", "Category 5.A"),
            "CATEGORY5_A": extract_text_between_ifra(full_text, "Category 5.A", "Category 5.B"),
            "CATEGORY5_B": extract_text_between_ifra(full_text, "Category 5.B", "Category 5.C"),
            "CATEGORY5_C": extract_text_between_ifra(full_text, "Category 5.C", "Category 5.D"),
            "CATEGORY5_D": extract_text_between_ifra(full_text, "Category 5.D", "Category 6"),
            "CATEGORY6": extract_text_between_ifra(full_text, "Category 6", "Category 7.A"),
            "CATEGORY7_A": extract_text_between_ifra(full_text, "Category 7.A", "Category 7.B"),
            "CATEGORY7_B": extract_text_between_ifra(full_text, "Category 7.B", "Category 8"),
            "CATEGORY8": extract_text_between_ifra(full_text, "Category 8", "Category 9"),
            "CATEGORY9": extract_text_between_ifra(full_text, "Category 9", "Category 10.A"),
            "CATEGORY10_A": extract_text_between_ifra(full_text, "Category 10.A", "Category 10.B"),
            "CATEGORY10_B": extract_text_between_ifra(full_text, "Category 10.B", "Category 11.A"),
            "CATEGORY11_A": extract_text_between_ifra(full_text, "Category 11.A", "Category 11.B"),
            "CATEGORY11_B": extract_text_between_ifra(full_text, "Category 11.B", "Category 12"),
            "CATEGORY12": extract_text_between_ifra(full_text, "Category 12", None)
        }
    elif mode == "HP":
        context = {
            "CUSTOMER": customer_name, "PRODUCT": product_name,
            "CATEGORY1": extract_text_between_ifra(full_text, "Category 1*", "Category 2"),
            "CATEGORY2": extract_text_between_ifra(full_text, "Category 2", "Category 3"),
            "CATEGORY3": extract_text_between_ifra(full_text, "Category 3", "Category 4"),
            "CATEGORY4": extract_text_between_ifra(full_text, "Category 4", "Category 5.A"),
            "CATEGORY5_A": extract_text_between_ifra(full_text, "Category 5.A", "Category 5.B"),
            "CATEGORY5_B": extract_text_between_ifra(full_text, "Category 5.B", "Category 5.C"),
            "CATEGORY5_C": extract_text_between_ifra(full_text, "Category 5.C", "Category 5.D"),
            "CATEGORY5_D": extract_text_between_ifra(full_text, "Category 5.D", "Category 6*"),
            "CATEGORY6": extract_text_between_ifra(full_text, "Category 6*", "Category 7.A"),
            "CATEGORY7_A": extract_text_between_ifra(full_text, "Category 7.A", "Category 7.B"),
            "CATEGORY7_B": extract_text_between_ifra(full_text, "Category 7.B", "Category 8"),
            "CATEGORY8": extract_text_between_ifra(full_text, "Category 8", "Category 9"),
            "CATEGORY9": extract_text_between_ifra(full_text, "Category 9", "Category 10.A"),
            "CATEGORY10_A": extract_text_between_ifra(full_text, "Category 10.A", "Category 10.B"),
            "CATEGORY10_B": extract_text_between_ifra(full_text, "Category 10.B", "Category 11.A"),
            "CATEGORY11_A": extract_text_between_ifra(full_text, "Category 11.A", "Category 11.B"),
            "CATEGORY11_B": extract_text_between_ifra(full_text, "Category 11.B", "Category 12"),
            "CATEGORY12": extract_text_between_ifra(full_text, "Category 12", None)
        }
    template_path = get_resource_path("IFRA templates/IFRA.docx")
    doc = DocxTemplate(template_path)
    doc.render(context)
    output_io = io.BytesIO()
    doc.save(output_io)
    output_io.seek(0)
    return output_io, f"{product_name} IFRA 51TH.docx"

# ==============================================================================
# [4번: MSDS 변환기 로직] (핵심 로직 보존)
# ==============================================================================
def get_master_data_path():
    try: base_path = sys._MEIPASS
    except Exception: base_path = os.path.abspath(".")
    p1 = os.path.join(base_path, "data", "master_data.xlsx")
    if os.path.exists(p1): return p1
    p2 = os.path.join(base_path, "master_data.xlsx")
    if os.path.exists(p2): return p2
    for root, dirs, files in os.walk(base_path):
        for f in files:
            if "master_data" in f and f.endswith(".xlsx") and not f.startswith("~"):
                return os.path.join(root, f)
    return None

FONT_STYLE = Font(name='굴림', size=8)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_TITLE = Alignment(horizontal='center', vertical='center', wrap_text=True)

def safe_write_force(ws, row, col, value, center=False):
    cell = ws.cell(row=row, column=col)
    try: cell.value = value
    except AttributeError:
        try:
            for rng in list(ws.merged_cells.ranges):
                if cell.coordinate in rng:
                    ws.unmerge_cells(str(rng))
                    cell = ws.cell(row=row, column=col)
                    break
            cell.value = value
        except: pass
    if cell.font.name != '굴림': cell.font = FONT_STYLE
    if center: cell.alignment = ALIGN_CENTER
    else: cell.alignment = ALIGN_LEFT

def get_description_smart(code, code_map):
    clean_code = str(code).replace(" ", "").upper().strip()
    if clean_code in code_map: return code_map[clean_code]
    if "+" in clean_code:
        parts = clean_code.split("+")
        found_texts = [code_map[p] for p in parts if p in code_map]
        if found_texts: return " ".join(found_texts)
    return ""

def calculate_smart_height_basic(text, mode="CFF(K)"): 
    if not text: return 19.2
    lines = str(text).split('\n')
    total_visual_lines = 0
    if "E" in mode:
        char_limit = 58.0
        for line in lines:
            if len(line) == 0: total_visual_lines += 1
            else:
                words = line.split(" ")
                current_len = 0; lines_for_this_paragraph = 1
                for word in words:
                    if current_len == 0: current_len = len(word)
                    elif current_len + 1 + len(word) <= char_limit: current_len += 1 + len(word)
                    else: lines_for_this_paragraph += 1; current_len = len(word)
                total_visual_lines += lines_for_this_paragraph
        clean_check = re.sub(r'[\s]+', '', str(text).lower())
        if "rinsecautiouslywithwater" in clean_check and "removecontactlenses" in clean_check:
            if total_visual_lines < 3: total_visual_lines = 3
        if total_visual_lines <= 1: return 18.75
        elif total_visual_lines == 2: return 25.5
        elif total_visual_lines == 3: return 36.0
        elif total_visual_lines == 4: return 44.0
        elif total_visual_lines == 5: return 54.0
        else: return 64.0 + (total_visual_lines - 6) * 10.0
    else:
        char_limit = 45.0
        for line in lines:
            if len(line) == 0: total_visual_lines += 1
            else: total_visual_lines += math.ceil(len(line) / char_limit)
        if total_visual_lines <= 1: return 19.2
        elif total_visual_lines == 2: return 26.0
        elif total_visual_lines == 3: return 36.0
        else: return 45.0

def format_and_calc_height_sec47(text, mode="CFF(K)"):
    if not text: return "", 19.2
    if "E" in mode:
        keywords = r"(IF|If|Get|When|Wash|Remove|Take|Prevent|Call|Move|Settle|Please|After|Should|Rescuer|For|Do|Wipe|Follow|Stop|Collect|Make|Absorb|Put|Since|Contaminated|Without|Empty|Keep|Store|The|It|Some|During|Containers)"
        formatted_text = re.sub(r'(?<=[a-z0-9\)\]\.\;])\s+(' + keywords + r'\b)', r'\n\1', text)
        formatted_text = re.sub(r'\.([A-Z])', r'.\n\1', formatted_text).replace("Follow\nStop", "Follow Stop")
        formatted_text = re.sub(r'\band\s*\n\s*([A-Z])', r'and \1', formatted_text)
        formatted_text = re.sub(r'unattended\.\s*\n\s*If', 'unattended. If', formatted_text)
        formatted_text = re.sub(r'minutes\.\s*\n\s*Remove', 'minutes. Remove', formatted_text)
        formatted_text = re.sub(r'do\.\s*\n\s*Continue', 'do. Continue', formatted_text)
        lines = [line.strip() for line in formatted_text.split('\n') if line.strip()]
        final_text = "\n".join(lines)
        total_visual_lines = 0
        for line in lines:
            if len(line) == 0: total_visual_lines += 1
            else:
                words = line.split(" "); current_len = 0; lines_for_this_paragraph = 1
                for word in words:
                    if current_len == 0: current_len = len(word)
                    elif current_len + 1 + len(word) <= 73.0: current_len += 1 + len(word)
                    else: lines_for_this_paragraph += 1; current_len = len(word)
                total_visual_lines += lines_for_this_paragraph
        if total_visual_lines == 0: total_visual_lines = 1
        height = max(total_visual_lines * 12.0, 24.0)
        return final_text, height
    else:
        formatted_text = re.sub(r'(?<!\d)\.(?!\d)(?!\n)', '.\n', text)
        lines = [line.strip() for line in formatted_text.split('\n') if line.strip()]
        final_text = "\n".join(lines)
        total_visual_lines = 0
        for line in lines:
            line_len = sum(2 if '가' <= ch <= '힣' else 1.1 for ch in line)
            visual_lines = max(math.ceil(line_len / 90.0), 1)
            total_visual_lines += visual_lines
        if total_visual_lines == 0: total_visual_lines = 1
        height = max((total_visual_lines * 10) + 10, 24.0)
        return final_text, height

def fill_fixed_range(ws, start_row, end_row, codes, code_map, mode="CFF(K)"):
    unique_codes = []; seen = set()
    for c in codes:
        clean = c.replace(" ", "").upper().strip()
        if clean not in seen: unique_codes.append(clean); seen.add(clean)
    limit = end_row - start_row + 1
    for i in range(limit):
        current_row = start_row + i
        if i < len(unique_codes):
            code = unique_codes[i]
            desc = get_description_smart(code, code_map)
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = calculate_smart_height_basic(desc, mode)
            safe_write_force(ws, current_row, 2, code, center=False)
            safe_write_force(ws, current_row, 4, desc, center=False)
        else:
            if "K" in mode and current_row in [25, 38, 50, 64, 70]:
                ws.row_dimensions[current_row].hidden = False
                safe_write_force(ws, current_row, 2, "")
                safe_write_force(ws, current_row, 4, "자료없음", center=False)
            elif "E" in mode and current_row in [24, 38, 50, 64, 70]:
                ws.row_dimensions[current_row].hidden = False
                safe_write_force(ws, current_row, 2, "")
                safe_write_force(ws, current_row, 4, "no data available", center=False)
            else:
                ws.row_dimensions[current_row].hidden = True
                safe_write_force(ws, current_row, 2, "") 
                safe_write_force(ws, current_row, 4, "")

def fill_composition_data(ws, comp_data, cas_to_name_map, mode="CFF(K)"):
    start_row = 80; end_row = 122 if "E" in mode else 123
    limit = end_row - start_row + 1
    for i in range(limit):
        current_row = start_row + i
        if i < len(comp_data):
            cas_no, concentration = comp_data[i]
            chem_name = cas_to_name_map.get(cas_no.replace(" ", "").strip(), "")
            ws.row_dimensions[current_row].hidden = False
            ws.row_dimensions[current_row].height = 26.7
            safe_write_force(ws, current_row, 1, chem_name, center=False)
            safe_write_force(ws, current_row, 4, cas_no, center=False) 
            safe_write_force(ws, current_row, 6, concentration if concentration else "", center=True)
        else:
            ws.row_dimensions[current_row].hidden = True
            safe_write_force(ws, current_row, 1, ""); safe_write_force(ws, current_row, 4, ""); safe_write_force(ws, current_row, 6, "")

def fill_regulatory_section(ws, start_row, end_row, substances, data_map, col_key, mode="CFF(K)"):
    limit = end_row - start_row + 1
    for i in range(limit):
        current_row = start_row + i
        if i < len(substances):
            substance_name = substances[i]
            safe_write_force(ws, current_row, 1, substance_name, center=False)
            cell_data = str(data_map.get(substance_name, {}).get(col_key, "")) if substance_name in data_map else ""
            if cell_data == "nan": cell_data = ""
            safe_write_force(ws, current_row, 2, cell_data, center=False)
            ws.row_dimensions[current_row].hidden = False
            _, h = format_and_calc_height_sec47(cell_data, mode=mode)
            ws.row_dimensions[current_row].height = max(h, 24.0)
        else:
            safe_write_force(ws, current_row, 1, ""); safe_write_force(ws, current_row, 2, "")
            ws.row_dimensions[current_row].hidden = True

def auto_crop(pil_img):
    try:
        if pil_img.mode != 'RGB':
            bg = PILImage.new('RGB', pil_img.size, (255, 255, 255))
            if pil_img.mode == 'RGBA': bg.paste(pil_img, mask=pil_img.split()[3])
            else: bg.paste(pil_img)
            pil_img = bg
        bbox = ImageChops.invert(pil_img).getbbox()
        return pil_img.crop(bbox) if bbox else pil_img
    except: return pil_img

def normalize_image_legacy(pil_img):
    try:
        if pil_img.mode in ('RGBA', 'LA') or (pil_img.mode == 'P' and 'transparency' in pil_img.info):
            background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
            if pil_img.mode == 'P': pil_img = pil_img.convert('RGBA')
            background.paste(pil_img, mask=pil_img.split()[3])
            pil_img = background
        else: pil_img = pil_img.convert('RGB')
        return pil_img.resize((32, 32)).convert('L')
    except: return pil_img.resize((32, 32)).convert('L')

def normalize_image_smart(pil_img):
    try: return auto_crop(pil_img).resize((64, 64)).convert('L')
    except: return pil_img.resize((64, 64)).convert('L')

def get_reference_images():
    img_folder = get_resource_path("reference_imgs")
    if not os.path.exists(img_folder): return {}, False
    try:
        ref_images = {}
        for fname in sorted(os.listdir(img_folder)):
            if fname.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.tif', '.tiff')):
                try: ref_images[fname] = PILImage.open(os.path.join(img_folder, fname))
                except: continue
        return ref_images, True
    except: return {}, False

def is_blue_dominant(pil_img):
    try:
        img = pil_img.resize((50, 50)).convert('RGB')
        data = np.array(img); r, g, b = data[:,:,0], data[:,:,1], data[:,:,2]
        blue_mask = (b > r + 30) & (b > g + 30)
        return (np.sum(blue_mask) / 2500) > 0.05
    except: return False

def is_square_shaped(width, height): return height != 0 and 0.8 < (width / height) < 1.2 

def find_best_match_name(src_img, ref_images, mode="CFF(K)"):
    best_score = float('inf'); best_name = None
    if mode in ["HP(K)", "HP(E)"]: src_norm = normalize_image_smart(src_img); threshold = 70 
    else: src_norm = normalize_image_legacy(src_img); threshold = 65
    try:
        src_arr = np.array(src_norm, dtype='int16')
        for name, ref_img in ref_images.items():
            ref_norm = normalize_image_smart(ref_img) if mode in ["HP(K)", "HP(E)"] else normalize_image_legacy(ref_img)
            diff = np.mean(np.abs(src_arr - np.array(ref_norm, dtype='int16')))
            if diff < best_score: best_score = diff; best_name = name
        return best_name if best_score < threshold else None
    except: return None

def extract_number(filename):
    nums = re.findall(r'\d+', filename)
    return int(nums[0]) if nums else 999

def extract_codes_ordered(text):
    return list(dict.fromkeys([c.replace(" ", "").upper() for c in re.findall(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)", text)]))

def get_clustered_lines(doc):
    all_lines = []
    noise_regexs = [
        r'^\s*\d+\s*/\s*\d+\s*$', r'물질안전보건자료', r'Material Safety Data Sheet', 
        r'PAGE', r'Ver\.\s*:?\s*\d+\.?\d*', r'발행일\s*:?.*', r'Date of issue',
        r'주식회사\s*고려.*', r'Cff', r'Corea\s*flavors.*', r'제\s*품\s*명\s*:?.*',
        r'according to the Global Harmonized System', r'Product Name', r'Date\s*:\s*\d{2}\.[a-zA-Z]{3}\.\d{4}'
    ]
    global_y_offset = 0
    for page in doc:
        page_h = page.rect.height
        clip_rect = fitz.Rect(0, 60, page.rect.width, page_h - 50)
        words = sorted(page.get_text("words", clip=clip_rect), key=lambda w: w[1]) 
        rows = []
        if words:
            current_row = [words[0]]; row_base_y = words[0][1]
            for w in words[1:]:
                if abs(w[1] - row_base_y) < 8: current_row.append(w)
                else: rows.append(sorted(current_row, key=lambda x: x[0])); current_row = [w]; row_base_y = w[1]
            if current_row: rows.append(sorted(current_row, key=lambda x: x[0]))
        for row in rows:
            line_text = " ".join([w[4] for w in row])
            if not any(re.search(pat, line_text, re.IGNORECASE) for pat in noise_regexs):
                avg_y = sum([w[1] for w in row]) / len(row)
                all_lines.append({
                    'text': line_text, 'global_y0': avg_y + global_y_offset,
                    'global_y1': (sum([w[3] for w in row]) / len(row)) + global_y_offset
                })
        global_y_offset += page_h
    return all_lines

def extract_section_smart(all_lines, start_kw, end_kw, mode="CFF(K)"):
    start_idx = end_idx = -1
    clean_start_kw = start_kw.replace(" ", "")
    for i, line in enumerate(all_lines):
        check = line['text'].replace(" ", "").lower() if "E" in mode else line['text'].replace(" ", "")
        target = clean_start_kw.lower() if "E" in mode else clean_start_kw
        if target in check: start_idx = i; break
    if start_idx == -1: return ""
    clean_end_kws = [k.replace(" ", "") for k in (end_kw if isinstance(end_kw, list) else [end_kw])]
    for i in range(start_idx + 1, len(all_lines)):
        line_clean = all_lines[i]['text'].replace(" ", "").lower() if "E" in mode else all_lines[i]['text'].replace(" ", "")
        if any((cek.lower() if "E" in mode else cek) in line_clean for cek in clean_end_kws): end_idx = i; break
    if end_idx == -1: end_idx = len(all_lines)
    target_lines_raw = all_lines[start_idx : end_idx]
    if not target_lines_raw: return ""
    first_line = target_lines_raw[0].copy()
    txt = first_line['text']
    match = re.search(re.escape(start_kw).replace(r"\ ", r"\s*"), txt, re.IGNORECASE)
    first_line['text'] = re.sub(r"^[:\.\-\s]+", "", txt[match.end():].strip()) if match else (txt.split(start_kw, 1)[1].strip() if start_kw in txt else "")
    target_lines = [first_line] if first_line['text'].strip() else []
    target_lines.extend(target_lines_raw[1:])
    if not target_lines: return ""

    if mode == "CFF(E)": garbage_heads = ["Classification of the substance or mixture", "Classification of the substance or", "mixture", "Precautionary statements", "Hazard pictograms", "Signal word", "Hazard statements", "Response", "Storage", "Disposal", "Other hazards", "General advice", "In case of eye contact", "In case of skin contact", "If inhaled", "If swallowed", "Special note for doctors", "Extinguishing media", "Special hazards arising from the", "Advice for firefighters", "Personal precautions, protective", "Environmental precautions", "Methods and materials for containment", "Precautions for safe handling", "Conditions for safe storage, including", "Internal regulations", "ACGIH regulations", "Biological exposure standards", "arising from the", ", protective", "precautions", "and materials for containment", "for safe handling", "for safe storage, including", "conditions for safe storage, including"]; sensitive_garbage_regex = []
    elif mode == "HP(E)": garbage_heads = ["Classification of the substance or mixture", "Classification of the substance or", "mixture", "Precautionary statements", "Hazard pictograms", "Signal word", "Hazard statements", "Response", "Storage", "Disposal", "Other hazards", "General advice", "In case of eye contact", "In case of skin contact", "If inhaled", "If swallowed", "Special note for doctors", "Extinguishing media", "Special hazards arising from the", "Advice for firefighters", "Personal precautions, protective", "Environmental precautions", "Methods and materials for containment", "Precautions for safe handling", "Conditions for safe storage, including", "Internal regulations", "ACGIH regulations", "Biological exposure standards", "arising from the", ", protective", "precautions", "and materials for containment", "for safe handling", "for safe storage, including", "conditions for safe storage, including", "equipment and emergency procedures", "and cleaning up", "any incompatibilities", "suitable (unsuitable) extinguishing media", "(unsuitable) extinguishing media", "specific hazards arising from the chemical", "specific hazards", "from the chemical", "special protective actions for firefighters", "special protective", "for firefighters", "handling", "incompatible materials", "safe storage", "contact with", ", including any incompatibilities", "including any incompatibilities"]; sensitive_garbage_regex = []
    elif mode == "HP(K)": garbage_heads = ["에 접촉했을 때", "에 들어갔을 때", "들어갔을 때", "접촉했을 때", "했을 때", "흡입했을 때", "먹었을 때", "주의사항", "내용물", "취급요령", "저장방법", "보호구", "조치사항", "제거 방법", "소화제", "유해성", "로부터 생기는", "착용할 보호구", "예방조치", "방법", "경고표지 항목", "그림문자", "화학물질", "의사의 주의사항", "기타 의사의 주의사항", "필요한 정보", "관한 정보", "보호하기 위해 필요한 조치사항", "또는 제거 방법", "시 착용할 보호구 및 예방조치", "시 착용할 보호구", "부터 생기는 특정 유해성", "사의 주의사항", "(부적절한) 소화제", "및", "요령", "때", "항의", "색상", "인화점", "비중", "굴절률", "에 의한 규제", "의한 규제", "- 색", "(및 부적절한) 소화제", "특정 유해성", "보호하기 위해 필요한 조치 사항 및 보호구", "저장 방법"]; sensitive_garbage_regex = [r"^시\s+", r"^또는\s+", r"^의\s+"]
    else: garbage_heads = ["에 접촉했을 때", "에 들어갔을 때", "들어갔을 때", "접촉했을 때", "했을 때", "흡입했을 때", "먹었을 때", "주의사항", "내용물", "취급요령", "저장방법", "보호구", "조치사항", "제거 방법", "소화제", "유해성", "로부터 생기는", "착용할 보호구", "예방조치", "방법", "경고표지 항목", "그림문자", "화학물질", "의사의 주의사항", "기타 의사의 주의사항", "필요한 정보", "관한 정보", "보호하기 위해 필요한 조치사항", "또는 제거 방법", "시 착용할 보호구 및 예방조치", "시 착용할 보호구", "부터 생기는 특정 유해성", "사의 주의사항", "(부적절한) 소화제", "및", "요령", "때", "항의", "색상", "인화점", "비중", "굴절률", "에 의한 규제", "의한 규제"]; sensitive_garbage_regex = [r"^시\s+", r"^또는\s+", r"^의\s+"]

    cleaned_lines = []
    for line in target_lines:
        txt = re.sub(r'^\s*-\s*', '', line['text'].strip()).strip() if mode == "HP(K)" else line['text'].strip()
        for _ in range(3):
            changed = False
            for gb in garbage_heads:
                if txt.lower().replace(" ","").startswith(gb.lower().replace(" ","")):
                    m = re.compile(r"^" + re.escape(gb).replace(r"\ ", r"\s*") + r"[\s\.:]*", re.IGNORECASE).match(txt)
                    if m: txt = txt[m.end():].strip(); changed = True
                    elif txt.lower().startswith(gb.lower()): txt = txt[len(gb):].strip(); changed = True
            for pat in sensitive_garbage_regex:
                m = re.search(pat, txt)
                if m: txt = txt[m.end():].strip(); changed = True
            txt = re.sub(r"^[:\.\)\s]+", "", txt)
            if not changed: break
        if txt:
            line['text'] = re.sub(r'^\s*-\s*', '', txt).strip() if mode == "HP(K)" else txt
            cleaned_lines.append(line)
    
    if not cleaned_lines: return ""
    final_text = cleaned_lines[0]['text']
    for i in range(1, len(cleaned_lines)):
        prev, curr = cleaned_lines[i-1], cleaned_lines[i]
        prev_txt, curr_txt = prev['text'].strip(), curr['text'].strip()
        if mode in ["CFF(E)", "HP(E)"]:
            final_text += ("\n" + curr_txt) if re.match(r"^(\-|•|\*|\d+\.)", curr_txt) or (curr['global_y0'] - prev['global_y1']) >= 3.0 else (" " + curr_txt)
        else: 
            if re.search(r"(\.|시오|음|함|것|임|있음|주의|금지|참조|따르시오|마시오)$", prev_txt) or re.match(r"^(\-|•|\*|\d+\.|[가-하]\.|\(\d+\))", curr_txt): final_text += "\n" + curr_txt
            elif (curr['global_y0'] - prev['global_y1']) < 3.0: 
                need_space = False
                if prev_txt and curr_txt and 0xAC00 <= ord(prev_txt[-1]) <= 0xD7A3 and 0xAC00 <= ord(curr_txt[0]) <= 0xD7A3:
                    if prev_txt[-1] in ['을','를','이','가','은','는','의','와','과','에','로','서','고','며','여','해','나','면','니','등','및','또는','경우',',',')','속'] or any(curr_txt.startswith(x) for x in ['및','또는','(','참고']): need_space = True
                final_text += (" " + curr_txt) if need_space else curr_txt
            else: final_text += "\n" + curr_txt
    return final_text

def parse_sec8_hp_content(text):
    if not text: return "자료없음"
    valid_lines = []
    for chunk in text.split("-"):
        clean_chunk = chunk.strip()
        if not clean_chunk: continue
        if ":" in clean_chunk:
            parts = clean_chunk.split(":", 1)
            name_part, value_part = parts[0].strip(), parts[1].strip()
            if "해당없음" in value_part: continue 
            valid_lines.append(f"{name_part.replace('[', '').replace(']', '').strip()} : {value_part.replace('[', '').replace(']', '').strip()}")
        elif "해당없음" not in clean_chunk:
            valid_lines.append(clean_chunk.replace("[", "").replace("]", "").strip())
    return "\n".join(valid_lines) if valid_lines else "자료없음"

def parse_pdf_final(doc, mode="CFF(K)"):
    all_lines = get_clustered_lines(doc)
    result = {"hazard_cls": [], "signal_word": "", "h_codes": [], "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [], "composition_data": [], "sec4_to_7": {}, "sec8": {}, "sec9": {}, "sec14": {}, "sec15": {}}
    sec9_lines = []; start_9 = end_9 = -1
    for i, line in enumerate(all_lines):
        if "9. PHYSICAL" in line['text'].upper() or "9. 물리화학" in line['text']: start_9 = i
        if "10. STABILITY" in line['text'].upper() or "10. 안정성" in line['text']: end_9 = i; break
    if start_9 != -1: sec9_lines = all_lines[start_9:(end_9 if end_9 != -1 else len(all_lines))]

    if mode == "HP(E)":
        b19_clean = re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "A. GHS Classification", "B. GHS label elements", mode)).strip()
        result["hazard_cls"] = [l.strip() for l in b19_clean.split('\n') if l.strip()]
        result["signal_word"] = re.sub(r'^(?:[sS]\b)?[\s\-\○•]+', '', extract_section_smart(all_lines, "Signal word", "Hazard statement", mode).strip()).strip()
        result["h_codes"] = extract_codes_ordered(extract_section_smart(all_lines, "Hazard statement", "Precautionary statement", mode))
        result["p_prev"] = extract_codes_ordered(extract_section_smart(all_lines, "1) Prevention", "2) Response", mode))
        result["p_resp"] = extract_codes_ordered(extract_section_smart(all_lines, "2) Response", "3) Storage", mode))
        result["p_stor"] = extract_codes_ordered(extract_section_smart(all_lines, "3) Storage", "4) Disposal", mode))
        result["p_disp"] = extract_codes_ordered(extract_section_smart(all_lines, "4) Disposal", "C. Other hazards", mode))

        comp_text = extract_section_smart(all_lines, "3. Composition", ["4. FIRST-AID", "4. First aid"], mode)
        regex_cas = re.compile(r'\b\d{2,7}-\d{2}-\d\b'); regex_conc = re.compile(r'\b(\d+(?:\.\d+)?)\s*(?:~|-)\s*(\d+(?:\.\d+)?)\b')
        cas_list = regex_cas.findall(comp_text); conc_list = []
        for match in regex_conc.finditer(regex_cas.sub(" ", comp_text)):
            if float(match.group(1)) <= 100 and float(match.group(2)) <= 100: conc_list.append(f"{match.group(1)} ~ {match.group(2)}")
        for i in range(max(len(cas_list), len(conc_list))):
            result["composition_data"].append((cas_list[i] if i < len(cas_list) else "", conc_list[i] if i < len(conc_list) else ""))

        sec5_lines = all_lines
        for i, line in enumerate(all_lines):
            if "5. FIREFIGHTING" in line['text'].upper(): sec5_lines = all_lines[i:]; break

        data = {
            "B126": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Eye contact", "Skin contact", mode)).strip(),
            "B127": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Skin contact", "Inhalation contact", mode)).strip(),
            "B128": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Inhalation contact", "Ingestion contact", mode)).strip(),
            "B129": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Ingestion contact", "Delayed and", mode)).strip(),
            "B132": re.sub(r'(?i)^\s*\(unsuitable\)\s*extinguishing\s*media\s*', '', re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(sec5_lines, "Suitable", "Specific hazards", mode)).strip()).strip(),
            "B134": re.sub(r'(?i)^\s*from\s*the\s*chemical\s*', '', re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(sec5_lines, "Specific hazards arising", "Special protective", mode)).strip()).strip(),
            "B136": re.sub(r'(?i)^\s*for\s*firefighters\s*', '', re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(sec5_lines, "Special protective actions", "6. ACCIDENTAL", mode)).strip()).strip(),
            "B140": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Personal precautions", "Environmental precautions", mode)).strip(),
            "B142": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Environmental precautions", "Methods and materials", mode)).strip(),
            "B144": re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Methods and materials for containment", "7. HANDLING", mode)).strip(),
            "B148": re.sub(r'(?i)^\s*handling\s*', '', re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Precautions for safe", "Conditions for safe", mode)).strip()).strip(),
            "B150": re.sub(r'(?i)^[\s,]*including\s*any\s*incompatibilities\s*', '', re.sub(r'(?m)^\s*-\s*', '', extract_section_smart(all_lines, "Conditions for safe storage", "8. EXPOSURE", mode)).strip()).strip()
        }
        result["sec4_to_7"] = data

        s8_clean = re.sub(r'[○•\-\*]+', '', re.sub(r'(?i)^.*TLV\s*', '', extract_section_smart(all_lines, "ACGIH", "OSHA", mode)).strip()).replace("[", "").replace("]", "")
        s8 = {}
        if "Not applicable" in s8_clean or "Not available" in s8_clean or not s8_clean:
            s8["B156"] = "no data available"; s8["B157"] = ""; s8["B158"] = ""
        else:
            lines = [l.strip() for l in s8_clean.split('\n') if l.strip()]
            s8["B156"] = lines[0] if len(lines) > 0 else "no data available"
            s8["B157"] = lines[1] if len(lines) > 1 else ""
            s8["B158"] = "\n".join(lines[2:]) if len(lines) > 2 else ""
        result["sec8"] = s8

        def find_val_in_sec9(lines, keyword):
            for l in lines:
                if keyword.lower() in l['text'].lower():
                    parts = l['text'].split(keyword, 1)
                    if len(parts) > 1:
                        val = re.sub(r'^[:\s\-\.]+', '', parts[1]).strip()
                        if val: return val
            return ""

        s9 = {}
        c_val = find_val_in_sec9(sec9_lines, "Color"); s9["B170"] = c_val.capitalize() if c_val else ""
        s9["B176"] = find_val_in_sec9(sec9_lines, "Flash point")
        g_m = re.search(r'([\d\.]+)', find_val_in_sec9(sec9_lines, "Specific gravity"))
        s9["B183"] = f"{g_m.group(1)} ± 0.010" if g_m else ""
        s9["B189"] = find_val_in_sec9(sec9_lines, "Refractive index").replace("(20℃)", "").strip()
        result["sec9"] = s9

        s14 = {}
        s14["UN"] = re.sub(r'\D', '', extract_section_smart(all_lines, "UN No.", "Proper shipping name", mode))
        s14["NAME"] = re.sub(r'\([^)]*\)', '', re.sub(r'(?i)shipping\s*name', '', re.sub(r'(?i)proper\s*shipping\s*name', '', extract_section_smart(all_lines, "Proper shipping name", ["C. Hazard Class", "Hazard Class"], mode)))).replace("-", "").strip()
        class_match = re.search(r'(\d)', extract_section_smart(all_lines, "C. Hazard Class", ["D. IMDG", "Packing group"], mode).replace("-", ""))
        s14["CLASS"] = class_match.group(1) if class_match else ""
        s14["PG"] = extract_section_smart(all_lines, "Packing group", "E. Marine pollutant", mode).replace("-", "").strip()
        s14["ENV"] = extract_section_smart(all_lines, "E. Marine pollutant", "F. Special precautions", mode).replace("-", "").strip()
        result["sec14"] = s14
        result["sec15"] = {"DANGER": ""}
        return result

    if mode == "CFF(E)":
        hazard_cls_lines = []
        for line in re.sub(r'(Category\s*\d+[A-Za-z]?)', r'\1\n', extract_section_smart(all_lines, "2. Hazards identification", "2.2 Labelling", mode)).split('\n'):
            line = line.strip()
            if not line: continue
            if "2.1 Classification" in line:
                line = line.replace("2.1 Classification of the substance or", "").replace("mixture", "").strip()
                if not line: continue 
            if line.lower() not in ["mixture", "mixture."]: hazard_cls_lines.append(line)
        result["hazard_cls"] = hazard_cls_lines

        full_text = "\n".join([l['text'] for l in all_lines])
        m_sig = re.search(r"Signal word\s*[:\-\s]*([A-Za-z]+)", full_text, re.IGNORECASE)
        if m_sig: result["signal_word"] = m_sig.group(1).capitalize()
        
        seen = set()
        for code_raw in re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)").findall(extract_section_smart(all_lines, "2. Hazards", "3. Composition", mode)):
            code = code_raw.replace(" ", "").upper()
            if code in seen: continue
            seen.add(code)
            if code.startswith("H"): result["h_codes"].append(code)
            elif code.startswith("P"):
                p = code.split("+")[0]
                if p.startswith("P2"): result["p_prev"].append(code)
                elif p.startswith("P3"): result["p_resp"].append(code)
                elif p.startswith("P4"): result["p_stor"].append(code)
                elif p.startswith("P5"): result["p_disp"].append(code)

        comp_text = extract_section_smart(all_lines, "3. Composition", "4. FIRST-AID", mode)
        regex_cas = re.compile(r'\b\d{2,7}-\d{2}-\d\b'); regex_conc = re.compile(r'\b(\d+(?:\.\d+)?)\s*(?:~|-)\s*(\d+(?:\.\d+)?)\b')
        cas_list = regex_cas.findall(comp_text); conc_list = []
        for match in regex_conc.finditer(regex_cas.sub(" ", comp_text)):
            if float(match.group(1)) <= 100 and float(match.group(2)) <= 100: conc_list.append(f"{match.group(1)} ~ {match.group(2)}")
        for i in range(max(len(cas_list), len(conc_list))):
            result["composition_data"].append((cas_list[i] if i < len(cas_list) else "", conc_list[i] if i < len(conc_list) else ""))

        data = {
            "B125": extract_section_smart(all_lines, "4.1 General advice", "4.2 In case of eye contact", mode),
            "B126": extract_section_smart(all_lines, "4.2 In case of eye contact", "4.3 In case of skin contact", mode),
            "B127": extract_section_smart(all_lines, "4.3 In case of skin contact", "4.4 If inhaled", mode),
            "B128": extract_section_smart(all_lines, "4.4 If inhaled", "4.5 If swallowed", mode),
            "B129": extract_section_smart(all_lines, "4.5 If swallowed", "4.6 Special note for doctors", mode).replace("Medical personnel, and to ensure that take protection measures is recognized for its substance", ""),
            "B132": extract_section_smart(all_lines, "5.1 Extinguishing media", "5.2 Special hazards", mode),
            "B134": extract_section_smart(all_lines, "5.2 Special hazards", "5.3 Advice for firefighters", mode).replace("substance or mixture", ""),
            "B136": extract_section_smart(all_lines, "5.3 Advice for firefighters", "6. Accidental", mode),
            "B140": extract_section_smart(all_lines, "6.1 Personal precautions", "6.2 Environmental", mode).replace("equipment and emergency procedures", ""),
            "B142": extract_section_smart(all_lines, "6.2 Environmental", "6.3 Methods", mode),
            "B144": extract_section_smart(all_lines, "6.3 Methods", "7. Handling", mode).replace("and cleaning up", ""),
            "B148": extract_section_smart(all_lines, "7.1 Precautions", "7.2 Conditions", mode),
            "B150": extract_section_smart(all_lines, "7.2 Conditions", "8. Exposure", mode).replace("any incompatibilities", "")
        }
        result["sec4_to_7"] = data

        result["sec8"] = {
            "B154": extract_section_smart(all_lines, "Internal regulations", "ACGIH regulations", mode).replace("[", "").replace("]", ""),
            "B156": extract_section_smart(all_lines, "ACGIH regulations", "Biological exposure", mode).replace("[", "").replace("]", "")
        }

        result["sec9"] = {
            "B170": extract_section_smart(sec9_lines, "Color", "Odor", mode),
            "B176": extract_section_smart(sec9_lines, "Flash point", "Evaporation rate", mode),
            "B183": extract_section_smart(sec9_lines, "Specific gravity", "Partition coefficient", mode).replace("(20/20℃)", "").replace("(Water=1)", "").strip(),
            "B189": extract_section_smart(sec9_lines, "Refractive index", "10. Stability", mode).replace("(20℃)", "").strip()
        }

        s14 = {}
        s14["UN"] = re.sub(r'\D', '', extract_section_smart(all_lines, "14.1 UN number", "14.2 Proper", mode))
        s14["NAME"] = re.sub(r'\([^)]*\)', '', re.sub(r'(?i)shipping\s*name', '', re.sub(r'(?i)proper\s*shipping\s*name', '', extract_section_smart(all_lines, "14.2 Proper", "14.3 Transport", mode)))).strip()
        class_match = re.search(r'(\d)', extract_section_smart(all_lines, "14.3 Transport hazard class", "14.4 Packing group", mode))
        s14["CLASS"] = class_match.group(1) if class_match else ""
        s14["PG"] = extract_section_smart(all_lines, "14.4 Packing group", "14.5 Environmental hazard", mode)
        s14["ENV"] = extract_section_smart(all_lines, "14.5 Environmental hazard", "IATA", mode)
        result["sec14"] = s14

        return result

    if mode == "CFF(K)":
        for i in range(len(all_lines)):
            if "적정선적명" in all_lines[i]['text'] and i > 0:
                prev_line = all_lines[i-1]
                if abs(prev_line['global_y0'] - all_lines[i]['global_y0']) < 20 and "적정선적명" not in prev_line['text'] and "유엔번호" not in prev_line['text']:
                    all_lines[i]['text'] += " " + prev_line['text']; all_lines[i-1]['text'] = ""
    
    limit_y = 999999
    for line in all_lines:
        if "3. 구성성분" in line['text'] or "3. 성분" in line['text']: limit_y = line['global_y0']; break
    full_text_hp = "\n".join([l['text'] for l in all_lines if l['global_y0'] < limit_y])
    
    signal_found = False
    if mode == "HP(K)":
        try:
            start_sig = full_text_hp.find("신호어"); end_sig = full_text_hp.find("유해", start_sig)
            if start_sig != -1 and end_sig != -1:
                m = re.search(r"[-•]\s*(위험|경고)", full_text_hp[start_sig:end_sig])
                if m: result["signal_word"] = m.group(1); signal_found = True
        except: pass
    if not signal_found:
        for line in full_text_hp.split('\n'):
            if "신호어" in line:
                val = line.replace("신호어", "").replace(":", "").strip()
                if val in ["위험", "경고"]: result["signal_word"] = val
            elif line.strip() in ["위험", "경고"] and not result["signal_word"]:
                result["signal_word"] = line.strip()
    
    if mode == "HP(K)":
        state = 0
        for l in full_text_hp.split('\n'):
            if "가. 유해성" in l: state=1; continue
            if "나. 예방조치" in l: state=0; continue
            if state==1 and l.strip() and "공급자" not in l and "회사명" not in l:
                clean_l = l.replace("-", "").strip()
                if clean_l: result["hazard_cls"].append(clean_l)
    else: 
        state = 0
        for l in full_text_hp.split('\n'):
            l_ns = l.replace(" ", "")
            if "2.유해성" in l_ns and "위험성" in l_ns: state = 1; continue 
            if "나.예방조치" in l_ns: state = 0; continue
            if state == 1 and l.strip():
                if "가.유해성" in l_ns and "분류" in l_ns:
                    check_header = re.sub(r'[가-하][\.\s]*유해성[\s\.\·ㆍ\-]*위험성[\s\.\·ㆍ\-]*분류[\s:]*', '', l).strip()
                    if not check_header: continue 
                    l = check_header
                if "공급자" not in l and "회사명" not in l: result["hazard_cls"].append(l.strip())

    all_matches = re.compile(r"([HP]\s?\d{3}(?:\s*\+\s*[HP]\s?\d{3})*)").findall(full_text_hp)
    seen = set()
    if "P321" in full_text_hp and "P321" not in all_matches: all_matches.append("P321")
    for code_raw in all_matches:
        code = code_raw.replace(" ", "").upper()
        if code in seen: continue
        seen.add(code)
        if code.startswith("H"): result["h_codes"].append(code)
        elif code.startswith("P"):
            p = code.split("+")[0]
            if p.startswith("P2"): result["p_prev"].append(code)
            elif p.startswith("P3"): result["p_resp"].append(code)
            elif p.startswith("P4"): result["p_stor"].append(code)
            elif p.startswith("P5"): result["p_disp"].append(code)

    regex_cas_strict = re.compile(r'\b(\d{2,7}\s*-\s*\d{2}\s*-\s*\d)\b')
    regex_cas_ec_kill = re.compile(r'\b\d{2,7}\s*-\s*\d{2,3}\s*-\s*\d\b')
    regex_tilde_range = re.compile(r'(\d+(?:\.\d+)?)\s*~\s*(\d+(?:\.\d+)?)') 
    
    in_comp = False
    for line in all_lines:
        txt = line['text']
        if "3." in txt and ("성분" in txt or "Composition" in txt): in_comp=True; continue
        if "4." in txt and ("응급" in txt or "First" in txt): in_comp=False; break
        if in_comp:
            if re.search(r'^\d+\.\d+', txt): continue 
            c_val = cn_val = ""
            if mode == "HP(K)":
                cas_found = regex_cas_strict.findall(txt)
                if cas_found:
                    c_val = cas_found[0].replace(" ", "")
                    txt_no_cas = re.sub(r'\b(?:(?:19|20)\d{2}-\d{1,2}-\d+|KE-\d+)\b', ' ', txt.replace(cas_found[0], " " * len(cas_found[0])), flags=re.IGNORECASE)
                    m_range = re.search(r'\b(\d+(?:\.\d+)?)\s*(?:-|~)\s*(\d+(?:\.\d+)?)\b', txt_no_cas)
                    if m_range:
                        s, e = m_range.group(1), m_range.group(2)
                        try:
                            if float(s) <= 100 and float(e) <= 100: cn_val = f"{'0' if s=='1' else s} ~ {e}"
                        except: pass
                    if not cn_val:
                        m_single = re.search(r'\b(\d+(?:\.\d+)?)\b', txt_no_cas)
                        if m_single:
                            try:
                                if float(m_single.group(1)) <= 100: cn_val = m_single.group(1)
                            except: pass
            else:
                cas_found = regex_cas_strict.findall(txt)
                if cas_found: c_val = cas_found[0].replace(" ", "")
                else:
                    cas_found_loose = regex_cas_ec_kill.findall(txt)
                    if cas_found_loose and re.match(r'\d{2,7}-\d{2}-\d', cas_found_loose[0].replace(" ", "")): c_val = cas_found_loose[0].replace(" ", "")
                txt_clean = re.sub(r'\b(?:(?:19|20)\d{2}-\d{1,2}-\d+|KE-\d+)\b', ' ', regex_cas_ec_kill.sub(" ", txt), flags=re.IGNORECASE)
                m_tilde = regex_tilde_range.search(txt_clean)
                if m_tilde:
                    s, e = m_tilde.group(1), m_tilde.group(2)
                    try:
                        if float(s) <= 100 and float(e) <= 100: cn_val = f"{'0' if s=='1' else s} ~ {e}"
                    except: pass
            if c_val or cn_val: result["composition_data"].append((c_val, cn_val))

    data = {}
    if mode == "HP(K)":
        data["B125"] = extract_section_smart(all_lines, "가. 눈에", "나. 피부", mode)
        data["B126"] = extract_section_smart(all_lines, "나. 피부", "다. 흡입", mode)
        data["B127"] = extract_section_smart(all_lines, "다. 흡입", "라. 먹었을", mode)
        data["B128"] = extract_section_smart(all_lines, "라. 먹었을", "마. 기타", mode)
        data["B129"] = extract_section_smart(all_lines, "마. 기타", ["5.", "폭발"], mode)
        b132_raw = extract_section_smart(all_lines, "가. 적절한", "나. 화학물질", mode)
        if "직사주수를 사용한" in b132_raw and "\n직사주수를" not in b132_raw: b132_raw = b132_raw.replace("직사주수를 사용한", "\n직사주수를 사용한")
        data["B132"] = b132_raw
        data["B133"] = re.sub(r'^(특정\s*유해성)\s*', '', extract_section_smart(all_lines, "나. 화학물질", "다. 화재진압", mode)).strip()
        data["B134"] = extract_section_smart(all_lines, "다. 화재진압", ["6.", "누출"], mode)
    else: 
        data["B125"] = extract_section_smart(all_lines, "나. 눈", "다. 피부", mode)
        data["B126"] = extract_section_smart(all_lines, "다. 피부", "라. 흡입", mode)
        data["B127"] = extract_section_smart(all_lines, "라. 흡입", "마. 먹었을", mode)
        data["B128"] = extract_section_smart(all_lines, "마. 먹었을", "바. 기타", mode)
        data["B129"] = extract_section_smart(all_lines, "바. 기타", ["5.", "폭발"], mode)
        data["B132"] = extract_section_smart(all_lines, "가. 적절한", "나. 화학물질", mode)
        data["B133"] = re.sub(r'^(특정\s*유해성)\s*', '', extract_section_smart(all_lines, "나. 화학물질", "다. 화재진압", mode)).strip()
        data["B134"] = extract_section_smart(all_lines, "다. 화재진압", ["6.", "누출"], mode)
    
    data["B138"] = extract_section_smart(all_lines, "가. 인체를", "나. 환경을", mode)
    data["B139"] = extract_section_smart(all_lines, "나. 환경을", "다. 정화", mode)
    data["B140"] = extract_section_smart(all_lines, "다. 정화", ["7.", "취급"], mode)
    data["B143"] = extract_section_smart(all_lines, "가. 안전취급", "나. 안전한", mode)
    data["B144"] = extract_section_smart(all_lines, "나. 안전한", ["8.", "노출"], mode)
    result["sec4_to_7"] = data

    sec8_lines = []; start_8 = end_8 = -1
    for i, line in enumerate(all_lines):
        if "8. 노출방지" in line['text']: start_8 = i
        if "9. 물리화학" in line['text']: end_8 = i; break
    if start_8 != -1: sec8_lines = all_lines[start_8:(end_8 if end_8 != -1 else len(all_lines))]
    
    if mode == "HP(K)":
        result["sec8"] = {"B148": parse_sec8_hp_content(extract_section_smart(sec8_lines, "국내노출기준", "ACGIH노출기준", mode)), "B150": parse_sec8_hp_content(extract_section_smart(sec8_lines, "ACGIH노출기준", "생물학적", mode))}
        result["sec9"] = {
            "B163": extract_section_smart(sec9_lines, "- 색", "나. 냄새", mode),
            "B169": extract_section_smart(sec9_lines, "인화점", "아. 증발속도", mode),
            "B176": extract_section_smart(sec9_lines, "비중", "거. n-옥탄올", mode),
            "B182": extract_section_smart(sec9_lines, "굴절률", ["10. 안정성", "10. 화학적"], mode)
        }
    else:
        result["sec8"] = {"B148": extract_section_smart(sec8_lines, "국내규정", "ACGIH", mode), "B150": extract_section_smart(sec8_lines, "ACGIH", "생물학적", mode)}
        result["sec9"] = {
            "B163": extract_section_smart(sec9_lines, "색상", "나. 냄새", mode),
            "B169": extract_section_smart(sec9_lines, "인화점", "아. 증발속도", mode),
            "B176": extract_section_smart(sec9_lines, "비중", "거. n-옥탄올", mode),
            "B182": extract_section_smart(sec9_lines, "굴절률", ["10. 안정성", "10. 화학적"], mode)
        }

    sec14_lines = []; start_14 = end_14 = -1
    for i, line in enumerate(all_lines):
        if "14. 운송에" in line['text']: start_14 = i
        if "15. 법적규제" in line['text']: end_14 = i; break
    if start_14 != -1: sec14_lines = all_lines[start_14:(end_14 if end_14 != -1 else len(all_lines))]
    
    if mode == "HP(K)":
        un_no = extract_section_smart(sec14_lines, "유엔번호", "나. 유엔", mode)
        ship_name = extract_section_smart(sec14_lines, "유엔 적정 선적명", ["다. 운송에서의", "다.운송에서의"], mode)
        class_raw = extract_section_smart(sec14_lines, "다. 운송에서의 위험성 등급", ["라. 용기등급", "라.용기등급"], mode)
        pg_raw = re.sub(r'\(\s*IMDG\s*CODE\s*/\s*IATA\s*DGR\s*\)', '', extract_section_smart(sec14_lines, "라. 용기등급", ["마. 해양오염물질", "마.해양오염물질"], mode), flags=re.IGNORECASE).replace("-", "").strip()
        env_raw = extract_section_smart(sec14_lines, "마. 해양오염물질", ["바. 사용자", "바.사용자"], mode).replace("-", "").strip()
    else:
        un_no = extract_section_smart(sec14_lines, "유엔번호", "나. 적정선적명", mode)
        ship_name = extract_section_smart(sec14_lines, "적정선적명", ["다. 운송에서의", "다.운송에서의"], mode)
        class_raw = extract_section_smart(sec14_lines, "다. 운송에서의 위험성 등급", ["라. 용기등급", "라.용기등급"], mode)
        pg_raw = extract_section_smart(sec14_lines, "라. 용기등급", "마. 환경유해성", mode)
        env_raw = extract_section_smart(sec14_lines, "마. 환경유해성", "IATA", mode)
        
    class_match = re.search(r'(\d)', class_raw)
    result["sec14"] = {"UN": un_no, "NAME": ship_name, "CLASS": class_match.group(1) if class_match else "", "PG": pg_raw, "ENV": env_raw}

    sec15_lines = []; start_15 = end_15 = -1
    for i, line in enumerate(all_lines):
        clean_txt = line['text'].replace(" ", "")
        if "15.법적" in clean_txt: start_15 = i
        if "16.그밖의" in clean_txt or "16.기타" in clean_txt: end_15 = i; break
    sec15_lines = all_lines[start_15:(end_15 if end_15 != -1 else len(all_lines))] if start_15 != -1 else all_lines
        
    danger_act_text = extract_section_smart(sec15_lines, "위험물안전관리법에 의한 규제", ["마. 폐기물", "마.폐기물"], mode) or extract_section_smart(sec15_lines, "위험물안전관리법", ["마. 폐기물", "마.폐기물"], mode)
    result["sec15"] = {"DANGER": danger_act_text, "FULL_TEXT": "\n".join([l['text'] for l in sec15_lines])}

    return result


def process_msds(uploaded_files, product_name_input, option, refractive_index_input, kor_excel_file, kor_form_version):
    master_data_path = get_master_data_path()
    if not master_data_path: return {"error": "내장된 중앙 데이터(master_data.xlsx)를 찾을 수 없습니다."}
    loaded_refs, _ = get_reference_images()
    
    template_path = get_resource_path(os.path.join("MSDS templates", "MSDS 영문.xlsx" if option in ["CFF(E)", "HP(E)"] else "MSDS 국문.xlsx"))
    if not os.path.exists(template_path): return {"error": f"템플릿을 찾을 수 없습니다: {template_path}"}
    
    new_files = []; new_download_data = {}
    code_map = {}; cas_name_map = {}; kor_data_map = {}; eng_data_map = {}
    
    try:
        xls = pd.ExcelFile(master_data_path)
        sheet_names = xls.sheet_names
        target_sheet = next((s for s in sheet_names if "위험" in s and "안전" in s), sheet_names[0])
        df_code = pd.read_excel(master_data_path, sheet_name=target_sheet)
        target_col_idx = 1 if "K" in option else 2
        for _, row in df_code.iterrows():
            if pd.notna(row.iloc[0]):
                code_map[str(row.iloc[0]).replace(" ","").upper().strip()] = str(row.iloc[target_col_idx]).strip() if pd.notna(row.iloc[target_col_idx]) else ""
        
        if "K" in option:
            sheet_kor = next((s for s in sheet_names if "국문" in s), sheet_names[1] if len(sheet_names) > 1 else sheet_names[0])
            for _, row in pd.read_excel(master_data_path, sheet_name=sheet_kor).iterrows():
                if pd.notna(row.iloc[0]):
                    c = str(row.iloc[0]).replace(" ", "").strip()
                    n = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                    cas_name_map[c] = n
                    if n:
                        kor_data_map[n] = {
                            'F': str(row.iloc[5]) if len(row) > 5 else "", 'G': str(row.iloc[6]) if len(row) > 6 else "", 
                            'H': str(row.iloc[7]) if len(row) > 7 else "", 'P': str(row.iloc[15]) if len(row) > 15 else "", 
                            'T': str(row.iloc[19]) if len(row) > 19 else "", 'U': str(row.iloc[20]) if len(row) > 20 else "", 
                            'V': str(row.iloc[21]) if len(row) > 21 else ""
                        }
        else:
            sheet_eng = next((s for s in sheet_names if "영문" in s), sheet_names[2] if len(sheet_names) > 2 else sheet_names[-1])
            for _, row in pd.read_excel(master_data_path, sheet_name=sheet_eng).iterrows():
                if pd.notna(row.iloc[0]):
                    c = str(row.iloc[0]).replace(" ", "").strip()
                    n = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
                    cas_name_map[c] = n
                    if n:
                        eng_data_map[n] = {
                            'F': str(row.iloc[5]) if len(row) > 5 else "", 'G': str(row.iloc[6]) if len(row) > 6 else "", 
                            'H': str(row.iloc[7]) if len(row) > 7 else "", 'P': str(row.iloc[15]) if len(row) > 15 else "", 
                            'Q': str(row.iloc[16]) if len(row) > 16 else "", 'T': str(row.iloc[19]) if len(row) > 19 else "", 
                            'U': str(row.iloc[20]) if len(row) > 20 else "", 'V': str(row.iloc[21]) if len(row) > 21 else ""
                        }
    except Exception as e: return {"error": f"데이터 로드 오류: {e}"}

    kor_override_data = None
    if option in ["CFF(E)", "HP(E)"] and kor_excel_file is not None:
        kor_excel_file.seek(0)
        kor_ws = load_workbook(io.BytesIO(kor_excel_file.read()), data_only=True).active
        kor_override_data = {"h_codes": [], "p_prev": [], "p_resp": [], "p_stor": [], "p_disp": [], "composition_data": []}
        
        def ext_codes(ws, s_r, e_r, col=2):
            res = []; regex = re.compile(r"([HP]\d{3}(?:\+[HP]\d{3})*)")
            for r in range(s_r, e_r + 1):
                if not ws.row_dimensions[r].hidden and ws.cell(row=r, column=col).value:
                    val_str = str(ws.cell(row=r, column=col).value).strip().upper()
                    if re.match(r'^[HP]\s?\d{3}', val_str):
                        for m in regex.findall(val_str.replace(" ", "")):
                            if m not in res: res.append(m)
            return res

        def ext_comp(ws, s_r, e_r, cas_col=4, conc_col=6):
            res = []; cas_regex = re.compile(r'(\d{2,7}\s*-\s*\d{2}\s*-\s*\d)')
            for r in range(s_r, e_r + 1):
                if not ws.row_dimensions[r].hidden:
                    cas = ws.cell(row=r, column=cas_col).value
                    if cas and str(cas).strip():
                        match = cas_regex.search(str(cas).strip())
                        if match: res.append((match.group(1).replace(" ", ""), str(ws.cell(row=r, column=conc_col).value).strip() if ws.cell(row=r, column=conc_col).value else ""))
            return res

        if "신버전" in kor_form_version:
            kor_override_data["h_codes"] = ext_codes(kor_ws, 25, 36)
            kor_override_data["p_prev"] = ext_codes(kor_ws, 38, 49)
            kor_override_data["p_resp"] = ext_codes(kor_ws, 50, 63)
            kor_override_data["p_stor"] = ext_codes(kor_ws, 64, 69)
            kor_override_data["p_disp"] = ext_codes(kor_ws, 70, 72)
            kor_override_data["composition_data"] = ext_comp(kor_ws, 80, 122)
        else:
            all_c = ext_codes(kor_ws, 25, 70)
            kor_override_data["h_codes"] = [c for c in all_c if c.startswith('H')]
            kor_override_data["p_prev"] = [c for c in all_c if c.startswith('P2')]
            kor_override_data["p_resp"] = [c for c in all_c if c.startswith('P3')]
            kor_override_data["p_stor"] = [c for c in all_c if c.startswith('P4')]
            kor_override_data["p_disp"] = [c for c in all_c if c.startswith('P5')]
            kor_override_data["composition_data"] = ext_comp(kor_ws, 25, 150)

    for uploaded_file in uploaded_files:
        try:
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            parsed_data = parse_pdf_final(doc, mode=option)
            current_prod_name = product_name_input.strip() if product_name_input.strip() else uploaded_file.name.rsplit('.', 1)[0]
            
            if option in ["CFF(E)", "HP(E)"] and kor_override_data:
                for k in ["h_codes", "p_prev", "p_resp", "p_stor", "p_disp", "composition_data"]: parsed_data[k] = kor_override_data[k]
            
            dest_wb = load_workbook(template_path); dest_ws = dest_wb.active; dest_wb.external_links = []
            for row in dest_ws.iter_rows():
                for cell in row:
                    if not isinstance(cell, MergedCell) and cell.column == 2 and cell.data_type == 'f': cell.value = ""

            if option == "HP(E)":
                for addr in ['A50', 'A64', 'A70']: dest_ws[addr].alignment = ALIGN_LEFT
                safe_write_force(dest_ws, 6, 2, current_prod_name, center=True)
                safe_write_force(dest_ws, 9, 2, current_prod_name, center=False)
                if parsed_data["hazard_cls"]:
                    safe_write_force(dest_ws, 19, 2, "\n".join(parsed_data["hazard_cls"]), center=False)
                    dest_ws.row_dimensions[19].height = len(parsed_data["hazard_cls"]) * 14.0
                if parsed_data["signal_word"]: safe_write_force(dest_ws, 23, 2, parsed_data["signal_word"], center=False)
                fill_fixed_range(dest_ws, 24, 36, parsed_data["h_codes"], code_map, mode=option)
                fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map, mode=option)
                fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map, mode=option)
                fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map, mode=option)
                fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map, mode=option)
                fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map, mode=option)
                active_substances = [cas_name_map[c[0].replace(" ", "").strip()] for c in parsed_data["composition_data"] if c[0].replace(" ", "").strip() in cas_name_map and cas_name_map[c[0].replace(" ", "").strip()]]
                
                sd = parsed_data["sec4_to_7"]
                cell_map_e = {
                    "B126": sd.get("B126",""), "B127": sd.get("B127",""), "B128": sd.get("B128",""), "B129": sd.get("B129",""), "B132": sd.get("B132",""),
                    "B134": sd.get("B134",""), "B136": sd.get("B136",""), "B140": sd.get("B140",""), "B142": sd.get("B142",""), "B144": sd.get("B144",""),
                    "B148": sd.get("B148",""), "B150": sd.get("B150",""), "B170": parsed_data["sec9"].get("B170",""), "B176": parsed_data["sec9"].get("B176",""), "B183": parsed_data["sec9"].get("B183","")
                }
                for addr, val in cell_map_e.items():
                    if val:
                        formatted, h = format_and_calc_height_sec47(val, mode=option)
                        r_idx = int(re.search(r'\d+', addr).group())
                        safe_write_force(dest_ws, r_idx, 2, formatted, center=False)
                        dest_ws.row_dimensions[r_idx].height = h
                for r in [170, 176, 183, 189]: dest_ws.row_dimensions[r].height = 18.4

                s8 = parsed_data["sec8"]
                for i, k in enumerate(["B156", "B157", "B158"]):
                    if k in s8:
                        safe_write_force(dest_ws, 156+i, 2, s8[k], center=False)
                        if i > 0: dest_ws.row_dimensions[156+i].hidden = not bool(s8[k])

                for sr, er, ck in [(202, 240, 'F'), (242, 279, 'G'), (281, 315, 'H'), (324, 358, 'P'), (360, 395, 'Q'), (401, 437, 'T'), (439, 478, 'U'), (480, 519, 'V')]:
                    fill_regulatory_section(dest_ws, sr, er, active_substances, eng_data_map, ck, mode=option)
                
                if refractive_index_input: safe_write_force(dest_ws, 189, 2, f"{refractive_index_input.strip()} ± 0.005", center=False)

                s14 = parsed_data["sec14"]
                for i, key in enumerate(["UN", "NAME", "CLASS", "PG", "ENV"]):
                    val = str(s14.get(key, "")).strip()
                    if not val or val.lower() == "not applicable": val = "no data available" if key in ["UN", "NAME"] else "Not applicable"
                    safe_write_force(dest_ws, 531+i, 2, val, center=False)
                safe_write_force(dest_ws, 544, 1, f"16.2 Date of Issue : {datetime.now().strftime('%d. %b. %Y').upper()}", center=False)

            elif option == "CFF(E)":
                for addr in ['A50', 'A64', 'A70']: dest_ws[addr].alignment = ALIGN_LEFT
                safe_write_force(dest_ws, 6, 2, current_prod_name, center=True); safe_write_force(dest_ws, 9, 2, current_prod_name, center=False)
                if parsed_data["hazard_cls"]:
                    safe_write_force(dest_ws, 19, 2, "\n".join(parsed_data["hazard_cls"]), center=False)
                    dest_ws.row_dimensions[19].height = len(parsed_data["hazard_cls"]) * 14.0
                if parsed_data["signal_word"]: safe_write_force(dest_ws, 23, 2, parsed_data["signal_word"], center=False)
                fill_fixed_range(dest_ws, 24, 36, parsed_data["h_codes"], code_map, mode=option)
                fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map, mode=option)
                fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map, mode=option)
                fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map, mode=option)
                fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map, mode=option)
                fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map, mode=option)
                active_substances = [cas_name_map[c[0].replace(" ", "").strip()] for c in parsed_data["composition_data"] if c[0].replace(" ", "").strip() in cas_name_map and cas_name_map[c[0].replace(" ", "").strip()]]
                
                sd = parsed_data["sec4_to_7"]
                cell_map_e = {
                    "B125": sd.get("B125",""), "B126": sd.get("B126",""), "B127": sd.get("B127",""), "B128": sd.get("B128",""), "B129": sd.get("B129",""),
                    "B132": sd.get("B132",""), "B134": sd.get("B134",""), "B136": sd.get("B136",""), "B140": sd.get("B140",""), "B142": sd.get("B142",""),
                    "B144": sd.get("B144",""), "B148": sd.get("B148",""), "B150": sd.get("B150",""),
                    "B170": parsed_data["sec9"].get("B170","").capitalize(), "B176": parsed_data["sec9"].get("B176",""), "B183": parsed_data["sec9"].get("B183",""), "B189": parsed_data["sec9"].get("B189","")
                }
                for addr, val in cell_map_e.items():
                    if val:
                        if addr in ["B183", "B189"] and "±" not in val:
                            num = re.search(r'([\d\.]+)', val)
                            if num: val = f"{num.group(1)} ± {'0.01' if addr == 'B183' else '0.005'}"
                        formatted, h = format_and_calc_height_sec47(val, mode=option)
                        r_idx = int(re.search(r'\d+', addr).group())
                        safe_write_force(dest_ws, r_idx, 2, formatted, center=False)
                        dest_ws.row_dimensions[r_idx].height = h

                s8 = parsed_data["sec8"]
                if s8["B154"]:
                    lines = s8["B154"].split('\n')
                    safe_write_force(dest_ws, 154, 2, lines[0].lower() if "no data" in lines[0].lower() else lines[0], center=False)
                    if len(lines) > 1:
                        safe_write_force(dest_ws, 155, 2, "\n".join(lines[1:]), center=False)
                        dest_ws.row_dimensions[155].hidden = False
                if s8["B156"]:
                    lines = s8["B156"].split('\n')
                    safe_write_force(dest_ws, 156, 2, lines[0].lower() if "no data" in lines[0].lower() else lines[0], center=False)
                    if len(lines) > 1:
                        safe_write_force(dest_ws, 157, 2, "\n".join(lines[1:]), center=False)
                        dest_ws.row_dimensions[157].hidden = False

                for sr, er, ck in [(202, 240, 'F'), (242, 279, 'G'), (281, 315, 'H'), (324, 358, 'P'), (360, 395, 'Q'), (401, 437, 'T'), (439, 478, 'U'), (480, 519, 'V')]:
                    fill_regulatory_section(dest_ws, sr, er, active_substances, eng_data_map, ck, mode=option)
                
                if refractive_index_input: safe_write_force(dest_ws, 189, 2, f"{refractive_index_input.strip()} ± 0.005", center=False)

                s14 = parsed_data["sec14"]
                for i, key in enumerate(["UN", "NAME", "CLASS", "PG", "ENV"]):
                    val = str(s14.get(key, "")).strip()
                    if not val or val.lower() == "not applicable": val = "no data available" if key in ["UN", "NAME"] else "Not applicable"
                    safe_write_force(dest_ws, 531+i, 2, val, center=False)
                safe_write_force(dest_ws, 544, 1, f"16.2 Date of Issue : {datetime.now().strftime('%d. %b. %Y').upper()}", center=False)

            else: # CFF(K) / HP(K)
                safe_write_force(dest_ws, 7, 2, current_prod_name, center=True); safe_write_force(dest_ws, 10, 2, current_prod_name, center=False)
                if parsed_data["hazard_cls"]:
                    clean_hazard_text = "\n".join([line for line in parsed_data["hazard_cls"] if line.strip()])
                    safe_write_force(dest_ws, 20, 2, clean_hazard_text, center=False)
                    dest_ws['B20'].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
                    valid_lines_count = len([line for line in parsed_data["hazard_cls"] if line.strip()])
                    if valid_lines_count > 0: dest_ws.row_dimensions[20].height = valid_lines_count * 14.0

                safe_write_force(dest_ws, 24, 2, parsed_data["signal_word"] if parsed_data["signal_word"] else "", center=False) 
                if option == "HP(K)":
                    for r, v in [(38, "예방"), (50, "대응"), (64, "저장"), (70, "폐기")]: safe_write_force(dest_ws, r, 1, v, center=False)

                fill_fixed_range(dest_ws, 25, 36, parsed_data["h_codes"], code_map, mode=option)
                fill_fixed_range(dest_ws, 38, 49, parsed_data["p_prev"], code_map, mode=option)
                fill_fixed_range(dest_ws, 50, 63, parsed_data["p_resp"], code_map, mode=option)
                fill_fixed_range(dest_ws, 64, 69, parsed_data["p_stor"], code_map, mode=option)
                fill_fixed_range(dest_ws, 70, 72, parsed_data["p_disp"], code_map, mode=option)
                fill_composition_data(dest_ws, parsed_data["composition_data"], cas_name_map, mode=option)
                
                active_substances = [cas_name_map[c[0].replace(" ", "").strip()] for c in parsed_data["composition_data"] if c[0].replace(" ", "").strip() in cas_name_map and cas_name_map[c[0].replace(" ", "").strip()]]

                for cell_addr, raw_text in parsed_data["sec4_to_7"].items():
                    formatted_txt, row_h = format_and_calc_height_sec47(raw_text, mode=option)
                    try:
                        col_idx = openpyxl.utils.column_index_from_string(re.match(r"([A-Z]+)", cell_addr).group(1))
                        row_num = int(re.search(r"(\d+)", cell_addr).group(1))
                        safe_write_force(dest_ws, row_num, col_idx, "")
                        if formatted_txt:
                            safe_write_force(dest_ws, row_num, col_idx, formatted_txt, center=False)
                            dest_ws.row_dimensions[row_num].height = row_h
                            try:
                                cell_a = dest_ws.cell(row=row_num, column=1)
                                if cell_a.value: cell_a.value = str(cell_a.value).strip()
                                cell_a.alignment = ALIGN_TITLE
                            except: pass
                    except: pass

                s8 = parsed_data["sec8"]
                val148 = s8["B148"].replace("해당없음", "자료없음")
                lines148 = [l.strip() for l in val148.split('\n') if l.strip()]
                safe_write_force(dest_ws, 148, 2, ""); safe_write_force(dest_ws, 149, 2, ""); dest_ws.row_dimensions[149].hidden = True
                if lines148:
                    safe_write_force(dest_ws, 148, 2, lines148[0], center=False)
                    if len(lines148) > 1:
                        safe_write_force(dest_ws, 149, 2, "\n".join(lines148[1:]), center=False)
                        dest_ws.row_dimensions[149].hidden = False
                safe_write_force(dest_ws, 150, 2, re.sub(r"^규정[:\s]*", "", s8["B150"].replace("해당없음", "자료없음")).strip(), center=False)

                s9 = parsed_data["sec9"]
                safe_write_force(dest_ws, 163, 2, s9["B163"], center=False)
                flash_num = re.findall(r'([<>]?\s*\d{2,3})' if option == "HP(K)" else r'(\d{2,3})', s9["B169"])
                safe_write_force(dest_ws, 169, 2, f"{flash_num[0]}℃" if flash_num else "", center=False)
                g_match = re.search(r'([\d\.]+)', s9["B176"].replace("(20℃)", "").replace("(물=1)", ""))
                safe_write_force(dest_ws, 176, 2, f"{g_match.group(1)} ± 0.01" if g_match else "", center=False)
                
                if option == "HP(K)" and refractive_index_input: safe_write_force(dest_ws, 182, 2, f"{refractive_index_input.strip()} ± 0.005", center=False)
                else:
                    r_match = re.search(r'([\d\.]+)', s9["B182"].replace("(20℃)", ""))
                    safe_write_force(dest_ws, 182, 2, f"{r_match.group(1)} ± 0.005" if r_match else "", center=False)

                for sr, er, ck in [(195, 226, 'F'), (228, 260, 'G'), (269, 300, 'H'), (316, 348, 'P'), (353, 385, 'P'), (392, 426, 'T'), (428, 460, 'U'), (465, 497, 'V')]:
                    fill_regulatory_section(dest_ws, sr, er, active_substances, kor_data_map, ck, mode=option)

                for r in list(range(261, 268)) + list(range(349, 352)) + [386] + list(range(461, 464)): dest_ws.row_dimensions[r].hidden = True

                s14 = parsed_data["sec14"]
                un_raw = str(s14.get("UN", "")).strip(); un_val = re.sub(r"\D", "", un_raw)
                if not un_val or "해당없음" in un_raw: un_val = "자료없음"
                name_raw = str(s14.get("NAME", "")).strip(); name_val = re.sub(r"\([^)]*\)", "", name_raw).strip()
                if not name_val or "해당없음" in name_raw: name_val = "자료없음"
                class_val = str(s14.get("CLASS", "")).strip(); pg_val = str(s14.get("PG", "")).strip(); env_val = str(s14.get("ENV", "")).strip()
                
                for i, v in enumerate([un_val, name_val, class_val if class_val and "해당없음" not in class_val else "해당없음", pg_val if pg_val and "해당없음" not in pg_val else "해당없음", env_val if env_val and "해당없음" not in env_val else "해당없음"]):
                    safe_write_force(dest_ws, 512+i, 2, v, center=False)

                s15 = parsed_data["sec15"]
                if option == "HP(K)":
                    clean_full = re.sub(r'[\s\-\,\.\:\(\)]+', '', s15.get("FULL_TEXT", ""))
                    if ("4류" in clean_full and "3석유류" in clean_full and "2000" in clean_full):
                        safe_write_force(dest_ws, 521, 2, "4류 제3석유류(비수용성) 2,000L", center=False)
                    else:
                        safe_write_force(dest_ws, 521, 2, "", center=False)
                        dest_ws.cell(row=521, column=2).fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                else: safe_write_force(dest_ws, 521, 2, s15.get("DANGER", "").strip(), center=False)

                safe_write_force(dest_ws, 542, 2, datetime.now().strftime("%Y.%m.%d"), center=False)

            # 이미지 삽입
            collected_pil_images = []
            for img_info in doc.get_page_images(0):
                try:
                    pil_img = PILImage.open(io.BytesIO(doc.extract_image(img_info[0])["image"]))
                    if option in ["HP(K)", "HP(E)"] and (is_blue_dominant(pil_img) or not is_square_shaped(pil_img.size[0], pil_img.size[1])): continue
                    if loaded_refs:
                        matched_name = find_best_match_name(pil_img, loaded_refs, mode=option)
                        if matched_name: collected_pil_images.append((extract_number(matched_name), loaded_refs[matched_name]))
                except: continue
            
            unique_images = {}
            for key, img in collected_pil_images:
                if key not in unique_images: unique_images[key] = img
            final_sorted_imgs = [item[1] for item in sorted(unique_images.items(), key=lambda x: x[0])]

            if final_sorted_imgs:
                unit_size = 67; icon_size = 60; padding_top = 4; padding_left = (unit_size - icon_size) // 2
                merged_img = PILImage.new('RGBA', (unit_size * len(final_sorted_imgs), unit_size), (255, 255, 255, 0))
                for idx, p_img in enumerate(final_sorted_imgs):
                    merged_img.paste(p_img.resize((icon_size, icon_size), PILImage.LANCZOS), ((idx * unit_size) + padding_left, padding_top))
                img_byte_arr = io.BytesIO()
                merged_img.save(img_byte_arr, format='PNG')
                img_byte_arr.seek(0)
                dest_ws.add_image(XLImage(img_byte_arr), 'B23' if option in ["HP(K)", "CFF(K)"] else 'B22') 

            dest_wb.external_links = []
            output = io.BytesIO()
            dest_wb.save(output)
            output.seek(0)
            
            final_name = f"{current_prod_name} GHS MSDS({'E' if 'E' in option else 'K'}).xlsx"
            if final_name in new_download_data: final_name = f"{current_prod_name}_{uploaded_file.name.split('.')[0]}.xlsx"
            new_download_data[final_name] = output.getvalue()
            new_files.append(final_name)
            
        except Exception as e:
            return {"error": f"파일 처리 중 오류 발생 ({uploaded_file.name}): {e}"}

    return {"files": new_files, "data": new_download_data}

# ==============================================================================
# [5번: OTHERS 변환기 로직]
# ==============================================================================
def process_others(customer_name, product_name):
    kst = timezone(timedelta(hours=9))
    current_time = datetime.now(kst)
    english_months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    current_date = f"{english_months[current_time.month - 1]} {current_time.strftime('%d, %Y')}"
    
    template_dir = get_resource_path("OTHERS templates")
    if not os.path.exists(template_dir): return None, f"'{template_dir}' 폴더를 찾을 수 없습니다."
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename in os.listdir(template_dir):
            if filename.endswith(".docx") and not filename.startswith("~"):
                file_path = os.path.join(template_dir, filename)
                try:
                    doc = DocxTemplate(file_path)
                    doc.render({'DATE': current_date, 'CUSTOMER': customer_name, 'PRODUCT': product_name})
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    zip_file.writestr(filename.replace("STH", product_name), doc_io.read())
                except Exception:
                    pass
    zip_buffer.seek(0)
    return zip_buffer, f"Documents_{customer_name}_{product_name}.zip"

# ==============================================================================
# [UI 레이아웃 구성]
# ==============================================================================
st.title("📄 통합 문서 양식 자동 변환기")

# --- 공통 정보 및 일괄 변환 ---
st.subheader("공통 정보 입력")
col_top1, col_top2, col_top3 = st.columns([2, 2, 1])
with col_top1:
    global_customer = st.text_input("고객사명 (CUSTOMER)")
with col_top2:
    global_product = st.text_input("제품명 (PRODUCT)")
with col_top3:
    st.write("")
    batch_run = st.button("🌟 일괄 변환 실행", use_container_width=True)
    include_others = st.checkbox("일괄 변환 시 OTHERS 포함", value=True)

st.divider()

# --- Section 1: SPEC ---
st.subheader("1. SPEC 양식 변환")
col1_1, col1_2, col1_3 = st.columns(3)
with col1_1:
    spec_up = st.file_uploader("원본 PDF 업로드", type=["pdf"], key="spec_up")
with col1_2:
    spec_mode = st.selectbox("모드 선택", ["CFF", "HP"], key="spec_mode")
    if st.button("SPEC 변환", use_container_width=True):
        if not spec_up or not global_product: st.warning("원본 파일과 제품명을 입력해주세요.")
        else:
            with st.spinner("SPEC 변환 중..."):
                try:
                    res, fname = process_spec(spec_up, global_product, spec_mode)
                    st.session_state['spec_res'] = res.getvalue()
                    st.session_state['spec_fname'] = fname
                    st.success("변환 성공!")
                except Exception as e: st.error(f"오류: {e}")
with col1_3:
    if st.session_state['spec_res']:
        st.download_button("📥 결과물 다운로드", data=st.session_state['spec_res'], file_name=st.session_state['spec_fname'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_spec")

st.divider()

# --- Section 2: ALLERGY ---
st.subheader("2. ALLERGY 양식 변환")
col2_1, col2_2, col2_3 = st.columns(3)
with col2_1:
    allergy_up = st.file_uploader("원본 Excel 업로드", type=['xlsx', 'xls'], key="allergy_up")
with col2_2:
    allergy_mode = st.selectbox("업체 타입 선택", ["CFF", "HP"], key="allergy_mode")
    if st.button("ALLERGY 변환", use_container_width=True):
        if not allergy_up or not global_customer or not global_product: st.warning("원본 파일, 고객사명, 제품명을 모두 입력해주세요.")
        else:
            with st.spinner("ALLERGY 변환 중..."):
                try:
                    input_df = pd.read_excel(allergy_up)
                    base_path = get_resource_path("ALLERGY templates")
                    if allergy_mode == "CFF":
                        res_83 = logic_cff_83(input_df, os.path.join(base_path, "83 CFF.xlsx"), global_customer, global_product)
                        res_26 = logic_cff_26(input_df, os.path.join(base_path, "26 통합.xlsx"), global_customer, global_product)
                    else:
                        res_83 = logic_hp_83(input_df, os.path.join(base_path, "83 HP.xlsx"), global_customer, global_product)
                        res_26 = logic_hp_26(input_df, os.path.join(base_path, "26 통합.xlsx"), global_customer, global_product)
                    st.session_state['allergy_res_83'] = to_excel(res_83)
                    st.session_state['allergy_res_26'] = to_excel(res_26)
                    st.session_state['allergy_fname_83'] = f"83 ALLERGENS {global_product}.xlsx"
                    st.session_state['allergy_fname_26'] = f"ALLERGEN {global_product}.xlsx"
                    st.success("변환 성공!")
                except Exception as e: st.error(f"오류: {e}")
with col2_3:
    if st.session_state['allergy_res_83'] and st.session_state['allergy_res_26']:
        st.download_button(f"📥 {st.session_state['allergy_fname_83']} 다운로드", data=st.session_state['allergy_res_83'], file_name=st.session_state['allergy_fname_83'], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_al_83")
        st.download_button(f"📥 {st.session_state['allergy_fname_26']} 다운로드", data=st.session_state['allergy_res_26'], file_name=st.session_state['allergy_fname_26'], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="dl_al_26")

st.divider()

# --- Section 3: IFRA ---
st.subheader("3. IFRA 양식 변환")
col3_1, col3_2, col3_3 = st.columns(3)
with col3_1:
    ifra_up = st.file_uploader("원본 PDF 업로드", type=["pdf"], key="ifra_up")
with col3_2:
    ifra_mode = st.selectbox("모드 선택", ["CFF", "HP"], key="ifra_mode")
    if st.button("IFRA 변환", use_container_width=True):
        if not ifra_up or not global_customer or not global_product: st.warning("원본 파일, 고객사명, 제품명을 모두 입력해주세요.")
        else:
            with st.spinner("IFRA 변환 중..."):
                try:
                    res, fname = process_ifra(ifra_up, global_customer, global_product, ifra_mode)
                    st.session_state['ifra_res'] = res.getvalue()
                    st.session_state['ifra_fname'] = fname
                    st.success("변환 성공!")
                except Exception as e: st.error(f"오류: {e}")
with col3_3:
    if st.session_state['ifra_res']:
        st.download_button("📥 결과물 다운로드", data=st.session_state['ifra_res'], file_name=st.session_state['ifra_fname'], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True, key="dl_ifra")

st.divider()

# --- Section 4: MSDS ---
st.subheader("4. MSDS 양식 변환")
col4_1, col4_2, col4_3 = st.columns(3)
with col4_1:
    msds_up = st.file_uploader("원본 PDF 업로드 (다중 선택 가능)", type=["pdf"], accept_multiple_files=True, key="msds_up")
with col4_2:
    msds_mode = st.selectbox("모드 선택", ["CFF(K)", "CFF(E)", "HP(K)", "HP(E)"], key="msds_mode")
    msds_ri = ""
    msds_kor_file = None
    msds_kor_ver = "신버전"
    
    if "HP" in msds_mode:
        msds_ri = st.text_input("굴절률 입력", key="msds_ri")
    if "E" in msds_mode:
        st.info("💡 영문 양식 생성 시 국문 파일 첨부")
        msds_kor_file = st.file_uploader("국문 엑셀 파일", type="xlsx", key="msds_kor_file")
        msds_kor_ver = st.radio("국문 양식 버전", ["신버전", "구버전"], key="msds_kor_ver")

    if st.button("MSDS 변환", use_container_width=True):
        if not msds_up: st.warning("원본 파일을 하나 이상 업로드해주세요.")
        else:
            with st.spinner("MSDS 변환 중..."):
                res_dict = process_msds(msds_up, global_product, msds_mode, msds_ri, msds_kor_file, msds_kor_ver)
                if "error" in res_dict:
                    st.error(res_dict["error"])
                else:
                    st.session_state['msds_res'] = [{'fname': f, 'data': res_dict["data"][f]} for f in res_dict["files"]]
                    st.success("변환 성공!")
with col4_3:
    if st.session_state['msds_res']:
        for i, item in enumerate(st.session_state['msds_res']):
            st.download_button(f"📥 {item['fname']} 다운로드", data=item['data'], file_name=item['fname'], mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key=f"dl_msds_{i}")

st.divider()

# --- Section 5: OTHERS ---
st.subheader("5. 원본 불필요 파일 (OTHERS) 변환")
col5_1, col5_2, col5_3 = st.columns(3)
with col5_1:
    st.info("이 양식은 원본 파일 첨부가 필요 없으며, 입력한 고객사명과 제품명으로 templates 폴더 내 모든 파일이 일괄 생성됩니다.")
with col5_2:
    if st.button("OTHERS 변환", use_container_width=True):
        if not global_customer or not global_product: st.warning("상단의 고객사명과 제품명을 모두 입력해주세요.")
        else:
            with st.spinner("OTHERS 변환 중..."):
                res, info = process_others(global_customer, global_product)
                if res:
                    st.session_state['others_res'] = res.getvalue()
                    st.session_state['others_fname'] = info
                    st.success("변환 성공!")
                else:
                    st.error(info)
with col5_3:
    if st.session_state['others_res']:
        st.download_button("📥 ZIP 파일 다운로드", data=st.session_state['others_res'], file_name=st.session_state['others_fname'], mime="application/zip", use_container_width=True, key="dl_others")

# ==============================================================================
# [일괄 변환 트리거]
# ==============================================================================
if batch_run:
    st.write("---")
    if not global_customer or not global_product:
        st.error("❗ 일괄 변환을 위해 상단의 [고객사명]과 [제품명]을 먼저 입력해주세요.")
    else:
        with st.spinner("🌟 일괄 변환을 진행 중입니다... 잠시만 기다려주세요."):
            # 1. SPEC
            if spec_up:
                try:
                    res, fname = process_spec(spec_up, global_product, spec_mode)
                    st.session_state['spec_res'] = res.getvalue(); st.session_state['spec_fname'] = fname
                except Exception as e: st.error(f"SPEC 일괄 변환 오류: {e}")
            
            # 2. ALLERGY
            if allergy_up:
                try:
                    input_df = pd.read_excel(allergy_up)
                    base_path = get_resource_path("ALLERGY templates")
                    if allergy_mode == "CFF":
                        r83 = logic_cff_83(input_df, os.path.join(base_path, "83 CFF.xlsx"), global_customer, global_product)
                        r26 = logic_cff_26(input_df, os.path.join(base_path, "26 통합.xlsx"), global_customer, global_product)
                    else:
                        r83 = logic_hp_83(input_df, os.path.join(base_path, "83 HP.xlsx"), global_customer, global_product)
                        r26 = logic_hp_26(input_df, os.path.join(base_path, "26 통합.xlsx"), global_customer, global_product)
                    st.session_state['allergy_res_83'] = to_excel(r83)
                    st.session_state['allergy_res_26'] = to_excel(r26)
                    st.session_state['allergy_fname_83'] = f"83 ALLERGENS {global_product}.xlsx"
                    st.session_state['allergy_fname_26'] = f"ALLERGEN {global_product}.xlsx"
                except Exception as e: st.error(f"ALLERGY 일괄 변환 오류: {e}")

            # 3. IFRA
            if ifra_up:
                try:
                    res, fname = process_ifra(ifra_up, global_customer, global_product, ifra_mode)
                    st.session_state['ifra_res'] = res.getvalue(); st.session_state['ifra_fname'] = fname
                except Exception as e: st.error(f"IFRA 일괄 변환 오류: {e}")
                
            # 4. MSDS
            if msds_up:
                res_dict = process_msds(msds_up, global_product, msds_mode, msds_ri, msds_kor_file, msds_kor_ver)
                if "error" in res_dict: st.error(f"MSDS 일괄 변환 오류: {res_dict['error']}")
                else: st.session_state['msds_res'] = [{'fname': f, 'data': res_dict["data"][f]} for f in res_dict["files"]]

            # 5. OTHERS
            if include_others:
                res, info = process_others(global_customer, global_product)
                if res:
                    st.session_state['others_res'] = res.getvalue(); st.session_state['others_fname'] = info
                else: st.error(f"OTHERS 일괄 변환 오류: {info}")
            
            st.success("✅ 파일이 첨부된 모든 항목의 일괄 변환이 완료되었습니다. 각 섹션의 우측에서 결과를 다운로드하세요!")
            st.rerun() # UI 리프레시를 위해 재실행
