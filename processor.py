import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook

def safe_str(x):
    return str(x).lower().strip() if pd.notna(x) else ""

def process_boQ(template_path, lms_paths, output_path):

    wb = load_workbook(template_path)
    ws = wb.active

    for idx, file in enumerate(lms_paths):

        try:
            df = pd.read_excel(file)
        except:
            continue

        items = {}

        for _, row in df.iterrows():
            item = safe_str(row[0])
            qty = row[1] if len(row) > 1 else 0

            if item and pd.notna(qty):
                items[item] = items.get(item, 0) + qty

        col = 6 + (idx * 2)

        for r in range(6, ws.max_row + 1):

            temp_item = safe_str(ws.cell(r,2).value)
            harga = ws.cell(r,4).value or 0

            best = None
            best_score = 0

            for k in items:
                score = fuzz.token_sort_ratio(temp_item, k)
                if score > best_score and score > 75:
                    best = k
                    best_score = score

            if best:
                qty = items[best]
                ws.cell(r,col).value = qty
                ws.cell(r,col+1).value = qty * harga

    wb.save(output_path)
