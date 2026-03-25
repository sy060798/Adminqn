import pandas as pd
from rapidfuzz import fuzz
from openpyxl import load_workbook

def smart_match(a, b):
    return fuzz.token_sort_ratio(a, b)

def process_boQ(template_path, lms_paths, output_path):

    wb = load_workbook(template_path)
    ws = wb.active

    for idx, file in enumerate(lms_paths):

        df = pd.read_excel(file)
        items = {}

        for _, row in df.iterrows():
            item = str(row[0]).lower()
            qty = row[1]

            if pd.notna(item) and pd.notna(qty):
                items[item] = items.get(item, 0) + qty

        col = 6 + (idx * 2)

        for r in range(6, ws.max_row):

            temp_item = str(ws.cell(r,2).value).lower()
            harga = ws.cell(r,4).value or 0

            best = None
            best_score = 0

            for k in items:
                score = smart_match(temp_item, k)
                if score > best_score and score > 75:
                    best = k
                    best_score = score

            if best:
                qty = items[best]
                ws.cell(r,col).value = qty
                ws.cell(r,col+1).value = qty * harga

    wb.save(output_path)
