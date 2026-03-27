import os
import csv
import re
from bs4 import BeautifulSoup
import codecs
from datetime import datetime

def convert_xls_to_csv_test(xls_path, csv_path):
    print(f"Processing {xls_path}...")
    base_filename = os.path.basename(xls_path)
    
    # 支援新舊副合檔名模式
    date_match = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", base_filename)
    if date_match:
        current_year = int(date_match.group(1))
        last_month = int(date_match.group(2))
        print(f"  Parsed End Date from filename -> Year: {current_year}, Month: {last_month}")
    else:
        current_year = datetime.now().year
        last_month = 12
        print("  Could not parse date from filename. Used default fallback.")

    with codecs.open(xls_path, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()

    soup = BeautifulSoup(content, 'html.parser')
    table = soup.find('table')
    if not table:
        print("  [警告] 找不到表格結構")
        return False

    rows = table.find_all('tr')
    matrix = []
    for r_idx, row in enumerate(rows):
        while len(matrix) <= r_idx:
            matrix.append([])
        cols = row.find_all(['th', 'td'])
        c_idx = 0
        for col in cols:
            while c_idx < len(matrix[r_idx]) and matrix[r_idx][c_idx] is not None:
                c_idx += 1
            text = col.get_text(strip=True)
            colspan = int(col.get('colspan', 1))
            rowspan = int(col.get('rowspan', 1))
            for i in range(rowspan):
                while len(matrix) <= r_idx + i:
                    matrix.append([])
                for j in range(colspan):
                    while len(matrix[r_idx + i]) <= c_idx + j:
                        matrix[r_idx + i].append(None)
                    matrix[r_idx + i][c_idx + j] = text
            c_idx += colspan

    header_depth = 1
    if rows:
        first_row_cols = rows[0].find_all(['th', 'td'])
        for col in first_row_cols:
            if int(col.get('colspan', 1)) > 1:
                header_depth = 2
                break

    flat_headers = []
    if header_depth == 2 and len(matrix) > 1:
        for i in range(len(matrix[0])):
            h1 = matrix[0][i] if matrix[0][i] else ""
            h2 = matrix[1][i] if len(matrix[1]) > i and matrix[1][i] else ""
            if h1 == h2 or not h2:
                flat_headers.append(h1)
            else:
                flat_headers.append(f"{h1}({h2})")
    else:
        flat_headers = matrix[0] if matrix else []

    with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(flat_headers)
        
        debug_printed = 0
        for row_data in matrix[header_depth:]:
            if not row_data:
                continue
            if row_data[0] == "交易日期" or row_data[0] == flat_headers[0]:
                continue
                
            date_str = row_data[0].strip() if row_data[0] else ""
            old_str = date_str
            
            # 使用正則表達式匹配 MM/DD 格式 (例如 03/26)
            match = re.match(r"^(\d{2})/(\d{2})$", date_str)
            if match:
                month = int(match.group(1))
                day = int(match.group(2))
                
                if month > last_month and (month - last_month >= 6):
                    current_year -= 1
                    
                last_month = month
                
                yy = str(current_year)[-2:]
                mm = str(month).zfill(2)
                dd = str(day).zfill(2)
                row_data[0] = f"'{yy}/{mm}/{dd}"
            
            if debug_printed < 3:
                print(f"    Raw: '{old_str}' -> Parsed: '{row_data[0]}'")
                debug_printed += 1

            writer.writerow(row_data)
            
    print(f"  Successfully wrote CSV: {csv_path}\n")

if __name__ == "__main__":
    convert_xls_to_csv_test("2026_03_27-2025_03_28_K_ChartFlow.xls", "2026_03_27-2025_03_28_K_ChartFlow.csv")
    convert_xls_to_csv_test("2025_03_28-2024_03_29_K_ChartFlow.xls", "2025_03_28-2024_03_29_K_ChartFlow.csv")
