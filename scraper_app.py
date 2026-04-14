import os
import sys

# 確保在 PyInstaller --noconsole (--windowed) 模式下不會因為 print() 找不到輸出管道而崩潰
if sys.stdout is None:
    sys.stdout = open(os.devnull, 'w', encoding='utf-8')
if sys.stderr is None:
    sys.stderr = open(os.devnull, 'w', encoding='utf-8')

import subprocess
import threading
import time
import random
import tkinter as tk
from tkinter import scrolledtext, messagebox
from datetime import datetime, timedelta
import csv

# 1. 環境自我診斷與安裝 (Environment Initialization)
def install_and_import(package, import_name=None):
    if import_name is None:
        import_name = package
    try:
        __import__(import_name)
    except ImportError:
        print(f"正在安裝 {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        __import__(import_name)

# 確保相依套件已安裝
install_and_import('selenium')
install_and_import('webdriver-manager', 'webdriver_manager')
install_and_import('beautifulsoup4', 'bs4')

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Goodinfo 多項目自動區間爬蟲工具")
        self.root.geometry("550x550")
        
        # UI Components
        # 股號輸入
        tk.Label(root, text="股號 (多筆以分號隔開):").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.stock_id_entry = tk.Entry(root)
        self.stock_id_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # 抓取年份上限設定
        tk.Label(root, text="抓取年數上限:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.year_limit_entry = tk.Entry(root)
        self.year_limit_entry.insert(0, "5") # 預設只抓 5 年
        self.year_limit_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # 抓取時間間隔最大值設定
        tk.Label(root, text="最大冷卻時間(秒):").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.max_delay_entry = tk.Entry(root)
        self.max_delay_entry.insert(0, "90") # 預設只抓 90 秒
        self.max_delay_entry.grid(row=2, column=1, padx=10, pady=10, sticky="w")

        # 擷取項目選擇 (Checkboxes)
        tk.Label(root, text="擷取項目:").grid(row=3, column=0, padx=10, pady=5, sticky="ne")
        self.cb_frame = tk.Frame(root)
        self.cb_frame.grid(row=3, column=1, sticky="w", pady=5)
        
        self.price_var = tk.BooleanVar(value=False)
        self.pbr_var = tk.BooleanVar(value=True)
        self.per_var = tk.BooleanVar(value=False)
        self.inst_var = tk.BooleanVar(value=False)
        self.bias_var = tk.BooleanVar(value=False)
        self.finance_var = tk.BooleanVar(value=False)
        
        self.cb_price = tk.Checkbutton(self.cb_frame, text="股價 (Price)", variable=self.price_var)
        self.cb_price.pack(anchor="w")
        self.cb_pbr = tk.Checkbutton(self.cb_frame, text="本淨比 (PBR)", variable=self.pbr_var)
        self.cb_pbr.pack(anchor="w")
        self.cb_per = tk.Checkbutton(self.cb_frame, text="本益比 (PER)", variable=self.per_var)
        self.cb_per.pack(anchor="w")
        self.cb_inst = tk.Checkbutton(self.cb_frame, text="法人買賣", variable=self.inst_var)
        self.cb_inst.pack(anchor="w")
        self.cb_bias = tk.Checkbutton(self.cb_frame, text="乖離率", variable=self.bias_var)
        self.cb_bias.pack(anchor="w")
        self.cb_finance = tk.Checkbutton(self.cb_frame, text="財報", variable=self.finance_var)
        self.cb_finance.pack(anchor="w")
        
        # 狀態提示標籤 (UI 提示使用者現在作業的期間)
        self.status_lbl = tk.Label(root, text="目前作業期間：尚未開始", fg="blue", font=("Arial", 11, "bold"))
        self.status_lbl.grid(row=4, column=0, columnspan=2, pady=5)

        # 執行按鈕
        self.start_btn = tk.Button(root, text="開始執行自動推迴圈爬蟲", command=self.start_scraping_thread, bg="#4CAF50", fg="white", font=("Arial", 12))
        self.start_btn.grid(row=5, column=0, columnspan=2, pady=10, ipadx=20)
        
        # 日誌區
        self.log_area = scrolledtext.ScrolledText(root, width=65, height=14, state='disabled', bg="#f0f0f0")
        self.log_area.grid(row=6, column=0, columnspan=2, padx=10, pady=10)

    def log(self, message):
        """將訊息輸出到唯讀的日誌區，也同步列印在 Console"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        full_message = f"[{timestamp}] {message}"
        print(full_message, flush=True)
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, full_message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        # 強制更新 UI
        self.root.update()

    def update_status(self, text):
        self.status_lbl.config(text=f"目前作業期間：{text}")
        self.root.update()

    def start_scraping_thread(self):
        """啟動背景執行緒來跑爬蟲，避免把 UI 執行緒卡死"""
        stock_ids_raw = self.stock_id_entry.get()
        stock_ids = [s.strip() for s in stock_ids_raw.split(';') if s.strip()]
        
        if not stock_ids:
            messagebox.showwarning("警告", "請輸入至少一個股號！")
            return
            
        year_limit_str = self.year_limit_entry.get().strip()
        try:
            year_limit = int(year_limit_str)
            if year_limit <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("警告", "抓取年份數必須是有效的正整數！")
            return
            
        max_delay_str = self.max_delay_entry.get().strip()
        try:
            max_delay = int(max_delay_str)
            if max_delay < 60:
                messagebox.showwarning("警告", "最大冷卻時間至少需 60 秒以上！")
                return
        except ValueError:
            messagebox.showwarning("警告", "最大冷卻時間必須是有效的整數！")
            return
            
        # 收集需要跑的項目
        tasks = []
        if self.price_var.get():
            tasks.append("個股K線")
        if self.pbr_var.get():
            tasks.append("本淨比")
        if self.per_var.get():
            tasks.append("本益比")
        if self.inst_var.get():
            tasks.append("法人買賣")
        if self.bias_var.get():
            tasks.append("乖離率")
        if self.finance_var.get():
            tasks.append("財報")
            
        if not tasks:
            messagebox.showwarning("警告", "請至少勾選一個抓取項目！")
            return
            
        self.start_btn.config(state='disabled')
        self.stock_id_entry.config(state='disabled')
        self.year_limit_entry.config(state='disabled')
        self.max_delay_entry.config(state='disabled')
        self.cb_price.config(state='disabled')
        self.cb_pbr.config(state='disabled')
        self.cb_per.config(state='disabled')
        self.cb_inst.config(state='disabled')
        self.cb_bias.config(state='disabled')
        self.cb_finance.config(state='disabled')
        self.log("="*30)
        self.log(f"啟動自動化區間多項目爬蟲任務 (任務目標: 往前 {year_limit} 年)...")
        self.log(f"最大冷卻預設間隔: 20 ~ {max_delay} 秒")
        self.log(f"批次追蹤股號清單: {', '.join(stock_ids)}")
        self.log(f"勾選的項目: {', '.join(tasks)}")
        self.log("="*30)
        
        threading.Thread(target=self.run_scraper, args=(stock_ids, year_limit, max_delay, tasks), daemon=True).start()

    def wait_for_downloads(self, download_dir, files_before_len, timeout=60):
        """等待檔案開始創建且完整下載至資料夾"""
        self.log("等待檔案下載安全寫入完成...")
        seconds = 0
        while seconds < timeout:
            time.sleep(1)
            files = os.listdir(download_dir)
            is_downloading = any(f.endswith('.crdownload') or f.endswith('.tmp') for f in files)
            
            if len(files) > files_before_len and not is_downloading:
                self.log("檔案下載安全寫入完成！")
                return True
                
            seconds += 1
            
        self.log("[警告] 下載等待逾時，檔案可能未完全寫入或網路延遲過高無法下載。")
        return False

    def convert_xls_to_csv(self, xls_path, csv_path):
        """將 Goodinfo 的 HTML Table 解析轉換為標準 CSV，完美修復 colspan/rowspan 表頭錯位，並刪除原檔"""
        self.log(f"實作二維矩陣演算法，對齊 HTML 表格欄位並生成純淨 CSV...")
        try:
            import codecs
            import re
            
            # 從檔名提取當前這批資料的結束年份與月份，供後續判定跨年 (檔名支援多種分隔符)
            base_filename = os.path.basename(csv_path)
            date_match = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", base_filename)
            if date_match:
                current_year = int(date_match.group(1))
                last_month = int(date_match.group(2))
            else:
                current_year = datetime.now().year
                last_month = 12

            with codecs.open(xls_path, "r", encoding="utf-8", errors="ignore") as f:
                content = f.read()
            
            soup = BeautifulSoup(content, 'html.parser')
            table = soup.find('table')
            if not table:
                self.log("[警告] 找不到表格結構，直接改檔名保留。")
                os.rename(xls_path, csv_path)
                return False

            # 將 HTML 表格依照 rowspan/colspan 轉換為真實對齊的 2D 矩陣
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
                    # 抓取跨欄或跨列屬性
                    colspan = int(col.get('colspan', 1))
                    rowspan = int(col.get('rowspan', 1))
                    
                    # 將這塊積木填滿到它涵蓋的範圍裡
                    for i in range(rowspan):
                        while len(matrix) <= r_idx + i:
                            matrix.append([])
                        for j in range(colspan):
                            while len(matrix[r_idx + i]) <= c_idx + j:
                                matrix[r_idx + i].append(None)
                            matrix[r_idx + i][c_idx + j] = text
                    c_idx += colspan

            # 判斷表頭層數 (如果第一列有 colspan > 1, 通常蘊含兩層表頭)
            header_depth = 1
            if rows:
                first_row_cols = rows[0].find_all(['th', 'td'])
                for col in first_row_cols:
                    if int(col.get('colspan', 1)) > 1:
                        header_depth = 2
                        break

            # 重組出平坦化的單一層完美表頭
            flat_headers = []
            if header_depth == 2 and len(matrix) > 1:
                for i in range(len(matrix[0])):
                    h1 = matrix[0][i] if matrix[0][i] else ""
                    h2 = matrix[1][i] if len(matrix[1]) > i and matrix[1][i] else ""
                    # 若上下層名字一樣，或下層為空，就取大標題
                    if h1 == h2 or not h2:
                        flat_headers.append(h1)
                    else:
                        # 將大標題跟小標題合併變成單一獨立標題 (例如: 本淨比換算價格(1x))
                        flat_headers.append(f"{h1}({h2})")
            else:
                flat_headers = matrix[0] if matrix else []

            # 寫出為標準相容的 CSV，並自帶萬能的 utf-8-sig (BOM) 給 Excel 用
            with open(csv_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(flat_headers)
                
                # 從實際數據行開寫 (自動過濾掉原網頁中每隔 50 列重複一次的夾層表頭)
                for row_data in matrix[header_depth:]:
                    if not row_data:
                        continue
                    if row_data[0] == "交易日期" or row_data[0] == flat_headers[0]:
                        continue
                        
                    # 處理日期格式 (將本淨比/本益比原始的 "MM/DD" 換算為統一的 "'YY/MM/DD")
                    date_str = row_data[0].strip() if row_data[0] else ""
                    match = re.match(r"^(\d{2})/(\d{2})$", date_str)
                    if match:
                        month = int(match.group(1))
                        day = int(match.group(2))
                        
                        # 檢查是否往前跨越年度 (例如資料從 1 月回到前一年的 12 月)
                        if month > last_month and (month - last_month >= 6):
                            current_year -= 1
                            
                        last_month = month
                        
                        yy = str(current_year)[-2:]
                        mm = str(month).zfill(2)
                        dd = str(day).zfill(2)
                        row_data[0] = f"'{yy}/{mm}/{dd}"

                    writer.writerow(row_data)

            self.log(f"成功轉換寫出完美對齊的 CSV : {os.path.basename(csv_path)}")
            # 功成身退，刪除原始假 XLS 檔
            os.remove(xls_path)
            self.log(f"已刪除帶有結構瑕疵的原始假 XLS 檔: {os.path.basename(xls_path)}")
            return True

        except Exception as e:
            self.log(f"[錯誤] CSV 表列二維重整過程發生異常: {str(e)}")
            return False

    def merge_csv_files(self, target_dir, task=""):
        """依選項分組合併資料夾內所有 CSV，去重並依日期由新到舊排序"""
        self.log(f"【{task}】開始歸納並分組合併資料夾內的 CSV: {target_dir}")
        all_files = [f for f in os.listdir(target_dir) if f.endswith('.csv')]
        if not all_files:
            self.log("[警告] 沒有找到可合併的 CSV 檔案。")
            return
            
        groups = {}
        for filename in all_files:
            # 檔名格式: {option_text}_{start_str}_{end_str}.csv 或 {option_text}_{end_str}_Result.csv
            parts = filename.replace('.csv', '').rsplit('_', 2)
            if len(parts) >= 3:
                group_name = parts[0]
            else:
                group_name = task # fallback
                
            if group_name not in groups:
                groups[group_name] = []
            groups[group_name].append(filename)
            
        today_str = datetime.now().strftime("%Y-%m-%d")
        
        for group_name, files in groups.items():
            header = None
            data_dict = {}
            for filename in files:
                filepath = os.path.join(target_dir, filename)
                with open(filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    try:
                        current_header = next(reader)
                        if header is None:
                            header = current_header
                        for row in reader:
                            if not row:
                                continue
                            date_val = row[0]
                            if date_val == "交易日期" or date_val == header[0]:
                                continue
                            data_dict[date_val] = row
                    except StopIteration:
                        pass
                        
            sorted_dates = sorted(data_dict.keys(), reverse=True)
            # 新規則: {option顯示的文字}_{結束時間}_Result.csv
            output_filename = f"{group_name}_{today_str}_Result.csv"
            output_filepath = os.path.join(target_dir, output_filename)
            
            with open(output_filepath, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(header)
                for d in sorted_dates:
                    writer.writerow(data_dict[d])
                    
            for filename in files:
                if filename != output_filename:
                    try:
                        os.remove(os.path.join(target_dir, filename))
                    except Exception as e:
                        self.log(f"[警告] 移除已合併檔案 {filename} 時發生錯誤: {e}")
                        
            self.log(f"✅ 群組 [{group_name}] 合併完成！保留 {len(sorted_dates)} 筆資料並清除切片。儲存為: {output_filename}")

    def download_xls_and_convert(self, driver, wait, task_dir, curr_start_str, curr_end_str, task):
        # 尋找是否具備下拉選單 selKCSheet
        select_elements = driver.find_elements(By.ID, "selKCSheet")
        if select_elements:
            total_options = len(Select(select_elements[0]).options)
            for i in range(total_options):
                # 重新抓取避免 DOM 失效
                sel_elem = wait.until(EC.presence_of_element_located((By.ID, "selKCSheet")))
                dropdown = Select(sel_elem)
                opt = dropdown.options[i]
                
                # 判斷是否可用
                if opt.is_enabled() and opt.get_attribute('disabled') is None:
                    opt_text = opt.text
                    self.log(f"【{task}】自動切換至選項: {opt_text}")
                    dropdown.select_by_index(i)
                    time.sleep(3) # Wait for Ajax loading
                    
                    div_details = driver.find_elements(By.ID, "divDetailBox")
                    if div_details and "查無相關資料" in div_details[0].text:
                        self.log(f"【{task}】選項 {opt_text} 查無資料，自動跳過。")
                        continue
                        
                    self._execute_single_download(driver, wait, task_dir, curr_start_str, curr_end_str, opt_text)
        else:
            self._execute_single_download(driver, wait, task_dir, curr_start_str, curr_end_str, task)

    def _execute_single_download(self, driver, wait, task_dir, curr_start_str, curr_end_str, opt_text):
        self.log(f"等待點擊下載 XLS 與轉檔 ({opt_text})...")
        xls_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='XLS']")))
        files_before = set(os.listdir(task_dir))
        driver.execute_script("arguments[0].click();", xls_btn)
        self.wait_for_downloads(task_dir, len(files_before))
        
        files_after = set(os.listdir(task_dir))
        new_files = list(files_after - files_before)
        if new_files:
            xls_filename = new_files[0]
            xls_path = os.path.join(task_dir, xls_filename)
            # 全新命名規則: {option顯示的文字}_{起始時間}_{結束時間}.csv
            csv_filename = f"{opt_text}_{curr_start_str}_{curr_end_str}.csv"
            csv_path = os.path.join(task_dir, csv_filename)
            self.convert_xls_to_csv(xls_path, csv_path)

    def _process_financial_report_scraper(self, driver, wait, stock_id, base_download_dir, year_limit, max_delay):
        task = "財報"
        task_dir = os.path.join(base_download_dir, stock_id, task)
        os.makedirs(task_dir, exist_ok=True)
        self.log(f"\n==========================================")
        self.log(f"🌟 正在執行專屬項目：【 {task} 】 🌟")
        self.log(f"==========================================")
        self.log(f"設定 【{task}】 下載目的地為: {task_dir}")
        self.update_status(f"【{task}】正在鎖定目標擷取區間...")
        
        # 設定 Chrome 下載行為給指定的 task_dir
        driver.execute_cdp_cmd('Page.setDownloadBehavior', {
            'behavior': 'allow',
            'downloadPath': task_dir
        })
        
        current_year_roc = datetime.now().year - 1911
        start_year = current_year_roc
        end_year = current_year_roc - year_limit + 1
        if end_year < 1:
            end_year = 1
            
        quarter_map = {"第一季": 1, "第二季": 2, "第三季": 3, "第四季": 4}
        
        for roc_year in range(start_year, end_year - 1, -1):
            self.log(f"【{task}】開始檢查 {roc_year} 年度財報...")
            
            # 先列出該年度需要抓取的所有季別，並提前逐一檢查檔案是否已存在
            missing_quarters = []
            for q_num in range(1, 5):
                expected_filename = f"財報_{roc_year}年_第{q_num}季.pdf"
                expected_filepath = os.path.join(task_dir, expected_filename)
                if os.path.exists(expected_filepath):
                    self.log(f"【{task}】 {expected_filename} 已存在，略過。")
                else:
                    missing_quarters.append(q_num)
            
            # 若所有可能的檔案(1~4季)都不需要下載，在此判斷直接不執行存取網頁動作
            if not missing_quarters:
                self.log(f"【{task}】 {roc_year} 年度財報已下載齊全，略過網頁存取。")
                continue
                
            self.update_status(f"【{task}】探索中: {roc_year} 年度")
            url = f"https://doc.twse.com.tw/server-java/t57sb01?step=1&colorchg=1&co_id={stock_id}&year={roc_year}&seamon=&mtype=A&"
            
            try:
                driver.get(url)
                time.sleep(2)
            except Exception as e:
                self.log(f"【{task}】讀取 {roc_year} 年度頁面失敗: {e}")
                continue
                
            try:
                # 檢查是否包含「查無所需資料」
                empty_msgs = driver.find_elements(By.XPATH, "//font[contains(text(), '查無所需資料')] | //h4[contains(text(), '查無所需資料')]")
                if empty_msgs:
                    self.log(f"【{task}】 {roc_year} 年度查無所需資料，略過。")
                    continue
            except Exception:
                pass
                
            try:
                rows = driver.find_elements(By.XPATH, "//tr[td]")
            except Exception as e:
                self.log(f"【{task}】解析 {roc_year} 年度表格失敗: {e}")
                continue
                
            download_targets = []
            for r_index in range(len(rows)):
                try:
                    # 重新尋找元素避免 stale element
                    current_rows = driver.find_elements(By.XPATH, "//tr[td]")
                    if r_index >= len(current_rows):
                        break
                    row = current_rows[r_index]
                    tds = row.find_elements(By.TAG_NAME, "td")
                    if len(tds) >= 8:
                        year_quarter = tds[1].text
                        desc = tds[5].text.strip()
                        
                        # 嚴格過濾只找合併資料(包含「合併財報」或「合併報表」)並剃除「英文版」和「個體」
                        if ("合併財報" in desc or "合併報表" in desc) and "英文版" not in desc and "個體" not in desc:
                            q_num = None
                            for q_text, num in quarter_map.items():
                                if q_text in year_quarter:
                                    q_num = num
                                    break
                                    
                            # 只把確定缺失 (在 missing_quarters 中) 的季度加入下載佇列
                            if q_num and (q_num in missing_quarters):
                                links = tds[7].find_elements(By.TAG_NAME, "a")
                                if links:
                                    download_targets.append({
                                        "elem_idx": r_index,
                                        "year": roc_year,
                                        "quarter": q_num
                                    })
                except Exception as e:
                    pass  # ignore row parsing errors
                    
            if not download_targets:
                self.log(f"【{task}】 {roc_year} 年度未發現符合條件的「合併財報」。")
                continue
                
            original_window = driver.current_window_handle
            
            for target in download_targets:
                expected_filename = f"財報_{target['year']}年_第{target['quarter']}季.pdf"
                expected_filepath = os.path.join(task_dir, expected_filename)
                
                self.log(f"【{task}】準備下載: {expected_filename}")
                self.update_status(f"【{task}】下載中: {expected_filename}")
                
                try:
                    current_rows = driver.find_elements(By.XPATH, "//tr[td]")
                    row = current_rows[target["elem_idx"]]
                    tds = row.find_elements(By.TAG_NAME, "td")
                    a_elem = tds[7].find_elements(By.TAG_NAME, "a")[0]
                    
                    # 點擊開啟新視窗
                    driver.execute_script("arguments[0].click();", a_elem)
                    time.sleep(2)
                    
                    new_windows = [w for w in driver.window_handles if w != original_window]
                    if not new_windows:
                        self.log(f"【{task}】點擊後未發現新分頁 ({expected_filename})")
                        continue
                        
                    driver.switch_to.window(new_windows[0])
                    
                    # 偵測是否被 TWSE 判定為「下載過量」
                    time.sleep(1) # 等待頁面短暫渲染
                    page_text = driver.page_source
                    if "下載過量" in page_text:
                        self.log(f"【{task}】🚨 遭到 TWSE 流量管制 (下載過量)！為保護您的 IP，將取消此檔案 ({expected_filename}) 的本次下載並強制暫停 30 秒。")
                        time.sleep(30)
                        raise Exception("TWSE_RATE_LIMIT (下載過量)")
                            
                    # 在新分頁找實際 PDF 連結
                    pdf_link = wait.until(EC.presence_of_element_located((By.XPATH, "//a[contains(translate(@href, 'PDF', 'pdf'), '.pdf')] | //a[contains(text(), '.pdf')]")))
                    
                    files_before = set(os.listdir(task_dir))
                    driver.execute_script("arguments[0].click();", pdf_link)
                    
                    # 等待下載完成
                    success = self.wait_for_downloads(task_dir, len(files_before), timeout=45)
                    
                    if success:
                        files_after = set(os.listdir(task_dir))
                        new_files = list(files_after - files_before)
                        if new_files:
                            # 找出字尾是 pdf 的檔案
                            pdf_files = [f for f in new_files if f.lower().endswith('.pdf')]
                            if pdf_files:
                                downloaded_file = os.path.join(task_dir, pdf_files[0])
                                if os.path.exists(expected_filepath):
                                    os.remove(expected_filepath)
                                os.rename(downloaded_file, expected_filepath)
                                self.log(f"【{task}】✅ 成功儲存: {expected_filename}")
                            else:
                                self.log(f"【{task}】未能辨識出下載的 PDF 檔案。")
                    else:
                        self.log(f"【{task}】下載超時或失敗: {expected_filename}")
                        
                except Exception as e:
                    self.log(f"【{task}】下載過程異常 ({expected_filename})，如果是超時可能是遭到封鎖: {e}")
                    if "driver" in locals() and len(driver.window_handles) > 1:
                        # 印出當下錯誤畫面的前 100 字幫助除錯
                        self.log(f"網頁內容 snippet: {driver.page_source[:100].replace(chr(10), ' ')}")
                finally:
                    # 無論成功失敗，都關閉新視窗並切回主視窗
                    try:
                        for window in driver.window_handles:
                            if window != original_window:
                                driver.switch_to.window(window)
                                driver.close()
                        driver.switch_to.window(original_window)
                    except Exception:
                        pass
                    
                    # 非常重要：在同一年度的不同季報之間，也必須插入冷卻時間避免被抓！
                    intra_sleep = random.randint(10, max_delay)
                    self.log(f"【{task}】為避免連續下載被鎖，套用最大冷卻時間機制，等待 {intra_sleep} 秒...")
                    time.sleep(intra_sleep)
            
        self.log(f"【{task}】 {stock_id} 所有指定年度爬取完畢！")

    def _process_single_stock(self, driver, wait, stock_id, base_download_dir, year_limit, max_delay, tasks):
        # 依照勾選的項目依序執行
        for task in tasks:
            if task == "財報":
                self._process_financial_report_scraper(driver, wait, stock_id, base_download_dir, year_limit, max_delay)
                continue
                
            self.log(f"\n==========================================")
            self.log(f"🌟 正在執行新項目：【 {task} 】 🌟")
            self.log(f"==========================================")
            
            # 建立該項目的專屬資料夾
            task_dir = os.path.join(base_download_dir, stock_id, task)
            os.makedirs(task_dir, exist_ok=True)
            self.log(f"設定 【{task}】 下載目的地為: {task_dir}")
            
            # 最強絕招：利用執行 Chrome DevTools Protocol (CDP) 瞬間改變下載路徑
            driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                'behavior': 'allow',
                'downloadPath': task_dir
            })
            
            # 分派該項目的入口網址
            if task == "個股K線":
                base_url = f"https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID={stock_id}&CHT_CAT=DATE"
            elif task == "本淨比":
                base_url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PBR&STOCK_ID={stock_id}&CHT_CAT=DATE"
            elif task == "本益比":
                base_url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PER&STOCK_ID={stock_id}&CHT_CAT=DATE"
            elif task == "法人買賣":
                base_url = f"https://goodinfo.tw/tw/ShowBuySaleChart.asp?STOCK_ID={stock_id}&CHT_CAT=DATE"
            elif task == "乖離率":
                base_url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=DR_3M&STOCK_ID={stock_id}&CHT_CAT=DATE"
            else:
                continue
                
            self.log(f"【{task}】導航至入口網址: {base_url}")
            driver.get(base_url)
            time.sleep(2) # 導航初次暖機
            
            curr_end_str = datetime.now().strftime("%Y-%m-%d")
            curr_start_str = None
            
            # 檢查是否有既有的 Result.csv (讀取最新一筆日期)
            existing_files = []
            for f in os.listdir(task_dir):
                if task == "法人買賣":
                    if re.match(r"^\d{4}-\d{2}-\d{2}_", f) is None and f.endswith(".csv"):
                        existing_files.append(f)
                else:
                    if f.endswith('_Result.csv'):
                        existing_files.append(f)

            if existing_files:
                existing_files.sort(reverse=True)
                latest_csv = os.path.join(task_dir, existing_files[0])
                try:
                    import csv
                    import re
                    with open(latest_csv, 'r', encoding='utf-8-sig') as f:
                        reader = csv.reader(f)
                        header = next(reader, None) # 略過 header
                        first_row = next(reader, None)
                        if first_row and first_row[0]:
                            date_str = first_row[0].replace("'", "").strip()
                            match = re.match(r"(\d+)[-/](\d+)[-/](\d+)", date_str)
                            if match:
                                y, m, d = match.groups()
                                y = int(y)
                                y = 2000 + y if y < 50 else (1900 + y if y < 100 else y)
                                curr_start_str = f"{y:04d}-{int(m):02d}-{int(d):02d}"
                except Exception as e:
                    self.log(f"【{task}】讀取舊有 Result.csv 失敗: {e}")
            
            if curr_start_str:
                self.log(f"【{task}】讀取到 Result.csv 最新資料日期為: {curr_start_str}")
                if curr_start_str == curr_end_str:
                    self.log(f"💡 【{task}】最新資料日期與今日 ({curr_end_str}) 相同，不需要重複抓取。")
                    self.update_status(f"【{task}】資料已是最新，跳過。")
                    continue
                    
                self.log(f"💡 發現舊紀錄，將自動設定起始日期為: {curr_start_str}，結束日期為當日: {curr_end_str} 進行更新抓取。")
                self.update_status(f"【{task}】正在更新區間: {curr_start_str} ~ {curr_end_str}")
                
                try:
                    start_input = wait.until(EC.presence_of_element_located((By.ID, "edtSTART_TIME")))
                    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", start_input, curr_start_str)
                    
                    end_input = wait.until(EC.presence_of_element_located((By.ID, "edtEND_TIME")))
                    driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", end_input, curr_end_str)
                    time.sleep(1)
                    
                    self.log("點擊「查詢」按鈕，獲取最新區間資料...")
                    try:
                        query_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='button' and @value='查詢']")))
                        driver.execute_script("arguments[0].click();", query_btn)
                    except:
                        self.log("尋找「查詢」按鈕超時，嘗試使用備案腳本提交表單...")
                        driver.execute_script("var btn = document.getElementById('btnQUERY'); if(btn) btn.click(); else document.forms[0].submit();")
                    
                    time.sleep(5)
                    
                    div_details = driver.find_elements(By.ID, "divDetailBox")
                    if div_details and "查無相關資料" in div_details[0].text:
                        self.log("💡 [提示] 此區間內查無新的交易紀錄。")
                    else:
                        self.download_xls_and_convert(driver, wait, task_dir, curr_start_str, curr_end_str, task)
                except Exception as e:
                    self.log(f"❌ 區間更新發生錯誤: {e}")
                    
                self.log(f"【{task}】資料更新下載完畢，進行合併排序作業...")
                self.merge_csv_files(task_dir, task)
                continue  # 此項目的更新完成，進入下一個 task
            
            self.log(f"【{task}】未發現歷史紀錄，將執行原始的多年度回推抓取作業。")
            
            iteration = 1
            while True:
                if iteration > year_limit:
                    self.log(f"✅ 【{task}】已順利達到您設定的抓取上限 ({year_limit} 年)，將在此終止，不繼續往前追溯。")
                    self.update_status(f"💡 【{task}】已滿 {year_limit} 年歷史資料 💡")
                    break

                self.update_status(f"【{task}】正在設定結束時間: {curr_end_str}...")
                self.log(f"\n>>>> [項目: {task} | 第 {iteration} 次資料查找] 準備區間結尾: {curr_end_str} <<<<")
                
                self.log("鎖定網頁裡面的結束日期輸入框...")
                end_input = wait.until(EC.presence_of_element_located((By.ID, "edtEND_TIME")))
                
                self.log(f"植入結束時間 '{curr_end_str}' ...")
                driver.execute_script("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('change'));", end_input, curr_end_str)
                time.sleep(1)
                
                self.log("點擊「查1年」按鈕，觸發系統自動推前一年的 Ajax...")
                query_year_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='查1年']")))
                driver.execute_script("arguments[0].click();", query_year_btn)
                
                self.log("暫停 5 秒讓資料庫有充足時間回應並渲染回前端...")
                time.sleep(5)
                
                self.log("讀取系統自動推算完畢的起始時間...")
                start_input = wait.until(EC.presence_of_element_located((By.ID, "edtSTART_TIME")))
                curr_start_str = driver.execute_script("return arguments[0].value;", start_input)
                
                range_text = f"{curr_start_str} ~ {curr_end_str}"
                self.update_status(f"【{task}】實際抓取區間: {range_text}")  
                self.log(f"系統實際抓取的區間為: {range_text}")
                
                self.log("解析網頁內文元素，判定資料盡頭...")
                div_details = driver.find_elements(By.ID, "divDetailBox")
                if div_details and "查無相關資料" in div_details[0].text:
                    self.log("🚨 [偵測停止訊號] 網頁文字包含「查無相關資料!!」")
                    self.log(f"🚨 說明這檔股票的歷史掛牌資料已撈取殆盡，結束【{task}】的迴圈流程。")
                    self.update_status(f"💡 【{task}】全數歷史資料已達極限 💡")
                    break
                
                self.log("確認具備表格資料，開始執行下拉選單爬取或常規下載...")
                self.download_xls_and_convert(driver, wait, task_dir, curr_start_str, curr_end_str, task)
                
                curr_end_str = curr_start_str
                iteration += 1
                
                sleep_time = random.randint(20, max_delay)
                self.log(f"✅ 【{task}】單次區間作業漂亮完成，將冷卻 {sleep_time} 秒鐘以避免發出過多網路請求被伺服器封鎖...")
                time.sleep(sleep_time)
                
            self.log(f"【{task}】區間下載已完成，開始進行合併與去重排序作業...")
            self.merge_csv_files(task_dir, task)

    def run_scraper(self, stock_ids, year_limit, max_delay, tasks):
        driver = None
        try:
            self.log("正在進行環境檢查與瀏覽器設定...")
            
            # 使用 sys.frozen 來判斷是否被 Pyinstaller 打包，避免把檔案抓進暫存資料夾
            if getattr(sys, 'frozen', False):
                current_dir = os.path.dirname(sys.executable)
            else:
                current_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else os.getcwd()
            base_download_dir = os.path.join(current_dir, "Download_Data")
            
            options = Options()
            options.add_argument("--disable-blink-features=AutomationControlled")
            options.add_experimental_option("excludeSwitches", ["enable-automation"])
            options.add_experimental_option('useAutomationExtension', False)
            
            # 使用初始的 Download Directory 打開 ChromeDriver
            prefs = {
                "download.default_directory": base_download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_setting_values.automatic_downloads": 1,
                "plugins.always_open_pdf_externally": True
            }
            options.add_experimental_option("prefs", prefs)
            
            self.log("正在啟動瀏覽器核心 (WebDriver)...")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.minimize_window() # 最小化視窗，不影響使用者工作
            wait = WebDriverWait(driver, 15)
            
            for stock_id in stock_ids:
                try:
                    self.log(f"\n{'='*50}")
                    self.log(f"🚀 開始執行批次股號: 【{stock_id}】 🚀")
                    self.log(f"{'='*50}")
                    self._process_single_stock(driver, wait, stock_id, base_download_dir, year_limit, max_delay, tasks)
                except Exception as e:
                    import traceback
                    tb = traceback.format_exc()
                    self.log(f"❌ 股號 【{stock_id}】 處理時發生異常，自動跳轉至下一檔股票:\n{tb}")
                    
            self.log("🎉🎉 全任務清單：所有股號自動化作業均已順利執行完畢！🎉🎉")
            self.update_status("💡 滿載而歸！全部工作皆已完成 💡")
            
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.log(f"❌ 執行期間發生未預料深層錯誤:\n{tb}")
            self.update_status("❌ 執行中斷 (發生未預知錯誤)")
            
        except BaseException as e:
            import traceback
            tb = traceback.format_exc()
            self.log(f"💥 捕捉到底層中斷錯誤 (BaseException):\n{tb}")
            self.update_status("❌ 執行強制中斷")
            
        finally:
            if driver:
                self.log("正在關機卸載瀏覽器資源...")
                driver.quit()
            
            self.start_btn.config(state='normal')
            self.stock_id_entry.config(state='normal')
            self.year_limit_entry.config(state='normal')
            self.max_delay_entry.config(state='normal')
            self.cb_price.config(state='normal')
            self.cb_pbr.config(state='normal')
            self.cb_per.config(state='normal')
            self.cb_inst.config(state='normal')
            self.cb_bias.config(state='normal')
            self.cb_finance.config(state='normal')
            self.log("系統已就緒釋放，可隨時嘗試新的股號與項目選項。\n" + "-"*50)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperApp(root)
    root.mainloop()
