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
from selenium.webdriver.support.ui import WebDriverWait
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
        tk.Label(root, text="股號:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
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
        
        self.cb_price = tk.Checkbutton(self.cb_frame, text="股價 (Price)", variable=self.price_var)
        self.cb_price.pack(anchor="w")
        self.cb_pbr = tk.Checkbutton(self.cb_frame, text="本淨比 (PBR)", variable=self.pbr_var)
        self.cb_pbr.pack(anchor="w")
        self.cb_per = tk.Checkbutton(self.cb_frame, text="本益比 (PER)", variable=self.per_var)
        self.cb_per.pack(anchor="w")
        
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
        stock_id = self.stock_id_entry.get().strip()
        if not stock_id:
            messagebox.showwarning("警告", "請輸入股號！")
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
            if max_delay < 20:
                messagebox.showwarning("警告", "最大冷卻時間至少需 20 秒以上！")
                return
        except ValueError:
            messagebox.showwarning("警告", "最大冷卻時間必須是有效的整數！")
            return
            
        # 收集需要跑的項目
        tasks = []
        if self.price_var.get():
            tasks.append("Price")
        if self.pbr_var.get():
            tasks.append("PBR")
        if self.per_var.get():
            tasks.append("PER")
            
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
        self.log("="*30)
        self.log(f"啟動自動化區間多項目爬蟲任務 (任務目標: 往前 {year_limit} 年)...")
        self.log(f"最大冷卻預設間隔: 20 ~ {max_delay} 秒")
        self.log(f"勾選的項目: {', '.join(tasks)}")
        self.log("="*30)
        
        threading.Thread(target=self.run_scraper, args=(stock_id, year_limit, max_delay, tasks), daemon=True).start()

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
        """合併資料夾內所有 CSV，去重並依日期由新到舊排序"""
        self.log(f"開始掃描並合併資料夾內的 CSV: {target_dir}")
        # 將舊有的 _Result.csv 也納入合併對象
        all_files = [f for f in os.listdir(target_dir) if f.endswith('.csv')]
        if not all_files:
            self.log("[警告] 沒有找到可合併的 CSV 檔案。")
            return
            
        header = None
        data_dict = {}
        for filename in all_files:
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
                        # 相同日期若已存在，直接覆蓋 (保留一筆)
                        data_dict[date_val] = row
                except StopIteration:
                    pass
                    
        # 日期格式為 'yy/mm/dd，字串降序排序剛好符合新到舊
        sorted_dates = sorted(data_dict.keys(), reverse=True)
        today_str = datetime.now().strftime("%Y-%m-%d")
        output_filename = f"{today_str}_{task.upper()}_Result.csv" if task else f"{today_str}_Result.csv"
        output_filepath = os.path.join(target_dir, output_filename)
        
        with open(output_filepath, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(header)
            for d in sorted_dates:
                writer.writerow(data_dict[d])
                
        # 加入: 刪除已被合併的碎檔以及舊有的 Result.csv
        for filename in all_files:
            if filename != output_filename:
                try:
                    os.remove(os.path.join(target_dir, filename))
                except Exception as e:
                    self.log(f"[警告] 移除已合併檔案 {filename} 時發生錯誤: {e}")
                
        self.log(f"✅ 合併完成！總共彙整 {len(sorted_dates)} 筆不重複資料，已儲存至: {output_filename}，且已清除舊檔/暫存檔。")

    def run_scraper(self, stock_id, year_limit, max_delay, tasks):
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
                "profile.default_content_setting_values.automatic_downloads": 1
            }
            options.add_experimental_option("prefs", prefs)
            
            self.log("正在啟動瀏覽器核心 (WebDriver)...")
            service = Service(ChromeDriverManager().install())
            driver = webdriver.Chrome(service=service, options=options)
            
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            driver.minimize_window() # 最小化視窗，不影響使用者工作
            wait = WebDriverWait(driver, 15)
            
            # 依照勾選的項目依序執行
            for task in tasks:
                self.log(f"\n==========================================")
                self.log(f"🌟 正在執行新項目：【 {task} 】 🌟")
                self.log(f"==========================================")
                
                # 建立該項目的專屬資料夾
                task_dir = os.path.join(base_download_dir, task)
                os.makedirs(task_dir, exist_ok=True)
                self.log(f"設定 【{task}】 下載目的地為: {task_dir}")
                
                # 最強絕招：利用執行 Chrome DevTools Protocol (CDP) 瞬間改變下載路徑
                # 不須重啟 Chrome 即可對應 Price, PBR, PER 資料夾存放
                driver.execute_cdp_cmd('Page.setDownloadBehavior', {
                    'behavior': 'allow',
                    'downloadPath': task_dir
                })
                
                # 分派該項目的入口網址
                if task == "Price":
                    base_url = f"https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID={stock_id}&CHT_CAT=DATE&PRICE_ADJ=F"
                elif task == "PBR":
                    base_url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PBR&STOCK_ID={stock_id}&CHT_CAT=DATE"
                elif task == "PER":
                    base_url = f"https://goodinfo.tw/tw/ShowK_ChartFlow.asp?RPT_CAT=PER&STOCK_ID={stock_id}&CHT_CAT=DATE"
                else:
                    continue
                    
                self.log(f"【{task}】導航至入口網址: {base_url}")
                driver.get(base_url)
                time.sleep(2) # 導航初次暖機
                
                curr_end_str = datetime.now().strftime("%Y-%m-%d")
                curr_start_str = None
                
                # 檢查是否有既有的 Result.csv (讀取最新一筆日期)
                existing_files = [f for f in os.listdir(task_dir) if f.endswith('_Result.csv')]
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
                                # 解析 "YY/MM/DD" 或 "YYYY/MM/DD"
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
                            # 嘗試點擊標準的查詢按鈕
                            query_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='button' and @value='查詢']")))
                            driver.execute_script("arguments[0].click();", query_btn)
                        except:
                            self.log("尋找「查詢」按鈕超時，嘗試使用備案腳本提交表單...")
                            # 備案: 直接戳 btnQUERY，或是送出表單
                            driver.execute_script("var btn = document.getElementById('btnQUERY'); if(btn) btn.click(); else document.forms[0].submit();")
                        
                        time.sleep(5)
                        
                        div_details = driver.find_elements(By.ID, "divDetailBox")
                        if div_details and "查無相關資料" in div_details[0].text:
                            self.log("💡 [提示] 此區間內查無新的交易紀錄。")
                        else:
                            xls_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='XLS']")))
                            files_before = set(os.listdir(task_dir))
                            driver.execute_script("arguments[0].click();", xls_btn)
                            self.wait_for_downloads(task_dir)
                            
                            files_after = set(os.listdir(task_dir))
                            new_files = list(files_after - files_before)
                            if new_files:
                                xls_filename = new_files[0]
                                xls_path = os.path.join(task_dir, xls_filename)
                                base_name = os.path.splitext(xls_filename)[0]
                                csv_filename = f"{curr_end_str}_{curr_start_str}_{base_name}.csv"
                                csv_path = os.path.join(task_dir, csv_filename)
                                self.convert_xls_to_csv(xls_path, csv_path)
                    except Exception as e:
                        self.log(f"❌ 區間更新發生錯誤: {e}")
                        
                    self.log(f"【{task}】資料更新下載完畢，進行合併排序作業...")
                    self.merge_csv_files(task_dir, task)
                    continue  # 此項目的更新完成，進入下一個 task
                
                self.log(f"【{task}】未發現歷史紀錄，將執行原始的多年度回推抓取作業。")
                
                # 開始自動推時域的無盡迴圈 (全新抓取模式)
                iteration = 1
                
                while True:
                    if iteration > year_limit:
                        self.log(f"✅ 【{task}】已順利達到您設定的抓取上限 ({year_limit} 年)，將在此終止，不繼續往前追溯。")
                        self.update_status(f"💡 【{task}】已滿 {year_limit} 年歷史資料 💡")
                        break

                    self.update_status(f"【{task}】正在設定結束時間: {curr_end_str}...")
                    self.log(f"\n>>>> [項目: {task} | 第 {iteration} 次資料查找] 準備區間結尾: {curr_end_str} <<<<")
                    
                    # 鎖定網頁裡面的結束日期輸入框...
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
                    
                    # 從網頁上抓回系統自動推算出來的起始時間
                    self.log("讀取系統自動推算完畢的起始時間...")
                    start_input = wait.until(EC.presence_of_element_located((By.ID, "edtSTART_TIME")))
                    curr_start_str = driver.execute_script("return arguments[0].value;", start_input)
                    
                    range_text = f"{curr_start_str} ~ {curr_end_str}"
                    self.update_status(f"【{task}】實際抓取區間: {range_text}")  # 更新 UI 給使用者看
                    self.log(f"系統實際抓取的區間為: {range_text}")
                    
                    # 判斷網頁是否顯示「查無相關資料!!」
                    self.log("解析網頁內文元素，判定資料盡頭...")
                    div_details = driver.find_elements(By.ID, "divDetailBox")
                    if div_details and "查無相關資料" in div_details[0].text:
                        self.log("🚨 [偵測停止訊號] 網頁文字包含「查無相關資料!!」")
                        self.log(f"🚨 說明這檔股票的歷史掛牌資料已撈取殆盡，結束【{task}】的迴圈流程。")
                        self.update_status(f"💡 【{task}】全數歷史資料已達極限 💡")
                        break
                    
                    self.log("確認具備表格資料，鎖定「XLS」下載按鈕...")
                    xls_btn = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@value='XLS']")))
                    
                    # 非常重要: 觀察【目前項目】的獨立目錄以捕捉新增的檔案數
                    files_before = set(os.listdir(task_dir))
                    files_before_len = len(files_before)
                    
                    self.log("點此觸發原生的檔案下載作業...")
                    driver.execute_script("arguments[0].click();", xls_btn)
                    
                    # 安全等待機制 (精準比對資料夾新增檔案數量)
                    self.wait_for_downloads(task_dir, files_before_len)
                    
                    # 揪出新生的假 XLS 檔案
                    files_after = set(os.listdir(task_dir))
                    new_files = list(files_after - files_before)
                    
                    if new_files:
                        xls_filename = new_files[0]
                        self.log(f"捕捉到實體生出的檔案: {xls_filename}")
                        xls_path = os.path.join(task_dir, xls_filename)
                        
                        # (ex: 2026-03-26_2025-03-26_原本的檔名.csv)
                        base_name = os.path.splitext(xls_filename)[0]
                        csv_filename = f"{curr_end_str}_{curr_start_str}_{base_name}.csv"
                        csv_path = os.path.join(task_dir, csv_filename)
                        
                        self.convert_xls_to_csv(xls_path, csv_path)
                    else:
                        self.log("[警告] 點擊後未發現任何目錄新增檔案，可能當下阻擋或被系統隔離！")
                    
                    # 將這一次系統算出來的起始時間，直接繼承為下一次的結束時間！
                    curr_end_str = curr_start_str
                    iteration += 1
                    
                    sleep_time = random.randint(20, max_delay)
                    self.log(f"✅ 【{task}】單次區間作業漂亮完成，將冷卻 {sleep_time} 秒鐘以避免發出過多網路請求被伺服器封鎖...")
                    time.sleep(sleep_time)
                    
                # 該項目所有年度抓取完畢 (while 迴圈結束)，執行合併作業
                self.log(f"【{task}】區間下載已完成，開始進行合併與去重排序作業...")
                self.merge_csv_files(task_dir, task)
                
                    
            self.log("🎉🎉 所有勾選項目的自動化區間作業，均已順利執行完畢！🎉🎉")
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
            self.log("系統已就緒釋放，可隨時嘗試新的股號與項目選項。\n" + "-"*50)

if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperApp(root)
    root.mainloop()
