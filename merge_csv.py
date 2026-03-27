import os
import csv
from datetime import datetime

def merge_csv_files(target_dir):
    print(f"開始合併資料夾: {target_dir}")
    if not os.path.exists(target_dir):
        print("資料夾不存在!")
        return
        
    all_files = [f for f in os.listdir(target_dir) if f.endswith('.csv') and not f.endswith('_Result.csv')]
    
    if not all_files:
        print("沒有找到可合併的 CSV 檔案。")
        return
        
    header = None
    data_dict = {}
    
    for filename in all_files:
        filepath = os.path.join(target_dir, filename)
        print(f"讀取: {filename}")
        
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
                    # 若為表頭重複則略過
                    if date_val == "交易日期" or date_val == header[0]:
                        continue
                        
                    # 相同日期若已存在，直接覆蓋 (保留一筆)
                    data_dict[date_val] = row
            except StopIteration:
                pass
                
    # 根據日期從大到小 (新到舊) 排序
    # 日期格式為 'yy/mm/dd，字串降序排序剛好符合新到舊
    sorted_dates = sorted(data_dict.keys(), reverse=True)
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    output_filename = f"{today_str}_Result.csv"
    output_filepath = os.path.join(target_dir, output_filename)
    
    with open(output_filepath, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)
        for d in sorted_dates:
            writer.writerow(data_dict[d])
            
    print(f"\n合併完成！總共 {len(sorted_dates)} 筆資料，已儲存至: {output_filename}")

if __name__ == "__main__":
    target = os.path.join("Download_Data", "PBR")
    merge_csv_files(target)
