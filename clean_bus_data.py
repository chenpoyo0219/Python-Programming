import pandas as pd
import glob
import os

# 設定您的資料夾路徑
DATA_FOLDER = r'Path'
OUTPUT_FILE = os.path.join(DATA_FOLDER, 'bus_time_15_final.xlsx')

# 1. 定義 132 公車的標準站名順序 (根據您提供的 20 個站點)
STATION_LIST = [
    "中央大學警衛室", "中大湖", "中央大學依仁堂", "中央大學後門", 
    "中央大學觀景台", "中央大學正門", "三民中正路口", "土地公廟", 
    "三民五興路口", "高雙里", "祐民醫院", "五權", "青果市場", 
    "仁愛新村", "廣興", "新明國中(民族路)", "舊社", "河川教育中心", 
    "第一銀行", "中壢公車站"
]

def process_bus_data():
    all_chunks = [] 
    
    # 搜尋符合 1218.xlsx 結尾的檔案
    search_pattern = os.path.join(DATA_FOLDER, "*1518.xlsx")
    files = glob.glob(search_pattern)
    files.sort() # 確保按日期排序
    
    if not files:
        print(f"錯誤：在路徑 {DATA_FOLDER} 下找不到符合 *1218.xlsx 的檔案。")
        return

    for file_path in files:
        file_name = os.path.basename(file_path)
        
        # 2. 提取日期
        try:
            date_code = file_name.split('bus_time_')[1][:4]
        except:
            date_code = "未知"
            
        print(f"正在處理：{file_name} (日期：{date_code})")
        
        try:
            # 3. 讀取 Excel 檔案
            df = pd.read_excel(file_path)
            
            # 4. 提取有抓到到站訊號 'O' 的站點時間
            arrived_df = df[df['目前公車位置'] == 'O'].copy()
            arrived_df = arrived_df.drop_duplicates(subset=['站名'], keep='first')
            
            # 建立一個時間對照表 {站名: 時間}
            recorded_times = dict(zip(arrived_df['站名'], arrived_df['時間']))
            
            # 5. 建立這一天完整的 20 個站點表格
            day_full_table = pd.DataFrame({'站名': STATION_LIST})
            
            # 將抓到的時間對應進去，沒抓到的會是空值 (NaN)
            day_full_table['時間'] = day_full_table['站名'].map(recorded_times)
            
            # 6. 【核心功能】補齊缺失時間
            # ffill(): 用上一個非空值填充 (沿用上一站時間)
            # bfill(): 如果第一站就沒時間，則用後一站時間補回 (確保表格沒有 NaN)
            day_full_table['時間'] = day_full_table['時間'].ffill().bfill()
            
            # 7. 插入日期欄位
            day_full_table.insert(0, '日期', date_code)
            
            # 8. 加入處理清單
            all_chunks.append(day_full_table)
            
            # 9. 加入一個「全空行」作為日期隔層
            empty_row = pd.DataFrame([[None] * 3], columns=['日期', '站名', '時間'])
            all_chunks.append(empty_row)
            
        except Exception as e:
            print(f"讀取檔案 {file_name} 時發生錯誤: {e}")
    
    # 10. 合併並匯出
    if all_chunks:
        if all_chunks[-1].isnull().all().all():
            all_chunks.pop() # 移除最後一個多餘的空行
            
        final_df = pd.concat(all_chunks, ignore_index=True)
        final_df.to_excel(OUTPUT_FILE, index=False)
        print(f"\n✅ 處理完成！所有缺失站點已補齊。")
        print(f"總表已儲存至：{OUTPUT_FILE}")
    else:
        print("❌ 未能提取到任何數據。")

if __name__ == "__main__":
    process_bus_data()