# Python-Programming
Unknown
# 公車即時到站時間蒐集與分析系統

本專案使用 **Python 自動化爬蟲** 定期蒐集桃園市公車動態資訊系統之即時資料，記錄各站點的「預估到站時間」與「實際進站時間」，並計算各站點的平均實際進站時間，輸出為 Excel 檔案以利後續分析與模型訓練使用。

***

## 功能說明

*   自動開啟桃園公車動態資訊系統網頁
*   擷取指定路線、指定方向之站序、站名、預估到站時間
*   當公車實際進站時記錄系統時間
*   以固定時間間隔（每分鐘）連續蒐集資料
*   將原始資料與計算後的平均實際進站時間輸出至 Excel

***

## 使用套件與環境需求

### 第三方 Python 套件（需自行安裝）

```txt
selenium
beautifulsoup4
pandas
openpyxl
```

套件用途說明：

*   **selenium**：自動化操作瀏覽器，取得動態載入之公車即時資料
*   **beautifulsoup4**：解析 HTML 網頁內容，擷取站點資訊
*   **pandas**：資料整理、清洗與平均時間計算
*   **openpyxl**：讀取與寫入 Excel（.xlsx）檔案

安裝指令：

```bash
pip install selenium beautifulsoup4 pandas openpyxl
```

***

### Python 內建標準函式庫（無需安裝）

```txt
time
datetime
os
```

*   **time / datetime**：時間取得、格式轉換、分鐘數計算
*   **os**：檔案與資料夾存在性處理

***

## 系統環境需求

*   **Python**：3.9 以上版本（建議）
*   **Google Chrome 瀏覽器**
*   **ChromeDriver**
    *   版本需與 Chrome 瀏覽器版本相符
    *   已設定於系統環境變數（PATH）中

***

## 資料來源

*   桃園市政府  
    **桃園公車動態資訊系統**  
    <https://ebus.tycg.gov.tw/>

***

## 輸出說明

Excel 檔案包含兩個工作表：

1.  **原始資料**
    *   站序
    *   站名
    *   預估到站時間
    *   是否偵測到公車位置
    *   實際進站時間（分鐘）

2.  **計算後時間**
    *   各站點平均實際進站時間（HH:MM）

***

## 備註

*   本程式使用 **Headless Chrome** 執行，不會顯示瀏覽器畫面
*   適合搭配排程工具（如 Windows 工作排程）定期執行
*   產出資料可作為後續時間序列分析或預測模型訓練資料使用

***

如果你要的是：

*   ✅ **README 英文版**
*   ✅ **requirements.txt**
*   ✅ **專題報告用版本（老師會喜歡那種）**

直接跟我說，我可以一次幫你補齊。
