# commute-screenshot-tool
自動擷取 Google Maps 通勤路線截圖（Python + Selenium）
# Google Maps 通勤路線自動截圖工具

這是一個使用 Python 開發的自動化工具，透過 Selenium 自動開啟 Google Maps，依據 Excel 資料中的起點與終點地址查詢通勤路線，擷取路線截圖並自動儲存與編號。專案主要用於輔助 ESG 報告中碳排放盤查作業，大幅減少人工截圖與紀錄的作業時間。

## 📌 專案功能特色

- 從 Excel 中讀取多筆起訖地址
- 自動開啟 Google Maps 並查詢通勤路線
- 擷取螢幕截圖並依序命名儲存
- 自動處理地圖介面異動與載入延遲問題
- 
- ## 🛠 使用技術
- Python 3.x
- Selenium
- WebDriver Manager
- pandas / openpyxl
- 正規表達式（re）與 urllib.parse 處理地址與格式

- ## 📦 安裝套件

使用 pip 安裝必要的 Python 套件：

```bash
pip install selenium pandas openpyxl webdriver-manager
```

📑 Excel 輸入格式說明
可參考demo_excel.xlsx 的檔案，資料表需包含以下欄位（從左至右）：

index	 交通方式	  起點地址	    終點地址
 1	  汽車	    台中市XX路	   台中市YY街
 2	  機車	    台中市XX路	   台中市YY街

🚀 執行方式
確認路徑設定：
請打開 google爬蟲demo.py，並修改此行為你的 Excel 檔案實際路徑：

file_path = r"C:存放demo_excel的資料夾\\demo_excel.xlsx"
執行程式：python google爬蟲demo.py

產出自動命名的通勤截圖於「公里數截圖」資料夾
匯出查詢結果為 公里數資料_result.csv

🖼 截圖命名規則
{index}_{交通方式}_{地址}.png
例如：A001_汽車_台中市南屯區OO路.png


🔐 注意事項
須安裝 Chrome 瀏覽器，且建議使用最新版以確保與 ChromeDriver 相容
執行過程中會自動開啟瀏覽器進行操作，請勿中途干擾
若部分路線無法擷取公里數，程式會自動略過
