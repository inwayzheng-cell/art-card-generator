# 🎨 作品說明卡批次生成器 (Art Card Generator)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io/) 

這是一個專門為藝術展覽開發的自動化工具。透過 **Python** 與 **Streamlit** 驅動，能讀取 Excel 清單並自動將資料填入 PDF 模板，實現一鍵批次排版。

## 🌟 重點特色
- **自動分檔**：依據「作者」名稱自動分類，每個作者生成一份獨立 PDF。
- **繁體中文支援**：內建 `標楷體 (kaiu.ttf)`，解決 PDF 中文亂碼問題。
- **即時排版微調**：
  - 調整文字 X/Y 座標位置。
  - 自訂作品名、年代、作者的字體大小。
  - 彈性調整行距間隔。
- **一鍵打包**：所有生成的 PDF 會自動打包成 ZIP 壓縮檔供下載。

## 🛠️ 技術架構
- **UI 框架**：Streamlit
- **PDF 繪製**：ReportLab
- **PDF 合成**：pikepdf (QPDF 引擎)
- **數據處理**：Pandas

## 📁 專案檔案結構
- `streamlit_app.py`: 網頁主程式邏輯。
- `requirements.txt`: Python 依賴庫清單。
- `packages.txt`: 系統級組件 (libqpdf-dev)。
- `kaiu.ttf`: 繁體中文字型檔。
- 
## 📋 Excel 檔案格式說明
請確保您的 Excel 檔案 (.xlsx) 欄位依照以下順序排列，程式將會自動讀取：
- 序號,作者,名稱,大小,創作年份
- 1,王大明,無敵鐵金剛,100x100x100,2025
- 2,王大明,變形金剛,50x50x50,2024

## 🚀 部署說明
1. 將本儲存庫連動至 [Streamlit Cloud](https://share.streamlit.io/)。
2. 確保儲存庫根目錄包含 `packages.txt` 以正確安裝 PDF 處理引擎。
3. 部署完成後即可透過瀏覽器在任何裝置上使用。

##線上存取
可以透過以下連結直接在線上使用此工具：
https://art-card-generator-9zteodp8hffljwmi2iuexx.streamlit.app/





