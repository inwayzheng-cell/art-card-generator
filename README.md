# 🎨 作品說明卡自動生成器 (Art Card Generator)

這是一個專門為藝術展覽設計的自動化工具，能將 Excel 裡的作品清單自動排版到指定的 PDF 說明卡模板上。

### 🌟 核心功能
* **批次處理**：自動根據「作者」分組，生成獨立的 PDF 檔案。
* **網頁操作**：基於 Streamlit 開發，免安裝環境，開箱即用。
* **彈性排版**：可即時調整作品名、大小、年代與作者的字體大小及間距。
* **一鍵打包**：產出的所有 PDF 會自動壓縮成 ZIP 檔供下載。

### 🛠️ 技術架構
* **Python**: 核心邏輯處理。
* **Streamlit**: 網頁介面。
* **ReportLab / pikepdf**: PDF 內容繪製與模板合成。

### 📁 檔案說明
* `streamlit_app.py`: 網頁程式主體。
* `requirements.txt`: 必要的 Python 套件清單。
* `packages.txt`: 雲端伺服器所需的 QPDF 系統組件。
* `kaiu.ttf`: 預設標楷體字型檔。
