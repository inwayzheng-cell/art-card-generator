import base64
import streamlit as st
import pandas as pd
import io, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pikepdf
from pdf2image import convert_from_bytes
import requests

FONT_NAME = "msjhbd"
FONT_PATH = "msjhbd.ttc"

def download_font():
    # 使用正確的 Raw 連結
    url = "https://github.com/inwayzheng-cell/art-card-generator/raw/main/msjhbd.ttc"
    
    # 如果檔案不存在，或者檔案太小 (小於 100KB 通常就是 LFS 指標檔)，就重新下載
    if not os.path.exists(FONT_PATH) or os.path.getsize(FONT_PATH) < 100000:
        with st.spinner("正在從雲端載入微軟正黑體 (約 15MB)，請稍候..."):
            try:
                response = requests.get(url, timeout=30)
                response.raise_for_status() # 確保下載成功
                with open(FONT_PATH, "wb") as f:
                    f.write(response.content)
                st.success("字體下載完成！")
            except Exception as e:
                st.error(f"下載字體時發生網路錯誤: {e}")

# 執行下載
download_font()

# 註冊字體
try:
    # 注意：針對 .ttc 檔，ReportLab 有時需要指定 subfontIndex
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
except Exception as e:
    st.error(f"⚠️ 字體註冊失敗: {e}")

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
else:
    st.error("⚠️ 找不到 msjhbd.ttc 字型檔，請確保檔案已上傳至 GitHub。")

def format_value(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.replace(" .", ".").replace(". ", ".")

LOGO_URL = "https://raw.githubusercontent.com/inwayzheng-cell/art-card-generator/main/static/logo.png"


import streamlit as st

# --- 1. 基礎頁面配置 ---
st.set_page_config(
    page_title="Art Card Generator",
    layout="wide",
    page_icon="🎨"
)

# --- 2. 注入 PWA  ---
st.markdown(
    """
    <style>
        /* 1. 解決標題消失問題 (截圖頂部的製作小卡) */
        h1, h2, h3, .stMarkdown p {
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
            opacity: 1 !important;
        }

        /* 2. 徹底修正上傳框 (截圖中黑壓壓的區塊) */
        [data-testid="stFileUploader"] {
            background-color: #ffffff !important; /* 強制框內背景白色 */
            border-radius: 10px;
            padding: 10px;
        }
        
        /* 修正上傳框內部的「Browse files」按鈕 */
        [data-testid="stFileUploader"] button {
            background-color: #007bff !important; /* 改成藍色背景 */
            color: #ffffff !important;           /* 白色文字 */
            -webkit-text-fill-color: #ffffff !important;
            opacity: 1 !important;
            border: none !important;
        }

        /* 3. 修正截圖中「Drag and drop file here」字體太淡的問題 */
        [data-testid="stFileUploader"] section div {
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
        }

        /* 4. 側邊欄與參數調整區 (排版參數黑糊糊修正) */
        [data-testid="stSidebar"] {
            background-color: #ffffff !important;
        }
        [data-testid="stSidebar"] * {
            color: #000000 !important;
            -webkit-text-fill-color: #000000 !important;
        }
        
        /* 5. 確保輸入框是白底黑字 */
        input {
            background-color: #ffffff !important;
            color: #000000 !important;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("製作小卡")

with st.sidebar:
    st.header("📏 排版設定")
    xl = st.number_input("左欄(往右偏加大,往左偏減小)", value=166.0)
    xr = st.number_input("右欄(往右偏加大,往左偏減小)", value=435.0)
    
    st.divider()
    st.subheader("🔠 字體大小")
    st_sz = st.slider("作品名 字體大小", 15.0, 25.0, 18.5)
    si_sz = st.slider("大小/年代 字體大小", 10.0, 20.0, 14.5)
    sa_sz = st.slider("作者 字體大小", 10.0, 20.0, 16.0)
    
    st.divider()
    st.subheader("↕️ 上下間距")
    g1 = st.number_input("作品名 -> 大小與年代間距", value=30.0)
    g2 = st.number_input("作品名 -> 作者間距", value=61.0)
    
    st.divider()
    st.subheader("↔️ 資訊欄精密微調")
    # 🚀 新增這兩行：
    gi = st.number_input("大小與年代 之間的間距 (藍色)", value=15.0)
    h_off = st.number_input("大小與年代 整體水平位移 (紅色)", value=24.0)
    
    st.caption("提示：位移正值往右，負值往左")

if "final_pdf_data" not in st.session_state:
    st.session_state.final_pdf_data = None


col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上傳 Excel", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. 上傳 PDF 模板", type=["pdf"])

if st.button("🚀 開始生成 PDF 並預覽", use_container_width=True):
    if uploaded_excel and uploaded_pdf:
        try:
            df = pd.read_excel(uploaded_excel)
            pdf_io = io.BytesIO()
            
            with pikepdf.Pdf.open(uploaded_pdf) as src:
                final_pdf = pikepdf.Pdf.new()
                
                
                for page_start in range(0, len(df), 10):
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=A4)
                    
# --- 在 run_process 方法內部的繪製迴圈中 ---
for i in range(10):
    idx = page_start + i
    if idx < len(df):
        row = df.iloc[idx]
        cx, cy = (xl if i < 5 else xr), y_list[i % 5]
        
        # 1. 繪製作品名 (置中)
        can.setFont("UserFont", st)
        can.drawCentredString(cx, cy, format_value(row.iloc[2]))
        
        # 2. 繪製「大小」與「年代」 (應用內部間距與整體位移)
        can.setFont("UserFont", si)
        size_txt, year_txt = format_value(row.iloc[3]), format_value(row.iloc[4])
        
        # 計算水平位置：基準點 cx 加上整體位移 h_off
        # 大小向左偏移間距的一半，年代向右偏移間距的一半
        base_x = cx + h_off
        can.drawRightString(base_x - (gi / 2), cy - g1, size_txt) # 大小 (向左偏)
        can.drawString(base_x + (gi / 2), cy - g1, year_txt)      # 年代 (向右偏)
        
        # 3. 繪製作者名 (置中)
        can.setFont("UserFont", sa)
        can.drawCentredString(cx, cy - g2, format_value(row.iloc[1]))
                    
                    can.save()
                    packet.seek(0)
                    
                    
                    with pikepdf.Pdf.open(packet) as overlay:
                        dst_page = final_pdf.add_blank_page(page_size=A4)
                        dst_page.add_underlay(src.pages[0])
                        dst_page.add_overlay(overlay.pages[0])
                
                final_pdf.save(pdf_io)
                st.session_state.final_pdf_data = pdf_io.getvalue()
                st.success(f"✅ 生成成功！共{len(df)}小張。")
                
        except Exception as e:
            st.error(f"發生錯誤: {e}")
    else:
        st.warning("⚠ 請先上傳 Excel 與 PDF 模板。")


if st.session_state.final_pdf_data:
    st.divider()
    
    
    st.download_button(
        label="📥 下載作品小卡 (PDF)",
        data=st.session_state.final_pdf_data,
        file_name="小卡總表.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    
    st.subheader("👁️即時預覽 (第一頁)")
    try:
        
        images = convert_from_bytes(st.session_state.final_pdf_data, first_page=1, last_page=1)
        if images:
            st.image(images[0], use_container_width=True)
    except:
        
        b64_pdf = base64.b64encode(st.session_state.final_pdf_data).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="600"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)
