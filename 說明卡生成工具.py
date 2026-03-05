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

# 1. 字型設定
FONT_NAME = "Kaiu"
FONT_PATH = "kaiu.ttf"

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
else:
    st.error("⚠️ 找不到 kaiu.ttf 字型檔，請確保檔案已上傳至 GitHub。")

def format_value(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.replace(" .", ".").replace(". ", ".")

# 1. 頁面配置
st.set_page_config(
    page_title="作品說明卡生成器", 
    page_icon="logo.png", 
    layout="wide"
)

# 2. 強制手機抓取新圖示 (加上 ?v=1 參數)
LOGO_URL = "https://raw.githubusercontent.com/inwayzheng-cell/art-card-generator/main/logo.png?v=1"

st.markdown(f"""
    <link rel="apple-touch-icon" href="{LOGO_URL}">
    <link rel="icon" type="image/png" href="{LOGO_URL}">
    <link rel="shortcut icon" href="{LOGO_URL}">
    """, unsafe_allow_html=True)

st.title("🎨作品小卡生成工具")

with st.sidebar:
    st.header("📏 排版參數")
    xl = st.number_input("左欄(往右偏加大,往左偏減小)", value=166)
    xr = st.number_input("右欄(往右偏加大,往左偏減小)", value=435)
    st.divider()
    st_sz = st.slider("作品名 字體大小", 10, 40, 25)
    si_sz = st.slider("大小/年代 字體大小", 10, 40, 15)
    sa_sz = st.slider("作者 字體大小", 10, 40, 12)
    st.divider()
    g1 = st.number_input("作品名 -> 大小年代間距", value=25)
    g2 = st.number_input("作品名 -> 作者間距", value=60)

# 初始化 Session State
if "final_pdf_data" not in st.session_state:
    st.session_state.final_pdf_data = None

# 檔案上傳區
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
                
                # 按照資料列，每 10 筆產出一頁 PDF
                for page_start in range(0, len(df), 10):
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=A4)
                    
                    for i in range(10):
                        idx = page_start + i
                        if idx < len(df):
                            row = df.iloc[idx]
                            # 計算位置：前 5 筆在左，後 5 筆在右
                            cx = xl if i < 5 else xr
                            cy = [765.0, 601.5, 436.5, 273.5, 108.5][i % 5]
                            
                            # 寫入文字
                            can.setFont(FONT_NAME, st_sz)
                            can.drawCentredString(cx, cy, format_value(row.iloc[2])) # 作品名
                            can.setFont(FONT_NAME, si_sz)
                            can.drawCentredString(cx, cy - g1, f"{format_value(row.iloc[3])} {format_value(row.iloc[4])}") # 大小 年份
                            can.setFont(FONT_NAME, sa_sz)
                            can.drawCentredString(cx, cy - g2, format_value(row.iloc[1])) # 作者
                    
                    can.save()
                    packet.seek(0)
                    
                    # 疊加模板與文字
                    with pikepdf.Pdf.open(packet) as overlay:
                        dst_page = final_pdf.add_blank_page(page_size=A4)
                        dst_page.add_underlay(src.pages[0])
                        dst_page.add_overlay(overlay.pages[0])
                
                final_pdf.save(pdf_io)
                st.session_state.final_pdf_data = pdf_io.getvalue()
                st.success(f"✅ 生成成功！共 {len(df)} 張。")
                
        except Exception as e:
            st.error(f"發生錯誤: {e}")
    else:
        st.warning("⚠️ 請先上傳 Excel 與 PDF 模板。")

# --- 顯示與下載區 ---
if st.session_state.final_pdf_data:
    st.divider()
    
    # 下載按鈕 (直接下載 PDF)
    st.download_button(
        label="📥 下載完整作品說明卡 (PDF)",
        data=st.session_state.final_pdf_data,
        file_name="作品說明卡總表.pdf",
        mime="application/pdf",
        use_container_width=True
    )

    # 預覽區
    st.subheader("👁️ 即時預覽 (第一頁)")
    try:
        # 手機顯示圖片預覽
        images = convert_from_bytes(st.session_state.final_pdf_data, first_page=1, last_page=1)
        if images:
            st.image(images[0], use_container_width=True)
    except:
        # 電腦顯示內嵌 PDF
        b64_pdf = base64.b64encode(st.session_state.final_pdf_data).decode('utf-8')
        pdf_display = f'<iframe src="data:application/pdf;base64,{b64_pdf}" width="100%" height="600"></iframe>'
        st.markdown(pdf_display, unsafe_allow_html=True)


