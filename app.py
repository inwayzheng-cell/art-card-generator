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

LOGO_URL = "https://raw.githubusercontent.com/inwayzheng-cell/art-card-generator/main/static/logo.png"


import streamlit as st

# --- 1. 基礎頁面配置 ---
st.set_page_config(
    page_title="Art Card Generator",
    layout="wide",
    page_icon="🎨"
)

# --- 2. 注入 PWA 與外觀修正 (強制亮色模式避免字體過淡) ---
st.markdown(
    """
    <style>
        html, body, [data-testid="stAppViewContainer"], .main {
            background-color: white !important;
        }
        h1, h2, h3, p, span, label, .stMarkdown {
            color: #31333F !important;
        }
    </style>
    <head>
        <link rel="manifest" href="./manifest.json">
        <link rel="icon" sizes="512x512" href="./static/logo.png">
        <link rel="apple-touch-icon" href="./static/logo.png">
        <meta name="theme-color" content="#ffffff">
    </head>
    """,
    unsafe_allow_html=True
)

st.title("製作小卡")

with st.sidebar:
    st.header("📏 排版")
    xl = st.number_input("左欄(往右偏加大,往左偏減小)", value=166)
    xr = st.number_input("右欄(往右偏加大,往左偏減小)", value=435)
    st.divider()
    st_sz = st.slider("作品名 字體大小", 10, 40, 25)
    si_sz = st.slider("大小/年代 字體大小", 10, 40, 15)
    sa_sz = st.slider("作者 字體大小", 10, 40, 12)
    st.divider()
    g1 = st.number_input("作品名 -> 大小年代間距", value=25)
    g2 = st.number_input("作品名 -> 作者間距", value=60)


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
                    
                    for i in range(10):
                        idx = page_start + i
                        if idx < len(df):
                            row = df.iloc[idx]
                            
                            cx = xl if i < 5 else xr
                            cy = [765.0, 601.5, 436.5, 273.5, 108.5][i % 5]
                            
                            
                            can.setFont(FONT_NAME, st_sz)
                            can.drawCentredString(cx, cy, format_value(row.iloc[2])) # 作品名
                            can.setFont(FONT_NAME, si_sz)
                            can.drawCentredString(cx, cy - g1, f"{format_value(row.iloc[3])} {format_value(row.iloc[4])}") # 大小 年份
                            can.setFont(FONT_NAME, sa_sz)
                            can.drawCentredString(cx, cy - g2, format_value(row.iloc[1])) # 作者
                    
                    can.save()
                    packet.seek(0)
                    
                    
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










