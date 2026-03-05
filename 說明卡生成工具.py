import base64
import streamlit as st
import pandas as pd
import io, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pikepdf

# 1. 自動偵測字型
FONT_NAME = "Kaiu"
FONT_PATH = "kaiu.ttf"

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
else:
    st.error("⚠️ 找不到 kaiu.ttf 字型檔，請確保檔案已上傳至 GitHub 儲存庫。")

def format_value(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.replace(" .", ".").replace(". ", ".")

# --- 網頁介面 ---
st.set_page_config(page_title="作品說明卡生成器", layout="wide")
st.title("🎨 作品說明卡網頁生成工具")

with st.sidebar:
    st.header("📏 排版參數")
    xl = st.number_input("左欄 X (往右數字加大)", value=166.0)
    xr = st.number_input("右欄 X (往右數字加大)", value=435.0)
    st.divider()
    st_sz = st.slider("作品名 字體大小", 10, 40, 19)
    si_sz = st.slider("大小/年代 字體大小", 10, 40, 15)
    sa_sz = st.slider("作者 字體大小", 10, 40, 12)
    st.divider()
    g1 = st.number_input("作品名 -> 大小年代間距", value=25.0)
    g2 = st.number_input("作品名 -> 作者間距", value=60.0)

# 主畫面
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上傳 Excel ", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. 上傳 PDF 模板背景", type=["pdf"])

if st.button("🚀 生成單一 PDF 並預覽", use_container_width=True):
    if uploaded_excel and uploaded_pdf:
        try:
            df = pd.read_excel(uploaded_excel)
            # 不分組，直接處理全體資料
            pdf_io = io.BytesIO()
            
            with pikepdf.Pdf.open(uploaded_pdf) as src:
                final_pdf = pikepdf.Pdf.new()
                # 每 10 筆資料一頁
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
                            can.drawCentredString(cx, cy, format_value(row.iloc[2]))
                            can.setFont(FONT_NAME, si_sz)
                            can.drawCentredString(cx, cy - g1, f"{format_value(row.iloc[3])} {format_value(row.iloc[4])}")
                            can.setFont(FONT_NAME, sa_sz)
                            can.drawCentredString(cx, cy - g2, format_value(row.iloc[1]))
                    can.save()
                    packet.seek(0)
                    with pikepdf.Pdf.open(packet) as overlay:
                        dst_page = final_pdf.add_blank_page(page_size=A4)
                        dst_page.add_underlay(src.pages[0])
                        dst_page.add_overlay(overlay.pages[0])
                final_pdf.save(pdf_io)
            
            result_pdf = pdf_io.getvalue()

            # --- 下載 PDF 按鈕 ---
            st.download_button(
                label="📥 下載完整 PDF 說明卡",
                data=result_pdf,
                file_name="全部作品說明卡.pdf",
                mime="application/pdf",
                use_container_width=True
            )

# --- 介面顯示：PDF 預覽區 ---
            if last_pdf_data:
                st.divider()
                st.subheader("👁️ 排版成果預覽")
                
                # 將 PDF 轉為 Base64
                base64_pdf = base64.b64encode(last_pdf_data).decode('utf-8')
                
                # 建立一個下載/預覽連結 (對手機更友善)
                pdf_link = f'<a href="data:application/pdf;base64,{base64_pdf}" target="_blank" style="text-decoration:none;"><button style="width:100%; padding:10px; background-color:#FF4B4B; color:white; border:none; border-radius:5px;">🔍 手機點此開啟全螢幕預覽</button></a>'
                st.markdown(pdf_link, unsafe_allow_html=True)
                
                # 原有的內嵌視窗 (保留給電腦版看，手機若空白可點上面的按鈕)
                pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="100%" height="800" type="application/pdf"></iframe>'
                st.markdown(pdf_display, unsafe_allow_html=True)

        except Exception as e:
            st.error(f"發生錯誤: {e}")

