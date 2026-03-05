import streamlit as st
import pandas as pd
import io, zipfile, os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
import pikepdf

# 1. 自動偵測字型
FONT_NAME = "Kaiu"
FONT_PATH = "kaiu.ttf" # 確保 GitHub 倉庫裡有這個檔案

if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_PATH))
else:
    st.error("⚠️ 找不到 kaiu.ttf 字型檔，請上傳至 GitHub 儲存庫。")

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
    xl = st.number_input("左(往右數字加大,往左數字減小)", value=166.0)
    xr = st.number_input("右(往右數字加大,往左數字減小)", value=435.0)
    st.divider()
    st_sz = st.slider("作品名 字體大小", 10, 30, 19)
    si_sz = st.slider("大小/年代 字體大小", 10, 30, 15)
    sa_sz = st.slider("作者 字體大小", 10, 30, 12)
    st.divider()
    g1 = st.number_input("作品名 -> 大小年代間距", value=25.0)
    g2 = st.number_input("作品名 -> 作者間距", value=60.0)

# 主畫面
col1, col2 = st.columns(2)
with col1:
    uploaded_excel = st.file_uploader("1. 上傳 Excel ", type=["xlsx"])
with col2:
    uploaded_pdf = st.file_uploader("2. 上傳 PDF 模板背景", type=["pdf"])

if st.button("🚀 開始批次生成所有 PDF", use_container_width=True):
    if uploaded_excel and uploaded_pdf:
        try:
            df = pd.read_excel(uploaded_excel)
            grouped = df.groupby(df.iloc[:, 1]) # 按作者分組
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a") as zip_file:
                for author_name, group_df in grouped:
                    safe_name = str(author_name).strip()
                    pdf_io = io.BytesIO()
                    
                    with pikepdf.Pdf.open(uploaded_pdf) as src:
                        final_pdf = pikepdf.Pdf.new()
                        for page_start in range(0, len(group_df), 10):
                            packet = io.BytesIO()
                            can = canvas.Canvas(packet, pagesize=A4)
                            for i in range(10):
                                idx = page_start + i
                                if idx < len(group_df):
                                    row = group_df.iloc[idx]
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
                    zip_file.writestr(f"{safe_name}.pdf", pdf_io.getvalue())
            
            st.success("✅ 生成完畢！")
            st.download_button("📥 點我下載所有 PDF (ZIP)", zip_buffer.getvalue(), "說明卡產出.zip", "application/zip", use_container_width=True)
        except Exception as e:
            st.error(f"發生錯誤: {e}")
    else:
        st.warning("⚠️ 請先上傳 Excel 與 PDF 模板。")