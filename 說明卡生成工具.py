import os
import sys
import hashlib
import io
import tkinter as tk
from tkinter import filedialog, messagebox

# --- 核心補丁 (解決部分環境 MD5 安全性限制) ---
original_md5 = hashlib.md5
def patched_md5(*args, **kwargs):
    kwargs.pop('usedforsecurity', None)
    return original_md5(*args, **kwargs)
hashlib.md5 = patched_md5

import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

try:
    import pikepdf
except ImportError:
    pass

DEFAULT_FONT_PATH = "C:\\Windows\\Fonts\\msjhbd.ttc"

def format_value(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.replace(" .", ".").replace(". ", ".")

class CardApp:
    def __init__(self, root):
        self.root = root
        self.root.title("作品說明卡整合工具 v1.0")
        self.root.geometry("880x720") # 稍微增加高度確保排版舒適

        # 1. 統一定義字體與風格
        self.ui_f_base = ("微軟正黑體", 12)
        self.ui_f_bold = ("微軟正黑體", 13, "bold")
        self.ui_f_entry = ("Consolas", 13) # 全域輸入框字體加大
        self.ui_f_title = ("微軟正黑體", 20, "bold")

        # 2. 變數設定
        self.excel_path = tk.StringVar()
        self.pdf_path = tk.StringVar()
        self.font_path = tk.StringVar(value=DEFAULT_FONT_PATH)
        self.x_left = tk.StringVar(value="166.0")
        self.x_right = tk.StringVar(value="435.0")
        self.y_vals = [tk.StringVar(value=str(v)) for v in [776.0, 612.5, 447.5, 284.5, 119.5]]
        
        self.sz_title = tk.StringVar(value="18.5")
        self.sz_info = tk.StringVar(value="14.5")
        self.sz_author = tk.StringVar(value="16.0")
        
        self.gap1 = tk.StringVar(value="30.0")
        self.gap2 = tk.StringVar(value="61.0")
        self.gap_info = tk.StringVar(value="15.0") 
        self.h_offset = tk.StringVar(value="24.0") 

        self.setup_ui()

    def setup_ui(self):
        self.root.configure(bg="#f0f2f5") 
        
        # --- 頂部總標題 ---
        header_label = tk.Label(self.root, text="作品說明卡整合工具", font=self.ui_f_title, 
                                bg="#f0f2f5", fg="#1a3a5f")
        header_label.pack(pady=(20, 10))

        # --- 1. 檔案設定 (淺藍色系) ---
        f_frame = tk.LabelFrame(self.root, text="", padx=20, pady=10, bg="#e3f2fd", relief="groove", bd=2)
        f_frame.pack(fill="x", padx=40, pady=15)
        
        # 終極墊片法標題
        f_title = tk.Label(self.root, text=" 📂 檔案路徑設定 ", font=self.ui_f_bold, bg="#e3f2fd", fg="#0d47a1")
        f_title.place(x=55, y=87)

        f_inner = tk.Frame(f_frame, bg="#e3f2fd")
        f_inner.pack(fill="x", pady=(15, 5))

        tk.Button(f_inner, text="選擇 Excel", command=self.select_excel, width=12, 
                  bg="#2196f3", fg="white", activebackground="#1976d2").grid(row=0, column=0, pady=5)
        tk.Entry(f_inner, textvariable=self.excel_path, font=self.ui_f_entry, width=75).grid(row=0, column=1, padx=10)
        
        tk.Button(f_inner, text="選擇PDF模板", command=self.select_pdf, width=12, 
                  bg="#2196f3", fg="white", activebackground="#1976d2").grid(row=1, column=0, pady=5)
        tk.Entry(f_inner, textvariable=self.pdf_path, font=self.ui_f_entry, width=75).grid(row=1, column=1, padx=10)

 # --- 2. 座標設定 (淺橘色系) ---
        # 建立框架，text 設為空字串避免原生標題干擾
        pos_frame = tk.LabelFrame(self.root, text="", padx=20, pady=10, 
                                  bg="#fff3e0", relief="groove", bd=2)
        pos_frame.pack(fill="x", padx=40, pady=15)

        # 🚀 改用絕對定位標題，背景色蓋過邊框線
        pos_title = tk.Label(self.root, text=" A4 座標排版 ", font=self.ui_f_bold, 
                             bg="#fff3e0", fg="#e65100")
        pos_title.place(x=55, y=233) # y 座標需根據你視窗的實際高度微調

        # 提示文字移到內部
        inner_pos = tk.Frame(pos_frame, bg="#fff3e0")
        inner_pos.pack(fill="x", pady=(15, 5))

        # 第一列：X 座標
        tk.Label(inner_pos, text="左欄 X:", bg="#fff3e0", font=self.ui_f_base).grid(row=0, column=0, sticky="e", pady=5)
        tk.Entry(inner_pos, textvariable=self.x_left, width=10, justify="center", font=self.ui_f_entry).grid(row=0, column=1, padx=10)
        
        tk.Label(inner_pos, text="右欄 X:", bg="#fff3e0", font=self.ui_f_base).grid(row=0, column=2, sticky="e", padx=(20, 0))
        tk.Entry(inner_pos, textvariable=self.x_right, width=10, justify="center", font=self.ui_f_entry).grid(row=0, column=3, padx=10)

        # 提示語（小字）放在 X 座標旁邊
        tk.Label(inner_pos, text="(X 向右加大 / Y 向上加大)", font=("微軟正黑體", 9), 
                 bg="#fff3e0", fg="#ef6c00").grid(row=0, column=4, padx=10)

        # 第二列：Y 座標
        y_box = tk.Frame(inner_pos, bg="#fff3e0")
        y_box.grid(row=1, column=0, columnspan=5, pady=(15, 5)) # columnspan 改為 5 以對齊
        for i in range(5):
            tk.Label(y_box, text=f"排{i+1} Y:", bg="#fff3e0", font=("Consolas", 12, "bold"), fg="#bf360c").pack(side="left", padx=(5, 2))
            tk.Entry(y_box, textvariable=self.y_vals[i], width=8, justify="center", font=self.ui_f_entry).pack(side="left", padx=(0, 10))

        # --- 3. 精密微調 (淺紫色系) ---
        adj_frame = tk.LabelFrame(self.root, text="", padx=20, pady=10, bg="#f3e5f5", relief="groove", bd=2)
        adj_frame.pack(fill="x", padx=40, pady=15)

        adj_title = tk.Label(self.root, text=" 字體與間距微調 ", font=self.ui_f_bold, bg="#f3e5f5", fg="#4a148c")
        adj_title.place(x=55, y=388) # 座標依視窗高度微調

        grid_adj = tk.Frame(adj_frame, bg="#f3e5f5")
        grid_adj.pack(fill="x", pady=(15, 5))

        # Row 0: 字體大小
        tk.Label(grid_adj, text="作品 Size:", bg="#f3e5f5", font=self.ui_f_base).grid(row=0, column=0, sticky="e", pady=5)
        tk.Entry(grid_adj, textvariable=self.sz_title, width=10, justify="center", font=self.ui_f_entry).grid(row=0, column=1, padx=5)
        tk.Label(grid_adj, text="大小&年代 Size:", bg="#f3e5f5", font=self.ui_f_base).grid(row=0, column=2, sticky="e", padx=(15, 0))
        tk.Entry(grid_adj, textvariable=self.sz_info, width=10, justify="center", font=self.ui_f_entry).grid(row=0, column=3, padx=5)
        tk.Label(grid_adj, text="作者 Size:", bg="#f3e5f5", font=self.ui_f_base).grid(row=0, column=4, sticky="e", padx=(15, 0))
        tk.Entry(grid_adj, textvariable=self.sz_author, width=10, justify="center", font=self.ui_f_entry).grid(row=0, column=5, padx=5)

        # Row 1: 間距
        tk.Label(grid_adj, text="作品↔大小&年代(上/下):", bg="#f3e5f5", font=self.ui_f_base, fg="#6a1b9a").grid(row=1, column=0, sticky="e", pady=5)
        tk.Entry(grid_adj, textvariable=self.gap1, width=10, justify="center", font=self.ui_f_entry).grid(row=1, column=1, padx=5)
        tk.Label(grid_adj, text="作品↔作者(上/下):", bg="#f3e5f5", font=self.ui_f_base, fg="#6a1b9a").grid(row=1, column=2, sticky="e", padx=(15, 0))
        tk.Entry(grid_adj, textvariable=self.gap2, width=10, justify="center", font=self.ui_f_entry).grid(row=1, column=3, padx=5)

        # Row 2: 水平修正
        tk.Label(grid_adj, text="大小&年代間距(藍):", bg="#f3e5f5", fg="#1565c0", font=self.ui_f_base).grid(row=2, column=0, sticky="e", pady=5)
        tk.Entry(grid_adj, textvariable=self.gap_info, width=10, justify="center", font=self.ui_f_entry).grid(row=2, column=1, padx=5)
        tk.Label(grid_adj, text="大小&年代位置(紅):", bg="#f3e5f5", fg="#c62828", font=self.ui_f_base).grid(row=2, column=2, sticky="e", padx=(15, 0))
        tk.Entry(grid_adj, textvariable=self.h_offset, width=10, justify="center", font=self.ui_f_entry).grid(row=2, column=3, padx=5)
        tk.Label(grid_adj, text="(向左減小/向右增大)", font=("微軟正黑體", 9), fg="#7b1fa2", bg="#f3e5f5").grid(row=2, column=4, columnspan=2, sticky="w", padx=5)

        # --- 生成按鈕 ---
        btn_run = tk.Button(self.root, text="🚀 產生作品說明卡 PDF", font=("微軟正黑體", 16, "bold"), 
                            bg="#43a047", fg="white", activebackground="#2e7d32", activeforeground="white",
                            width=35, height=2, relief="raised", bd=3, command=self.run_process)
        btn_run.pack(pady=25)

        self.status_label = tk.Label(self.root, text="● 準備就緒", fg="#546e7a", bg="#f0f2f5", font=("微軟正黑體", 11, "bold"))
        self.status_label.pack()

    def select_excel(self): 
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path: self.excel_path.set(path)

    def select_pdf(self): 
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if path: self.pdf_path.set(path)

    def run_process(self):
        e_file, p_template, f_file = self.excel_path.get(), self.pdf_path.get(), self.font_path.get()
        if not e_file or not p_template:
            messagebox.showwarning("提示", "請確認 Excel 與 PDF 模板皆已選擇")
            return

        try:
            xl, xr = float(self.x_left.get()), float(self.x_right.get())
            y_list = [float(v.get()) for v in self.y_vals]
            st, si, sa = float(self.sz_title.get()), float(self.sz_info.get()), float(self.sz_author.get())
            g1, g2 = float(self.gap1.get()), float(self.gap2.get())
            gi, h_off = float(self.gap_info.get()), float(self.h_offset.get())

            pdfmetrics.registerFont(TTFont("UserFont", f_file))
            df = pd.read_excel(e_file, engine='openpyxl')
            df = df.dropna(subset=[df.columns[1]])
            df = df.sort_values(by=df.columns[1]).reset_index(drop=True)

            output_path = os.path.join(os.path.dirname(e_file), "作品說明卡-總表.pdf")
            self.status_label.config(text="正在處理中...", fg="blue")
            self.root.update()

            with pikepdf.Pdf.open(p_template) as src:
                final_pdf = pikepdf.Pdf.new()
                for page_start in range(0, len(df), 10):
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=A4)
                    for i in range(10):
                        idx = page_start + i
                        if idx < len(df):
                            row = df.iloc[idx]
                            cx, cy = (xl if i < 5 else xr), y_list[i % 5]
                            can.setFont("UserFont", st)
                            can.drawCentredString(cx, cy, format_value(row.iloc[2]))
                            can.setFont("UserFont", si)
                            size_txt, year_txt = format_value(row.iloc[3]), format_value(row.iloc[4])
                            can.drawRightString(cx + h_off - gi/2, cy - g1, size_txt)
                            can.drawString(cx + h_off + gi/2, cy - g1, year_txt)
                            can.setFont("UserFont", sa)
                            can.drawCentredString(cx, cy - g2, format_value(row.iloc[1]))
                    can.save()
                    packet.seek(0)
                    with pikepdf.Pdf.open(packet) as overlay_pdf:
                        dst_page = final_pdf.add_blank_page(page_size=A4)
                        dst_page.add_underlay(src.pages[0])
                        dst_page.add_overlay(overlay_pdf.pages[0])
                final_pdf.save(output_path)

            self.status_label.config(text="⭐ 生成成功！", fg="green")
            os.startfile(output_path)
            
        except Exception as e:
            messagebox.showerror("錯誤", f"發生問題：\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    CardApp(root)
    root.mainloop()