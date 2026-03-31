import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from datetime import datetime

class PersonnelSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("人事管理與教師統計系統")
        self.root.geometry("1000x800")
        
        # 檔案路徑
        self.file_retrieval = "人事資料調閱紀錄.xlsx"
        self.file_teacher = "各系所教師_跑程式用.xlsx"
        
        # 建立 Notebook (分頁控制)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)
        
        # 建立分頁框架
        self.tab1 = tk.Frame(self.notebook)
        self.tab2 = tk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text=" 人事借閱系統 ")
        self.notebook.add(self.tab2, text=" 教師人數統計 ")
        
        # 初始化分頁內容
        self.setup_tab1()
        self.setup_tab2()

    # ========================== 分頁 1: 人事借閱系統 ==========================
    def setup_tab1(self):
        self.columns = ["日期", "調閱人", "單位", "被調閱名單", "歸還", "備註"]
        self.init_excel_retrieval()

        # 標題
        tk.Label(self.tab1, text="人事資料借閱系統", font=("Microsoft JhengHei", 20, "bold"), pady=15).pack()

        # 輸入區
        input_frame = tk.LabelFrame(self.tab1, text="新增調閱紀錄", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        input_frame.pack(fill="x", padx=20, pady=5)

        self.entries = {}
        fields = [("日期", 0, 0, datetime.now().strftime("%Y%m%d")), ("調閱人", 0, 2, ""),
                  ("單位", 1, 0, ""), ("被調閱名單", 1, 2, ""), ("備註", 2, 0, "")]

        for text, row, col, default in fields:
            tk.Label(input_frame, text=text + ":", font=("Microsoft JhengHei", 10)).grid(row=row, column=col, sticky="e", padx=5, pady=5)
            entry = tk.Entry(input_frame, font=("Microsoft JhengHei", 10), width=25)
            entry.insert(0, default)
            entry.grid(row=row, column=col+1, padx=5, pady=5)
            self.entries[text] = entry

        tk.Label(input_frame, text="歸還:", font=("Microsoft JhengHei", 10)).grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.return_status = ttk.Combobox(input_frame, values=["Y", "N"], width=23, state="readonly")
        self.return_status.set("Y")
        self.return_status.grid(row=2, column=3, padx=5, pady=5)

        btn_frame = tk.Frame(self.tab1)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="新增紀錄", command=self.add_record, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="清除輸入", command=self.clear_fields, font=("Microsoft JhengHei", 10), padx=20).grid(row=0, column=1, padx=10)

        # 搜尋區
        search_frame = tk.LabelFrame(self.tab1, text="快速查詢", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        search_frame.pack(fill="x", padx=20, pady=5)
        tk.Label(search_frame, text="請輸入關鍵字:", font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, font=("Microsoft JhengHei", 10), width=30)
        self.search_entry.pack(side="left", padx=5)
        tk.Button(search_frame, text="搜尋", command=self.perform_search, bg="#2196F3", fg="white").pack(side="left", padx=5)
        tk.Button(search_frame, text="顯示全部", command=self.load_data_retrieval).pack(side="left", padx=5)

        # 表格
        table_frame = tk.Frame(self.tab1)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)
        self.load_data_retrieval()

    def init_excel_retrieval(self):
        if not os.path.exists(self.file_retrieval):
            pd.DataFrame(columns=self.columns).to_excel(self.file_retrieval, index=False)

    def load_data_retrieval(self):
        self.search_entry.delete(0, tk.END)
        for item in self.tree.get_children(): self.tree.delete(item)
        try:
            df = pd.read_excel(self.file_retrieval).fillna("")
            for _, row in df.iterrows(): self.tree.insert("", "end", values=list(row))
        except: pass

    def add_record(self):
        data = {k: v.get() for k, v in self.entries.items()}
        data["歸還"] = self.return_status.get()
        if not data["調閱人"] or not data["被調閱名單"]:
            messagebox.showwarning("警告", "請填寫調閱人與名單")
            return
        try:
            df = pd.read_excel(self.file_retrieval)
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
            df.to_excel(self.file_retrieval, index=False)
            messagebox.showinfo("成功", "紀錄已新增")
            self.clear_fields()
            self.load_data_retrieval()
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def clear_fields(self):
        for k, e in self.entries.items():
            e.delete(0, tk.END)
            if k == "日期": e.insert(0, datetime.now().strftime("%Y%m%d"))
        self.return_status.set("Y")

    def perform_search(self):
        query = self.search_entry.get().strip().lower()
        if not query: return self.load_data_retrieval()
        for item in self.tree.get_children(): self.tree.delete(item)
        try:
            df = pd.read_excel(self.file_retrieval).fillna("")
            mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(query)).any(axis=1)
            for _, row in df[mask].iterrows(): self.tree.insert("", "end", values=list(row))
        except: pass

    # ========================== 分頁 2: 教師人數統計系統 ==========================
    def setup_tab2(self):
        self.teacher_df = None
        if not os.path.exists(self.file_teacher):
            tk.Label(self.tab2, text="找不到 各系所教師_跑程式用.xlsx", fg="red").pack(pady=20)
            return

        # 頂部控制
        top_frame = tk.Frame(self.tab2, pady=10)
        top_frame.pack(fill="x", padx=20)
        tk.Label(top_frame, text="選擇工作表 (系所):", font=("Microsoft JhengHei", 11, "bold")).pack(side="left")
        
        xl = pd.ExcelFile(self.file_teacher)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(top_frame, textvariable=self.sheet_var, values=xl.sheet_names, state="readonly", width=35)
        self.sheet_combo.pack(side="left", padx=10)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_teacher_sheet_selected)

        # 主要內容區 (三欄佈局)
        content_frame = tk.Frame(self.tab2)
        content_frame.pack(fill="both", expand=True, padx=20)

        # 1. 人員類別欄 (左)
        self.cat_frame = tk.LabelFrame(content_frame, text="人員類別 (勾選)", font=("Microsoft JhengHei", 10, "bold"))
        self.cat_frame.place(relx=0, rely=0, relwidth=0.3, relheight=0.8)
        self.cat_canvas = tk.Canvas(self.cat_frame); self.cat_scroll = ttk.Scrollbar(self.cat_frame, command=self.cat_canvas.yview)
        self.cat_list_frame = tk.Frame(self.cat_canvas)
        self.cat_canvas.create_window((0,0), window=self.cat_list_frame, anchor="nw")
        self.cat_canvas.configure(yscrollcommand=self.cat_scroll.set)
        self.cat_canvas.pack(side="left", fill="both", expand=True); self.cat_scroll.pack(side="right", fill="y")

        # 2. 職稱欄 (中)
        self.title_frame = tk.LabelFrame(content_frame, text="職稱 (勾選)", font=("Microsoft JhengHei", 10, "bold"))
        self.title_frame.place(relx=0.32, rely=0, relwidth=0.3, relheight=0.8)
        self.title_canvas = tk.Canvas(self.title_frame); self.title_scroll = ttk.Scrollbar(self.title_frame, command=self.title_canvas.yview)
        self.title_list_frame = tk.Frame(self.title_canvas)
        self.title_canvas.create_window((0,0), window=self.title_list_frame, anchor="nw")
        self.title_canvas.configure(yscrollcommand=self.title_scroll.set)
        self.title_canvas.pack(side="left", fill="both", expand=True); self.title_scroll.pack(side="right", fill="y")

        # 3. 統計結果欄 (右)
        self.res_frame = tk.LabelFrame(content_frame, text="統計結果", font=("Microsoft JhengHei", 10, "bold"))
        self.res_frame.place(relx=0.64, rely=0, relwidth=0.36, relheight=0.8)
        self.res_text = tk.Text(self.res_frame, font=("Microsoft JhengHei", 10), state="disabled", bg="#f9f9f9")
        self.res_text.pack(fill="both", expand=True, padx=5, pady=5)

        # 底端按鈕
        tk.Button(self.tab2, text="開始計算", command=self.calculate_teacher_stats, bg="#4CAF50", fg="white", 
                  font=("Microsoft JhengHei", 12, "bold"), pady=10).pack(side="bottom", fill="x", padx=20, pady=20)

        self.cat_vars = {}; self.title_vars = {}

    def on_teacher_sheet_selected(self, event):
        sheet = self.sheet_var.get()
        try:
            self.teacher_df = pd.read_excel(self.file_teacher, sheet_name=sheet)
            # 模糊識別欄位
            self.col_cat = self.find_col(self.teacher_df, ["人員類別", "鈭箏憵", "人員"])
            self.col_title = self.find_col(self.teacher_df, ["職稱", "瑞迂"])
            
            # 更新勾選框
            self.build_checkboxes(self.cat_list_frame, self.cat_vars, self.col_cat, self.cat_canvas)
            self.build_checkboxes(self.title_list_frame, self.title_vars, self.col_title, self.title_canvas)
        except Exception as e: messagebox.showerror("錯誤", str(e))

    def find_col(self, df, names):
        for n in names:
            for c in df.columns:
                if n in str(c): return c
        return None

    def build_checkboxes(self, frame, var_dict, col_name, canvas):
        for w in frame.winfo_children(): w.destroy()
        var_dict.clear()
        if col_name and self.teacher_df is not None:
            items = sorted(self.teacher_df[col_name].dropna().unique().astype(str))
            for item in items:
                v = tk.BooleanVar()
                var_dict[item] = v
                tk.Checkbutton(frame, text=item, variable=v, font=("Microsoft JhengHei", 9)).pack(anchor="w")
        frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

    def calculate_teacher_stats(self):
        if self.teacher_df is None: return
        sel_cats = [k for k, v in self.cat_vars.items() if v.get()]
        sel_titles = [k for k, v in self.title_vars.items() if v.get()]
        
        if not sel_cats or not sel_titles:
            messagebox.showwarning("提示", "請至少勾選一項人員類別與一項職稱")
            return

        df = self.teacher_df.copy()
        df[self.col_cat] = df[self.col_cat].astype(str)
        df[self.col_title] = df[self.col_title].astype(str)
        
        filtered = df[(df[self.col_cat].isin(sel_cats)) & (df[self.col_title].isin(sel_titles))]
        count = len(filtered)

        # 顯示結果
        self.res_text.config(state="normal")
        self.res_text.delete(1.0, tk.END)
        self.res_text.insert(tk.END, f"【統計報告】\n\n")
        self.res_text.insert(tk.END, f"系所: {self.sheet_var.get()}\n")
        self.res_text.insert(tk.END, f"--------------------------\n")
        self.res_text.insert(tk.END, f"已勾選類別:\n {', '.join(sel_cats)}\n\n")
        self.res_text.insert(tk.END, f"已勾選職稱:\n {', '.join(sel_titles)}\n")
        self.res_text.insert(tk.END, f"--------------------------\n")
        self.res_text.insert(tk.END, f" 總計人數: {count} 人", "bold")
        self.res_text.tag_configure("bold", font=("Microsoft JhengHei", 12, "bold"), foreground="blue")
        self.res_text.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = PersonnelSystem(root)
    root.mainloop()
