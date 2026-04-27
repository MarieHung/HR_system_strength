import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import pandas as pd
import os
from datetime import datetime

class PersonnelSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("人事資料管理系統")
        self.root.geometry("1200x900")
        
        # 檔案路徑
        self.file_path_tab1 = "人事資料調閱紀錄.xlsx"
        self.file_path_tab2 = "新進教師登記_20250801.xlsx"
        self.file_path_tab3 = "各系所教師_跑程式用.xlsx"
        
        # 欄位定義
        self.columns_tab1 = ["日期", "調閱人", "單位", "被調閱名單", "歸還", "備註"]
        self.columns_tab2 = [
            "姓名", "服務單位", "一級單位", "身分證統一編號", "護照號碼", 
            "英文姓名", "職稱", "出生年月日", "到校日期", "戶籍地址", 
            "名冊地址", "現居地址", "通訊電話", "校外電子信箱", "國籍", 
            "學術專長及研究【以35字為限(含標點符號)】"
        ]
        self.columns_tab3 = ["姓名", "服務單位", "職稱", "證書字號", "發證日期"]
        
        # 確保檔案存在
        self.init_excels()

        # --- 使用 Notebook 製作分頁 ---
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab1 = tk.Frame(self.notebook)
        self.tab2 = tk.Frame(self.notebook)
        self.tab3 = tk.Frame(self.notebook)

        self.notebook.add(self.tab1, text="人事資料借閱系統")
        self.notebook.add(self.tab2, text="新進教師登記")
        self.notebook.add(self.tab3, text="教師證書查詢")

        # 初始化各個分頁內容
        self.setup_tab1()
        self.setup_tab2()
        self.setup_tab3()

    def init_excels(self):
        """初始化 Excel 檔案"""
        for path, cols in [(self.file_path_tab1, self.columns_tab1), (self.file_path_tab2, self.columns_tab2)]:
            if not os.path.exists(path):
                df = pd.DataFrame(columns=cols)
                df.to_excel(path, index=False)

    # --- 分頁 1 邏輯 ---
    def setup_tab1(self):
        title_label = tk.Label(self.tab1, text="人事資料借閱系統", font=("Microsoft JhengHei", 20, "bold"), pady=10)
        title_label.pack()

        input_frame = tk.LabelFrame(self.tab1, text="新增調閱紀錄", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        input_frame.pack(fill="x", padx=20, pady=5)

        self.entries_tab1 = {}
        fields = [
            ("日期", 0, 0, datetime.now().strftime("%Y%m%d")),
            ("調閱人", 0, 2, ""),
            ("單位", 1, 0, ""),
            ("被調閱名單", 1, 2, ""),
            ("備註", 2, 0, "")
        ]

        for text, row, col, default in fields:
            tk.Label(input_frame, text=text + ":", font=("Microsoft JhengHei", 10)).grid(row=row, column=col, sticky="e", padx=5, pady=5)
            entry = tk.Entry(input_frame, font=("Microsoft JhengHei", 10), width=25)
            entry.insert(0, default)
            entry.grid(row=row, column=col+1, padx=5, pady=5)
            self.entries_tab1[text] = entry

        tk.Label(input_frame, text="歸還:", font=("Microsoft JhengHei", 10)).grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.return_status = ttk.Combobox(input_frame, values=["Y", "N"], width=23, state="readonly")
        self.return_status.set("Y")
        self.return_status.grid(row=2, column=3, padx=5, pady=5)

        btn_frame = tk.Frame(self.tab1)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="新增紀錄", command=self.add_record_tab1, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="歸檔註記", command=self.mark_as_archived, bg="#FF9800", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).grid(row=0, column=1, padx=10)
        tk.Button(btn_frame, text="清除輸入", command=self.clear_fields_tab1, font=("Microsoft JhengHei", 10), padx=20).grid(row=0, column=2, padx=10)

        search_frame = tk.LabelFrame(self.tab1, text="快速查詢", font=("Microsoft JhengHei", 12), padx=20, pady=5)
        search_frame.pack(fill="x", padx=20, pady=5)
        self.search_entry_tab1 = tk.Entry(search_frame, font=("Microsoft JhengHei", 10), width=30)
        self.search_entry_tab1.pack(side="left", padx=5)
        tk.Button(search_frame, text="搜尋", command=lambda: self.perform_search(self.tree1, self.file_path_tab1, self.search_entry_tab1, self.columns_tab1), bg="#2196F3", fg="white").pack(side="left", padx=5)
        tk.Button(search_frame, text="顯示全部", command=lambda: self.load_data_to_tree(self.tree1, self.file_path_tab1, self.columns_tab1)).pack(side="left", padx=5)

        self.tree1 = self.create_treeview(self.tab1, self.columns_tab1)
        self.load_data_to_tree(self.tree1, self.file_path_tab1, self.columns_tab1)

    # --- 分頁 2 邏輯 ---
    def setup_tab2(self):
        title_label = tk.Label(self.tab2, text="新進教師登記管理", font=("Microsoft JhengHei", 20, "bold"), pady=10)
        title_label.pack()

        # 輸入區域 (網格佈局，分兩欄顯示 16 個欄位)
        input_frame = tk.LabelFrame(self.tab2, text="手動登記新進教師", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        input_frame.pack(fill="x", padx=20, pady=5)

        self.entries_tab2 = {}
        for i, col_name in enumerate(self.columns_tab2):
            row = i // 2
            col_pos = (i % 2) * 2
            tk.Label(input_frame, text=col_name + ":", font=("Microsoft JhengHei", 9)).grid(row=row, column=col_pos, sticky="e", padx=5, pady=2)
            
            # 對於最後一個字數限制欄位，加長顯示
            width = 40 if "學術專長" in col_name else 25
            entry = tk.Entry(input_frame, font=("Microsoft JhengHei", 9), width=width)
            entry.grid(row=row, column=col_pos+1, sticky="w", padx=5, pady=2)
            self.entries_tab2[col_name] = entry

        # 按鈕
        btn_frame = tk.Frame(self.tab2)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="新增教師紀錄", command=self.add_record_tab2, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="清除輸入", command=self.clear_fields_tab2, font=("Microsoft JhengHei", 10), padx=20).grid(row=0, column=1, padx=10)

        # 搜尋區域
        search_frame = tk.LabelFrame(self.tab2, text="資料查詢", font=("Microsoft JhengHei", 12), padx=20, pady=5)
        search_frame.pack(fill="x", padx=20, pady=5)
        self.search_entry_tab2 = tk.Entry(search_frame, font=("Microsoft JhengHei", 10), width=30)
        self.search_entry_tab2.pack(side="left", padx=5)
        tk.Button(search_frame, text="搜尋", command=lambda: self.perform_search(self.tree2, self.file_path_tab2, self.search_entry_tab2, self.columns_tab2), bg="#2196F3", fg="white").pack(side="left", padx=5)
        tk.Button(search_frame, text="重新讀取檔案", command=lambda: self.load_data_to_tree(self.tree2, self.file_path_tab2, self.columns_tab2)).pack(side="left", padx=5)

        self.tree2 = self.create_treeview(self.tab2, self.columns_tab2)
        self.load_data_to_tree(self.tree2, self.file_path_tab2, self.columns_tab2)

    # --- 通用工具函式 ---
    def create_treeview(self, parent, columns):
        container = tk.Frame(parent)
        container.pack(fill="both", expand=True, padx=20, pady=10)
        
        tree = ttk.Treeview(container, columns=columns, show="headings")
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100, anchor="center")
        
        vsb = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        return tree

    def load_data_to_tree(self, tree, file_path, columns):
        for item in tree.get_children():
            tree.delete(item)
        try:
            if not os.path.exists(file_path): return
            df = pd.read_excel(file_path)
            df = df.fillna("")
            # 確保欄位一致性
            for col in columns:
                if col not in df.columns: df[col] = ""
            
            for index, row in df.iterrows():
                tree.insert("", "end", iid=str(index), values=list(row[columns]))
        except Exception as e:
            messagebox.showerror("錯誤", f"讀取 {file_path} 失敗: {e}")

    def perform_search(self, tree, file_path, search_entry, columns):
        query = search_entry.get().strip().lower()
        if not query:
            self.load_data_to_tree(tree, file_path, columns)
            return
        
        for item in tree.get_children(): tree.delete(item)
        try:
            df = pd.read_excel(file_path).fillna("")
            mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(query)).any(axis=1)
            filtered_df = df[mask]
            for index, row in filtered_df.iterrows():
                tree.insert("", "end", iid=str(index), values=list(row[columns]))
        except Exception as e:
            messagebox.showerror("錯誤", f"搜尋失敗: {e}")

    # --- Tab 1 新增邏輯 ---
    def add_record_tab1(self):
        data = {col: self.entries_tab1[col].get() for col in self.entries_tab1}
        data["歸還"] = self.return_status.get()
        if not data["調閱人"] or not data["被調閱名單"]:
            messagebox.showwarning("警告", "請填寫調閱人與名單！")
            return
        self.save_to_excel(self.file_path_tab1, data, self.tree1, self.columns_tab1)
        self.clear_fields_tab1()

    # --- Tab 2 新增邏輯 ---
    def add_record_tab2(self):
        # 檢查字數限制
        research_col = "學術專長及研究【以35字為限(含標點符號)】"
        if len(self.entries_tab2[research_col].get()) > 35:
            messagebox.showwarning("警告", "『學術專長及研究』請勿超過 35 字！")
            return
        
        data = {col: self.entries_tab2[col].get() for col in self.columns_tab2}
        if not data["姓名"]:
            messagebox.showwarning("警告", "『姓名』為必填欄位！")
            return
            
        self.save_to_excel(self.file_path_tab2, data, self.tree2, self.columns_tab2)
        messagebox.showinfo("成功", "新進教師資料已成功新增！")
        self.clear_fields_tab2()

    def save_to_excel(self, file_path, data, tree, columns):
        try:
            # 嘗試讀取現有檔案
            if os.path.exists(file_path):
                df = pd.read_excel(file_path, engine='openpyxl')
            else:
                df = pd.DataFrame(columns=columns)
            
            new_df = pd.DataFrame([data])
            df = pd.concat([df, new_df], ignore_index=True)
            
            # 儲存檔案
            df.to_excel(file_path, index=False, engine='openpyxl')
            self.load_data_to_tree(tree, file_path, columns)
            return True
        except PermissionError:
            messagebox.showerror("儲存失敗", f"無法寫入檔案：\n{file_path}\n\n請確認該 Excel 檔案是否已關閉！")
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存時發生非預期錯誤: {e}")
        return False

    def clear_fields_tab1(self):
        for entry in self.entries_tab1.values(): entry.delete(0, tk.END)
        self.entries_tab1["日期"].insert(0, datetime.now().strftime("%Y%m%d"))
        self.return_status.set("Y")

    def clear_fields_tab2(self):
        for entry in self.entries_tab2.values(): entry.delete(0, tk.END)

    # --- 分頁 3 邏輯 (教師證書查詢) ---
    def setup_tab3(self):
        title_label = tk.Label(self.tab3, text="教師證書查詢系統", font=("Microsoft JhengHei", 20, "bold"), pady=10)
        title_label.pack()

        search_frame = tk.LabelFrame(self.tab3, text="姓名搜尋", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        search_frame.pack(fill="x", padx=20, pady=5)

        tk.Label(search_frame, text="請輸入教師姓名:", font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        self.search_entry_tab3 = tk.Entry(search_frame, font=("Microsoft JhengHei", 10), width=30)
        self.search_entry_tab3.pack(side="left", padx=5)
        self.search_entry_tab3.bind("<Return>", lambda e: self.perform_search_tab3())

        tk.Button(search_frame, text="搜尋", command=self.perform_search_tab3, bg="#2196F3", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).pack(side="left", padx=10)
        tk.Button(search_frame, text="清空結果", command=lambda: self.load_data_to_tree(self.tree3, self.file_path_tab3, self.columns_tab3, sheet_name="Sheet0")).pack(side="left", padx=5)

        self.tree3 = self.create_treeview(self.tab3, self.columns_tab3)
        # 初始不載入全部資料，避免過慢，或只載入前 100 筆
        # self.load_data_to_tree(self.tree3, self.file_path_tab3, self.columns_tab3, sheet_name="Sheet0")

    def perform_search_tab3(self):
        query = self.search_entry_tab3.get().strip()
        if not query:
            messagebox.showwarning("警告", "請輸入姓名！")
            return
        
        for item in self.tree3.get_children(): self.tree3.delete(item)
        
        try:
            # 遍歷所有 Sheet 尋找該姓名
            xls = pd.ExcelFile(self.file_path_tab3)
            found = False
            
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name).fillna("")
                
                # 尋找姓名欄位
                name_col = next((c for c in df.columns if "姓名" in str(c)), None)
                if not name_col: continue
                
                # 篩選
                mask = df[name_col].astype(str).str.contains(query)
                results = df[mask]
                
                if not results.empty:
                    found = True
                    # 嘗試提取需要的資訊 (職級, 證書字號等)
                    cert_col = next((c for c in results.columns if any(k in str(c) for k in ["證書字號", "證號", "證書號"])), "無資料")
                    rank_col = next((c for c in results.columns if any(k in str(c) for k in ["職稱", "職級", "等級"])), "無資料")
                    date_col = next((c for c in results.columns if any(k in str(c) for k in ["發證日期", "起資日期", "審定日期"])), "無資料")
                    dept_col = next((c for c in results.columns if any(k in str(c) for k in ["服務單位", "系所", "單位"])), "無資料")

                    for _, row in results.iterrows():
                        display_values = [
                            row[name_col],
                            row[rank_col] if rank_col != "無資料" else "",
                            row[cert_col] if cert_col != "無資料" else "",
                            row[date_col] if date_col != "無資料" else "",
                            row[dept_col] if dept_col != "無資料" else ""
                        ]
                        self.tree3.insert("", "end", values=display_values)
            
            if not found:
                messagebox.showinfo("搜尋結果", f"找不到關於「{query}」的證書資料。")
                
        except Exception as e:
            messagebox.showerror("錯誤", f"搜尋教師證書失敗: {e}")

    def mark_as_archived(self):
        selected = self.tree1.selection()
        if not selected:
            messagebox.showwarning("警告", "請先選擇紀錄！")
            return
        new_remark = simpledialog.askstring("歸檔註記", "請輸入歸檔備註:")
        if new_remark is None: return
        try:
            df = pd.read_excel(self.file_path_tab1)
            for item_id in selected:
                idx = int(item_id)
                df.at[idx, "歸還"] = "Y"
                df.at[idx, "備註"] = new_remark
            df.to_excel(self.file_path_tab1, index=False)
            self.load_data_to_tree(self.tree1, self.file_path_tab1, self.columns_tab1)
        except Exception as e:
            messagebox.showerror("錯誤", f"更新失敗: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PersonnelSystem(root)
    root.mainloop()
