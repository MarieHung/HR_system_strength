import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from datetime import datetime

class PersonnelSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("人事資料借閱系統")
        self.root.geometry("900x600")
        
        self.file_path = "人事資料調閱紀錄.xlsx"
        self.columns = ["日期", "調閱人", "單位", "被調閱名單", "歸還", "備註"]
        
        # 確保檔案存在
        self.init_excel()

        # --- UI 佈局 ---
        # 標題
        title_label = tk.Label(root, text="人事資料借閱系統", font=("Microsoft JhengHei", 20, "bold"), pady=20)
        title_label.pack()

        # 輸入區域容器
        input_frame = tk.LabelFrame(root, text="新增調閱紀錄", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        input_frame.pack(fill="x", padx=20, pady=10)

        # 建立輸入欄位
        self.entries = {}
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
            self.entries[text] = entry

        # 歸還狀態 (下拉選單)
        tk.Label(input_frame, text="歸還:", font=("Microsoft JhengHei", 10)).grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.return_status = ttk.Combobox(input_frame, values=["Y", "N"], width=23, state="readonly")
        self.return_status.set("Y")
        self.return_status.grid(row=2, column=3, padx=5, pady=5)

        # 按鈕
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=5)
        
        tk.Button(btn_frame, text="新增紀錄", command=self.add_record, bg="#4CAF50", fg="white", font=("Microsoft JhengHei", 10, "bold"), padx=20).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="清除輸入", command=self.clear_fields, font=("Microsoft JhengHei", 10), padx=20).grid(row=0, column=1, padx=10)

        # --- 查詢區域 ---
        search_frame = tk.LabelFrame(root, text="快速查詢", font=("Microsoft JhengHei", 12), padx=20, pady=10)
        search_frame.pack(fill="x", padx=20, pady=5)

        tk.Label(search_frame, text="請輸入關鍵字:", font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        self.search_entry = tk.Entry(search_frame, font=("Microsoft JhengHei", 10), width=30)
        self.search_entry.pack(side="left", padx=5)
        
        tk.Button(search_frame, text="搜尋", command=self.perform_search, bg="#2196F3", fg="white", font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)
        tk.Button(search_frame, text="顯示全部", command=self.load_data, font=("Microsoft JhengHei", 10)).pack(side="left", padx=5)

        # 表格區域
        table_frame = tk.Frame(root)
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)

        self.tree = ttk.Treeview(table_frame, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        
        # 滾動條
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

        self.load_data()

    def init_excel(self):
        """如果 Excel 不存在，建立一個新的"""
        if not os.path.exists(self.file_path):
            df = pd.DataFrame(columns=self.columns)
            df.to_excel(self.file_path, index=False)

    def load_data(self):
        """讀取 Excel 並顯示在表格中"""
        self.search_entry.delete(0, tk.END) # 重設搜尋框
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            df = pd.read_excel(self.file_path)
            # 確保欄位一致，補齊缺失欄位
            for col in self.columns:
                if col not in df.columns:
                    df[col] = ""
            
            df = df.fillna("") # 處理 NaN
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=list(row[self.columns]))
        except Exception as e:
            messagebox.showerror("錯誤", f"無法讀取檔案: {e}")

    def perform_search(self):
        """根據關鍵字過濾資料"""
        query = self.search_entry.get().strip().lower()
        if not query:
            self.load_data()
            return

        for item in self.tree.get_children():
            self.tree.delete(item)

        try:
            df = pd.read_excel(self.file_path)
            df = df.fillna("")
            
            # 全欄位關鍵字搜尋
            mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(query)).any(axis=1)
            filtered_df = df[mask]

            for _, row in filtered_df.iterrows():
                self.tree.insert("", "end", values=list(row[self.columns]))
            
            if len(filtered_df) == 0:
                messagebox.showinfo("搜尋結果", "找不到符合條件的紀錄。")
        except Exception as e:
            messagebox.showerror("錯誤", f"搜尋時發生錯誤: {e}")

    def add_record(self):
        """新增資料到 Excel"""
        data = {
            "日期": self.entries["日期"].get(),
            "調閱人": self.entries["調閱人"].get(),
            "單位": self.entries["單位"].get(),
            "被調閱名單": self.entries["被調閱名單"].get(),
            "歸還": self.return_status.get(),
            "備註": self.entries["備註"].get()
        }

        if not data["調閱人"] or not data["被調閱名單"]:
            messagebox.showwarning("警告", "請填寫『調閱人』與『被調閱名單』！")
            return

        try:
            df = pd.read_excel(self.file_path)
            new_df = pd.DataFrame([data])
            df = pd.concat([df, new_df], ignore_index=True)
            df.to_excel(self.file_path, index=False)
            
            messagebox.showinfo("成功", "紀錄已成功新增！")
            self.clear_fields()
            self.load_data()
        except Exception as e:
            messagebox.showerror("錯誤", f"寫入失敗: {e}")

    def clear_fields(self):
        """清空輸入框"""
        for text, entry in self.entries.items():
            entry.delete(0, tk.END)
            if text == "日期":
                entry.insert(0, datetime.now().strftime("%Y%m%d"))
        self.return_status.set("Y")

if __name__ == "__main__":
    root = tk.Tk()
    app = PersonnelSystem(root)
    root.mainloop()
