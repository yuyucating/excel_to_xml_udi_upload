import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class UDIUploadUI:
    def __init__(self, root):
        self.root = root
        self.root.title("UDI Bulk Upload XML 產生器")
        self.root.geometry("720x420")
        self.root.minsize(640, 360)

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value="p_zta")
        self.output_dir = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title = ttk.Label(
            container,
            text="UDI Bulk Upload XML 產生器",
            font=("Microsoft JhengHei", 16, "bold")
        )
        title.pack(anchor="w", pady=(0, 12))

        form = ttk.LabelFrame(container, text="基本設定", padding=12)
        form.pack(fill="x")

        # Excel 檔案
        ttk.Label(form, text="Excel 檔案：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.excel_path, width=62).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_excel).grid(row=0, column=2, pady=6)

        # 工作表
        ttk.Label(form, text="工作表名稱：").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.sheet_name, width=20).grid(row=1, column=1, sticky="w", padx=8)

        # 輸出資料夾
        ttk.Label(form, text="輸出資料夾：").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_dir, width=62).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_output_dir).grid(row=2, column=2, pady=6)

        form.columnconfigure(1, weight=1)

        # 開始按鈕
        button_frame = ttk.Frame(container)
        button_frame.pack(fill="x", pady=16)

        self.start_button = ttk.Button(button_frame, text="開始", command=self.start_process)
        self.start_button.pack(anchor="center", ipadx=18, ipady=8)

        # 狀態 / 記錄區
        log_frame = ttk.LabelFrame(container, text="執行狀態", padding=12)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=10, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log("UI 已啟動，請選擇 Excel 檔案與輸出資料夾。")

    def select_excel(self):
        import os
        path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            self.excel_path.set(path)
            self.log(f"已選擇 Excel：{path}")

            # 預設將輸出資料夾帶入 Excel 檔案所在位置
            default_output_dir = os.path.dirname(path)
            self.output_dir.set(default_output_dir)
            self.log(f"輸出資料夾已自動帶入：{default_output_dir}")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="選擇輸出資料夾")
        if path:
            self.output_dir.set(path)
            self.log(f"已選擇輸出資料夾：{path}")

    def start_process(self):
        if not self.excel_path.get().strip():
            messagebox.showwarning("缺少資料", "請先選擇 Excel 檔案。")
            return

        if not self.output_dir.get().strip():
            messagebox.showwarning("缺少資料", "請先選擇輸出資料夾。")
            return

        # 這裡先做 UI 示範，後續再串接你的既有程式邏輯
        self.log("開始執行...")
        self.log(f"Excel 檔案：{self.excel_path.get()}")
        self.log(f"工作表名稱：{self.sheet_name.get()}")
        self.log(f"輸出資料夾：{self.output_dir.get()}")
        self.log("目前為 UI 雛形版本，下一步可直接串接 excel_to_df() / df_to_dict() / df_to_xml_files()。")
        messagebox.showinfo("完成", "UI 雛形已建立，開始按鈕可正常觸發。")

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")


if __name__ == "__main__":
    root = tk.Tk()
    app = UDIUploadUI(root)
    root.mainloop()
