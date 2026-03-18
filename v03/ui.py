import os
import sys
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from transfer_data import export_excel_to_xml


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

        ttk.Label(form, text="Excel 檔案：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.excel_path, width=62).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_excel).grid(row=0, column=2, pady=6)

        ttk.Label(form, text="工作表名稱：").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.sheet_name, width=20).grid(row=1, column=1, sticky="w", padx=8)

        ttk.Label(form, text="輸出資料夾：").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_dir, width=62).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_output_dir).grid(row=2, column=2, pady=6)

        form.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(container)
        button_frame.pack(fill="x", pady=16)

        button_group = ttk.Frame(button_frame)
        button_group.pack(anchor="center")

        self.start_button = ttk.Button(button_group, text="開始", command=self.start_process)
        self.start_button.pack(side="left", padx=6, ipadx=18, ipady=8)

        self.open_output_button = ttk.Button(button_group, text="開啟輸出資料夾", command=self.open_output_dir)
        self.open_output_button.pack(side="left", padx=6, ipadx=12, ipady=8)

        log_frame = ttk.LabelFrame(container, text="執行狀態", padding=12)
        log_frame.pack(fill="both", expand=True)

        self.log_text = tk.Text(log_frame, height=10, wrap="word", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)
        self.log("UI 已啟動，請選擇 Excel 檔案與輸出資料夾。")

    def select_excel(self):
        path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if path:
            self.excel_path.set(path)
            self.log(f"已選擇 Excel：{path}")

            base_dir = os.path.dirname(path)
            default_output_dir = os.path.join(base_dir, "output_xml")
            self.output_dir.set(default_output_dir)
            self.log(f"輸出資料夾已自動帶入：{default_output_dir}")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="選擇輸出資料夾")
        if path:
            self.output_dir.set(path)
            self.log(f"已選擇輸出資料夾：{path}")

    def open_output_dir(self):
        output_dir = self.output_dir.get().strip()
        if not output_dir:
            messagebox.showwarning("缺少資料", "目前沒有可開啟的輸出資料夾。")
            return

        if not os.path.exists(output_dir):
            messagebox.showwarning("資料夾不存在", "輸出資料夾尚未建立，請先執行轉檔或重新選擇資料夾。")
            return

        try:
            if sys.platform.startswith("win"):
                os.startfile(output_dir)
            elif sys.platform == "darwin":
                subprocess.run(["open", output_dir], check=False)
            else:
                subprocess.run(["xdg-open", output_dir], check=False)

            self.log(f"已開啟輸出資料夾：{output_dir}")
        except Exception as e:
            self.log(f"開啟輸出資料夾失敗：{e}")
            messagebox.showerror("錯誤", f"無法開啟輸出資料夾：\n{e}")

    def start_process(self):
        if not self.excel_path.get().strip():
            messagebox.showwarning("缺少資料", "請先選擇 Excel 檔案。")
            return

        if not self.output_dir.get().strip():
            messagebox.showwarning("缺少資料", "請先選擇輸出資料夾。")
            return

        try:
            self.start_button.config(state="disabled")
            self.log("開始執行...")
            self.log(f"Excel 檔案：{self.excel_path.get()}")
            self.log(f"工作表名稱：{self.sheet_name.get()}")
            self.log(f"輸出資料夾：{self.output_dir.get()}")

            result = export_excel_to_xml(
                self.excel_path.get(),
                self.output_dir.get(),
                self.sheet_name.get().strip() or None
            )

            self.log(f'已讀取資料，共 {result["df_count"]} 筆符合條件。')
            self.log(f'已轉換裝置資料，共 {result["device_count"]} 筆。')
            self.log(f'XML 檔案輸出完成：{result["output_path"]}')
            messagebox.showinfo("完成", f'XML 檔案已成功產生：\n{result["output_path"]}')

        except Exception as e:
            self.log("執行失敗。")
            self.log(str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("錯誤", f"執行失敗：\n{e}")
        finally:
            self.start_button.config(state="normal")

    def log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")


if __name__ == "__main__":
    root = tk.Tk()
    app = UDIUploadUI(root)
    root.mainloop()