import os
import sys
import json
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

        self.app_dir = os.path.dirname(os.path.abspath(__file__))
        self.settings_file = os.path.join(self.app_dir, "app_settings.json")
        self.settings = self.load_settings()

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value=self.settings.get("default_sheet_name", "p_zta"))
        self.output_dir = tk.StringVar()

        self._build_ui()

    def load_settings(self):
        default_settings = {
            "sender_actor_code": "",
            "sender_node_id": "",
            "service_access_token": "",
            "default_sheet_name": "p_zta"
        }

        if not os.path.exists(self.settings_file):
            return default_settings

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                saved = json.load(f)
            default_settings.update(saved)
        except Exception:
            pass

        return default_settings

    def save_settings(self, settings_data):
        with open(self.settings_file, "w", encoding="utf-8") as f:
            json.dump(settings_data, f, ensure_ascii=False, indent=2)

        self.settings = settings_data

        if self.settings.get("default_sheet_name"):
            self.sheet_name.set(self.settings["default_sheet_name"])

    def _build_ui(self):
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill="both", expand=True)

        title_bar = ttk.Frame(container)
        title_bar.pack(fill="x", pady=(0, 12))

        title = ttk.Label(
            title_bar,
            text="UDI Bulk Upload XML 產生器",
            font=("Microsoft JhengHei", 16, "bold")
        )
        title.pack(side="left")

        self.settings_button = ttk.Button(
            title_bar,
            text="基本資料設定",
            command=self.open_settings_window
        )
        self.settings_button.pack(side="right")

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

        self.open_output_button = ttk.Button(
            button_group,
            text="開啟輸出資料夾",
            command=self.open_output_dir
        )
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

        config = {
            "sender_actor_code": self.settings.get("sender_actor_code", "").strip(),
            "sender_node_id": self.settings.get("sender_node_id", "").strip(),
            "service_access_token": self.settings.get("service_access_token", "").strip()
        }

        if not config["sender_actor_code"]:
            messagebox.showwarning("缺少設定", "請先到「基本資料設定」填寫 Sender Actor Code。")
            return

        if not config["sender_node_id"]:
            messagebox.showwarning("缺少設定", "請先到「基本資料設定」填寫 Sender Node ID。")
            return

        if not config["service_access_token"]:
            messagebox.showwarning("缺少設定", "請先到「基本資料設定」填寫 Service Access Token。")
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
                self.sheet_name.get().strip() or None,
                config=config
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

    def open_settings_window(self):
        window = tk.Toplevel(self.root)
        window.title("基本資料設定")
        window.geometry("460x280")
        window.resizable(False, False)
        window.transient(self.root)
        window.grab_set()

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)

        sender_actor_code = tk.StringVar(value=self.settings.get("sender_actor_code", ""))
        sender_node_id = tk.StringVar(value=self.settings.get("sender_node_id", ""))
        service_access_token = tk.StringVar(value=self.settings.get("service_access_token", ""))
        default_sheet_name = tk.StringVar(value=self.settings.get("default_sheet_name", "p_zta"))

        ttk.Label(frame, text="Sender Actor Code：").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=sender_actor_code, width=34).grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="Sender Node ID：").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=sender_node_id, width=34).grid(row=1, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="Service Access Token：").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=service_access_token, width=34, show="*").grid(row=2, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="預設工作表：").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=default_sheet_name, width=34).grid(row=3, column=1, sticky="ew", padx=8)

        frame.columnconfigure(1, weight=1)

        button_bar = ttk.Frame(frame)
        button_bar.grid(row=4, column=0, columnspan=2, pady=(20, 0))

        def save_and_close():
            settings_data = {
                "sender_actor_code": sender_actor_code.get().strip(),
                "sender_node_id": sender_node_id.get().strip(),
                "service_access_token": service_access_token.get().strip(),
                "default_sheet_name": default_sheet_name.get().strip() or "p_zta"
            }

            self.save_settings(settings_data)
            self.log("基本資料已儲存。")
            messagebox.showinfo("完成", "基本資料已儲存。", parent=window)
            window.destroy()

        ttk.Button(button_bar, text="儲存", command=save_and_close).pack(side="left", padx=6, ipadx=12)
        ttk.Button(button_bar, text="取消", command=window.destroy).pack(side="left", padx=6, ipadx=12)


if __name__ == "__main__":
    root = tk.Tk()
    app = UDIUploadUI(root)
    root.mainloop()