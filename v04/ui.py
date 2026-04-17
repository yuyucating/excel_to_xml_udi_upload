import os
import pandas as pd
import sys
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from transfer_data import export_excel_to_xml

def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

class UDIUploadUI:
    def __init__(self, root):
        #主畫面

        self.root = root
        self.root.title("UDI Bulk Upload XML 產生器")
        self.root.geometry("720x420")
        self.root.minsize(640, 360)

        self.app_dir = get_app_dir()
        self.settings_file = os.path.join(self.app_dir, "app_settings.json")
        self.settings = self.load_settings()

        self.excel_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value=self.settings.get("default_sheet_name", "p_zta"))
        self.output_dir = tk.StringVar()
        self.available_sheets = []

        # XML 輸出模式：DEVICE.POST / UDI_DI.POST
        self.export_mode = tk.StringVar(value="DEVICE_POST")

        self._build_ui()

    def load_settings(self):

        # 系統原廠預設欄位對應
        default_field_mapping = {
            "COMMON": {
                "reg_type": "tc_jsb030",
                "risk_class": "tc_jsb080",
                "model": "tc_jsb200",
                "certificate_no": "tc_jsb170",
                "animal_tissues_cells": "tc_jsb360",
                "human_tissues_cells": "tc_jsb350",
                "medicinal_product": "tc_jsb370",
                "human_product": "tc_jsb380",
                "active": "tc_jsb120",
                "administering": "tc_jsb130",
                "implantable": "tc_jsb090",
                "measuring": "tc_jsb100",
                "reusable": "tc_jsb110",
                "udi_di": "tc_jsb001",
                "udi_status": "tc_jsb270",
                "emdn_code": "tc_jsb190",
                "pi_lot_number": "tc_jsb2401",
                "pi_serial_number": "tc_jsb2411",
                "pi_manufacturing_date": "tc_jsb2421",
                "pi_expiration_date": "tc_jsb2431",
                "product_number": "tc_jsb000",
                "sterile": "tc_jsb743",
                "sterilization": "tc_jsb742",
                "trade_name": "tc_jsb200",
                "trade_name_lang": "tc_jsb210",
                "number_of_reuses": "tc_jsb744",
                "base_quantity": "tc_jsb230",
                "latex": "tc_jsb550",
                "reprocessed": "tc_jsb620",
                "spec_value": "tc_jsb430",
                "spec_unit": "tc_jsb440",
                "first_market": "tc_jsb390",
                "marketing_status": "tc_jsb400",
            },
            "MDD": {
                "basicudi_di": "tc_jsb630", # 2026-03-26// The original default was tc_jsb070, but it has been changed to tc_jsb630
                "critical_warning": "tc_jsb730",
                "certificate_revision": "tc_jsb180",
                "certificate_expiry": "tc_jsb710"
            },
            "MDR": {
                "basicudi_di": "tc_jsb070"
            }
        }

        default_settings = {
            "sender_actor_code": "",
            "sender_node_id": "",
            "service_access_token": "",
            "default_sheet_name": "p_zta",
            "MFActorCode": "",
            "ARActorCode": "",
            "field_mapping_defaults": json.loads(json.dumps(default_field_mapping)),
            "field_mapping": json.loads(json.dumps(default_field_mapping))
        }
        if not os.path.exists(self.settings_file):
            return default_settings

        try:
            with open(self.settings_file, "r", encoding="utf-8") as f:
                saved = json.load(f)

            default_settings.update(saved)


            for section in ("COMMON", "MDD", "MDR"):
                default_settings["field_mapping_defaults"][section].update(
                    saved.get("field_mapping_defaults", {}).get(section, {})
                )
                default_settings["field_mapping"][section].update(
                    saved.get("field_mapping", {}).get(section, {})
                )

        except Exception:
            messagebox.showwarning("設定檔異常", "設定檔讀取失敗，已載入預設值。")
            return default_settings

        return default_settings
    
    def load_sheet_names(self, excel_path):
        try:
            xls = pd.ExcelFile(excel_path)
            self.available_sheets = xls.sheet_names
            self.sheet_combo["values"] = self.available_sheets

            if not self.available_sheets:
                self.sheet_name.set("")
                self.log("找不到任何工作表。")
                return

            default_sheet = self.settings.get("default_sheet_name", "").strip()

            if default_sheet and default_sheet in self.available_sheets:
                self.sheet_name.set(default_sheet)
                self.log(f"已載入工作表，預設選擇：{default_sheet}")
            else:
                self.sheet_name.set(self.available_sheets[0])
                self.log(f"已載入工作表，預設選擇第一張：{self.available_sheets[0]}")

        except Exception as e:
            self.available_sheets = []
            self.sheet_combo["values"] = []
            self.sheet_name.set("")
            self.log(f"讀取工作表失敗：{e}")
            messagebox.showerror("錯誤", f"無法讀取 Excel 工作表：\n{e}")

    def save_settings(self, settings_data):
        try:
            with open(self.settings_file, "w", encoding="utf-8") as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)
            self.settings = settings_data
            if "default_sheet_name" in settings_data and self.settings.get("default_sheet_name"):
                self.sheet_name.set(self.settings["default_sheet_name"])
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存設定失敗：\n{e}")
            raise

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

        button_bar_top = ttk.Frame(title_bar)
        button_bar_top.pack(side="right")

        self.excel_fields_button = ttk.Button(
            button_bar_top,
            text="Excel 欄位設定",
            command=self.open_excel_fields_window
        )
        self.excel_fields_button.pack(side="right")

        self.settings_button = ttk.Button(
            button_bar_top,
            text="基本資料設定",
            command=self.open_settings_window
        )
        self.settings_button.pack(side="right", padx=(0, 8))

        form = ttk.LabelFrame(container, text="基本設定", padding=12)
        form.pack(fill="x")

        ttk.Label(form, text="Excel 檔案：").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.excel_path, width=62).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_excel).grid(row=0, column=2, pady=6)

        ttk.Label(form, text="工作表名稱：").grid(row=1, column=0, sticky="w", pady=6)

        self.sheet_combo = ttk.Combobox(
            form,
            textvariable=self.sheet_name,
            width=30,
            state="readonly"
        )
        self.sheet_combo.grid(row=1, column=1, sticky="w", padx=8)

        ttk.Label(form, text="輸出資料夾：").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(form, textvariable=self.output_dir, width=62).grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(form, text="瀏覽", command=self.select_output_dir).grid(row=2, column=2, pady=6)

        ttk.Label(form, text="輸出模式：").grid(row=3, column=0, sticky="nw", pady=6)
        mode_frame = ttk.Frame(form)
        mode_frame.grid(row=3, column=1, sticky="w", padx=8, pady=6)

        ttk.Radiobutton(
            mode_frame,
            text="新增 Basic UDI",
            variable=self.export_mode,
            value="DEVICE_POST"
        ).pack(side="left", padx=(0, 16))

        ttk.Radiobutton(
            mode_frame,
            text="更新 Basic UDI",
            variable=self.export_mode,
            value="UDI_DI_POST"
        ).pack(side="left")

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

            self.load_sheet_names(path)

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

        if self.available_sheets and self.sheet_name.get().strip() not in self.available_sheets:
            messagebox.showwarning("工作表錯誤", "所選工作表不存在於目前 Excel 檔案中。")
            return

        if not self.output_dir.get().strip():
            messagebox.showwarning("缺少資料", "請先選擇輸出資料夾。")
            return

        config = {
            "sender_actor_code": self.settings.get("sender_actor_code", "").strip(),
            "sender_node_id": self.settings.get("sender_node_id", "").strip(),
            "service_access_token": self.settings.get("service_access_token", "").strip(),
            "MFActorCode": self.settings.get("MFActorCode", "").strip(),
            "ARActorCode": self.settings.get("ARActorCode", "").strip()
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
        
        if not config["MFActorCode"]:
            messagebox.showwarning("缺少設定", "請先到「基本資料設定」填寫 MF Actor Code。")
            return
        
        if not config["ARActorCode"]:
            messagebox.showwarning("缺少設定", "請先到「基本資料設定」填寫 AR Actor Code。")
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
                config=config,
                field_mapping=self.settings.get("field_mapping", {}),
                export_mode=self.export_mode.get()
            )

            self.log(f'已讀取資料，共 {result["df_count"]} 筆符合條件。')
            self.log(f'已轉換裝置資料，共 {result["device_count"]} 筆。')
            self.log(f'XML 檔案輸出完成：{result["output_path"]}')
            self.log(f"輸出模式：{self.export_mode.get()}")
            messagebox.showinfo("完成", f'XML 檔案已成功產生：\n{result["output_path"]}')

        except Exception as e:
            self.log("執行失敗。")
            self.log(str(e))
            self.log(traceback.format_exc())
            messagebox.showerror("錯誤", f"執行失敗：\n{e}")
        finally:
            self.start_button.config(state="normal")

    def log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

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
        ARActorCode = tk.StringVar(value=self.settings.get("ARActorCode", ""))
        MFActorCode = tk.StringVar(value=self.settings.get("MFActorCode", ""))

        ttk.Label(frame, text="Sender Actor Code：").grid(row=0, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=sender_actor_code, width=34).grid(row=0, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="Sender Node ID：").grid(row=1, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=sender_node_id, width=34).grid(row=1, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="Service Access Token：").grid(row=2, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=service_access_token, width=34, show="*").grid(row=2, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="MF Actor Code：").grid(row=3, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=MFActorCode, width=34).grid(row=3, column=1, sticky="ew", padx=8)


        ttk.Label(frame, text="AR Actor Code：").grid(row=4, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=ARActorCode, width=34).grid(row=4, column=1, sticky="ew", padx=8)

        ttk.Label(frame, text="預設工作表：").grid(row=5, column=0, sticky="w", pady=8)
        ttk.Entry(frame, textvariable=default_sheet_name, width=34).grid(row=5, column=1, sticky="ew", padx=8)

        frame.columnconfigure(1, weight=1)

        button_bar = ttk.Frame(frame)
        button_bar.grid(row=6, column=0, columnspan=2, pady=(20, 0))

        def save_and_close():
            settings_data = dict(self.settings)
            settings_data.update({
                "sender_actor_code": sender_actor_code.get().strip(),
                "sender_node_id": sender_node_id.get().strip(),
                "service_access_token": service_access_token.get().strip(),
                "MFActorCode": MFActorCode.get().strip(),
                "ARActorCode": ARActorCode.get().strip(),
                "default_sheet_name": default_sheet_name.get().strip() or "p_zta"
            })

            self.save_settings(settings_data)
            self.log("基本資料已儲存。")
            messagebox.showinfo("完成", "基本資料已儲存。", parent=window)
            window.destroy()

        ttk.Button(button_bar, text="儲存", command=save_and_close).pack(side="left", padx=6, ipadx=12)
        ttk.Button(button_bar, text="取消", command=window.destroy).pack(side="left", padx=6, ipadx=12)

    def open_excel_fields_window(self):
        window = tk.Toplevel(self.root)
        window.title("Excel 欄位設定")
        window.geometry("760x620")
        window.minsize(720, 560)
        window.transient(self.root)
        window.grab_set()

        frame = ttk.Frame(window, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="請設定 Excel 欄位對應（tc_jsbXXX）",
            font=("Microsoft JhengHei", 11, "bold")
        ).pack(anchor="w", pady=(0, 4))

        ttk.Label(
            frame,
            text="共用欄位會提供給 MDD / MDR 共同使用；專用欄位只會在對應法規使用。",
        ).pack(anchor="w", pady=(0, 12))

        notebook = ttk.Notebook(frame)
        notebook.pack(fill="both", expand=True)

        field_mapping = self.settings.get("field_mapping", {})
        field_mapping_defaults = self.settings.get("field_mapping_defaults", {})

        common_mapping = field_mapping.get("COMMON", {})
        mdd_mapping = field_mapping.get("MDD", {})
        mdr_mapping = field_mapping.get("MDR", {})

        common_defaults = field_mapping_defaults.get("COMMON", {})
        mdd_defaults = field_mapping_defaults.get("MDD", {})
        mdr_defaults = field_mapping_defaults.get("MDR", {})

        common_vars = {}
        mdd_vars = {}
        mdr_vars = {}

        common_fields = [
            ("reg_type", "法規類型"),
            ("risk_class", "風險等級"),
            ("model", "Model"),
            ("certificate_no", "Certificate Number"),
            ("animal_tissues_cells", "Animal Tissues / Cells"),
            ("human_tissues_cells", "Human Tissues / Cells"),
            ("medicinal_product", "Medicinal Product"),
            ("human_product", "Human Product"),
            ("active", "Active"),
            ("administering", "Administering"),
            ("implantable", "Implantable"),
            ("measuring", "Measuring"),
            ("reusable", "Reusable"),
            ("udi_di", "UDI-DI"),
            ("udi_status", "UDI Status"),
            ("emdn_code", "EMDN Code"),
            ("pi_lot_number", "PI - Lot Number"),
            ("pi_serial_number", "PI - Serial Number"),
            ("pi_manufacturing_date", "PI - Manufacturing Date"),
            ("pi_expiration_date", "PI - Expiration Date"),
            ("product_number", "Product Number"),
            ("sterile", "Sterile"),
            ("sterilization", "Sterilization"),
            ("trade_name", "Trade Name"),
            ("trade_name_lang", "Trade Name Language"),
            ("number_of_reuses", "Number Of Reuses"),
            ("base_quantity", "Base Quantity"),
            ("latex", "Latex"),
            ("reprocessed", "Reprocessed"),
            ("spec_value", "Spec Value"),
            ("spec_unit", "Spec Unit"),
            ("first_market", "First Marketing Status"),
            ("marketing_status", "Marketing Status Decription"),
        ]

        mdd_fields = [
            ("basicudi_di", "Basic UDI-DI"),
            ("critical_warning", "Critical Warning"),
            ("certificate_revision", "Certificate Revision"),
            ("certificate_expiry", "Certificate Expiry"),
        ]

        mdr_fields = [
            ("basicudi_di", "Basic UDI-DI"),
        ]

        def build_scrollable_form(parent, fields, value_dict, default_dict, var_dict):
            outer = ttk.Frame(parent)
            outer.pack(fill="both", expand=True)

            canvas = tk.Canvas(outer, highlightthickness=0)
            scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            for i, (key, label_text) in enumerate(fields):
                ttk.Label(scrollable_frame, text=f"{label_text}：").grid(
                    row=i, column=0, sticky="w", pady=6
                )
                var = tk.StringVar(value=value_dict.get(key, ""))
                var_dict[key] = var
                ttk.Entry(scrollable_frame, textvariable=var, width=40).grid(
                    row=i, column=1, sticky="ew", padx=8, pady=6
                )

                ttk.Button(
                    scrollable_frame,
                    text="預設",
                    command=lambda v=var, k=key: v.set(default_dict.get(k, ""))
                ).grid(row=i, column=2, sticky="w", padx=(4, 0), pady=6)

            scrollable_frame.columnconfigure(1, weight=1)
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            def _bind_mousewheel(_event):
                canvas.bind_all("<MouseWheel>", _on_mousewheel)

            def _unbind_mousewheel(_event):
                canvas.unbind_all("<MouseWheel>")

            canvas.bind("<Enter>", _bind_mousewheel)
            canvas.bind("<Leave>", _unbind_mousewheel)
            scrollable_frame.bind("<Enter>", _bind_mousewheel)
            scrollable_frame.bind("<Leave>", _unbind_mousewheel)



        common_tab = ttk.Frame(notebook, padding=12)
        notebook.add(common_tab, text="共用欄位")
        build_scrollable_form(common_tab, common_fields, common_mapping, common_defaults, common_vars)

        mdd_tab = ttk.Frame(notebook, padding=12)
        notebook.add(mdd_tab, text="MDD")
        build_scrollable_form(mdd_tab, mdd_fields, mdd_mapping, mdd_defaults, mdd_vars)

        mdr_tab = ttk.Frame(notebook, padding=12)
        notebook.add(mdr_tab, text="MDR")
        build_scrollable_form(mdr_tab, mdr_fields, mdr_mapping, mdr_defaults, mdr_vars)

        button_bar = ttk.Frame(frame)
        button_bar.pack(fill="x", pady=(12, 0))

        def save_field_mapping():
            updated_settings = dict(self.settings)
            updated_settings["field_mapping"] = {
                "COMMON": {key: var.get().strip() for key, var in common_vars.items()},
                "MDD": {key: var.get().strip() for key, var in mdd_vars.items()},
                "MDR": {key: var.get().strip() for key, var in mdr_vars.items()},
            }

            self.save_settings(updated_settings)
            self.log("Excel 欄位設定已儲存。")
            messagebox.showinfo("完成", "Excel 欄位設定已儲存。", parent=window)
            window.destroy()

        ttk.Button(button_bar, text="儲存", command=save_field_mapping).pack(side="right", padx=6, ipadx=12)
        ttk.Button(button_bar, text="取消", command=window.destroy).pack(side="right", padx=6, ipadx=12)

if __name__ == "__main__":
    root = tk.Tk()
    app = UDIUploadUI(root)
    root.mainloop()