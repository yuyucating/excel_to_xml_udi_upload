# Excel to XML UDI Upload Tool

這是一個用 Python 開發的桌面工具，用來將 ERP 導出 Excel 資料轉換為符合 EUDAMED Bulk Upload 格式的 XML 檔案，並提供簡單的圖形化介面，方便使用者選擇檔案、設定輸出資料夾、管理基本參數，以及產生 XML。

## 功能特色

- 讀取 Excel 檔案並指定工作表
- 自動篩選符合條件的 MDD / MDR 資料
- 將資料轉換成 EUDAMED Bulk Upload XML 格式
- 提供桌面 UI 介面操作
- 支援預設輸出資料夾
- 支援設定基本資料並保存到本機 JSON
- 可開啟輸出資料夾快速查看結果

## 調整輸入參數

目前此工具配合 TIPTOP 客製軟體 cimi290 之輸出 excel 作為格式轉換之輸入，若需要調整 excel 相應的欄位，請修改 `device_mapper.py`。

## 未來方向

- excel 欄位與 xml 標籤對應的調整將在 UI 可以直接進行調整。
- 第二個轉換格式，TUDID 批次上傳格式。