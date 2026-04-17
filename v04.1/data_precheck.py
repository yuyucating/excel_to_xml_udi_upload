import os
from datetime import datetime

import pandas as pd

from transfer_data import excel_to_df
from device_mapper import df_to_dict


def level_to_status(level):
    mapping = {
        "ERROR": "失敗",
        "WARNING": "警告",
        "INFO": "通過",
    }
    return mapping.get(level, "未知")


def make_issue(
    level,
    stage,
    message,
    row=None,
    field=None,
    value=None,
    udi_di=None,
    reg_type=None,
    status=None,
    suggestion=None,
):
    return {
        "level": level,
        "stage": stage,
        "row": row,
        "field": field,
        "value": value,
        "udi_di": udi_di,
        "reg_type": reg_type,
        "status": status or level_to_status(level),
        "message": message,
        "suggestion": suggestion,
    }


def precheck_data_format(
    excel_path,
    sheet_name=None,
    output_dir=None,
    field_mapping=None,
    export_mode=None,
    settings=None,
):
    settings = settings or {}
    field_mapping = field_mapping or {}

    errors = []
    warnings = []
    infos = []

    actor_codes = {
        "MFActorCode": (settings.get("MFActorCode") or "").strip(),
        "ARActorCode": (settings.get("ARActorCode") or "").strip(),
    }

    try:
        df = excel_to_df(
            excel_path,
            sheet_name=sheet_name,
            field_mapping=field_mapping
        )
        infos.append(make_issue(
            "INFO",
            "excel",
            f"Excel 讀取成功，預檢資料列數：{len(df)}",
            suggestion="可繼續進行 dict 階段預檢。"
        ))
    except Exception as e:
        errors.append(make_issue(
            "ERROR",
            "excel",
            f"Excel 預檢失敗：{e}",
            suggestion="請確認 Excel 檔案、工作表名稱與欄位設定是否正確。"
        ))
        return build_result(
            errors,
            warnings,
            infos,
            devices=[],
            excel_path=excel_path,
            sheet_name=sheet_name,
            export_mode=export_mode,
        )

    stage1_errors, stage1_warnings, stage1_infos = validate_dataframe_stage(
        df=df,
        field_mapping=field_mapping,
        export_mode=export_mode,
        settings=settings
    )
    errors.extend(stage1_errors)
    warnings.extend(stage1_warnings)
    infos.extend(stage1_infos)

    try:
        devices = df_to_dict(
            df,
            actor_codes,
            field_mapping=field_mapping,
            export_mode=export_mode
        )
        infos.append(make_issue(
            "INFO",
            "dict",
            f"dict 轉換成功，產生 {len(devices)} 筆資料",
            suggestion="可依報告檢查每筆資料的轉換結果。"
        ))
    except Exception as e:
        errors.append(make_issue(
            "ERROR",
            "dict",
            f"dict 轉換失敗：{e}",
            suggestion="請檢查欄位內容格式、對應設定與轉換函式。"
        ))
        return build_result(
            errors,
            warnings,
            infos,
            devices=[],
            excel_path=excel_path,
            sheet_name=sheet_name,
            export_mode=export_mode,
        )

    if len(devices) != len(df):
        warnings.append(make_issue(
            "WARNING",
            "dict",
            f"DataFrame 有 {len(df)} 筆，但 dict 只產出 {len(devices)} 筆；可能有資料列的 reg_type 未被轉入。",
            suggestion="請檢查 reg_type 是否為 MDR 或 MDD，並確認轉換邏輯是否有略過資料。"
        ))

    stage2_errors, stage2_warnings, stage2_infos = validate_dict_stage(
        devices=devices,
        export_mode=export_mode
    )
    errors.extend(stage2_errors)
    warnings.extend(stage2_warnings)
    infos.extend(stage2_infos)

    return build_result(
        errors,
        warnings,
        infos,
        devices=devices,
        excel_path=excel_path,
        sheet_name=sheet_name,
        export_mode=export_mode,
    )


def build_result(errors, warnings, infos, devices, excel_path="", sheet_name="", export_mode=""):
    issues = []
    issues.extend(errors)
    issues.extend(warnings)
    issues.extend(infos)

    pass_count = max(len(devices) - count_unique_problem_records(errors, warnings), 0)

    return {
        "success": len(errors) == 0,
        "error_count": len(errors),
        "warning_count": len(warnings),
        "info_count": len(infos),
        "pass_count": pass_count,
        "errors": errors,
        "warnings": warnings,
        "infos": infos,
        "issues": issues,
        "device_count": len(devices),
        "excel_path": excel_path,
        "sheet_name": sheet_name or "",
        "export_mode": export_mode or "",
        "message": f"預檢完成：{len(errors)} 個錯誤，{len(warnings)} 個警告，{len(infos)} 個提示",
    }


def count_unique_problem_records(errors, warnings):
    keys = set()

    for issue in list(errors) + list(warnings):
        if issue.get("udi_di"):
            keys.add(("udi_di", issue.get("udi_di")))
        elif issue.get("row") is not None:
            keys.add(("row", issue.get("row")))
        else:
            keys.add(("global", issue.get("message")))

    return len(keys)


def validate_dataframe_stage(df, field_mapping, export_mode, settings):
    errors = []
    warnings = []
    infos = []

    common = field_mapping.get("COMMON", {})
    mdd = field_mapping.get("MDD", {})
    mdr = field_mapping.get("MDR", {})

    def col(key, section="COMMON"):
        if section == "COMMON":
            return common.get(key)
        if section == "MDD":
            return mdd.get(key)
        if section == "MDR":
            return mdr.get(key)
        return None

    def cell_value(row_obj, column_name):
        if column_name and column_name in df.columns:
            return row_obj.get(column_name)
        return None

    for key in ("udi_di", "trade_name", "udi_status"):
        c = col(key, "COMMON")
        if c and c not in df.columns:
            errors.append(make_issue(
                "ERROR",
                "excel",
                f"欄位對應存在，但 Excel 找不到欄位：{c}",
                field=key,
                suggestion=f"請到『Excel 欄位設定』確認 {key} 的對應欄位名稱。"
            ))

    if export_mode == "DEVICE_POST":
        mdr_basic = col("basicudi_di", "MDR")
        mdd_basic = col("basicudi_di", "MDD")
        if not mdr_basic and not mdd_basic:
            warnings.append(make_issue(
                "WARNING",
                "excel",
                "DEVICE.POST 模式下，MDR/MDD 的 Basic UDI-DI 欄位對應都未設定。",
                suggestion="請到『Excel 欄位設定』補上 MDR 或 MDD 的 basicudi_di 對應欄位。"
            ))

    reg_col = col("reg_type", "COMMON")
    udi_col = col("udi_di", "COMMON")
    status_col = col("udi_status", "COMMON")
    market_col = col("marketing_status", "COMMON")

    for idx, row in df.iterrows():
        excel_row = idx + 2
        reg_value = cell_value(row, reg_col)
        reg_text = str(reg_value).strip().upper() if pd.notna(reg_value) else ""
        udi_value = cell_value(row, udi_col)
        udi_text = "" if pd.isna(udi_value) else str(udi_value).strip()

        if reg_col in df.columns and reg_text and reg_text not in ("MDR", "MDD"):
            warnings.append(make_issue(
                "WARNING",
                "excel",
                f"法規類型 {reg_text} 目前未支援，後續可能不會轉入 dict。",
                row=excel_row,
                field=reg_col,
                value=reg_text,
                udi_di=udi_text,
                reg_type=reg_text,
                suggestion="請確認 reg_type 是否應為 MDR 或 MDD。"
            ))

        if udi_col in df.columns and not udi_text:
            errors.append(make_issue(
                "ERROR",
                "excel",
                "UDI-DI 不可空白。",
                row=excel_row,
                field=udi_col,
                udi_di=udi_text,
                reg_type=reg_text,
                suggestion="請補上 UDI-DI。"
            ))

        if status_col in df.columns:
            status_value = cell_value(row, status_col)
            if pd.notna(status_value):
                norm = str(status_value).strip().upper().replace("EU ", "").replace(" ", "_")
                allowed = {
                    "ON_THE_MARKET",
                    "NO_LONGER_PLACED_ON_THE_MARKET",
                    "AVAILABLE",
                    "UNAVAILABLE",
                }
                if norm not in allowed:
                    warnings.append(make_issue(
                        "WARNING",
                        "excel",
                        f"UDI Status 可能不是預期值：{status_value}",
                        row=excel_row,
                        field=status_col,
                        value=status_value,
                        udi_di=udi_text,
                        reg_type=reg_text,
                        suggestion="請確認 UDI Status 是否可正確轉成 EUDAMED 代碼。"
                    ))

        if market_col in df.columns:
            market_value = cell_value(row, market_col)
            if pd.notna(market_value) and str(market_value).strip():
                txt = str(market_value)
                if "From " in txt and " To " not in txt:
                    warnings.append(make_issue(
                        "WARNING",
                        "excel",
                        "marketing_status 看起來有 From 但缺少 To。",
                        row=excel_row,
                        field=market_col,
                        value=txt,
                        udi_di=udi_text,
                        reg_type=reg_text,
                        suggestion="請檢查 marketing_status 的日期區間格式。"
                    ))

        if export_mode == "DEVICE_POST":
            if reg_text == "MDR":
                basic_col = col("basicudi_di", "MDR")
            elif reg_text == "MDD":
                basic_col = col("basicudi_di", "MDD")
            else:
                basic_col = None

            if basic_col and basic_col in df.columns:
                basic_value = row.get(basic_col)
                if pd.isna(basic_value) or str(basic_value).strip() == "":
                    errors.append(make_issue(
                        "ERROR",
                        "excel",
                        "DEVICE.POST 模式下 Basic UDI-DI 不可空白。",
                        row=excel_row,
                        field=basic_col,
                        udi_di=udi_text,
                        reg_type=reg_text,
                        suggestion="請補上 Basic UDI-DI 後再重新預檢。"
                    ))

    infos.append(make_issue(
        "INFO",
        "excel",
        "Excel 階段預檢完成。",
        suggestion="可繼續檢視 dict 階段的轉換結果。"
    ))
    return errors, warnings, infos


def validate_dict_stage(devices, export_mode):
    errors = []
    warnings = []
    infos = []

    seen_udi = set()

    for i, item in enumerate(devices, start=1):
        try:
            if export_mode == "DEVICE_POST":
                root = item.get("device:Device", {})
                check_device_post_dict(root, i, errors, warnings, seen_udi)
            elif export_mode == "UDI_DI_POST":
                root = item.get("udidiDatas:UDIDIData", {})
                check_udi_di_post_dict(root, i, errors, warnings, seen_udi)
            else:
                errors.append(make_issue(
                    "ERROR",
                    "dict",
                    f"不支援的 export_mode：{export_mode}",
                    suggestion="請確認 UI 的輸出模式設定。"
                ))
                break
        except Exception as e:
            errors.append(make_issue(
                "ERROR",
                "dict",
                f"第 {i} 筆 dict 驗證失敗：{e}",
                row=i,
                suggestion="請檢查該筆資料的轉換結果與欄位格式。"
            ))

    infos.append(make_issue(
        "INFO",
        "dict",
        "dict 階段預檢完成。",
        suggestion="可查看 Details 工作表了解每筆資料的狀況。"
    ))
    return errors, warnings, infos


def check_device_post_dict(root, index, errors, warnings, seen_udi):
    basic = root.get("device:MDRBasicUDI") or root.get("device:MDEUDI") or {}
    udidi = root.get("device:MDRUDIDIData") or root.get("device:MDEUData") or {}

    reg_type = "MDR" if "device:MDRUDIDIData" in root else "MDD"

    if not basic:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "DEVICE.POST 缺少 Basic UDI 區塊，無法產生完整裝置主檔。",
            row=index,
            reg_type=reg_type,
            suggestion="請檢查 Basic UDI-DI 與相關欄位是否成功轉換。"
        ))

    if not udidi:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "DEVICE.POST 缺少 UDI-DI 區塊，無法產生 XML。",
            row=index,
            reg_type=reg_type,
            suggestion="請檢查 UDI-DI、狀態與基本識別欄位是否成功轉換。"
        ))
        return

    identifier = udidi.get("udidi:identifier", {})
    di_code = identifier.get("commondi:DICode")
    issuing = identifier.get("commondi:issuingEntityCode")

    if not di_code:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "DEVICE.POST 缺少 UDI-DI identifier。",
            row=index,
            field="udidi:identifier/commondi:DICode",
            reg_type=reg_type,
            suggestion="請確認 Excel 的 UDI-DI 欄位是否有值，且轉換邏輯有正確帶入。"
        ))
    else:
        if di_code in seen_udi:
            warnings.append(make_issue(
                "WARNING",
                "dict",
                f"發現重複 UDI-DI：{di_code}",
                row=index,
                field="commondi:DICode",
                value=di_code,
                udi_di=di_code,
                reg_type=reg_type,
                suggestion="請確認是否重複匯出同一筆產品。"
            ))
        else:
            seen_udi.add(di_code)

    if not issuing:
        warnings.append(make_issue(
            "WARNING",
            "dict",
            "UDI-DI issuingEntityCode 為空。",
            row=index,
            field="commondi:issuingEntityCode",
            udi_di=di_code,
            reg_type=reg_type,
            suggestion="請確認 issuing entity 是否已設定並正確轉入。"
        ))

    status_code = ((udidi.get("udidi:status") or {}).get("commondi:code"))
    if not status_code:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "缺少 UDI Status 代碼，EUDAMED 送件可能失敗。",
            row=index,
            udi_di=di_code,
            reg_type=reg_type,
            suggestion="請檢查 Excel 的 UDI Status 值是否可成功轉換。"
        ))

    market_infos = udidi.get("udidi:marketInfos")
    if not market_infos:
        warnings.append(make_issue(
            "WARNING",
            "dict",
            "未產生 marketInfos，若此產品需要上市資訊，送件可能失敗。",
            row=index,
            udi_di=di_code,
            reg_type=reg_type,
            suggestion="請確認 first_market 或 marketing_status 是否有填寫。"
        ))


def check_udi_di_post_dict(root, index, errors, warnings, seen_udi):
    identifier = root.get("udidi:identifier", {})
    di_code = identifier.get("commondevice:DICode")
    issuing = identifier.get("commondevice:issuingEntityCode")

    if not di_code:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "UDI_DI.POST 缺少 UDI-DI identifier。",
            row=index,
            field="udidi:identifier/commondevice:DICode",
            suggestion="請確認 UDI-DI 是否成功轉入更新資料。"
        ))
    else:
        if di_code in seen_udi:
            warnings.append(make_issue(
                "WARNING",
                "dict",
                f"發現重複 UDI-DI：{di_code}",
                row=index,
                field="commondevice:DICode",
                value=di_code,
                udi_di=di_code,
                reg_type="MDR",
                suggestion="請確認是否重複匯出同一筆 UDI-DI 更新資料。"
            ))
        else:
            seen_udi.add(di_code)

    if not issuing:
        warnings.append(make_issue(
            "WARNING",
            "dict",
            "UDI_DI.POST issuingEntityCode 為空。",
            row=index,
            field="commondevice:issuingEntityCode",
            udi_di=di_code,
            reg_type="MDR",
            suggestion="請確認 issuing entity 是否已正確帶入。"
        ))

    status_code = ((root.get("udidi:status") or {}).get("commondevice:code"))
    if not status_code:
        errors.append(make_issue(
            "ERROR",
            "dict",
            "缺少 UDI Status 代碼，EUDAMED 更新送件可能失敗。",
            row=index,
            udi_di=di_code,
            reg_type="MDR",
            suggestion="請檢查 UDI Status 的來源值與轉換結果。"
        ))

    basic_identifier = root.get("udidi:basicUDIIdentifier", {})
    if not basic_identifier.get("commondevice:DICode"):
        errors.append(make_issue(
            "ERROR",
            "dict",
            "缺少 Basic UDI-DI，無法建立 UDI_DI.POST 更新資料。",
            row=index,
            udi_di=di_code,
            reg_type="MDR",
            suggestion="請補上 Basic UDI-DI。"
        ))


def build_report_rows(result, run_time, export_mode, excel_path="", sheet_name=""):
    rows = []

    for issue in result.get("issues", []):
        rows.append({
            "執行時間": run_time,
            "Excel檔案": excel_path,
            "工作表": sheet_name,
            "輸出模式": export_mode,
            "UDI-DI": issue.get("udi_di", ""),
            "Excel列號": issue.get("row", ""),
            "法規類型": issue.get("reg_type", ""),
            "狀態": issue.get("status") or level_to_status(issue.get("level")),
            "層級": issue.get("level", ""),
            "階段": issue.get("stage", ""),
            "問題欄位": issue.get("field", ""),
            "問題值": issue.get("value", ""),
            "問題說明": issue.get("message", ""),
            "建議處理": issue.get("suggestion", ""),
        })

    return rows


def export_precheck_report(
    result,
    output_dir,
    excel_path="",
    sheet_name="",
    export_mode=""
):
    if not output_dir:
        return None

    os.makedirs(output_dir, exist_ok=True)

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    file_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(output_dir, f"PRECHECK_REPORT_{file_time}.xlsx")

    detail_rows = build_report_rows(
        result=result,
        run_time=run_time,
        export_mode=export_mode,
        excel_path=excel_path,
        sheet_name=sheet_name,
    )

    if detail_rows:
        details_df = pd.DataFrame(detail_rows)
    else:
        details_df = pd.DataFrame([{
            "執行時間": run_time,
            "Excel檔案": excel_path,
            "工作表": sheet_name,
            "輸出模式": export_mode,
            "UDI-DI": "",
            "Excel列號": "",
            "法規類型": "",
            "狀態": "通過",
            "層級": "INFO",
            "階段": "",
            "問題欄位": "",
            "問題值": "",
            "問題說明": "未發現任何問題",
            "建議處理": "",
        }])

    summary_df = pd.DataFrame([
        {"項目": "執行時間", "內容": run_time},
        {"項目": "Excel檔案", "內容": excel_path},
        {"項目": "工作表", "內容": sheet_name},
        {"項目": "輸出模式", "內容": export_mode},
        {"項目": "資料筆數", "內容": result.get("device_count", 0)},
        {"項目": "通過筆數", "內容": result.get("pass_count", 0)},
        {"項目": "錯誤數", "內容": result.get("error_count", 0)},
        {"項目": "警告數", "內容": result.get("warning_count", 0)},
        {"項目": "提示數", "內容": result.get("info_count", 0)},
        {"項目": "預檢結果", "內容": result.get("message", "")},
    ])

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        details_df.to_excel(writer, sheet_name="Details", index=False)

        if "層級" in details_df.columns:
            error_df = details_df[details_df["層級"] == "ERROR"]
            warning_df = details_df[details_df["層級"] == "WARNING"]
            info_df = details_df[details_df["層級"] == "INFO"]

            if not error_df.empty:
                error_df.to_excel(writer, sheet_name="Errors", index=False)

            if not warning_df.empty:
                warning_df.to_excel(writer, sheet_name="Warnings", index=False)

            if not info_df.empty:
                info_df.to_excel(writer, sheet_name="Infos", index=False)

    return report_path
