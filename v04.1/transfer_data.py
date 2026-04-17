import os
from datetime import datetime
import pandas as pd

from device_mapper import df_to_dict
from xml_builder import wrap_with_push, dict_to_xml_string, save_xml


# REQUIRED_COLUMNS = [
#     "tc_jsb030", "tc_jsb080", "tc_jsb150"
# ]
DEFAULT_REQUIRED_FIELD_KEYS = [
    ("COMMON", "reg_type"),
    ("COMMON", "risk_class"),
]
DEFAULT_FIELD_MAPPING = {
    "COMMON": {
        "reg_type": "tc_jsb030",
        "risk_class": "tc_jsb080",
    },
    "MDD": {},
    "MDR": {},
}

EXPORT_MODE_CONFIG = {
    "DEVICE_POST": {
        "service_id": "DEVICE",
        "service_operation": "POST",
        "payload_root": "device:Device",
    },
    "UDI_DI_POST": {
        "service_id": "UDI_DI",
        "service_operation": "POST",
        "payload_root": "udidiDatas:UDIDIData",
    },
}

def merge_required_mapping(field_mapping=None):
    # 先定義預設值為 DEFAULT_FIELD_MAPPING 的內容，如果 UI 有帶入 field_mapping 就用 UI 的值覆蓋預設值
    merged = {
        "COMMON": dict(DEFAULT_FIELD_MAPPING["COMMON"]),
        "MDD": {},
        "MDR": {},
    }
    if field_mapping:
        for section in ("COMMON", "MDD", "MDR"):
            merged[section].update(field_mapping.get(section, {}))
    return merged


# 驗證是否具備必要欄位
def validate_required_columns(df, field_mapping=None):
    # mapping 取得 UI 預設值，如果 UI 沒有帶入則使用 DEFAULT_FIELD_MAPPING 的預設值
    mapping = merge_required_mapping(field_mapping)

    required_columns = []
    for section, key in DEFAULT_REQUIRED_FIELD_KEYS:
        col = mapping.get(section, {}).get(key)
        if col:
            required_columns.append(col)

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {', '.join(missing)}")

# 將為完整的資料列移除
def excel_to_df(file_path, sheet_name=None, field_mapping=None):
    if sheet_name is None:
        sheet_name = "p_zta"

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        dtype={"tc_jsb001": str, "tc_jsb710": str} # 避免 UDI 編碼失去前導零
    )

    validate_required_columns(df, field_mapping)

    reg_col = field_mapping.get("COMMON", {}).get("reg_type", "tc_jsb030") if field_mapping else "tc_jsb030"
    risk_col = field_mapping.get("COMMON", {}).get("risk_class", "tc_jsb080") if field_mapping else "tc_jsb080"

    subset_cols = [c for c in [reg_col, risk_col] if c in df.columns]

    # 紀錄被移除的資料
    df_removed = df[df[subset_cols].isna().any(axis=1)]

    # 👉 標記缺少欄位（推薦）
    def find_missing_reason(row):
        missing = [col for col in subset_cols if pd.isna(row[col])]
        return ", ".join(missing)

    if not df_removed.empty:
        df_removed["missing_fields"] = df_removed.apply(find_missing_reason, axis=1)

        # 👉 產生 errorlog 檔名
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        error_file = f"{base_name}_errorlog.xlsx"

        # 👉 輸出 Excel
        df_removed.to_excel(error_file, index=False)

        print(f"⚠️ 已輸出錯誤資料至: {error_file}")

    # 只保留在必要欄位都有值的資料列
    df_eu = df.dropna(subset=subset_cols)

    return df_eu, df_removed


def generate_output_path(output_dir):
    os.makedirs(output_dir, exist_ok=True)

    today = datetime.now().strftime("%Y%m%d")
    file_name = f"BULK_UPLOAD_DATA_{today}"
    ext = ".xml"
    output_path = os.path.join(output_dir, file_name + ext)

    i = 1
    while os.path.exists(output_path):
        output_path = os.path.join(output_dir, f"{file_name} ({i}){ext}")
        i += 1

    return output_path


def df_to_xml_files(devices, output_dir, config, export_mode="DEVICE_POST"):
    mode_config = EXPORT_MODE_CONFIG.get(export_mode, EXPORT_MODE_CONFIG["DEVICE_POST"])
    payload_root = mode_config["payload_root"]
    service_id = mode_config["service_id"]
    service_operation = mode_config["service_operation"]

    if export_mode == "UDI_DI_POST":
        device_nodes = [d["udidiDatas:UDIDIData"] for d in devices]
    elif export_mode == "DEVICE_POST":
        device_nodes = [d["device:Device"] for d in devices]

    push_dict = wrap_with_push(
        {payload_root: device_nodes},
        service_id=service_id,
        service_operation=service_operation,
        service_access_token=config["service_access_token"],
        sender_actor_code=config["sender_actor_code"],
        sender_node_id=config["sender_node_id"],
        xsd_version="3.0.28",
        export_mode=export_mode
    )

    xml_str = dict_to_xml_string(push_dict, export_mode=export_mode)
    output_path = generate_output_path(output_dir)
    save_xml(xml_str, output_path)

    return output_path

# ui.py 的進入點
def export_excel_to_xml(file_path, output_dir, sheet_name=None, config=None,
    field_mapping=None, export_mode="DEVICE_POST"):
    # config 是從 UI 設定帶出來的資料

    ActorCodes = {k: v for k, v in config.items() if k.endswith("ActorCode")}

    df, df_removed = excel_to_df(file_path, sheet_name, field_mapping=field_mapping)
    devices = df_to_dict(df, ActorCodes, field_mapping=field_mapping, export_mode=export_mode)
    output_path = df_to_xml_files(devices, output_dir, config, export_mode=export_mode)

    return {
        "df_count": len(df),
        "df_removed_count": len(df_removed),
        "device_count": len(devices),
        "output_path": output_path,
        "export_mode": export_mode,
    }


if __name__ == "__main__":
    result = export_excel_to_xml(
        r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\cimi290 匯出範例_test0324.xlsx",
        r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\output_xml",
        "p_zta"
    )
    print(result)