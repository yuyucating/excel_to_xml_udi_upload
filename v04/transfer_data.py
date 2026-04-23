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
    merged = {
        "COMMON": dict(DEFAULT_FIELD_MAPPING["COMMON"]),
        "MDD": {},
        "MDR": {},
    }
    if field_mapping:
        for section in ("COMMON", "MDD", "MDR"):
            merged[section].update(field_mapping.get(section, {}))
    return merged


def validate_required_columns(df, field_mapping=None):
    mapping = merge_required_mapping(field_mapping)

    required_columns = []
    for section, key in DEFAULT_REQUIRED_FIELD_KEYS:
        col = mapping.get(section, {}).get(key).lower()
        if col:
            required_columns.append(col)

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {', '.join(missing)}")

def validate_marketing_status_fields(df, field_mapping=None):
    """
    驗證 tc_jsb390 (first_market) 和 tc_jsb400 (marketing_status) 
    不能為 "N"
    
    Args:
        df: pandas DataFrame
        field_mapping: 欄位對應字典
    
    Raises:
        ValueError: 當欄位值為 "N" 時
    """
    mapping = field_mapping.get("COMMON", {}) if field_mapping else {}
    
    first_market_col = mapping.get("first_market", "tc_jsb390")
    marketing_status_col = mapping.get("marketing_status", "tc_jsb400")
    
    error_messages = []
    
    # 檢查 tc_jsb390 (first_market)
    if first_market_col in df.columns:
        # 將值轉換為字符串並轉大寫進行比較
        invalid_mask = df[first_market_col].astype(str).str.strip().str.upper() == "N"
        invalid_rows = df[invalid_mask]
        if not invalid_rows.empty:
            # 加 2 是因為 DataFrame 索引從 0 開始，但 Excel 列從 1 開始，加上標題列
            row_numbers = (invalid_rows.index + 2).tolist()
            error_messages.append(
                f"欄位 '{first_market_col}' (First Marketing Status) 包含無效值 'N'，"
                f"位於 Excel 第 {row_numbers} 列。此欄位不能為 'N'。"
            )
    
    # 檢查 tc_jsb400 (marketing_status)
    if marketing_status_col in df.columns:
        invalid_mask = df[marketing_status_col].astype(str).str.strip().str.upper() == "N"
        invalid_rows = df[invalid_mask]
        if not invalid_rows.empty:
            row_numbers = (invalid_rows.index + 2).tolist()
            error_messages.append(
                f"欄位 '{marketing_status_col}' (Marketing Status Description) 包含無效值 'N'，"
                f"位於 Excel 第 {row_numbers} 列。此欄位不能為 'N'。"
            )
    
    if error_messages:
        raise ValueError("\n".join(error_messages))

def excel_to_df(file_path, sheet_name=None, field_mapping=None):
    if sheet_name is None:
        sheet_name = "p_zta"

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        dtype={"tc_jsb001": str, "tc_jsb710": str, "TC_JSB001": str, "TC_JSB710": str}
    )

    df.columns = df.columns.str.lower()

    validate_required_columns(df, field_mapping)

    validate_marketing_status_fields(df, field_mapping)

    reg_col = field_mapping.get("COMMON", {}).get("reg_type", "tc_jsb030") if field_mapping else "tc_jsb030"
    risk_col = field_mapping.get("COMMON", {}).get("risk_class", "tc_jsb080") if field_mapping else "tc_jsb080"

    subset_cols = [c for c in [reg_col, risk_col] if c in df.columns]
    df_eu = df.dropna(subset=subset_cols)
    return df_eu


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


def export_excel_to_xml(file_path, output_dir, sheet_name=None, config=None,
    field_mapping=None, export_mode="DEVICE_POST"):
    ActorCodes = {k: v for k, v in config.items() if k.endswith("ActorCode")}

    df = excel_to_df(file_path, sheet_name, field_mapping=field_mapping)
    devices = df_to_dict(df, ActorCodes, field_mapping=field_mapping, export_mode=export_mode)
    output_path = df_to_xml_files(devices, output_dir, config, export_mode=export_mode)

    print(df)

    return {
        "df_count": len(df),
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