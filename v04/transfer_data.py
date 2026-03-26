import os
from datetime import datetime
import pandas as pd

from device_mapper import df_to_dict
from xml_builder import wrap_with_push, dict_to_xml_string, save_xml


REQUIRED_COLUMNS = [
    "tc_jsb030", "tc_jsb080", "tc_jsb150"
]
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
        col = mapping.get(section, {}).get(key)
        if col:
            required_columns.append(col)

    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {', '.join(missing)}")


def excel_to_df(file_path, sheet_name=None, field_mapping=None):
    if sheet_name is None:
        sheet_name = "p_zta"

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        dtype={"tc_jsb001": str, "tc_jsb710": str}
    )

    validate_required_columns(df, field_mapping)

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


def df_to_xml_files(devices, output_dir, config):
    device_nodes = [d["device:Device"] for d in devices]

    push_dict = wrap_with_push(
        {"device:Device": device_nodes},
        service_id="DEVICE",
        service_operation="POST",
        service_access_token=config["service_access_token"],
        sender_actor_code=config["sender_actor_code"],
        sender_node_id=config["sender_node_id"],
        xsd_version="3.0.25",
    )

    xml_str = dict_to_xml_string(push_dict)
    output_path = generate_output_path(output_dir)
    save_xml(xml_str, output_path)

    return output_path


def export_excel_to_xml(file_path, output_dir, sheet_name=None, config=None, field_mapping=None):
    df = excel_to_df(file_path, sheet_name, field_mapping=field_mapping)
    devices = df_to_dict(df, field_mapping=field_mapping)
    output_path = df_to_xml_files(devices, output_dir, config)

    return {
        "df_count": len(df),
        "device_count": len(devices),
        "output_path": output_path
    }


if __name__ == "__main__":
    result = export_excel_to_xml(
        r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\cimi290 匯出範例_test0312.xlsx",
        r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\output_xml",
        "p_zta"
    )
    print(result)