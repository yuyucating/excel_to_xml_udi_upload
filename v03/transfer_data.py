import os
from datetime import datetime
import pandas as pd

from device_mapper import df_to_dict
from xml_builder import wrap_with_push, dict_to_xml_string, save_xml


REQUIRED_COLUMNS = [
    "tc_jsb030", "tc_jsb080", "tc_jsb150"
]


def validate_required_columns(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {', '.join(missing)}")


def excel_to_df(file_path, sheet_name=None):
    if sheet_name is None:
        sheet_name = "p_zta"

    df = pd.read_excel(
        file_path,
        sheet_name=sheet_name,
        dtype={"tc_jsb001": str, "tc_jsb710": str}
    )

    validate_required_columns(df)

    df_eu = df.dropna(subset=["tc_jsb030", "tc_jsb080", "tc_jsb150"])
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


def df_to_xml_files(devices, output_dir):
    device_nodes = [d["device:Device"] for d in devices]

    push_dict = wrap_with_push(
        {"device:Device": device_nodes},
        service_id="DEVICE",
        service_operation="POST",
        service_access_token="YOUR_SECURITY_TOKEN",
        sender_actor_code="YOUR_SRN_OR_ACTOR_ID",
        sender_node_id="YOUR_PARTY_ID",
        xsd_version="3.0.25",
    )

    xml_str = dict_to_xml_string(push_dict)
    output_path = generate_output_path(output_dir)
    save_xml(xml_str, output_path)

    return output_path


def export_excel_to_xml(file_path, output_dir, sheet_name=None):
    df = excel_to_df(file_path, sheet_name)
    devices = df_to_dict(df)
    output_path = df_to_xml_files(devices, output_dir)

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