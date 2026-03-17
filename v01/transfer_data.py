import pandas as pd
import uuid
from datetime import datetime, timezone, timedelta
import xml.etree.ElementTree as ET
from xml.dom import minidom

def excel_to_df(file_path, sheet_name=None):
    # 若未指定 sheet，預設讀取 p_zta 工作表
    if sheet_name is None:
        sheet_name = "p_zta"

    # 讀取 Excel 資料，以 str 格式讀取 tc_jsb001 欄位 (UDI-DI code)
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'tc_jsb001': str, 'tc_jsb710': str})
    # 移除非MDD/MDR產品 (tc_jsb030: 標記註冊類型, tc_jsb080: 等級, tc_jsb150: 證書號碼)
    df_eu = df.dropna(subset=["tc_jsb030", "tc_jsb080", "tc_jsb150"])
    return df_eu

def df_to_xml_files(devices, output_dir):
    import os
    os.makedirs(output_dir, exist_ok=True)
    # 待 Bulk Upload 驗證 @2026-03-17
    # push_dict = wrap_with_push(
    #     devices,
    #     service_id="DEVICE",
    #     service_operation="POST",
    #     service_access_token="YOUR_SECURITY_TOKEN",
    #     sender_actor_code="YOUR_SRN_OR_ACTOR_ID",
    #     sender_node_id="YOUR_PARTY_ID",
    #     xsd_version="3.0.25",
    # )

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

    today = datetime.now().strftime("%Y%m%d")
    file_name = f"BULK_UPLOAD_DATA_{today}"
    ext = ".xml"
    output_path = os.path.join(output_dir, file_name+ext)

    i = 1
    while os.path.exists(output_path):
        output_path = os.path.join(output_dir, f"{file_name} ({i}){ext}")
        i += 1

    save_xml(xml_str, output_path)

    print(f"已輸出: {output_path}")


def df_to_dict(df):
    device_dict_list = []
    for i in range(len(df)):
        row = df.iloc[i]

        if row["tc_jsb030"] == "MDD":
            print(row["tc_jsb001"], "MDD產品")
            dict = row_to_dict_MDD(row)
            # print(dict, end="\n\n")
        elif row["tc_jsb030"] == "MDR": 
            print(row["tc_jsb001"], "MDR產品")
            dict = row_to_dict_MDR(row)
            # print(dict, end="\n\n")
        else: continue

        device_dict_list.append(dict)
        print(device_dict_list)
    return device_dict_list


def row_to_dict_MDR(row):
    # 回傳一個 dict，代表 MDR 產品的結構
    riskClass = row["tc_jsb080"].upper().replace(" ", "_")  # e.g., CLASS_I, CLASS_IIA, CLASS_IIB, CLASS_III
    model = row["tc_jsb200"] #????
    B_DICode = row["tc_jsb070"]
    B_Entity = "GS1"
    is_animalTissuesCells = "true" if row["tc_jsb360"] == "Y" else "false"
    is_humanTissuesCells = "true" if row["tc_jsb350"] == "Y" else "false"
    MFActorCode = "TW-MF-000017454"
    is_medicinalProduct = "true" if row["tc_jsb370"] == "Y" else "false"
    is_humanProduct = "true" if row["tc_jsb380"] == "Y" else "false"
    productType = "DEVICE" # enum: DEVICE, SYSTEM, PROCEDURE_PACK, KIT, COMPONENT, SOFTWARE
    is_active = "true" if row["tc_jsb120"] == "Y" else "false"
    is_administering = "true" if row["tc_jsb130"] == "Y" else "false"
    is_implantable = "true" if row["tc_jsb090"] == "Y" else "false"
    is_measuring = "true" if row["tc_jsb100"] == "Y" else "false"
    is_reusable = "true" if row["tc_jsb110"] == "Y" else "false"
    i_DICode = row["tc_jsb001"] if pd.notna(row["tc_jsb001"]) else None
    i_Entity = "GS1"
    # enum: ON_EU_MARKET, NO_LONGER_PLACED_ON_EU_MARKET, NOT_INTENDED_FOR_EU_MARKET
    udi_status = row["tc_jsb270"].upper().replace("EU ","").replace(" ", "_") if pd.notna(row["tc_jsb270"]) else None
    emdn_code = row["tc_jsb190"] if pd.notna(row["tc_jsb190"]) else None
    is_pi_lotNumber = "true" if row["tc_jsb2401"] == "Y" else "false"
    is_pi_serialNumber = "true" if row["tc_jsb2411"] == "Y" else "false"
    is_pi_manufacturingDate = "true" if row["tc_jsb2421"] == "Y" else "false"
    is_pi_expirationDate = "true" if row["tc_jsb2431"] == "Y" else "false"
    productNumber = row["tc_jsb000"]
    is_sterile = "true" if row["tc_jsb743"] == "Y" else "false"
    is_sterilization = "true" if row["tc_jsb742"] == "Y" else "false"
    tradeName = row["tc_jsb200"]
    tradeName_lang = row["tc_jsb210"]
    numberOfReuses = row["tc_jsb744"] if isinstance(row["tc_jsb744"], int) else 0
    baseQuantity = row["tc_jsb230"] if pd.notna(row["tc_jsb230"]) else None
    is_latex = "true" if row["tc_jsb550"] == "Y" else "false"
    is_reprocessed = "true" if row["tc_jsb620"] == "Y" else "false"
    spec = row["tc_jsb430"]+" "+row["tc_jsb440"] if pd.notna(row["tc_jsb430"]) and pd.notna(row["tc_jsb440"]) else None

    return {
        "device:Device": {
            "@xsi:type": "device:MDRDeviceType",
            # 💡 <device:MDRBasicUDI> … </device:MDRBasicUDI>
            "device:MDRBasicUDI":{
                "@xsi:type": "device:MDRBasicUDIType",
                "basicudi:riskClass": riskClass,
                "basicudi:model": model,
                "basicudi:identifier":{
                    "commondi:DICode": B_DICode,
                    "commondi:issuingEntityCode": B_Entity
                },
                "basicudi:animalTissuesCells": is_animalTissuesCells,
                "basicudi:humanTissuesCells": is_humanTissuesCells,
                "basicudi:MFActorCode": MFActorCode,
                "basicudi:humanProductCheck": is_humanProduct,
                "basicudi:medicinalProductCheck": is_medicinalProduct,
                "basicudi:type": productType,
                "commondi:active": is_active,
                "commondi:administeringMedicine": is_administering,
                "commondi:implantable": is_implantable,
                "commondi:measuringFunction": is_measuring,
                "commondi:reusable": is_reusable
            },
            # 💡 <device:MDRUDIDIData> … </device:MDRUDIDIData>
            "device:MDRUDIDIData":{
                "@xsi:type": "device:MDRUDIDIDataType",
                "udidi:identifier":{
                    "commondi:DICode": i_DICode,
                    "commondi:issuingEntityCode": i_Entity
                },
                "udidi:status": udi_status,
                "udidi:basicUDIIdentifier":{
                    "commondi:DICode": B_DICode,
                    "commondi:issuingEntityCode": B_Entity
                },
                "udidi:MDNCodes": emdn_code,
                "udidi:productionIdentifier": {
                    "udidi:lotNumber": is_pi_lotNumber,
                    "udidi:serialNumber": is_pi_serialNumber,
                    "udidi:manufacturingDate": is_pi_manufacturingDate,
                    "udidi:expirationDate": is_pi_expirationDate
                },
                "udidi:referenceNumber": productNumber,
                "udidi:sterile": is_sterile,
                "udidi:sterilization": is_sterilization,
                "udidi:tradeNames": {
                    "lsn:name":{
                        "lsn:language": tradeName_lang,
                        "lsn:textValue": tradeName
                    }
                },
                "udidi:numberOfReuses": numberOfReuses,
                "udidi:baseQuantity": baseQuantity,
                "udidi:latex": is_latex,
                "udidi:reprocessed": is_reprocessed,
                "udidi:clinicalSizes": {
                    "commondi:clinicalSize": {
                        "@xsi:type": "commondi:ValueClinicalSizeType",
                        "commondi:clinicalSizeType": "CST19",
                        "commondi:text": spec
    }}}}}

def row_to_dict_MDD(row):
    # 回傳一個 dict，代表 MDR 產品的結構
    riskClass = row["tc_jsb080"].upper().replace(" ", "_")  # e.g., CLASS_I, CLASS_IIA, CLASS_IIB, CLASS_III
    model = row["tc_jsb200"] #????
    B_DICode = row["tc_jsb630"] # only for MDD, MDR is tc_jsb070
    B_Entity = "GS1"
    is_animalTissuesCells = "true" if row["tc_jsb360"] == "Y" else "false"
    is_humanTissuesCells = "true" if row["tc_jsb350"] == "Y" else "false"
    MFActorCode = "TW-MF-000017454"
    is_medicinalProduct = "true" if row["tc_jsb370"] == "Y" else "false"
    is_humanProduct = "true" if row["tc_jsb380"] == "Y" else "false"
    productType = "DEVICE" # enum: DEVICE, SYSTEM, PROCEDURE_PACK, KIT, COMPONENT, SOFTWARE
    is_active = "true" if row["tc_jsb120"] == "Y" else "false"
    is_administering = "true" if row["tc_jsb130"] == "Y" else "false"
    is_implantable = "true" if row["tc_jsb090"] == "Y" else "false"
    is_measuring = "true" if row["tc_jsb100"] == "Y" else "false"
    is_reusable = "true" if row["tc_jsb110"] == "Y" else "false"
    i_DICode = row["tc_jsb001"] if pd.notna(row["tc_jsb001"]) else None
    i_Entity = "GS1"
    # enum: ON_EU_MARKET, NO_LONGER_PLACED_ON_EU_MARKET, NOT_INTENDED_FOR_EU_MARKET
    udi_status = row["tc_jsb270"].upper().replace("EU ","").replace(" ", "_") if pd.notna(row["tc_jsb270"]) else None
    emdn_code = row["tc_jsb190"] if pd.notna(row["tc_jsb190"]) else None
    is_pi_lotNumber = "true" if row["tc_jsb2401"] == "Y" else "false"
    is_pi_serialNumber = "true" if row["tc_jsb2411"] == "Y" else "false"
    is_pi_manufacturingDate = "true" if row["tc_jsb2421"] == "Y" else "false"
    is_pi_expirationDate = "true" if row["tc_jsb2431"] == "Y" else "false"
    productNumber = row["tc_jsb000"]
    is_sterile = "true" if row["tc_jsb743"] == "Y" else "false"
    is_sterilization = "true" if row["tc_jsb742"] == "Y" else "false"
    tradeName = row["tc_jsb200"]
    tradeName_lang = row["tc_jsb210"]
    numberOfReuses = row["tc_jsb744"] if isinstance(row["tc_jsb744"], int) else 0
    baseQuantity = row["tc_jsb230"] if pd.notna(row["tc_jsb230"]) else None
    is_latex = "true" if row["tc_jsb550"] == "Y" else "false"
    is_reprocessed = "true" if row["tc_jsb620"] == "Y" else "false"
    spec = row["tc_jsb430"]+" "+row["tc_jsb440"] if pd.notna(row["tc_jsb430"]) and pd.notna(row["tc_jsb440"]) else None
    criticalWarnings = row["tc_jsb730"] if pd.notna(row["tc_jsb730"]) else None
    warningValue = "CW010" if criticalWarnings == "Consult Instruction for Use" else None
    certificate_MDD = row["tc_jsb170"] if pd.notna(row["tc_jsb170"]) else None
    certificate_expiry_MDD = row["tc_jsb710"].split(" ")[0] if pd.notna(row["tc_jsb710"]) else None
    MNBctorCode = "2195"
    certificate_revision = row["tc_jsb180"] if pd.notna(row["tc_jsb180"]) else None
    certificate_type = "MDD_"+row["tc_jsb080"].split()[1]
    

    return {
        "device:Device": {
            "@xsi:type": "device:MDEUDeviceType",
            "device:MDEUData": {
                "udidi:identifier":{
                    "commondi:DICode": i_DICode,
                    "commondi:issuingEntityCode": i_Entity
                },
                "udidi:status": {
                    "commondi:code": udi_status
                },
                "udidi:basicUDIIdentifier": {
                    "commondi:DICode": B_DICode,
                    "commondi:issuingEntityCode": B_Entity
                },
                "udidi:MDNCodes": emdn_code,
                "udidi:referenceNumber": productNumber,
                "udidi:sterile": is_sterile,
                "udidi:sterilization": is_sterilization,
                "udidi:tradeNames": {
                    "lsn:name":{
                        "lsn:language": tradeName_lang,
                        "lsn:textValue": tradeName
                    }
                },
                "udidi:criticalWarnings": {
                    "commondi:warning": {
                        "commondi:comments": {
                                "lsn:name":{
                                    "lsn:language": "ANY",
                                    "lsn:textValue": criticalWarnings
                                }
                        },
                        "commondi:warningValue": warningValue
                    }
                },
                "udidi:numberOfReuses": numberOfReuses,
                "udidi:marketInfos": {
                    "marketinfo:marketInfo": {
                        "marketinfo:country": "ES",
                        "marketinfo:originalPlacedOnTheMarket": "false"
                    }
                },
    #             "udidi:baseQuantity": baseQuantity,
                "udidi:latex": is_latex,
                "udidi:reprocessed": is_reprocessed,
                "udidi:clinicalSizes": {
                    "commondi:clinicalSize": {
                        "@xsi:type": "commondi:TextClinicalSizeType",
                        "commondi:clinicalSizeType": "CST12",
                        "commondi:text": spec
                    }
                }},
            "device:MDEUDI": {
                "basicudi:riskClass": riskClass,
                "basicudi:model": model,
                "basicudi:identifier":{
                    "commondi:DICode": B_DICode,
                    "commondi:issuingEntityCode": B_Entity
                },
                "basicudi:animalTissuesCells": is_animalTissuesCells,
                "basicudi:humanTissuesCells": is_humanTissuesCells,
                "basicudi:MFActorCode": MFActorCode,
                "basicudi:deviceCertificateLinks": {
                    "links:deviceCertificateLink": {
                        "links:certificateNumber": certificate_MDD,
                        "links:expiryDate": certificate_expiry_MDD,
                        "links:NBActorCode": MNBctorCode,
                        "links:certificateRevisionNumber": certificate_revision,
                        "links:certificateType": certificate_type
                    }},
                "basicudi:humanProductCheck": is_humanProduct,
                "basicudi:medicinalProductCheck": is_medicinalProduct,
                "basicudi:type": productType,
                "commondi:active": is_active,
                "commondi:administeringMedicine": is_administering,
                "commondi:implantable": is_implantable,
                "commondi:measuringFunction": is_measuring,
                "commondi:reusable": is_reusable,
                "eudi:applicableLegislation": "MDD"           
    }}}

# ========= 2) 小工具 =========
def qname(tag: str) -> str:
    """
    'm:Push' -> '{uri}Push'
    """
    if ":" not in tag:
        return tag
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def is_empty(value) -> bool:
    """
    決定這個欄位要不要輸出
    """
    return value is None or value == "" or value == {} or value == []

def dict_to_xml(parent: ET.Element, data):
    """
    支援：
    - dict
    - list
    - @attribute
    - #text
    """
    if isinstance(data, dict):
        for key, value in data.items():
            if is_empty(value):
                continue

            # 屬性
            if key.startswith("@"):
                attr_name = key[1:]
                if ":" in attr_name:
                    parent.set(qname(attr_name), str(value))
                else:
                    parent.set(attr_name, str(value))
                continue

            # 文字
            if key == "#text":
                parent.text = str(value)
                continue

            # list -> 重複 tag
            if isinstance(value, list):
                for item in value:
                    child = ET.SubElement(parent, qname(key))
                    dict_to_xml(child, item)
            else:
                child = ET.SubElement(parent, qname(key))
                dict_to_xml(child, value)

    else:
        parent.text = str(data)

# ========= 3) 包成 m:Push =========
def wrap_with_push(
    payload_dict: dict,
    *,
    service_id: str = "DEVICE",
    service_operation: str = "POST",
    service_access_token: str = "YOUR_SECURITY_TOKEN",
    sender_actor_code: str = "YOUR_SRN_OR_ACTOR_ID",
    sender_node_id: str = "YOUR_PARTY_ID",
    xsd_version: str = "3.0.25",
) -> dict:
    """
    把你現在 row_to_dict_* 回傳的 payload dict，包成完整 m:Push dict
    """
    now = datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")

    return {
        "m:Push": {
            "@version": xsd_version,
            "@xsi:schemaLocation": (
                "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1 "
                "https://webgate.ec.europa.eu/tools/eudamed/dtx/service/Message.xsd"
            ),

            "m:conversationID": str(uuid.uuid4()),
            "m:correlationID": str(uuid.uuid4()),
            "m:creationDateTime": now,
            "m:messageID": str(uuid.uuid4()),

            "m:recipient": {
                "m:node": {
                    "s:nodeActorCode": "EUDAMED",
                    "s:nodeID": "eDelivery:EUDAMED",
                },
                "m:service": {
                    "s:serviceAccessToken": service_access_token,
                    "s:serviceID": service_id,
                    "s:serviceOperation": service_operation,
                },
            },

            "m:payload": payload_dict,

            "m:sender": {
                "m:node": {
                    "s:nodeActorCode": sender_actor_code,
                    "s:nodeID": sender_node_id,
                },
                "m:service": {
                    "s:serviceID": service_id,
                    "s:serviceOperation": service_operation,
                },
            },
        }
    }


# ========= 4) dict -> XML 字串 =========
def dict_to_xml_string(data: dict) -> str:
    root_key = next(iter(data))
    root_val = data[root_key]

    root = ET.Element(qname(root_key))
    dict_to_xml(root, root_val)

    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    pretty = minidom.parseString(xml_bytes).toprettyxml(indent="  ", encoding="utf-8")
    return pretty.decode("utf-8")


# ========= 5) 存檔 =========
def save_xml(xml_str: str, output_path: str):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(xml_str)



# ========= 1) Namespace 設定 =========
NS = {
    "m": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1",
    "s": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Service/v1",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "device": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/v1",
    "basicudi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/BasicUDI/v1",
    "udidi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/UDIDI/v1",
    "commondi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/CommonDevice/v1",
    "marketinfo": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/MktInfo/MarketInfo/v1",
    "links": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Links/v1",
    "lsn": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Common/LanguageSpecific/v1",
    "eudi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/LegacyDevice/EUDI/v1"
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)



# 讀取 cimi290 匯出檔案，轉成 DataFrame，然後轉成 dict
# df = excel_to_df(r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\cimi290 匯出範例_una.xlsx")
df = excel_to_df(r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\cimi290 匯出範例_test0312.xlsx")
devices = df_to_dict(df)
# print(devices)
df_to_xml_files(devices, r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\output_xml")