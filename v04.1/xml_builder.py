import uuid
from datetime import datetime, timezone, timedelta
import xml.etree.ElementTree as ET
from xml.dom import minidom

ALL_NAMESPACES = {
    "m": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1",
    "message": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1",
    "s": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Service/v1",
    "service": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Service/v1",
    "xsi": "http://www.w3.org/2001/XMLSchema-instance",
    "device": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/v1",
    "udidiDatas": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/v1",
    "basicudi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/BasicUDI/v1",
    "udidi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/UDIDI/v1",
    "commondi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/CommonDevice/v1",
    "commondevice": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/CommonDevice/v1",
    "marketinfo": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/MktInfo/MarketInfo/v1",
    "mi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/MktInfo/MarketInfo/v1",
    "links": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Links/v1",
    "lsn": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Common/LanguageSpecific/v1",
    "eudi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/LegacyDevice/EUDI/v1",
    "e": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/v1",
}


def get_namespaces(export_mode: str) -> dict[str, str]:
    """
    Return the namespaces required for the selected export mode.

    Recommended modes:
    - DEVICE_POST

    - UDI_DI_POST
    """
    # 對 ElementTree 來說，少宣告會炸；多宣告通常沒事。
    # 所以這裡採「固定核心集合」，避免某次因 xsi:type 只出現在 attribute value 裡而漏 prefix。
    core = ["xsi", "udidi", "lsn"]

    mode_specific = {
        "DEVICE_POST": core + ["s", "m", "device", "basicudi", "marketinfo", "links", "eudi", "commondi"],
        "UDI_DI_POST": core + ["service", "message", "udidiDatas", "mi"],
    }

    prefixes = mode_specific.get(export_mode, core + ["device", "basicudi", "marketinfo", "links", "eudi", "commondevice"])
    return {prefix: ALL_NAMESPACES[prefix] for prefix in prefixes}


def register_namespaces(ns_map: dict[str, str]) -> None:
    """
    Register prefixes for pretty/consistent serialization.
    """
    for prefix, uri in ns_map.items():
        ET.register_namespace(prefix, uri)


def qname(tag: str, ns_map: dict[str, str]) -> str:
    """
    Convert 'prefix:local' to Clark notation '{uri}local'.
    """
    if ":" not in tag:
        return tag

    prefix, local = tag.split(":", 1)
    if prefix not in ns_map:
        raise KeyError(f"Namespace prefix '{prefix}' is not defined for this export mode.")

    return f"{{{ns_map[prefix]}}}{local}"


def is_empty(value) -> bool:
    return value is None or value == "" or value == {} or value == []


def dict_to_xml(parent: ET.Element, data, ns_map: dict[str, str]) -> None:
    if isinstance(data, dict):
        for key, value in data.items():
            if is_empty(value):
                continue

            if key.startswith("@"):
                attr_name = key[1:]
                if ":" in attr_name:
                    parent.set(qname(attr_name, ns_map), str(value))
                else:
                    parent.set(attr_name, str(value))
                continue

            if key == "#text":
                parent.text = str(value)
                continue

            if isinstance(value, list):
                for item in value:
                    child = ET.SubElement(parent, qname(key, ns_map))
                    dict_to_xml(child, item, ns_map)
            else:
                child = ET.SubElement(parent, qname(key, ns_map))
                dict_to_xml(child, value, ns_map)
    else:
        parent.text = str(data)


def wrap_with_push(
    payload_dict: dict,
    *,
    service_id: str = "DEVICE",
    service_operation: str = "POST",
    service_access_token: str = "YOUR_SECURITY_TOKEN",
    sender_actor_code: str = "YOUR_SRN_OR_ACTOR_ID",
    sender_node_id: str = "YOUR_PARTY_ID",
    xsd_version: str = "3.0.28",
    export_mode: str = "DEVICE_POST"
) -> dict:
    now = datetime.now(timezone(timedelta(hours=8))).isoformat(timespec="seconds")
    if export_mode == "UDI_DI_POST":
        service = "service"
        message = "message"
    elif export_mode == "DEVICE_POST":
        service = "s"
        message = "m"

    return {
        "m:Push": {
            "@version": xsd_version,
            "@xsi:schemaLocation": (
                "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1 "
                "https://webgate.ec.europa.eu/tools/eudamed/dtx/service/Message.xsd"
            ),
            message+":conversationID": str(uuid.uuid4()),
            message+":correlationID": str(uuid.uuid4()),
            message+":creationDateTime": now,
            message+":messageID": str(uuid.uuid4()),
            message+":recipient": {
                message+":node": {
                    service+":nodeActorCode": "EUDAMED",
                     service+":nodeID": "eDelivery:EUDAMED",
                },
                message+":service": {
                     service+":serviceAccessToken": service_access_token,
                     service+":serviceID": service_id,
                     service+":serviceOperation": service_operation,
                },
            },
            message+":payload": payload_dict,
            message+":sender": {
                message+":node": {
                    service+":nodeActorCode": sender_actor_code,
                    service+":nodeID": sender_node_id,
                },
                message+":service": {
                    service+":serviceID": service_id,
                    service+":serviceOperation": service_operation,
                },
            },
        }
    }


def dict_to_xml_string(data: dict, export_mode: str) -> str:
    ns_map = get_namespaces(export_mode)
    register_namespaces(ns_map)

    root_key = next(iter(data))
    root_val = data[root_key]

    root = ET.Element(qname(root_key, ns_map))
    dict_to_xml(root, root_val, ns_map)

    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    pretty = minidom.parseString(xml_bytes).toprettyxml(indent="  ", encoding="utf-8")
    return pretty.decode("utf-8")


def save_xml(xml_str: str, output_path: str):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(xml_str)