import uuid
from datetime import datetime, timezone, timedelta
import xml.etree.ElementTree as ET
from xml.dom import minidom


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
    "eudi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/LegacyDevice/EUDI/v1",
    "e": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/v1",
    "message": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Message/v1",
    "mi": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/MktInfo/MarketInfo/v1",
    "service": "https://ec.europa.eu/tools/eudamed/dtx/servicemodel/Service/v1",
    "lngs": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Common/LanguageSpecific/v1",
    "udidiDatas": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/v1",
    "commondevice": "https://ec.europa.eu/tools/eudamed/dtx/datamodel/Entity/Device/CommonDevice/v1"
}

for prefix, uri in NS.items():
    ET.register_namespace(prefix, uri)


def qname(tag: str) -> str:
    if ":" not in tag:
        return tag
    prefix, local = tag.split(":", 1)
    return f"{{{NS[prefix]}}}{local}"


def is_empty(value) -> bool:
    return value is None or value == "" or value == {} or value == []


def dict_to_xml(parent: ET.Element, data):
    if isinstance(data, dict):
        for key, value in data.items():
            if is_empty(value):
                continue

            if key.startswith("@"):
                attr_name = key[1:]
                if ":" in attr_name:
                    parent.set(qname(attr_name), str(value))
                else:
                    parent.set(attr_name, str(value))
                continue

            if key == "#text":
                parent.text = str(value)
                continue

            if isinstance(value, list):
                for item in value:
                    child = ET.SubElement(parent, qname(key))
                    dict_to_xml(child, item)
            else:
                child = ET.SubElement(parent, qname(key))
                dict_to_xml(child, value)
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
) -> dict:
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


def dict_to_xml_string(data: dict) -> str:
    root_key = next(iter(data))
    root_val = data[root_key]

    root = ET.Element(qname(root_key))
    dict_to_xml(root, root_val)

    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    pretty = minidom.parseString(xml_bytes).toprettyxml(indent="  ", encoding="utf-8")
    return pretty.decode("utf-8")


def save_xml(xml_str: str, output_path: str):
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(xml_str)