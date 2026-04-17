import pandas as pd
from test_marketing_status import text_to_marketing_status_list, toISOcountry

DEFAULT_FIELD_MAPPING = {
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
        "basicudi_di": "tc_jsb070",
        "critical_warning": "tc_jsb730",
        "certificate_revision": "tc_jsb180",
        "certificate_expiry": "tc_jsb710",
    },
    "MDR": {
        "basicudi_di": "tc_jsb070",
    },
}


def merge_field_mapping(field_mapping=None):
    merged = {
        "COMMON": dict(DEFAULT_FIELD_MAPPING["COMMON"]),
        "MDD": dict(DEFAULT_FIELD_MAPPING["MDD"]),
        "MDR": dict(DEFAULT_FIELD_MAPPING["MDR"]),
    }

    if not field_mapping:
        return merged

    for section in ("COMMON", "MDD", "MDR"):
        merged[section].update(field_mapping.get(section, {}))

    return merged


def get_mapped_value(row, mapping, section, key, default=None):
    col_name = mapping.get(section, {}).get(key)
    if not col_name:
        return default
    return row[col_name] if col_name in row.index else default


def yn_to_bool_str(value):
    if pd.isna(value):
        return "false"
    return "true" if str(value).strip().upper() == "Y" else "false"


def safe_str(value):
    return value if pd.notna(value) else None


def safe_int(value, default=0):
    if pd.isna(value):
        return default
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def normalize_date(value):
    if pd.isna(value):
        return None
    return str(value).split(" ")[0]


def normalize_risk_class(raw_value):
    """
    將 'Ⅲ', 'III', 'Ⅱ', 'Ⅰ', 'II', 'I' 統一成 API 需要的 III / II / I
    """
    if pd.isna(raw_value):
        return None

    text = str(raw_value).strip().upper()
    text = text.replace("Ⅲ", "III").replace("Ⅱ", "II").replace("Ⅰ", "I")
    text = text.replace(" ", "")

    if text in {"I", "II", "III"}:
        return text

    # 保底：只保留 I 計數
    i_count = text.count("I")
    if i_count > 0:
        return "I" * i_count

    return text


def build_spec(row, mapping):
    spec_value = get_mapped_value(row, mapping, "COMMON", "spec_value")
    spec_unit = get_mapped_value(row, mapping, "COMMON", "spec_unit")

    if pd.notna(spec_value) and pd.notna(spec_unit):
        return [spec_value, spec_unit]
    return None


def determine_pi_code(row, mapping):
    if yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_lot_number")) == "true":
        return "BATCH_NUMBER"
    if yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_serial_number")) == "true":
        return "SERIALISATION_NUMBER"
    if yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_expiration_date")) == "true":
        return "EXPIRATION_DATE"
    if yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_manufacturing_date")) == "true":
        return "MANUFACTURING_DATE"
    return None


def build_common_context(row, mapping):
    risk_class_raw = get_mapped_value(row, mapping, "COMMON", "risk_class")
    udi_status_raw = get_mapped_value(row, mapping, "COMMON", "udi_status")

    first_market = toISOcountry(get_mapped_value(row, mapping, "COMMON", "first_market"))
    marketing_status_description = get_mapped_value(row, mapping, "COMMON", "marketing_status")
    marketing_status_list = (
        text_to_marketing_status_list(marketing_status_description)
        if marketing_status_description
        else []
    )

    return {
        "riskClass": normalize_risk_class(risk_class_raw),
        "model": get_mapped_value(row, mapping, "COMMON", "model"),
        "certificate_no": safe_str(get_mapped_value(row, mapping, "COMMON", "certificate_no")),
        "is_animalTissuesCells": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "animal_tissues_cells")),
        "is_humanTissuesCells": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "human_tissues_cells")),
        "is_medicinalProduct": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "medicinal_product")),
        "is_humanProduct": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "human_product")),
        "productType": "DEVICE",
        "is_active": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "active")),
        "is_administering": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "administering")),
        "is_implantable": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "implantable")),
        "is_measuring": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "measuring")),
        "is_reusable": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "reusable")),
        "i_DICode": safe_str(get_mapped_value(row, mapping, "COMMON", "udi_di")),
        "i_Entity": "GS1",
        "udi_status": str(udi_status_raw).upper().replace("EU ", "").replace(" ", "_") if pd.notna(udi_status_raw) else None,
        "emdn_code": safe_str(get_mapped_value(row, mapping, "COMMON", "emdn_code")),
        "pi_code": determine_pi_code(row, mapping),
        "productNumber": get_mapped_value(row, mapping, "COMMON", "product_number"),
        "is_sterile": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "sterile")),
        "is_sterilization": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "sterilization")),
        "tradeName": get_mapped_value(row, mapping, "COMMON", "trade_name"),
        "tradeName_lang": get_mapped_value(row, mapping, "COMMON", "trade_name_lang"),
        "numberOfReuses": safe_int(get_mapped_value(row, mapping, "COMMON", "number_of_reuses"), 0),
        "baseQuantity": safe_int(get_mapped_value(row, mapping, "COMMON", "base_quantity"), 0),
        "is_latex": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "latex")),
        "is_reprocessed": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "reprocessed")),
        "spec": build_spec(row, mapping),
        "first_market": first_market,
        "marketing_status_list": marketing_status_list,
    }


def build_context(row, mapping, actor_codes, reg_type, export_mode):
    common = build_common_context(row, mapping)

    if reg_type == "MDR":
        basicudi_code = get_mapped_value(row, mapping, "MDR", "basicudi_di")
        basicudi_entity = "GS1"
        nb_actor_code = "2696"
        certificate_type = "MDR_TYPE_EXAMINATION"
    elif reg_type == "MDD":
        basicudi_code = get_mapped_value(row, mapping, "MDD", "basicudi_di")
        basicudi_entity = "EUDAMED"
        nb_actor_code = "2195"

        risk_lv = common["riskClass"]
        certificate_type = f"MDD_{risk_lv}" if risk_lv else None
    else:
        raise ValueError(f"Unsupported reg_type: {reg_type}")

    return {
        **common,
        "reg_type": reg_type,
        "export_mode": export_mode,
        "MFActorCode": actor_codes.get("MFActorCode"),
        "ARActorCode": actor_codes.get("ARActorCode"),
        "basicudi_code": safe_str(basicudi_code),
        "basicudi_entity": basicudi_entity, # MDR->"GS1", MDD->"EUDAMED"
        "nb_actor_code": nb_actor_code,
        "certificate_type": certificate_type,
        "risk_lv": common["riskClass"],
        "market_infos": market_infos_to_dict(
            common.get("marketing_status_list", []),
            common.get("first_market"),
            prefix="mi" if export_mode == "UDI_DI_POST" else "marketinfo",
        ),
        "certificate_revision": safe_str(get_mapped_value(row, mapping, "MDD", "certificate_revision")),
        "certificate_expiry": normalize_date(get_mapped_value(row, mapping, "MDD", "certificate_expiry")),
        "critical_warning": safe_str(get_mapped_value(row, mapping, "MDD", "critical_warning")),
    }


def compact_dict(data):
    if isinstance(data, dict):
        return {
            k: compact_dict(v)
            for k, v in data.items()
            if v is not None and v != {} and v != []
        }
    if isinstance(data, list):
        return [compact_dict(v) for v in data if v is not None and v != {} and v != []]
    return data


def build_trade_names(prefix, lang, text):
    if not text or text == "N":
        return None

    return {
        f"{prefix}:tradeNames": {
            "lsn:name": {
                "lsn:language": lang,
                "lsn:textValue": text,
            }
        }
    }


def build_clinical_sizes(root_prefix, common_prefix, spec):
    if not spec or len(spec) < 2:
        return None

    return {
        f"{root_prefix}:clinicalSizes": {
            f"{common_prefix}:clinicalSize": {
                "@xsi:type": f"{common_prefix}:TextClinicalSizeType",
                f"{common_prefix}:clinicalSizeType": "CST999",
                f"{common_prefix}:clinicalSizeDescription": {
                    "lsn:name": {
                        "lsn:language": "EN",
                        "lsn:textValue": spec[1],
                    }
                },
                f"{common_prefix}:text": spec[0],
            }
        }
    }


def build_mdd_critical_warnings(ctx):
    if not ctx.get("critical_warning"):
        return None

    warning_value = "CW010" if ctx["critical_warning"] == "Consult Instruction for Use" else None

    return {
        "udidi:criticalWarnings": {
            "commondi:warning": {
                "commondi:comments": {
                    "lsn:name": {
                        "lsn:language": "ANY",
                        "lsn:textValue": ctx["critical_warning"],
                    }
                },
                "commondi:warningValue": warning_value,
            }
        }
    }


def build_basicudi_common(ctx):
    cert_links = build_certificate_links(ctx) or {}

    return compact_dict({
        "basicudi:riskClass": f'CLASS_{ctx["risk_lv"]}' if ctx.get("risk_lv") else None,
        "basicudi:model": ctx["model"],
        "basicudi:identifier": {
            "commondi:DICode": ctx["basicudi_code"],
            "commondi:issuingEntityCode": ctx["basicudi_entity"],
        },
        "basicudi:animalTissuesCells": ctx["is_animalTissuesCells"],
        "basicudi:ARActorCode": ctx["ARActorCode"],
        "basicudi:humanTissuesCells": ctx["is_humanTissuesCells"],
        "basicudi:MFActorCode": ctx["MFActorCode"],
        **cert_links,
        "basicudi:humanProductCheck": ctx["is_humanProduct"],
        "basicudi:medicinalProductCheck": ctx["is_medicinalProduct"],
        "basicudi:type": ctx["productType"],
        "commondi:active": ctx["is_active"],
        "commondi:administeringMedicine": ctx["is_administering"],
        "commondi:implantable": ctx["is_implantable"],
        "commondi:measuringFunction": ctx["is_measuring"],
        "commondi:reusable": ctx["is_reusable"],
    })


def build_certificate_links(ctx):
    cert_type = ctx.get("certificate_type")
    if not cert_type:
        return None

    if cert_type in {"MDD_I", "MDR_I"}:
        return None

    link = {
        "links:certificateNumber": ctx["certificate_no"],
        "links:NBActorCode": ctx["nb_actor_code"],
        "links:certificateType": cert_type,
    }

    # if ctx["reg_type"] == "MDD":
    #     #link["links:expiryDate"] = ctx["certificate_expiry"]
    #     link["links:certificateRevisionNumber"] = ctx["certificate_revision"]

    return {
        "basicudi:deviceCertificateLinks": {
            "links:deviceCertificateLink": compact_dict(link)
        }
    }


def build_udidi_common(ctx, *, root_prefix, common_prefix, include_pi=True, include_base_quantity=True):
    clinical_sizes = build_clinical_sizes(root_prefix, common_prefix, ctx["spec"]) or {}
    trade_names = build_trade_names(root_prefix, ctx["tradeName_lang"], ctx["tradeName"]) or {}

    # MDD
    critical_warnings = build_mdd_critical_warnings(ctx) or {}

    data = {
        f"{root_prefix}:identifier": {
            f"{common_prefix}:DICode": ctx["i_DICode"],
            f"{common_prefix}:issuingEntityCode": ctx["i_Entity"],
        },
        f"{root_prefix}:status": {
            f"{common_prefix}:code": ctx["udi_status"]
        },
        f"{root_prefix}:basicUDIIdentifier": {
            f"{common_prefix}:DICode": ctx["basicudi_code"],
            f"{common_prefix}:issuingEntityCode": ctx["basicudi_entity"],
        },
        f"{root_prefix}:MDNCodes": ctx["emdn_code"],
        f"{root_prefix}:productionIdentifier": ctx["pi_code"] if include_pi else None,
        f"{root_prefix}:referenceNumber": ctx["productNumber"],
        f"{root_prefix}:sterile": ctx["is_sterile"],
        f"{root_prefix}:sterilization": ctx["is_sterilization"],
        **trade_names,
        **critical_warnings,
        f"{root_prefix}:numberOfReuses": ctx["numberOfReuses"],
        f"{root_prefix}:marketInfos": ctx["market_infos"] or None,
        f"{root_prefix}:baseQuantity": ctx["baseQuantity"] if include_base_quantity else None,
        f"{root_prefix}:latex": ctx["is_latex"],
        f"{root_prefix}:reprocessed": ctx["is_reprocessed"],
        **clinical_sizes,
    }

    return compact_dict(data)


def market_infos_to_dict(marketing_status_list, first_market, prefix="marketinfo"):
    market_info_list = []

    for item in marketing_status_list:
        country = item.get("country")
        if country == "N":
            continue

        market_info_list.append({
            f"{prefix}:country": country,
            f"{prefix}:originalPlacedOnTheMarket": "false",
        })

    if first_market and first_market != "N":
        existed = False

        for market_info in market_info_list:
            if market_info[f"{prefix}:country"] == first_market:
                market_info[f"{prefix}:originalPlacedOnTheMarket"] = "true"
                existed = True
                break

        if not existed:
            market_info_list.append({
                f"{prefix}:country": first_market,
                f"{prefix}:originalPlacedOnTheMarket": "true",
            })

    return {f"{prefix}:marketInfo": market_info_list} if market_info_list else {}


def build_mdr_device_post(ctx):
    basicudi = build_basicudi_common(ctx)

    return {
        "device:Device": {
            "@xsi:type": "device:MDRDeviceType",
            "device:MDRBasicUDI": compact_dict({
                "@xsi:type": "device:MDRBasicUDIType",
                **basicudi,
            }),
            "device:MDRUDIDIData": compact_dict({
                "@xsi:type": "device:MDRUDIDIDataType",
                **build_udidi_common(
                    ctx,
                    root_prefix="udidi",
                    common_prefix="commondi",
                    include_pi=True,
                    include_base_quantity=True,
                ),
            }),
        }
    }

def build_mdr_udidi_post(ctx):
    return {
        "udidiDatas:UDIDIData": compact_dict({
            "@xsi:type": "udidi:MDRUDIDIDataType",
            **build_udidi_common(
                ctx,
                root_prefix="udidi",
                common_prefix="commondevice",
                include_pi=True,
                include_base_quantity=True,
            ),
        })
    }


def build_mdd_device_post(ctx):
    mdeu_data = build_udidi_common(
        ctx,
        root_prefix="udidi",
        common_prefix="commondi",
        include_pi=False,
        include_base_quantity=False,
    )

    basicudi = build_basicudi_common(ctx)
    basicudi["eudi:applicableLegislation"] = "MDD"

    return {
        "device:Device": {
            "@xsi:type": "device:MDEUDeviceType",
            "device:MDEUData": compact_dict(mdeu_data),
            "device:MDEUDI": compact_dict(basicudi),
        }
    }


BUILDERS = {
    ("MDR", "DEVICE_POST"): build_mdr_device_post,
    ("MDR", "UDI_DI_POST"): build_mdr_udidi_post,
    ("MDD", "DEVICE_POST"): build_mdd_device_post,
}


def df_to_dict(df, actor_codes, field_mapping=None, export_mode="DEVICE_POST"):
    mapping = merge_field_mapping(field_mapping)
    device_dict_list = []

    for _, row in df.iterrows():
        reg_type = get_mapped_value(row, mapping, "COMMON", "reg_type")
        builder = BUILDERS.get((reg_type, export_mode))
        if not builder:
            continue

        ctx = build_context(row, mapping, actor_codes, reg_type, export_mode)
        device_data = builder(ctx)
        device_dict_list.append(device_data)

    return device_dict_list