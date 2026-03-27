import pandas as pd

DEFAULT_FIELD_MAPPING = {
    "COMMON": {
        "reg_type": "tc_jsb030",
        "risk_class": "tc_jsb080",
        "model": "tc_jsb200",
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
        "spec_unit": "tc_jsb440"
    },
    "MDD": {
        "basicudi_di": "tc_jsb070",
        "critical_warning": "tc_jsb730",
        "certificate_no": "tc_jsb170",
        "certificate_revision": "tc_jsb180",
        "certificate_expiry": "tc_jsb710"
    },
    "MDR": {
        "basicudi_di": "tc_jsb070"
    }
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

# Below are the original helper functions and mapping logic, which can be refactored to use the above mapping system for better flexibility. For now, they remain unchanged.

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


def build_spec(row, mapping):
    spec_value = get_mapped_value(row, mapping, "COMMON", "spec_value")
    spec_unit = get_mapped_value(row, mapping, "COMMON", "spec_unit")

    if pd.notna(spec_value) and pd.notna(spec_unit):
        spec = [spec_value, spec_unit]
        return spec
    return None


def build_common_fields(row, mapping):

    risk_class = get_mapped_value(row, mapping, "COMMON", "risk_class")
    udi_status = get_mapped_value(row, mapping, "COMMON", "udi_status")

    if yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_lot_number")) == "true":
        productionIdentifier = "LOT_NUMBER"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_serial_number")) == "true":
        productionIdentifier = "SERIALISATION_NUMBER"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_expiration_date")) == "true":
        productionIdentifier = "EXPIRATION_DATE"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_manufacturing_date")) == "true":
        productionIdentifier = "MANUFACTURING_DATE"
    else:
        productionIdentifier = None

    return {
        "riskClass": str(risk_class).upper().replace(" ", "_") if pd.notna(risk_class) else None,
        "model": get_mapped_value(row, mapping, "COMMON", "model"),
        "is_animalTissuesCells": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "animal_tissues_cells")),
        "is_humanTissuesCells": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "human_tissues_cells")),
        "MFActorCode": "TW-MF-000017454",
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
        "udi_status": str(udi_status).upper().replace("EU ", "").replace(" ", "_") if pd.notna(udi_status) else None,
        "emdn_code": safe_str(get_mapped_value(row, mapping, "COMMON", "emdn_code")),
        "is_pi_lotNumber": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_lot_number")), # can be removed if pi_code can work
        "is_pi_serialNumber": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_serial_number")), # can be removed if pi_code can work
        "is_pi_manufacturingDate": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_manufacturing_date")), # can be removed if pi_code can work
        "is_pi_expirationDate": yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_expiration_date")), # can be removed if pi_code can work
        "pi_code": productionIdentifier, # 2026-03-27
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
    }


def row_to_dict_MDR(row, mapping):
    c = build_common_fields(row, mapping)
    b_di_code = get_mapped_value(row, mapping, "MDR", "basicudi_di")
    b_entity = "GS1"
    risk_lv = 'I'*c["riskClass"].count('I')

    return {
        "device:Device": {
            "@xsi:type": "device:MDRDeviceType",
            "device:MDRBasicUDI": {
                "@xsi:type": "device:MDRBasicUDIType",
                "basicudi:riskClass": 'CLASS_'+risk_lv,
                "basicudi:model": c["model"],
                "basicudi:identifier": {
                    "commondi:DICode": b_di_code,
                    "commondi:issuingEntityCode": b_entity
                },
                "basicudi:animalTissuesCells": c["is_animalTissuesCells"],
                "basicudi:humanTissuesCells": c["is_humanTissuesCells"],
                "basicudi:MFActorCode": c["MFActorCode"],
                "basicudi:humanProductCheck": c["is_humanProduct"],
                "basicudi:medicinalProductCheck": c["is_medicinalProduct"],
                "basicudi:type": c["productType"],
                "commondi:active": c["is_active"],
                "commondi:administeringMedicine": c["is_administering"],
                "commondi:implantable": c["is_implantable"],
                "commondi:measuringFunction": c["is_measuring"],
                "commondi:reusable": c["is_reusable"]
            },
            "device:MDRUDIDIData": {
                "@xsi:type": "device:MDRUDIDIDataType",
                "udidi:identifier": {
                    "commondi:DICode": c["i_DICode"],
                    "commondi:issuingEntityCode": c["i_Entity"]
                },
                "udidi:status": {
                    "commondi:code": c["udi_status"]
                },
                "udidi:basicUDIIdentifier": {
                    "commondi:DICode": b_di_code,
                    "commondi:issuingEntityCode": b_entity
                },
                "udidi:MDNCodes": c["emdn_code"],
                # "udidi:productionIdentifier": {
                #     "udidi:lotNumber": c["is_pi_lotNumber"],
                #     "udidi:serialNumber": c["is_pi_serialNumber"],
                #     "udidi:manufacturingDate": c["is_pi_manufacturingDate"],
                #     "udidi:expirationDate": c["is_pi_expirationDate"]
                # },
                "udidi:productionIdentifier": c["pi_code"], # 2026-03-27 modified to use pi_code for production identifier
                "udidi:referenceNumber": c["productNumber"],
                "udidi:sterile": c["is_sterile"],
                "udidi:sterilization": c["is_sterilization"],
                **(
                    {
                        "udidi:tradeNames": {
                            "lsn:name": {
                                "lsn:language": c["tradeName_lang"],
                                "lsn:textValue": c["tradeName"]
                        }},
                    } if c.get("tradeName") and c["tradeName"] != 'N' else {}
                ),
                "udidi:numberOfReuses": c["numberOfReuses"],
                "udidi:baseQuantity": c["baseQuantity"],
                "udidi:latex": c["is_latex"],
                "udidi:reprocessed": c["is_reprocessed"],
                **( # dict unpack 寫法：只有當 c["spec"] 不為 None 時才會加入 clinicalSizes 的內容 - 2026-03-20
                    {
                        "udidi:clinicalSizes": {
                            "commondi:clinicalSize": {
                                "@xsi:type": "commondi:TextClinicalSizeType",
                                "commondi:clinicalSizeType": "CST999",
                                "commondi:clinicalSizeDescription":{
                                    "lsn:name": {
                                        "lsn:language": "EN",
                                        "lsn:textValue": c["spec"][1]
                                }},
                                "commondi:text": c["spec"][0]
                    }}}
                    if c.get("spec") and len(c["spec"]) >= 2 else {}
                )
    }}}


def row_to_dict_MDD(row, mapping):
    c = build_common_fields(row, mapping)
    b_di_code = get_mapped_value(row, mapping, "MDD", "basicudi_di") # modified to use tc_jsb070 for basicudi DICode - 2026-03-19
    b_entity = "EUDAMED" # modified to use EUDAMED as issuing entity for MDD - 2026-03-26
    ARActorCode = "DE-AR-000006218" # modified for mdi Europa GmbH - 2026-03-19
    basicudi = safe_str(get_mapped_value(row, mapping, "MDD", "basicudi_di"))

    critical_warnings = safe_str(get_mapped_value(row, mapping, "MDD", "critical_warning"))
    warning_value = "CW010" if critical_warnings == "Consult Instruction for Use" else None
    certificate_mdd = safe_str(get_mapped_value(row, mapping, "MDD", "certificate_no"))
    
    certificate_expiry_raw = get_mapped_value(row, mapping, "MDD", "certificate_expiry")
    certificate_expiry_mdd = str(certificate_expiry_raw).split(" ")[0] if pd.notna(certificate_expiry_raw) else None
    
    mnb_actor_code = "2195"
    certificate_revision = safe_str(get_mapped_value(row, mapping, "MDD", "certificate_revision"))
    
    if c["riskClass"] == "Ⅲ":
        risk_lv = 'III' # modified to handle case where risk class is 'Ⅲ' - 2026-03-19
    else:
        risk_lv = 'I'*(c["riskClass"].count('I')+c["riskClass"].count('Ⅰ')) # modified to count both 'I' and 'Ⅰ' - 2026-03-19

    certificate_type = "MDD_"+risk_lv # modified to use certificate type MDD - 2026-03-19

    return {
        "device:Device": {
            "@xsi:type": "device:MDEUDeviceType",
            "device:MDEUData": {
                "udidi:identifier": {
                    "commondi:DICode": c["i_DICode"], # UDI-DI
                    "commondi:issuingEntityCode": c["i_Entity"]
                },
                "udidi:status": {
                    "commondi:code": c["udi_status"]
                },
                "udidi:basicUDIIdentifier": {
                    "commondi:DICode": basicudi, # modified to use Basic UDI - 2026-03-19
                    "commondi:issuingEntityCode": b_entity
                },
                
                "udidi:MDNCodes": c["emdn_code"],
                "udidi:referenceNumber": c["productNumber"],
                "udidi:sterile": c["is_sterile"],
                "udidi:sterilization": c["is_sterilization"],
                **(
                    {
                        "udidi:tradeNames": {
                            "lsn:name": {
                                "lsn:language": c["tradeName_lang"],
                                "lsn:textValue": c["tradeName"]
                            }
                        },
                    } if c.get("tradeName") and c["tradeName"] != 'N' else {}
                ),
                "udidi:criticalWarnings": {
                    "commondi:warning": {
                        "commondi:comments": {
                            "lsn:name": {
                                "lsn:language": "ANY",
                                "lsn:textValue": critical_warnings
                            }
                        },
                        "commondi:warningValue": warning_value
                    }
                },
                "udidi:numberOfReuses": c["numberOfReuses"],
                "udidi:marketInfos": {
                    "marketinfo:marketInfo": {
                        "marketinfo:country": "ES",
                        "marketinfo:originalPlacedOnTheMarket": "true"
                    }
                },
                "udidi:latex": c["is_latex"],
                "udidi:reprocessed": c["is_reprocessed"],
                **( # dict unpack 寫法：只有當 c["spec"] 不為 None 時才會加入 clinicalSizes 的內容 - 2026-03-20
                    {
                        "udidi:clinicalSizes": {
                            "commondi:clinicalSize": {
                                "@xsi:type": "commondi:TextClinicalSizeType",
                                "commondi:clinicalSizeType": "CST999",
                                "commondi:clinicalSizeDescription":{
                                    "lsn:name": {
                                        "lsn:language": "EN",
                                        "lsn:textValue": c["spec"][1]
                                    }
                                },
                                "commondi:text": c["spec"][0]
                            }
                        }
                    }
                    if c.get("spec") and len(c["spec"]) >= 2 else {}
                )
            },
            "device:MDEUDI": {
                "basicudi:riskClass": 'CLASS_'+risk_lv,
                "basicudi:model": c["model"],
                "basicudi:identifier": {
                    "commondi:DICode": b_di_code,
                    "commondi:issuingEntityCode": b_entity
                },
                "basicudi:animalTissuesCells": c["is_animalTissuesCells"],
                "basicudi:ARActorCode": ARActorCode,
                "basicudi:humanTissuesCells": c["is_humanTissuesCells"],
                "basicudi:MFActorCode": c["MFActorCode"],
                **(
                    {
                        "basicudi:deviceCertificateLinks": {
                            "links:deviceCertificateLink": {
                                "links:certificateNumber": certificate_mdd,
                                "links:expiryDate": certificate_expiry_mdd,
                                "links:NBActorCode": mnb_actor_code,
                                "links:certificateRevisionNumber": certificate_revision,
                                "links:certificateType": certificate_type
                            }
                        }
                    }
                    if certificate_type != "MDD_I" else {} 
                ),
                "basicudi:humanProductCheck": c["is_humanProduct"],
                "basicudi:medicinalProductCheck": c["is_medicinalProduct"],
                "basicudi:type": c["productType"],
                "commondi:active": c["is_active"],
                "commondi:administeringMedicine": c["is_administering"],
                "commondi:implantable": c["is_implantable"],
                "commondi:measuringFunction": c["is_measuring"],
                "commondi:reusable": c["is_reusable"],
                "eudi:applicableLegislation": "MDD"
            }
        }
    }


def df_to_dict(df, field_mapping=None):
    device_dict_list = []
    mapping = merge_field_mapping(field_mapping)

    for _, row in df.iterrows():
        reg_type = get_mapped_value(row, mapping, "COMMON", "reg_type")

        if reg_type == "MDD":
            device_data = row_to_dict_MDD(row, mapping)
        elif reg_type == "MDR":
            device_data = row_to_dict_MDR(row, mapping)
        else:
            continue

        device_dict_list.append(device_data)

    return device_dict_list