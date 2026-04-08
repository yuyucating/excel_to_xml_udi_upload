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
        productionIdentifier = "BATCH_NUMBER"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_serial_number")) == "true":
        productionIdentifier = "SERIALISATION_NUMBER"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_expiration_date")) == "true":
        productionIdentifier = "EXPIRATION_DATE"
    elif yn_to_bool_str(get_mapped_value(row, mapping, "COMMON", "pi_manufacturing_date")) == "true":
        productionIdentifier = "MANUFACTURING_DATE"
    else:
        productionIdentifier = None

    first_market = toISOcountry(get_mapped_value(row, mapping, "COMMON", "first_market"))
    marketing_status_description = get_mapped_value(row, mapping, "COMMON", "marketing_status")
    print("☆☆☆☆☆☆\n"+str(marketing_status_description)+"\n☆☆☆☆☆☆")
    marketing_status_list = text_to_marketing_status_list(marketing_status_description) if marketing_status_description else []

    return {
        "riskClass": str(risk_class).upper().replace(" ", "_") if pd.notna(risk_class) else None,
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
        "first_market": first_market,
        "marketing_status_list": marketing_status_list
    }


def row_to_dict_MDR_DEVICE_POST(row, mapping, ActorCodes, export_mode="DEVICE_POST"):
    print("進到 row_to_dict_MDR_DEVICE_POST()")
    c = build_common_fields(row, mapping)
    print("marketing_status_list =", c.get("marketing_status_list"))
    print("first_market =", c.get("first_market"))
    market_infos = marketInfos_to_dict(c.get("marketing_status_list", []), c.get("first_market"))

    b_di_code = get_mapped_value(row, mapping, "MDR", "basicudi_di")
    b_entity = "GS1"
    MFActorCode = ActorCodes.get("MFActorCode") # modified for mdi Europa GmbH - 2026-03-19
    ARActorCode = ActorCodes.get("ARActorCode") # modified for mdi Europa GmbH - 2026-03-19
    mnb_actor_code = "2696"

    if c["riskClass"] == "Ⅲ":
        risk_lv = 'III'
    else:
        risk_lv = 'I'*(c["riskClass"].count('I')+c["riskClass"].count('Ⅰ'))

    # certificate_type = "MDR_QUALITY_MANAGEMENT_SYSTEM"
    certificate_type = "MDR_TYPE_EXAMINATION"


    mdr_dict = {
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
                "basicudi:ARActorCode": ARActorCode,
                "basicudi:humanTissuesCells": c["is_humanTissuesCells"],
                "basicudi:MFActorCode": MFActorCode,
                **(
                    {
                        "basicudi:deviceCertificateLinks": {
                            "links:deviceCertificateLink": {
                                "links:certificateNumber": c["certificate_no"],
                                "links:NBActorCode": mnb_actor_code,
                                "links:certificateType": certificate_type
                            }
                        }
                    }
                    if certificate_type != "MDR_I" else {} 
                ),


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
                # NO MARKETINFOS
                "udidi:numberOfReuses": c["numberOfReuses"],
                **(
                    {
                        "udidi:marketInfos": market_infos
                    } if market_infos else {}
                ),
                "udidi:baseQuantity": c["baseQuantity"],
                # NO MARKETINFOS
                "udidi:latex": c["is_latex"],
                # NO MARKETINFOS
                "udidi:reprocessed": c["is_reprocessed"],
                # NO MARKETINFOS
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
                ),   
    }}}

    if export_mode == "UDI_DI_POST":
        mdr_dict["device:Device"].pop("device:MDRBasicUDI", None)

    return mdr_dict


def row_to_dict_MDD_DEVICE_POST(row, mapping, ActorCodes, export_mode="DEVICE_POST"):
    c = build_common_fields(row, mapping)
    b_di_code = get_mapped_value(row, mapping, "MDD", "basicudi_di") # modified to use tc_jsb070 for basicudi DICode - 2026-03-19
    b_entity = "EUDAMED" # modified to use EUDAMED as issuing entity for MDD - 2026-03-26
    MFActorCode = ActorCodes.get("MFActorCode") # modified for mdi Europa GmbH - 2026-03-19
    ARActorCode = ActorCodes.get("ARActorCode") # modified for mdi Europa GmbH - 2026-03-19
    basicudi = safe_str(get_mapped_value(row, mapping, "MDD", "basicudi_di"))

    critical_warnings = safe_str(get_mapped_value(row, mapping, "MDD", "critical_warning"))
    warning_value = "CW010" if critical_warnings == "Consult Instruction for Use" else None
    
    certificate_expiry_raw = get_mapped_value(row, mapping, "MDD", "certificate_expiry")
    certificate_expiry_mdd = str(certificate_expiry_raw).split(" ")[0] if pd.notna(certificate_expiry_raw) else None
    
    mnb_actor_code = "2195"
    certificate_revision = safe_str(get_mapped_value(row, mapping, "MDD", "certificate_revision"))
    
    if c["riskClass"] == "Ⅲ":
        risk_lv = 'III' # modified to handle case where risk class is 'Ⅲ' - 2026-03-19
    else:
        risk_lv = 'I'*(c["riskClass"].count('I')+c["riskClass"].count('Ⅰ')) # modified to count both 'I' and 'Ⅰ' - 2026-03-19

    certificate_type = "MDD_"+risk_lv # modified to use certificate type MDD - 2026-03-19

    mdd_dict = {
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
                **(
                    {
                        "udidi:marketInfos": {
                            "marketinfo:marketInfo": [
                                {
                                    "marketinfo:country": item["country"],
                                    "marketinfo:originalPlacedOnTheMarket": "true",
                                    # "marketinfo:startDate": item["datestart"]
                                }
                                for item in c.get("marketing_status_list", [])
                            ]
                        }
                    }
                    if c.get("marketing_status_list") else {}
                ),
                # "udidi:marketInfos": {
                #     "marketinfo:marketInfo": {
                #         "marketinfo:country": "ES",
                #         "marketinfo:originalPlacedOnTheMarket": "true"
                #     }
                # },
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
                "basicudi:MFActorCode": MFActorCode,
                **(
                    {
                        "basicudi:deviceCertificateLinks": {
                            "links:deviceCertificateLink": {
                                "links:certificateNumber": c["certificate_no"],
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

    return mdd_dict

def row_to_dict_MDR_UDIDI_POST(row, mapping, ActorCodes, export_mode="UDIDI_POST"):
    c = build_common_fields(row, mapping)
    market_infos = marketInfos_to_dict(c.get("marketing_status_list", []), c.get("first_market"), export_mode)

    mdr_dict = {
        "udidiDatas:UDIDIData": {
            "@xsi:type": "udidi:MDRUDIDIDataType",
            # "@xmlns:xs": "http://www.w3.org/2001/XMLSchema-instance",
            "udidi:identifier": {
                "commondevice:DICode": c["i_DICode"],
                "commondevice:issuingEntityCode": c["i_Entity"]
            },
            "udidi:status": {
                "commondevice:code": c["udi_status"]
            },
            "udidi:basicUDIIdentifier": {
                "commondevice:DICode": c["i_DICode"],
                "commondevice:issuingEntityCode": c["i_Entity"]
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
                    }
                }
                if c.get("tradeName") and c["tradeName"] != 'N' else {}
            ),
            "udidi:numberOfReuses": c["numberOfReuses"],
            **(
                {
                    "udidi:marketInfos": market_infos
                } if market_infos else {}
            ),
            "udidi:baseQuantity": c["baseQuantity"],
            "udidi:latex": c["is_latex"],
            "udidi:reprocessed": c["is_reprocessed"],
            "udidi:clinicalSizes": {
                "commondevice:clinicalSize": {
                    "@xsi:type": "commondevice:TextClinicalSizeType",
                    "commondevice:clinicalSizeType": "CST999",
                    "commondevice:clinicalSizeDescription":{
                        "lsn:name": {
                            "lsn:language": "EN",
                            "lsn:textValue": c["spec"][1] if c.get("spec") and len(c["spec"]) >= 2 else None
                        }
                    },
                    "commondevice:text": c["spec"][0] if c.get("spec") and len(c["spec"]) >= 2 else None
                }
            }
        }
    }

    return mdr_dict

def marketInfos_to_dict(marketing_status_list, first_market, export_mode="DEVICE_POST"):
    if export_mode == "UDI_DI_POST":
        miKey = "mi"
    elif export_mode == "DEVICE_POST":
        miKey = "marketinfo"
    market_info_list = []
    for item in marketing_status_list:
        if item["country"] == "N": continue
        market_info = {
            f"{miKey}:country": item["country"],
            f"{miKey}:originalPlacedOnTheMarket": "false",
            # f"{miKey}:startDate": item["datestart"]
        }
        market_info_list.append(market_info)

    if first_market and first_market != "N":

        if first_market not in [item[f"{miKey}:country"] for item in market_info_list]:
            market_info_list.append({
                f"{miKey}:country": first_market,
                f"{miKey}:originalPlacedOnTheMarket": "true",
                # f"{miKey}:startDate": item["datestart"]
            })
        else:
            for market_info in market_info_list:
                if market_info[f"{miKey}:country"] == first_market:
                    market_info[f"{miKey}:originalPlacedOnTheMarket"] = "true"
                    break

    return {f"{miKey}:marketInfo": market_info_list} if market_info_list else {}

def df_to_dict(df, ActorCodes, field_mapping=None, export_mode="DEVICE_POST"): #TODO ActorCodes
    device_dict_list = []
    mapping = merge_field_mapping(field_mapping)

    for _, row in df.iterrows():
        reg_type = get_mapped_value(row, mapping, "COMMON", "reg_type")
        print("★★★★★★★ reg_type =", reg_type)

        if reg_type == "MDD":
            device_data = row_to_dict_MDD_DEVICE_POST(row, mapping, ActorCodes, export_mode=export_mode)
        elif reg_type == "MDR":
            device_data = row_to_dict_MDR_DEVICE_POST(row, mapping, ActorCodes, export_mode=export_mode) if export_mode == "DEVICE_POST" else row_to_dict_MDR_UDIDI_POST(row, mapping, ActorCodes, export_mode=export_mode)
        else:
            continue
        print("★★★★★★★ reg_type =", reg_type, "DONE")
        print("★★★★★★★ device_data =", device_data)

        device_dict_list.append(device_data)

    return device_dict_list