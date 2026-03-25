import pandas as pd


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


def build_spec(row):
    if pd.notna(row["tc_jsb430"]) and pd.notna(row["tc_jsb440"]):
        spec = [row["tc_jsb430"], row["tc_jsb440"]] # tc_jsb430 is number, tc_jsb440 is unit
        return spec
    return None


def build_common_fields(row):
    return {
        "riskClass": row["tc_jsb080"].upper().replace(" ", "_"),
        "model": row["tc_jsb200"],
        "is_animalTissuesCells": yn_to_bool_str(row["tc_jsb360"]),
        "is_humanTissuesCells": yn_to_bool_str(row["tc_jsb350"]),
        "MFActorCode": "TW-MF-000017454",
        "is_medicinalProduct": yn_to_bool_str(row["tc_jsb370"]),
        "is_humanProduct": yn_to_bool_str(row["tc_jsb380"]),
        "productType": "DEVICE",
        "is_active": yn_to_bool_str(row["tc_jsb120"]),
        "is_administering": yn_to_bool_str(row["tc_jsb130"]),
        "is_implantable": yn_to_bool_str(row["tc_jsb090"]),
        "is_measuring": yn_to_bool_str(row["tc_jsb100"]),
        "is_reusable": yn_to_bool_str(row["tc_jsb110"]),
        "i_DICode": safe_str(row["tc_jsb001"]),
        "i_Entity": "GS1",
        "udi_status": row["tc_jsb270"].upper().replace("EU ", "").replace(" ", "_") if pd.notna(row["tc_jsb270"]) else None,
        "emdn_code": safe_str(row["tc_jsb190"]),
        "is_pi_lotNumber": yn_to_bool_str(row["tc_jsb2401"]),
        "is_pi_serialNumber": yn_to_bool_str(row["tc_jsb2411"]),
        "is_pi_manufacturingDate": yn_to_bool_str(row["tc_jsb2421"]),
        "is_pi_expirationDate": yn_to_bool_str(row["tc_jsb2431"]),
        "productNumber": row["tc_jsb000"],
        "is_sterile": yn_to_bool_str(row["tc_jsb743"]),
        "is_sterilization": yn_to_bool_str(row["tc_jsb742"]),
        "tradeName": row["tc_jsb200"],
        "tradeName_lang": row["tc_jsb210"],
        "numberOfReuses": safe_int(row["tc_jsb744"], 0),
        "baseQuantity": safe_str(row["tc_jsb230"]),
        "is_latex": yn_to_bool_str(row["tc_jsb550"]),
        "is_reprocessed": yn_to_bool_str(row["tc_jsb620"]),
        "spec": build_spec(row),
    }


def row_to_dict_MDR(row):
    c = build_common_fields(row)
    b_di_code = row["tc_jsb070"]
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
                "udidi:status": c["udi_status"],
                "udidi:basicUDIIdentifier": {
                    "commondi:DICode": b_di_code,
                    "commondi:issuingEntityCode": b_entity
                },
                "udidi:MDNCodes": c["emdn_code"],
                "udidi:productionIdentifier": {
                    "udidi:lotNumber": c["is_pi_lotNumber"],
                    "udidi:serialNumber": c["is_pi_serialNumber"],
                    "udidi:manufacturingDate": c["is_pi_manufacturingDate"],
                    "udidi:expirationDate": c["is_pi_expirationDate"]
                },
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
                "udidi:numberOfReuses": c["numberOfReuses"],
                "udidi:baseQuantity": c["baseQuantity"],
                "udidi:latex": c["is_latex"],
                "udidi:reprocessed": c["is_reprocessed"],
                "udidi:clinicalSizes": {
                    "commondi:clinicalSize": {
                        "@xsi:type": "commondi:ValueClinicalSizeType",
                        "commondi:clinicalSizeType": "CST999" if c["spec"] is not None else None, # modified - 2026-03-19
                        "commondi:text": c["spec"]
                    }
                }
            }
        }
    }


def row_to_dict_MDD(row):
    c = build_common_fields(row)
    b_di_code = row["tc_jsb070"] # modified to use tc_jsb070 for basicudi DICode - 2026-03-19
    b_entity = "GS1"
    ARActorCode = "DE-AR-000006218" # modified for mdi Europa GmbH - 2026-03-19
    basicudi = safe_str(row["tc_jsb070"]) # modified to use basicudi - 2026-03-19
    # basicudi = safe_str(row["tc_jsb630"]) # 0323

    critical_warnings = safe_str(row["tc_jsb730"])
    warning_value = "CW010" if critical_warnings == "Consult Instruction for Use" else None
    certificate_mdd = safe_str(row["tc_jsb170"])
    certificate_expiry_mdd = row["tc_jsb710"].split(" ")[0] if pd.notna(row["tc_jsb710"]) else None
    mnb_actor_code = "2195"
    certificate_revision = safe_str(row["tc_jsb180"])
    
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
                # "udidi:basicUDIIdentifier": { # removed Basic UDI for temporatory -- 2026-03-25
                #     "commondi:DICode": basicudi, # modified to use Basic UDI - 2026-03-19
                #     "commondi:issuingEntityCode": b_entity
                # },
                
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
                                # "commondi:clinicalSizeDescription": c["spec"]
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
                # "basicudi:identifier": { # removed Basic UDI for temporatory -- 2026-03-25
                #     "commondi:DICode": b_di_code,
                #     "commondi:issuingEntityCode": b_entity
                # },
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


def df_to_dict(df):
    device_dict_list = []

    for _, row in df.iterrows():
        reg_type = row["tc_jsb030"]

        if reg_type == "MDD":
            device_data = row_to_dict_MDD(row)
        elif reg_type == "MDR":
            device_data = row_to_dict_MDR(row)
        else:
            continue

        device_dict_list.append(device_data)

    return device_dict_list