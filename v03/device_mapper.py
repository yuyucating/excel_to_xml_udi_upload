import pandas as pd


def yn_to_bool_str(value):
    return "true" if value == "Y" else "false"


def safe_str(value):
    return value if pd.notna(value) else None


def build_spec(row):
    if pd.notna(row["tc_jsb430"]) and pd.notna(row["tc_jsb440"]):
        return f'{row["tc_jsb430"]} {row["tc_jsb440"]}'
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
        "numberOfReuses": row["tc_jsb744"] if isinstance(row["tc_jsb744"], int) else 0,
        "baseQuantity": safe_str(row["tc_jsb230"]),
        "is_latex": yn_to_bool_str(row["tc_jsb550"]),
        "is_reprocessed": yn_to_bool_str(row["tc_jsb620"]),
        "spec": build_spec(row),
    }


def row_to_dict_MDR(row):
    c = build_common_fields(row)
    b_di_code = row["tc_jsb070"]
    b_entity = "GS1"

    return {
        "device:Device": {
            "@xsi:type": "device:MDRDeviceType",
            "device:MDRBasicUDI": {
                "@xsi:type": "device:MDRBasicUDIType",
                "basicudi:riskClass": c["riskClass"],
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
                "udidi:tradeNames": {
                    "lsn:name": {
                        "lsn:language": c["tradeName_lang"],
                        "lsn:textValue": c["tradeName"]
                    }
                },
                "udidi:numberOfReuses": c["numberOfReuses"],
                "udidi:baseQuantity": c["baseQuantity"],
                "udidi:latex": c["is_latex"],
                "udidi:reprocessed": c["is_reprocessed"],
                "udidi:clinicalSizes": {
                    "commondi:clinicalSize": {
                        "@xsi:type": "commondi:ValueClinicalSizeType",
                        "commondi:clinicalSizeType": "CST19",
                        "commondi:text": c["spec"]
                    }
                }
            }
        }
    }


def row_to_dict_MDD(row):
    c = build_common_fields(row)
    b_di_code = row["tc_jsb630"]
    b_entity = "GS1"

    critical_warnings = safe_str(row["tc_jsb730"])
    warning_value = "CW010" if critical_warnings == "Consult Instruction for Use" else None
    certificate_mdd = safe_str(row["tc_jsb170"])
    certificate_expiry_mdd = row["tc_jsb710"].split(" ")[0] if pd.notna(row["tc_jsb710"]) else None
    mnb_actor_code = "2195"
    certificate_revision = safe_str(row["tc_jsb180"])
    certificate_type = "MDD_" + row["tc_jsb080"].split()[1]

    return {
        "device:Device": {
            "@xsi:type": "device:MDEUDeviceType",
            "device:MDEUData": {
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
                "udidi:referenceNumber": c["productNumber"],
                "udidi:sterile": c["is_sterile"],
                "udidi:sterilization": c["is_sterilization"],
                "udidi:tradeNames": {
                    "lsn:name": {
                        "lsn:language": c["tradeName_lang"],
                        "lsn:textValue": c["tradeName"]
                    }
                },
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
                        "marketinfo:originalPlacedOnTheMarket": "false"
                    }
                },
                "udidi:latex": c["is_latex"],
                "udidi:reprocessed": c["is_reprocessed"],
                "udidi:clinicalSizes": {
                    "commondi:clinicalSize": {
                        "@xsi:type": "commondi:TextClinicalSizeType",
                        "commondi:clinicalSizeType": "CST12",
                        "commondi:text": c["spec"]
                    }
                }
            },
            "device:MDEUDI": {
                "basicudi:riskClass": c["riskClass"],
                "basicudi:model": c["model"],
                "basicudi:identifier": {
                    "commondi:DICode": b_di_code,
                    "commondi:issuingEntityCode": b_entity
                },
                "basicudi:animalTissuesCells": c["is_animalTissuesCells"],
                "basicudi:humanTissuesCells": c["is_humanTissuesCells"],
                "basicudi:MFActorCode": c["MFActorCode"],
                "basicudi:deviceCertificateLinks": {
                    "links:deviceCertificateLink": {
                        "links:certificateNumber": certificate_mdd,
                        "links:expiryDate": certificate_expiry_mdd,
                        "links:NBActorCode": mnb_actor_code,
                        "links:certificateRevisionNumber": certificate_revision,
                        "links:certificateType": certificate_type
                    }
                },
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