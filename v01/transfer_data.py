import pandas as pd

def excel_to_df(file_path, sheet_name=None):
    if sheet_name is None:
        sheet_name = "p_zta"

    # 取出所有 data
    df = pd.read_excel(file_path, sheet_name=sheet_name, dtype={'tc_jsb001': str})
    # 移除非MDD/MDR產品
    df_eu = df.dropna(subset=["tc_jsb030", "tc_jsb080", "tc_jsb150"])
    return df_eu

def df_to_dict(df):
    device_dict_list = []
    for i in range(len(df)):
        row = df.iloc[i]

        if row["tc_jsb030"] == "MDD":
            print(row["tc_jsb001"], "MDD產品")
            dict = row_to_dict_MDD(row)
            print(dict, end="\n\n")
        elif row["tc_jsb030"] == "MDR": 
            print(row["tc_jsb001"], "MDR產品")
            dict = row_to_dict_MDR(row)
            print(dict, end="\n\n")
        
        # print(dict)
        # device_dict_list.append(dict)


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
    udi_status = row["tc_jsb270"].upper().replace("THE","").replace(" ", "_") if pd.notna(row["tc_jsb270"]) else None
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
                        "lsn:value": tradeName
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
                    }
                }
            }
        }
    }

def row_to_dict_MDD(row):
    # 回傳一個 dict，代表 MDD 產品的結構
    riskClass = "CLASS_" + row["tc_jsb080"]

    # 這裡可以根據需要定義如何將一行資料轉成 dict
    # 例如：
    return {
        "product_id": row["tc_jsb001"],
        "regulation_type": row["tc_jsb030"],
        "other_field": row["tc_jsb080"],
        # ... 根據實際欄位添加更多
    }

df = excel_to_df(r"I:\研發中心\法規課\private\A0_個人資料夾\Una Kuo\31. UDI Upload\cimi290 匯出範例_una.xlsx")
df_to_dict(df)