import pandas as pd

def text_to_marketing_status_list(marketing_status_description):
    marketing_status_list = marketing_status_description.split("。") if pd.notna(marketing_status_description) else []
    marketing_status_list = [x for x in marketing_status_list if x.strip()]  # 移除空字符串
    print(marketing_status_list)
    result = []
    for status in marketing_status_list:
        datestart = None
        dateend = None
        if "," not in status:
            country = toISOcountry(status.strip())  
        else:
            country = toISOcountry(status.split(",")[0].strip())

            if "from" in status.lower():
                datestart = status.split(",")[1].split("To")[0].replace("From ", "").strip() if len(status.split(",")) > 1 else None

            if "to" in status.lower():
                dateend = status.split(",")[1].split("To")[1].strip() if len(status.split(",")) > 1 else None

        result.append({"country": country, "datestart": datestart, "dateend": dateend})
    
    return result

def toISOcountry(country_name):
    if not country_name:
        return ""

    # 統一格式
    name = str(country_name).strip()

    # ISO 已經是兩碼 → 直接回傳（大寫）
    if len(name) == 2 and name.isalpha():
        return name.upper()

    # mapping（可自行擴充）
    country_mapping = {
        # 中文 → ISO
        "西班牙": "ES",
        "葡萄牙": "PT",
        "法國": "FR",
        "德國": "DE",
        "義大利": "IT",
        "荷蘭": "NL",
        "比利時": "BE",
        "瑞典": "SE",
        "丹麥": "DK",
        "芬蘭": "FI",
        "愛爾蘭": "IE",
        "波蘭": "PL",
        "奧地利": "AT",
        "澳洲": "AU",
        "澳大利亞": "AU",
        "加拿大": "CA",
        "美國": "US",
        "日本": "JP",
        "韓國": "KR",
        "中國": "CN",
        "台灣": "TW",
        "新加坡": "SG",
        "新西蘭": "NZ",
        "捷克": "CZ",
        "匈牙利": "HU",
        "希臘": "GR",
        "羅馬尼亞": "RO",
        "保加利亞": "BG",
        "克羅埃西亞": "HR",
        "斯洛伐克": "SK",
        "斯洛維尼亞": "SI",
        "立陶宛": "LT",
        "拉脫維亞": "LV",
        "愛沙尼亞": "EE",
        "盧森堡": "LU",
        "馬爾他": "MT",
        "賽普勒斯": "CY",

        # 英文 → ISO
        "spain": "ES",
        "portugal": "PT",
        "france": "FR",
        "germany": "DE",
        "italy": "IT",
        "netherlands": "NL",
        "belgium": "BE",
        "sweden": "SE",
        "denmark": "DK",
        "finland": "FI",
        "ireland": "IE",
        "poland": "PL",
        "austria": "AT",
        "australia": "AU",
        "canada": "CA",
        "united states": "US",
        "usa": "US",
        "new zealand": "NZ",
        "japan": "JP",
        "south korea": "KR",
        "china": "CN",
        "taiwan": "TW",
        "czechia": "CZ",
        "czech republic": "CZ",
        "hungary": "HU",
        "greece": "GR",
        "romania": "RO",
        "bulgaria": "BG",
        "croatia": "HR",
        "slovakia": "SK",
        "slovenia": "SI",
        "lithuania": "LT",
        "latvia": "LV",
        "estonia": "EE",
        "luxembourg": "LU",
        "malta": "MT",
        "cyprus": "CY",
    }

    # 英文轉小寫比對
    key = name.lower()

    print(country_mapping.get(key, country_mapping.get(name, name)), "為 ISO 國家碼")
    return country_mapping.get(key, country_mapping.get(name, name))

if __name__ == "__main__":
    marketing_status_description = "西班牙,From 2023-02-22 To 2027-02-21。葡萄牙,From 2024-01-23 To 2026-01-22。"

    marketinfo_list = text_to_marketing_status_list(marketing_status_description)
    print(marketinfo_list)