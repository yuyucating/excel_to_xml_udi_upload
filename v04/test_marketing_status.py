import pandas as pd

def text_to_marketing_status_list(marketing_status_description):
    print(marketing_status_description)
    marketing_status_list = marketing_status_description.split(";") if pd.notna(marketing_status_description) else []
    marketing_status_list = [x for x in marketing_status_list if x.strip()]  # 移除空字符串

    result = []
    for status in marketing_status_list:
        if "," not in status:
            continue
        country = toISOcountry(status.split(",")[0].strip())
        date_range = status.split(",")[1].replace("From ", "").replace(" To ", "/")
        datestart, dateend = date_range.split("/")
        result.append({"country": country, "datestart": datestart, "dateend": dateend})
    
    print(result)
    return result

def toISOcountry(country_name):
    country_mapping = {
        "西班牙": "ES",
        "葡萄牙": "PT",
        # 可以在這裡添加更多國家名稱和對應的 ISO 代碼
    }
    return country_mapping.get(country_name, country_name)  # 如果沒有對應，返回原始名稱

if __name__ == "__main__":
    marketing_status_description = "西班牙,From 2023-02-22 To 2027-02-21。葡萄牙,From 2024-01-23 To 2026-01-22。"

    marketinfo_list = text_to_marketing_status_list(marketing_status_description)
    print(marketinfo_list)