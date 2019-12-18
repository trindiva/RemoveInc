import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

def get_company_list():
    df = pd.read_excel("C:\\Users\\Ivan Trindev\\OneDrive - ACT Capital Advisors\\Copy of DnS - Buyer Tracking List - 12.13.2019 - CHS.xlsx", sheet_name="Complete List")
    df = df["Firm Name"].astype(str)

    companies = df.tolist()

    return companies

def clean_list():
    dirty_list = get_company_list()
    clean_list = []

    count = 0

    while count < len(dirty_list):
        name = dirty_list[count]

        name = name.replace("Incorporated", "")
        name = name.replace("Corporation", "")
        name = name.replace("Limited", "")
        name = name.replace("Inc.", "")
        name = name.replace("Inc", "")
        name = name.replace("LLC", "")
        name = name.replace("L.L.C.", "")
        name = name.replace("Ltd.", "")
        name = name.replace("Co.", "")
        name = name.replace("Corp.", "")
        
        name = name.rstrip()

        if name[len(name) - 1] == ",":
            name = name[:len(name) - 1]

        clean_list.append(name)
        count += 1

    return clean_list

def remove_duplicates():
    original_list = clean_list()
    new_list = []

    for i in original_list:
        if i not in new_list:
            new_list.append(i)

    final_df = pd.DataFrame(list(new_list), columns = ["Company Names"])

    final_df.to_excel("C:\\Users\\Ivan Trindev\\OneDrive - ACT Capital Advisors\\New DnS - Buyer Tracking List - 08.26.2019 - SKM V1.5.xlsx", index=False)

remove_duplicates()
