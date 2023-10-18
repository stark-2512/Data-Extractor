import re
from bs4 import BeautifulSoup
import openpyxl
import requests
import pandas as pd
import xlsxwriter
import category_maker

whole_tables_data = []
sheet_names = ['MS-R', 'HSD-R', 'HSD-D', 'HSD', 'LPG']


def extraction():
    with open('rawData.Html', 'r') as file:
        html = file.read()

    parsed = BeautifulSoup(html, 'html.parser')

    tables = parsed.find_all('table')
    dataframes = []

    for table in tables:
        table_data = []
        rows = table.find_all('tr')
        for row in rows:
            cells = row.find_all('td')
            row_data = []
            for cell in cells:
                data = cell.text
                cleaned_data = re.sub(r'[\t\n]', '', data)
                if(cleaned_data.isdigit()):
                    row_data.append(int(cleaned_data))
                else:
                    row_data.append(cleaned_data)
            
            table_data.append(row_data)
        
        
        df = pd.DataFrame(table_data)
        dataframes.append(df)
        
        whole_tables_data.append(table_data)
    
        # Writing RO Shares
    df_ro_share_raw = pd.read_excel("ROShareRaw.xlsx")
    # df_ro_share_raw.rename(columns={
    #     0 : 'IOC',
    #     1 : 'BPC',
    #     2 : 'HPC',
    #     3 : 'PVT',
    #     4 : 'Total'
    # })
    # df_ro_share_raw.drop(0, inplace=True)
    
    calculatedRoShare = pd.DataFrame(columns=["IOC RO Share", "BPC RO Share", "HPC RO Share", "PVT RO Share"])
    ioc_ro_share = []
    bpc_ro_share = []
    hpc_ro_share = []
    pvt_ro_share = []
    i = 0
    for _,row in df_ro_share_raw.iterrows():
        # print(type(row[1]), row[1])
        if(row[5] == 0):
            ioc_ro_share.append(0)
            bpc_ro_share.append(0)
            hpc_ro_share.append(0)
            pvt_ro_share.append(0)    
            continue
            
        ioc_ro_share.append((row[1] / row[5]) * 100)
        bpc_ro_share.append((row[2] / row[5]) * 100)
        hpc_ro_share.append((row[3] / row[5]) * 100)
        pvt_ro_share.append((row[4] / row[5]) * 100)
    df_ro_share_raw.insert(0, "IOC RO Share", ioc_ro_share)
    df_ro_share_raw.insert(0, "BPC RO Share", bpc_ro_share)
    df_ro_share_raw.insert(0, "HPC RO Share", hpc_ro_share)
    df_ro_share_raw.insert(0, "PVT RO Share", pvt_ro_share)
    df_ro_share_raw.to_excel("ROShare.xlsx", engine='openpyxl', index=False) 

extraction()

i = 0
for table in whole_tables_data:
    category_maker.make(raw_category_table=table, type=sheet_names[i])
    if(i == 4):
        break
    i = i+1

print("done")
# print(whole_tables_data)