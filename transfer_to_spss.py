import pandas as pd

df = pd.read_excel("source\\VINCENZO 2024 - CODEFRAME - Final - updated 10052024.xlsx", engine="openpyxl",sheet_name="CODEFRAME CSAT - proccess")
df.set_index(["LV1","LV2","LV3","LV4","LV5"], inplace=True)
print(df)

