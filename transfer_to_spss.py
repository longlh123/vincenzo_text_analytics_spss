import win32com.client as w32
import numpy as np
import pandas as pd
import sys
sys.path.append(r'C:\Users\toan.ngo\Documents\Git Projects\vincenzo_text_analytics_spss\libs')
from metadata import SPSSObject
from metadata import Metadata
from metadata import SPSSObject_Dataframe

mdd_path = r"Metadata\VINCENZO_RGB_W12M5_TCB_BM_OE_BATCH_1_29052024_v1_checkv1.mdd"
ddf_path = r"Metadata\VINCENZO_RGB_W12M5_TCB_BM_OE_BATCH_1_29052024_v1_checkv1.ddf"

# sql_query = "SELECT * FROM VDATA WHERE InstanceID = 15435420 or InstanceID = 15464679"
# sql_query = "SELECT * FROM VDATA WHERE InstanceID = 15435420"
sql_query = "SELECT * FROM VDATA"
questions = ["InstanceID", "_Q9a_Codes"] 

spssObject = SPSSObject(mdd_path, ddf_path, sql_query, questions)
# spssObject = SPSSObject_Dataframe(mdd_path, ddf_path, sql_query, questions)
# spssObject = Metadata(mdd_path = mdd_path, ddf_path = ddf_path, sql_query = sql_query)
df = pd.DataFrame(data=spssObject.records, columns=spssObject.varNames)
# df = spssObject.convertToDataFrame(questions=questions)

column = df.columns.drop("InstanceID")

df_excel = pd.read_excel("source\\2024 - 066 - VINCENZO - Format data SPSS_RBG Q224.xlsx", engine="openpyxl",sheet_name="CODEFRAME-NPS")
df_excel.set_index(["CODE","LV1","ATTITUDE","LV2"], inplace=True)
SPSS_column_Q9a = pd.read_excel("source\\2024 - 066 - VINCENZO - Format data SPSS_RBG Q224.xlsx", engine="openpyxl",sheet_name="Variable View_OE")
SPSS_column_Q9a_selected = list(SPSS_column_Q9a["Column"])
Codelist = [];
for i in list(df_excel.index):
    Codelist.append(i)

df_spss = pd.DataFrame()
df_spss.insert(0, "InstanceID", df["InstanceID"])
for i in range(len(SPSS_column_Q9a_selected)):
    df_spss.insert(i+1, SPSS_column_Q9a_selected[i],0)

column_selected = []
nan_values = df.isna()

for i in column:
    for x in range(len(df)): 
        if nan_values.loc[x,i]:
            pass
        else:
            # column_selected.append(i)
            for j in Codelist:
                if int(j[0]) == int(df[i][x]):
                    df_spss.loc[x, j[1]] = 1
                    df_spss.loc[x, str(j[1] + j[2])] = 1
                    df_spss.loc[x, str(j[1] + j[2] + j[3])] = 1

# for i in column_selected:
#     for j in Codelist:
#         for x in range(len(df)):
#             if nan_values.loc[x,i]:
#                 pass
#             else:            
#                 if int(j[0]) == int(df[i][x]):
#                     df_spss.loc[x, j[1]] = 1
#                     df_spss.loc[x, str(j[1] + j[2])] = 1
#                     df_spss.loc[x, str(j[1] + j[2] + j[3])] = 1

report = pd.ExcelWriter(r'C:\Users\toan.ngo\Documents\Git Projects\vincenzo_text_analytics_spss\abc.xlsx', engine = 'xlsxwriter')
df_spss.to_excel(report, sheet_name="Results")
report.close()
