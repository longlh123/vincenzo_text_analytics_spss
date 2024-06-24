import win32com.client as w32
import numpy as np
import pandas as pd
import sys
sys.path.append(r'C:\Users\toan.ngo\Documents\Git Projects\vincenzo_text_analytics_spss\libs')
from metadata import SPSSObject
from metadata import Metadata
from metadata import SPSSObject_Dataframe

mdd_path = r"Metadata\VINCENZO_RGB_W12_M6_TCB_BM_OE_Q9a_FULL_20062024_v3_SPSSv1.mdd"
ddf_path = r"Metadata\VINCENZO_RGB_W12_M6_TCB_BM_OE_Q9a_FULL_20062024_v3_SPSSv1.ddf"
# sql_query = "SELECT * FROM VDATA WHERE _Wave = {_12M04} OR _Wave = {_12M05} OR _Wave = {_12M06}"
# sql_query = "SELECT * FROM VDATA WHERE Respondent.ID = 16748991"
sql_query = "SELECT * FROM VDATA"
questions_CE = ["Respondent.ID","_Sttcusid", "_Q1","_Q2","_Q2a","_Q4","_Q5","_Q4Bis","_Phase","_Q9bis","_Q20a_Q6bis","_Q25","_Q23a_Q7bis"]
questions = ["_Sttcusid", "_Q9a_Codes"] 

spssObject = SPSSObject(mdd_path, ddf_path, sql_query, questions)
spssObject_CE = SPSSObject(mdd_path, ddf_path, sql_query, questions_CE)
df = pd.DataFrame(data=spssObject.records, columns=spssObject.varNames)
df_CE = pd.DataFrame(data=spssObject_CE.records, columns=spssObject_CE.varNames)

# spssObject = Metadata(mdd_file = mdd_path, ddf_file = ddf_path, sql_query = sql_query)

# df = spssObject.convertToDataFrame(questions_CE)

column = df.columns.drop("Sttcusid")

df_excel = pd.read_excel("source\\2024 - 066 - VINCENZO - Format data SPSS_RBG Q224.xlsx", engine="openpyxl",sheet_name="CODEFRAME-NPS")
df_excel.set_index(["CODE","LV1","ATTITUDE","LV2"], inplace=True)
SPSS_column_Q9a = pd.read_excel("source\\2024 - 066 - VINCENZO - Format data SPSS_RBG Q224.xlsx", engine="openpyxl",sheet_name="Variable View_OE")
SPSS_column_Q9a_selected = list(SPSS_column_Q9a["Column"])
Codelist = [];
for i in list(df_excel.index):
    Codelist.append(i)

df_spss = pd.DataFrame()
df_spss.insert(0, "RespondentID", df_CE["Respondent_ID"])
df_spss.insert(1, "Sttcusid", df_CE["Sttcusid"])
df_spss.insert(2, "Q1", df_CE["Q1"])
df_spss.insert(3, "Q2", df_CE["Q2"])
df_spss.insert(4, "Q2a", df_CE["Q2a"])
df_spss.insert(5, "Q4", df_CE["Q4"])
df_spss.insert(6, "Q5", df_CE["Q5"])
df_spss.insert(7, "Q4bis", df_CE["Q4Bis"])
df_spss.insert(8, "Q9", df_CE["_Q9_1"])
df_spss.insert(9, "Q9bis_R1", df_CE["_Q9bis_Codes_1"])
df_spss.insert(10, "Q9bis_R2", df_CE["_Q9bis_Codes_2"])
df_spss.insert(11, "Q9bis_R3", df_CE["_Q9bis_Codes_3"])
df_spss.insert(12, "Q9bis_R4", df_CE["_Q9bis_Codes_4"])
df_spss.insert(13, "Q9bis_R5", df_CE["_Q9bis_Codes_5"])
df_spss.insert(14, "Q20a_P1", df_CE["_Q20a_Q6bis_Codes_1"])
df_spss.insert(15, "Q20a_P2", df_CE["_Q20a_Q6bis_Codes_2"])
df_spss.insert(16, "Q20a_P3", df_CE["_Q20a_Q6bis_Codes_3"])
df_spss.insert(17, "Q20a_P4", df_CE["_Q20a_Q6bis_Codes_4"])
df_spss.insert(18, "Q20a_P5", df_CE["_Q20a_Q6bis_Codes_5"])
df_spss.insert(19, "Q20a_P6", df_CE["_Q20a_Q6bis_Codes_6"])
df_spss.insert(20, "Q20a_P7", df_CE["_Q20a_Q6bis_Codes_7"])
df_spss.insert(21, "Q20a_P8", df_CE["_Q20a_Q6bis_Codes_8"])
df_spss.insert(22, "Q25", df_CE["Q25"])
df_spss.insert(23, "Q23a_C1", df_CE["_Q23a_Q7bis_Codes_1"])
df_spss.insert(24, "Q23a_C2", df_CE["_Q23a_Q7bis_Codes_2"])
df_spss.insert(25, "Q23a_C3", df_CE["_Q23a_Q7bis_Codes_3"])
df_spss.insert(26, "Q23a_C4", df_CE["_Q23a_Q7bis_Codes_4"])
df_spss.insert(27, "Q23a_C5", df_CE["_Q23a_Q7bis_Codes_5"])
df_spss.insert(28, "Q23a_C6", df_CE["_Q23a_Q7bis_Codes_6"])
df_spss.insert(29, "Q23a_C7", df_CE["_Q23a_Q7bis_Codes_7"])
df_spss.insert(30, "Q23a_C8", df_CE["_Q23a_Q7bis_Codes_8"])
df_spss.insert(31, "Q9a_text", df_CE["_Q9a_Text_1"])

for i in range(len(SPSS_column_Q9a_selected)):
    df_spss.insert(i+32, SPSS_column_Q9a_selected[i],0)

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


report = pd.ExcelWriter(r'C:\Users\toan.ngo\Documents\Git Projects\vincenzo_text_analytics_spss\abc.xlsx', engine = 'xlsxwriter')
df_spss.to_excel(report, sheet_name="Results")
report.close()
