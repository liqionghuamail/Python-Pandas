import pandas as pd
from openpyxl import load_workbook

Filepath_sn024="I:\\Theatres\\RNS Theatres Data Mgr\\KPI-MOH\\2022\\2022-02\\DATA_KPI_FEBRUARY_2022.xlsx"
FilePath_HAI='I:\\Theatres\\RNS Theatres Data Mgr\\DATA Corrections\\2022\\2022-03\\2022-03-31\\HAI.xlsx'


df=pd.read_excel(Filepath_sn024, sheet_name='SN024 - Casemix', header=0)
df=df.dropna(how='all')
df_cardio=df[df['SPECIALTY']=='Cardiothoracic SN']
df_orthor=df[df['SPECIALTY']=='Orthopaedic SN']

with pd.ExcelWriter(FilePath_HAI) as writer:  
    df_cardio.to_excel(writer, sheet_name='Cardiothoracic Raw', index=False, header=1)
    df_orthor.to_excel(writer, sheet_name='Orthopaedic Raw', index=False,header=1)
    print('done')

