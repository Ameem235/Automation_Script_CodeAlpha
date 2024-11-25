import pandas as pd
import openpyxl
from openpyxl.styles import fonts
from openpyxl.styles.colors import Color



def automate_excel (filename,filepath,thershold):
    wb = openpyxl.load_workbook(filepath)
    data_read=pd.read_excel(filepath,sheet_name=None)
    filtered_data=data_read[data_read['Column Name']>thershold]
    filtered_data['new Column Name']=filtered_data['new Column Name'] *2
    with pd.ExcelWriter(filepath) as writer:
        filtered_data.to_excel(writer,index=False)
    new_Sheet=wb['New Column Name']
    for cell in new_Sheet[1]:
        cell.font=Font(bold=True)
        cell.data_type=str(filtered_data['new Column Name'][0])
        cell.value=str(filtered_data['new Column Name'][1])
    for cell in new_Sheet['A']:
        cell.font=Font(color='red')
    wb.save(filepath)
    wb.close()


automate_excel('Example',r'C:\Users\samraanahmed\Downloads.xlsx',10)



