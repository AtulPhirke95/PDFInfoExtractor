import os
from os import path
import glob
import shutil
from datetime import datetime
import xlrd
import openpyxl
import pandas as pd
import xlrd
import numpy as np
import re
def return_formating_excel(project_path,excel_file,timestamp):
    excelfilenameDF = pd.ExcelFile(excel_file)
    main_excel_file_DF = excelfilenameDF.parse("Data")

    requiredDF = main_excel_file_DF[main_excel_file_DF.columns[0:12]]

    requiredDF = requiredDF[requiredDF['id2'].notnull()]

    requiredDF = requiredDF[requiredDF.columns[6:12]]

    #requiredDF = requiredDF.font.str.replace(to_replace="\w+\+",value=r"")

    dict_all = {}
    list_text = []
    list_data_of_text = []
    dict_text = {}
    tempString = ""

    font_text = ""
    #bbox5_text = ""
    colourSpace_text = ""
    size_text = ""

    for row in requiredDF.itertuples(index=True,name='Pandas'):
        if str(getattr(row,"text")) != 'nan':
            tempString = tempString + getattr(row,"text")
            font_text = getattr(row,"font")
            font_text=re.sub("\w+\+","",font_text)
            #bbox5_text = getattr(row,"bbox5")
            #colourSpace_text = getattr(row,"colourspace")
            #ncolour_text = getattr(row,"ncolour")
            size_text = getattr(row,"size")
        else:
            dict_text["text"] = tempString
            dict_text["font_info"] = [{"font":font_text,"size":size_text}]
            list_text.append(dict_text)
            dict_text = {}
            tempString = ""
            font_text = ""
            #bbox5_text = ""
            #colourSpace_text = ""
            #ncolour_text = ""
            size_text = ""
    #print(list_text)

    rows = []
    for data in list_text:
        data_row = data['font_info']
        time = data['text']
        for row in data_row:
            row['text'] = time
            rows.append(row)
    df=pd.DataFrame(rows)

    cols = df.columns.tolist()
    cols = cols[-1:] + cols[:-1]
    df = df[cols]

    #df = pd.dataFrame.form_dict(list_text)
    #dict = dict.transpose()

    #df = df.fillna(np.nap)
    #df = df.dropna(how='all').apply(lambda x: pd.Series(x.dropna().values,1).fillna(''))

    writer = pd.ExcelWriter(project_path+"\\excel_files\\"+timestamp+".xlsx")
    df.to_excel(writer,sheet_name = "Sheet1")
    writer.close()

    return timestamp + ".xlsx"
    
            
            
