import os
from os import path
import glob
import shutil
from datetime import datetime
import xlrd
import openpyxl
import pandas as pd
import pdf_data_analyzer
import re
import argparse
import json
import json_file_creation

ap = argparse.ArgumentParser()
ap.add_argument("-proj","--path",required=True,
                help="path of the project till \pdf")
ap.add_argument("-p","--page",required=True,
                help="page number of pdf")
ap.add_argument("-pdf","--pdf",required=True,
                help="pdf name")
ap.add_argument("-s","--search",required=True,
                help="text to search for")
args = vars(ap.parse_args())

#current date and time
now = datetime.now()
timestamp = datetime.timestamp(now)
timestamp = str(timestamp).replace('.','_')
print("timestamp = ", timestamp)

project_path = args["path"]
page_no = args["page"]
xml_file_name = "xml_file_"+timestamp
pdf_file_name = args["pdf"]
pdf_file_path = project_path + "\\pdf_files\\"

if os.path.exists(pdf_file_path + pdf_file_name):
    #changing directory to project location
    os.chdir(project_path+"\\modules")
    #running pdf2txt to genarate xml file
    os.system('python pdf2txt.py -p {} -o {}.xml {}{}'.format(page_no,xml_file_name,pdf_file_path,pdf_file_name))

    #checking if xml file is generated or not
    if os.path.exists(project_path+"\\modules\\"+xml_file_name+".xml"):
        flag=False
        file = open(project_path+"\\modules\\"+xml_file_name+".xml","r")
        if file.readlines()[-1] != "</pages>":
            flag=True
        file.close()

        if flag==True:
            file = open(project_path+"\\modules\\"+xml_file_name+".xml","a+")
            file.write("</pages>")
            file.close()
        
        #moving xml file to generated_xml_files folder
        shutil.move(project_path+"\\modules\\"+xml_file_name+".xml",project_path+"\\generated_xml_files")
        #calling vbs file
        os.system('xml_Initializer.vbs {} {}.xml {}'.format(project_path+"\\modules\\xml_vba.xlsm",project_path+"\\generated_xml_files\\"+xml_file_name,project_path+"\\template\\xml2xlsx_"+timestamp+".xlsx"))
        #calling return _formating_excel which returns excel sheet containing formating of text
        formating_data_file = pdf_data_analyzer.return_formating_excel(project_path,project_path+"\\template\\xml2xlsx_"+timestamp+".xlsx",timestamp)

        #checking if formating text excel sheet is generated or not?
        if os.path.exists(project_path+"\\excel_files\\"+formating_data_file):
            xfile = openpyxl.load_workbook(project_path+"\\excel_files\\"+formating_data_file)
            sheet = xfile.get_sheet_by_name('Sheet1')
            sheet['A1'] = 'index'
            xfile.save(project_path+"\\excel_files\\"+formating_data_file)

            words = args["search"].split(" ")
            #print(words)
            excelfilenameDF = pd.ExcelFile(project_path+"\\excel_files\\"+formating_data_file)
            main_excel_file_DF = excelfilenameDF.parse("Sheet1")
            fontOf = ""
            size = ""
            list_no = []
            list_text = []
            str_text = ""
            dict_tuples = {}
            dict_values = {}

            for row in main_excel_file_DF.itertuples(index=True,name='Pandas'):
                for _ in words:
                    if str(getattr(row,"text")) == _:
                        fontOf = getattr(row,"font")
                        size = getattr(row,"size")

                        list_no.append(int(getattr(row,"index")))
                        dict_tuples[int(getattr(row,"index"))] = _
                        dict_values[int(getattr(row,"index"))] = [_,fontOf,size]
                        str_text = str_text + " " + str(getattr(row,"text"))
            print(dict_values)

            if args["search"] in str_text:
                print("Yes")
                temp_list = []
                temp_list1 = []
                count = 0
                dict_holding_temp_values = {}
                flag_for_text_search = True
                if(len(dict_values)) == 1:
                    flag_for_text_search = False
                    json_path = json_file_creation.return_json(project_path,words,dict_values,flag_for_text_search,timestamp)

                elif len(dict_values) > 1:
                    flag_for_text_search = True
                    json_path = json_file_creation.return_json(project_path,words,dict_values,flag_for_text_search,timestamp)
        else:
            print("Excel file is not genearated")
    else:
        print("Xml file is not generated")
else:
    print("Pdf file not exists")
        
        
        
