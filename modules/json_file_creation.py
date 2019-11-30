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

def json_save(project_path,temp_list1,timestamp):
    with open(project_path+"\\json_files\\"+timestamp+'.json','w') as fout:
        json.dump(temp_list1,fout)
    return project_path+"\\json_files\\"+timestamp+'.json'

def return_json(project_path,words,dict_values,flag_for_text_search,timestamp):
    temp_list = []
    temp_list1 = []
    count = 1
    dict_holding_temp_values = {}

    if flag_for_text_search == False:
        temp_list1.append(dict_values)
        json_path = json_save(project_path,temp_list1,timestamp)
        if os.path.exists(json_path):
            return json_path

    elif flag_for_text_search == True:
        for i in range(len(dict_values)-1):
            if list(dict_values)[i]+1 == list(dict_values)[i+1]:
                if list(dict_values)[i] not in dict_holding_temp_values:
                    dict_holding_temp_values[list(dict_values)[i]] = list(dict_values.values())[i]
                if list(dict_values)[i+1] not in dict_holding_temp_values:
                    dict_holding_temp_values[list(dict_values)[i+1]] = list(dict_values.values())[i+1]
            else:
                if len(dict_holding_temp_values)== len(words):
                    temp_list1.append(dict_holding_temp_values)
                    count += 1
                dict_holding_temp_values={}
        count = 0
        flag = False

        for _ in temp_list1:
            for index in range(len(_)-1):
                if list(_.values())[index][0] == words[count]:
                    flag=True
                else:
                    flag = False
                    break
                count += 1
            count = 0
        if flag == True:
            json_path = json_save(project_path,temp_list1,timestamp)

            if os.path.exists(json_path):
                return json_path
        
    
