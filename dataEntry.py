import pandas as pd
import os
from IPython.core.display import HTML
import re
from pandas import DataFrame
import numpy as np
from numpy import nan as NA
from IPython.core.display import HTML
from IPython.core.display import Image 
import matplotlib.pylab as pylab
from itertools import chain
import datetime, itertools, os, collections, re
import math
import time
from time import mktime
from datetime import datetime, timedelta
import ExcelExtraction
from time import mktime
import xlsxwriter


dma_name_list = [ 'Marwan', 'Mithu', 'Raju', 'Farhan', 'Fazle' ]

def new_file():
    dfs_dic = dict(enumerate(all_dfs))
    req_dfs = {}
    for i,df in dfs_dic.items():
        if df.sheet.unique()[0] in [ 'received_sheet']:
            req_dfs[len(req_dfs)] = df.copy()

    dfs_concat = pd.concat(req_dfs, ignore_index=True)
    dfs_concat = dfs_concat.loc[dfs_concat.format.isin(['handwritten copy','scanned copy'])].copy()
    dfs_concat.reset_index(inplace=True)
    dfs_concat.drop('index', axis=1, inplace=True)

    dfs_concat['factory_code'] = pd.Series([ elem.split('.')[0].split('_')[-1] for elem in dfs_concat.file] ,index=dfs_concat.index)

    new_dfs = pd.DataFrame(columns={'sl_no','factory_code','report_name','project_phase','start_time','end_time','entry_type'})
    for ind in dfs_concat.index:
        s_ = dfs_concat.loc[ind][[ 'factory_code','report_name','project_phase','start_time','end_time']]
        for i in range(3):
            s_['sl_no'] = len(new_dfs)+1
            if i==0:
                s_['entry_type'] = "1st entry"
            if i==1:
                s_['entry_type'] = "2nd entry"
            if i==2:
                s_['entry_type'] = "recon"
            
            new_dfs.loc[len(new_dfs)] = s_    
    for col in ['status','given_date','given_time','dma_name','workings','received_date','received_time','total_working_mins','dma_code']:
        new_dfs[col] = np.nan

    new_dfs = new_dfs[[   'sl_no','factory_code','report_name', 'project_phase','start_time',  'end_time',
                    'status', 'given_date','given_time',  
                   'dma_name','entry_type','workings', 
                   'received_date', 'received_time',
                    'total_working_mins', 'dma_code']]

    #print(new_dfs)

    #print(new_dfs)
    file_name ='../data entry spread sheet_new1'+'.xlsx'

    workbook   = xlsxwriter.Workbook(file_name)

    worksheet = workbook.add_worksheet('spread_sheet')

    for col in new_dfs.columns:
        new_dfs[col] = new_dfs[col].replace(np.nan,'.')

    new_dfs['status'] = new_dfs['status'].replace('.','not entered')

    ###################################################################################

    green_format = workbook.add_format()
    yellow_format = workbook.add_format()
    red_format = workbook.add_format()

    green_format.set_pattern(1)  # This is optional when using a solid fill.
    yellow_format.set_pattern(1) 
    red_format.set_pattern(1) 

    green_format.set_bg_color('green')
    yellow_format.set_bg_color('yellow')
    red_format.set_bg_color('red')


    font_size_format = workbook.add_format()
    font_size_format.set_font_size(15)


    # Write some data headers.
    for i,elem in enumerate(list(new_dfs.columns)):
        worksheet.write(0,i,elem,font_size_format)


    for i in range(len(new_dfs)):
        for j in range(len(new_dfs.columns)):
            worksheet.write(i+1,j, new_dfs.loc[i][j])

    #worksheet.autofilter(0, 0, len(concat_all), len(concat_all.columns)-1)  # Same as above.

    for i in range(2, len(new_dfs)+2):
        worksheet.data_validation('G'+str(i), {'validate': 'list','source': ['completed', 'ongoing', 'not entered']})

    for i in range(2, len(new_dfs)+2):
        worksheet.data_validation('J'+str(i), {'validate': 'list','source': dma_name_list})


    worksheet.conditional_format('G2:G'+str(len(new_dfs)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'completed',
                                       'format':   green_format})

    worksheet.conditional_format('G2:G'+str(len(new_dfs)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'ongoing',
                                       'format':   yellow_format})

    worksheet.conditional_format('G2:G'+str(len(new_dfs)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'not entered',
                                       'format':   red_format})

 



    workbook.close()


    ###################################################################
    #frames = {'spread_sheet': new_dfs}
    #file_name ='../data entry spread sheet_new1'+'.xlsx'
    #writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
 
    #for sheet, frame in  frames.items(): 
    #    frame.to_excel(writer, sheet_name = sheet)
    #writer.save()   

def update_file():
    old_dfs = old_version[0].copy()

    old_dfs.drop(['path','file','sheet'], axis=1, inplace=True)

    dfs_dic = dict(enumerate(all_dfs))
    req_dfs = {}
    for i,df in dfs_dic.items():
        if df.sheet.unique()[0] in [ 'received_sheet']:
            req_dfs[len(req_dfs)] = df.copy()

    dfs_concat = pd.concat(req_dfs, ignore_index=True)
    dfs_concat = dfs_concat.loc[dfs_concat.format.isin(['handwritten copy','scanned copy'])].copy()
    dfs_concat.reset_index(inplace=True)
    dfs_concat.drop('index', axis=1, inplace=True)

    dfs_concat['factory_code'] = pd.Series([ elem.split('.')[0].split('_')[-1] for elem in dfs_concat.file] ,index=dfs_concat.index)

    dfs_concat['id'] = pd.Series([str(e1)+str(e2)+str(e3) for e1,e2,e3 in zip(dfs_concat.factory_code, dfs_concat.report_name, dfs_concat.project_phase)] ,index=dfs_concat.index)

    old_dfs['id'] = pd.Series([str(e1)+str(e2)+str(e3) for e1,e2,e3 in zip(old_dfs.factory_code, old_dfs.report_name, old_dfs.project_phase)] ,index=old_dfs.index)

    drop_ind = list(dfs_concat.loc[dfs_concat.id.isin(old_dfs.id.unique())].index)

    dfs_concat.drop(drop_ind, inplace=True)
    dfs_concat.reset_index(inplace=True)
    dfs_concat.drop('index', axis=1, inplace=True)

    dfs_concat.drop('id', axis=1, inplace=True)
    old_dfs.drop('id', axis=1, inplace=True)

    old_dfs.given_date = pd.Series([ elem.date().strftime("%d-%b-%Y") if elem!='.' else elem for elem in old_dfs.given_date] ,index=old_dfs.index)

    old_dfs.given_time = old_dfs.given_time.astype(str)

    new_dfs = pd.DataFrame(columns={'sl_no','factory_code','report_name','project_phase','start_time','end_time','entry_type'})
    for ind in dfs_concat.index:
        s_ = dfs_concat.loc[ind][[ 'factory_code','report_name','project_phase','start_time','end_time']]
        for i in range(3):
            s_['sl_no'] = len(new_dfs)+1
            if i==0:
                s_['entry_type'] = "1st entry"
            if i==1:
                s_['entry_type'] = "2nd entry"
            if i==2:
                s_['entry_type'] = "recon"
            
            new_dfs.loc[len(new_dfs)] = s_    
    for col in ['status','given_date','given_time','dma_name','workings','received_date','received_time','total_working_mins','dma_code']:
        new_dfs[col] = np.nan

    new_dfs = new_dfs[[   'sl_no','factory_code','report_name', 'project_phase','start_time',  'end_time',
                    'status', 'given_date','given_time',  
                   'dma_name','entry_type','workings', 
                   'received_date', 'received_time',
                    'total_working_mins', 'dma_code']]

    frames_ = [ old_dfs, new_dfs]
    all_dic = dict(enumerate(frames_))
    concat_all = pd.concat(all_dic, ignore_index=True)

    concat_all = concat_all [[ 'sl_no','factory_code','report_name', 'project_phase','start_time',  'end_time',
                    'status', 'given_date','given_time',  
                   'dma_name','entry_type','workings', 
                   'received_date', 'received_time',
                    'total_working_mins', 'dma_code']]
    concat_all.sl_no = pd.Series([ elem+1 for elem in concat_all.index] ,index=concat_all.index)

    concat_all.factory_code = concat_all.factory_code.astype(str)

    #worksheet.data_validation(start_row, start_column, end_row, end_column, {'validate': 'list', 'source': options }

    #print(new_dfs)
    file_name ='../data entry spread sheet_new1'+'.xlsx'

    workbook   = xlsxwriter.Workbook(file_name)

    worksheet = workbook.add_worksheet('spread_sheet')

    for col in concat_all.columns:
        concat_all[col] = concat_all[col].replace(np.nan,'.')

    concat_all['status'] = concat_all['status'].replace('.','not entered')

    ##########################################################################

    green_format = workbook.add_format()
    yellow_format = workbook.add_format()
    red_format = workbook.add_format()

    green_format.set_pattern(1)  # This is optional when using a solid fill.
    yellow_format.set_pattern(1) 
    red_format.set_pattern(1) 

    green_format.set_bg_color('green')
    yellow_format.set_bg_color('yellow')
    red_format.set_bg_color('red')

    font_size_format = workbook.add_format()
    font_size_format.set_font_size(15)

    # Write some data headers.
    for i,elem in enumerate(list(concat_all.columns)):
        worksheet.write(0,i,elem,font_size_format)


    for i in range(len(concat_all)):
        for j in range(len(concat_all.columns)):
            worksheet.write(i+1,j, concat_all.loc[i][j])

    #worksheet.autofilter(0, 0, len(concat_all), len(concat_all.columns)-1)  # Same as above.

    for i in range(2, len(concat_all)+2):
        worksheet.data_validation('G'+str(i), {'validate': 'list','source': ['completed', 'ongoing', 'not entered']})

    for i in range(2, len(concat_all)+2):
        worksheet.data_validation('J'+str(i), {'validate': 'list','source': dma_name_list})


    worksheet.conditional_format('G2:G'+str(len(concat_all)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'completed',
                                       'format':   green_format})

    worksheet.conditional_format('G2:G'+str(len(concat_all)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'ongoing',
                                       'format':   yellow_format})

    worksheet.conditional_format('G2:G'+str(len(concat_all)+2), {'type':     'text',
                                       'criteria': 'containing',
                                       'value':    'not entered',
                                       'format':   red_format})




    workbook.close()
###################################################################
    #frames = {'spread_sheet': concat_all}
    #file_name ='../data entry spread sheet_new1'+'.xlsx'
    #writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
 
    #for sheet, frame in  frames.items(): 
    #    frame.to_excel(writer, sheet_name = sheet)

    #writer.filter_column_list('status', ['complete', 'inprogress', 'not complete'])
    #writer.save() 




all_dfs = ExcelExtraction.extract_all_files(r"../source")

old_version = ExcelExtraction.extract_all_files(r"../old_version")

if len(old_version)==0:
    new_file()
else:
    update_file()   

