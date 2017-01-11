#!/usr/bin/python3
# -*- coding: utf-8 -*-

# This module is designed as an opening xls|xlsx files, comparing the old and new version of the file (different days) and inserted into a new file all that more important and not solved

import os
import win32com.client as win32
import datetime

def extract_data_from_excel(filename=None):
  '''Extract data from .xls or .xlsx file and writing this data to filename + ".txt" file.'''
  xl = win32.gencache.EnsureDispatch('Excel.Application') #for compile can be used: xl = win32.Dispatch('Excel.Application')
  if filename:
    ss = xl.Workbooks.Open(filename)
  else:
    ss = xl.Workbooks.Add() # Adds a new workbook.
  
  sheets_data = []
  for sheet in range(3):
    if sheet == 0:
      sh = ss.Worksheets('PON')
    if sheet == 1:
      sh = ss.Worksheets('АТШ')
    if sheet == 2:
      sh = ss.Worksheets('ADSL')
      # sh = ss.ActiveSheet # sheet2 = ss.Sheets("Sheet2")

    xl.Visible = False

    extracted_data = []
    
    number_of_rows = sh.UsedRange.Rows.Count # define the last of row in xls file
    for i in range(1, number_of_rows+1): # from 1 to (number_of_rows + 1) rows
      for j in range(1,5): # from 1 to 4 column
        str_data = sh.Cells(i,j).Value # Extracts information
        str_data = check_inverted_commas(str_data)
        try:
          extracted_data.append(str(str_data))
        except UnicodeEncodeError:
          extracted_data.append(str_data.encode('utf-8'))
        try:
          print(str_data) 
        except UnicodeEncodeError:
          print(str_data.encode('utf-8'))
    sheets_data.append(extracted_data)
  
  
  # A handler for Stats worksheet. Need to smoke manuals!
  # sh = ss.Worksheets('Stats') 
  # extracted_data = []
  stats = []
    
  # number_of_rows = sh.UsedRange.Rows.Count # define the last of row in xls file
  # for i in range(1, number_of_rows+1): # from 1 to (number_of_rows + 1) rows
    # for j in range(1,7): # from 1 to 6 column
      # if i == 0:
        # str_data = sh.Cells(i,j).Value # Extracts information
        # try:
          # extracted_data.append(str(str_data))
        # except UnicodeEncodeError:
          # extracted_data.append(str_data.encode('utf-8'))
        # try:
          # print(str_data) 
        # except UnicodeEncodeError:
          # print(str_data.encode('utf-8'))
      # elif j % 2 == 1:
        # str_data = sh.Cells(i,j).Formula # Extracts information
        # try:
          # extracted_data.append(str(str_data))
        # except UnicodeEncodeError:
          # extracted_data.append(str_data.encode('utf-8'))
        # try:
          # print(str_data) 
        # except UnicodeEncodeError:
          # print(str_data.encode('utf-8'))
      # else:
        # str_data = sh.Cells(i,j).Value # Extracts information
        # try:
          # extracted_data.append(str(str_data))
        # except UnicodeEncodeError:
          # extracted_data.append(str_data.encode('utf-8'))
        # try:
          # print(str_data) 
        # except UnicodeEncodeError:
          # print(str_data.encode('utf-8'))
    
  # stats.append(extracted_data)

  
  ss.Close(False)
  xl.Application.Quit()
  return [sheets_data, stats]

def check_inverted_commas(data):
  ''' Method for checking case then in excel cell was used inverted commas mark and for changing it to neutral, single quotes mark.'''
  try:
    new_data_list = list(data)
    for i in ['<', '>', '«', '»', '–']:
      for j in range(len(data)):
        if i == data[j]:
          if i == '–':
            new_data_list[j]="-"
          else:
            new_data_list[j]="'"
    new_data = ''.join(new_data_list)
  except:
    new_data = data
    
  return new_data
  
def transponse(data):
  # transponse data from single-order list to double-order list
  width = 4
  count = 0
  new_data_line = []
  new_data = []
  for i in data:
    if count < width:
      if i.startswith('\n'):
        new_data_line.append(i[1:])
      else:
        new_data_line.append(i)
      count+=1
    else:
      new_data.append(new_data_line)
      new_data_line = []
      if i.startswith('\n'):
        new_data_line.append(i[1:])
      else:
        new_data_line.append(i)
      count=1
    
  new_data.append(new_data_line)
  return new_data

  
def data_comparison(old_data, current_data):
  '''A method for comparing data from old and current data (earlier was extracted from Excel files). Returns parsed_data_lines list for new, result Excel file.'''
  
  sheets_data = []
  for sheet in range(len(old_data)):
    old_data_2_dimensions = transponse(old_data[sheet])
    current_data_2_dimensions = transponse(current_data[sheet])
    new_ies_etc = current_data_2_dimensions
        
    for i in range(len(current_data_2_dimensions)):
      for j in range(len(old_data_2_dimensions)):
        if current_data_2_dimensions[i][0] == old_data_2_dimensions[j][0]:
          print(current_data_2_dimensions[i][0])
          
          # input("fdsfsd")
          new_ies_etc[i][3] = old_data_2_dimensions[j][3]
    
    sheets_data.append(new_ies_etc)   
  parsed_data_lines = sheets_data
  return parsed_data_lines

def save_xlsx_file(parsed_data_lines, filename, stats_data):
  '''Method for saving parsed_data_lines structurwe to Excel file with filename.'''
  
  if filename.endswith('.xls'):
    filename = filename[:-4]
  elif filename.endswith('.xlsx'):
    filename = filename[:-5]
  new_filename = filename + '_result.xlsx'
  xl = win32.gencache.EnsureDispatch('Excel.Application') #for compile can be used: xl = win32.Dispatch('Excel.Application')
  ss = xl.Workbooks.Add() # Adds a new workbook.
  # sh = ss.ActiveSheet # sheet2 = ss.Sheets("Sheet2")
  
  sh = ss.Worksheets.Add()
  sh.Name = "Stats"
  # for i in range(len(stats_data)):
    # for j in range(len(stats_data[i])):
      # if stats_data[i][j] == "None":
        # sh.Cells(i+1,j+1).Value = ''
      # else:
        # if i == 0:
          # sh.Cells(i+1,j+1).Value = stats_data[i][j]
        # elif j % 2 == 1:
          # sh.Cells(i+1,j+1).Formula = stats_data[i][j]
        # else:
          # sh.Cells(i+1,j+1).Value = stats_data[i][j]

  xl.Visible = False
   
  ### defining current day and month
  current_day = datetime.date.today().day
  months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
  current_month = months[datetime.date.today().month-1]
  day_and_month = str(current_day) + ' ' + str(current_month)
  ###
  
  for sheet in range(-len(parsed_data_lines)+1,1):
    print(len(parsed_data_lines[sheet]))
    if sheet == 0:
      sh = ss.Worksheets.Add()
      sh.Name = "PON"
    elif abs(sheet) == 1:
      sh = ss.Worksheets.Add()
      sh.Name = "ADSL"
    elif abs(sheet) == 2:
      sh = ss.Worksheets.Add()
      sh.Name = "АТШ"
    
    for i in range(len(parsed_data_lines[sheet])):
      for j in range(len(parsed_data_lines[sheet][i])):
        if parsed_data_lines[sheet][i][j] == "None":
          sh.Cells(i+1,j+1).BorderAround()
        else:
          if parsed_data_lines[sheet][i][1] == "None" and parsed_data_lines[sheet][i][2] == "None" and parsed_data_lines[sheet][i][3] == "None":
            sh.Range("A"+str(i+1) + ":" + "D" + str(i+1)).Select()
            xl.Selection.Merge()
            xl.Selection.HorizontalAlignment = win32.constants.xlCenter # centration
            xl.Selection.Font.Bold = True
            xl.Selection.BorderAround()
          try:
            if j == 2: # case for column "Время события"
              if parsed_data_lines[sheet][i][j].startswith('20'): # case for cells with date and time of EIs
                print(i,j, parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).Value = str(parsed_data_lines[sheet][i][j][:-9]) # [-9] - It's for removing last 9 signs of time (seconds and ms)
                sh.Cells(i+1,j+1).BorderAround()
              else: # case for header "Время события"
                print(i,j, parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).Value = str(parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).BorderAround()
            elif j == 3:
              if i == 0: # header case
                print(i,j, parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).Value = str(parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).BorderAround()
              else:
                print(i,j, parsed_data_lines[sheet][i][j])
                sh.Cells(i+1,j+1).Value = str(parsed_data_lines[sheet][i][j] + "\nна {0}: ".format(day_and_month)) # '\n' - New line for current day's comment
                sh.Cells(i+1,j+1).BorderAround()
            else:
              print(i,j, parsed_data_lines[sheet][i][j])
              sh.Cells(i+1,j+1).Value = str(parsed_data_lines[sheet][i][j])
              sh.Cells(i+1,j+1).BorderAround()
          except:
            pass
    
    if sh.Cells(1,1).Value.startswith("ЕИ") == False:
      sh.Range("A1").Select()
      xl.Selection.Insert()
      sh.Cells(1,1).Value = "ЕИ" 
      sh.Cells(1,1).BorderAround()
      sh.Cells(1,2).Value = "Группа исполнителя" 
      sh.Cells(1,2).BorderAround()
      sh.Cells(1,3).Value = "Время события" 
      sh.Cells(1,3).BorderAround()
      sh.Cells(1,4).Value = "Комментарий" 
      sh.Cells(1,4).BorderAround()
    else:
      input("dsasa")
      pass
    sh.Rows(1).Font.Bold =  True
    sh.Rows(1).HorizontalAlignment = win32.constants.xlCenter
    # sh.Columns(1).AutoFit()
    sh.Columns(1).ColumnWidth = 10
    sh.Columns(2).ColumnWidth = 30
    sh.Columns(3).ColumnWidth = 15
    sh.Columns(4).ColumnWidth = 160
    
    sh = ss.Worksheets.Item(5).Delete() # delete last worksheet with name "Лист1".
  sh = ss.ActiveSheet # sheet2 = ss.Sheets("Sheet2")
  sh.SaveAs(new_filename)

  ss.Close(False)
  xl.Application.Quit()

  
def define_old_file():
  curr_month = datetime.date.today().month
  months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
  minus1day = datetime.timedelta(-1)
  if datetime.date.today().day == 1:
    prev_date = months[curr_month-2] + str((datetime.date.today() + minus1day).day)
  else:
    prev_date = months[curr_month-1] + str((datetime.date.today() + minus1day).day)
  oldfile_name = '\TTmore48hrs-' + prev_date + '.xlsx'
  x = ''
  if oldfile_name[1:] in os.listdir(os.path.join('.', '_excel_files')):
    return oldfile_name
  else:
    print("We wait a file: {0}".format(oldfile_name[1:]))
    print("Copy this file to '_excel_files' folder and restart this app again!")
    return False    
      
    
def execution_logic(path_to_old_file, path_to_current_file, path_to_result_file):
  '''This method describes the start sequence of methods.'''
  
  old_dataset = extract_data_from_excel(path_to_old_file)
  stats_data = old_dataset[1]
  data = data_comparison(old_dataset[0], extract_data_from_excel(path_to_current_file)[0])
  save_xlsx_file(data, path_to_result_file, stats_data)



def main():
  if define_old_file():
    path_to_old_file = os.path.abspath(os.path.join('.', '_excel_files')) + define_old_file()
    # print(path_to_old_file)
    path_to_current_file = os.path.abspath(os.path.join('.', '_excel_files')) + '\\pp_result.xlsx'
    print(path_to_current_file)
    path_to_result_file = os.path.abspath(os.path.join('.', '_excel_files')) + '\\checker_file.xlsx'

    # print(extract_data_from_excel(path_to_old_file))
    execution_logic(path_to_old_file, path_to_current_file, path_to_result_file)
  else: 
    pass
  
  
if __name__ == '__main__':
    main()
    

# TODO: 
# 1) Need to add more descriptions
# 2) Need to add dinamically defining of old and current xls files!
# 3) Need do work page_parser module
# 4) Make settings to separate file or avoid need for do settings!
# 5) Need to create Stats worksheet content

# Instruction for use:
# 1) need to close MS Excel app (End task Excel.exe in Task manager)
# 2) define path to old (last) excel file (of previous day) and for current day's file (comprising current EIs).