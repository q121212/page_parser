#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Two main goals:
# 1. Extract from pages EIs and an appropriate info
# 2. Save this info to Excel file
# 3. Compare new and old files and copy comments from old file to new file, if new file have the same tickets! was implemented in checker.py

from urllib.request import urlopen
import win32com.client as win32
import os
import re
import checker
import encrypting

  
def extract_data_from_page(url, days_eis):
  rawd = urlopen(url).read()
  rawdata = ''
  
  ##### Start of magic part of parsing #####
  # This is a "magic" part of parsing. In result we have 3 output lists: eis_res, stages, dates
  if "target='_blank'>" in str(rawd.decode('utf-8', 'ignore')):
    rawdata = rawd.decode('utf-8', 'ignore')
  
  # this line of code at first stage read data from settings file, at second stage - extract (decrypt) encrypted data
  d = encrypting.decrypt(encrypting.read_settings())
  try:
    rawdata = d[1][27:-69]
    
    pattern = 'E\d{8}'
    eis = re.findall(pattern, rawdata)
    eis_res = []
    for i in range(len(eis)):
      if i %2 == 0:
        eis_res.append(eis[i])
    # eis_res - list of eis
    
    pattern = 'td>(.*)</td'
    groups = re.findall(pattern, rawdata)
    
    # times - list of creation's time of ei for ies
    times = re.findall('(\d{4}-\d{2}-\d{2} \d{2}:\d{2})',rawdata)
    dates = times
    
    x='\t'.join(rawdata.split('<td>'))
    y='\t'.join(x.split('</td>'))
    a=y.split('\t')
    b=[]
    for i in a:
      if i=='':
        pass
      else:
        b.append(i)
    # print('b={0}'.format(b))
    c=[]
    for i in range(len(b)):
      if i == 2:
        c.append(b[i])
      elif (i+5) % 7 == 0 and i !=1:
        c.append(b[i])
      elif i == len(b)-4:
        c.append(b[i])
    print('c={0}'.format(c))
    groups = re.findall('\t(.+)\t',rawdata)
    print('len_c {0}'.format(len(c)))
    
    print(len(eis_res))
    for i in range(len(eis_res)):
      print(eis_res[i])
    print(len(eis_res))
    # for i in range(len(groups)):
      # print(groups[i]+'\n')
    print('\n')
    print(times)
    print(len(times))
    stages = c
  except:
    
    eis_res = []
    stages = []
    dates = []
    
    pass  
  ##### end of magic part of parsing #####

  
  # Pasring URL
  if url.endswith('PON'):
    if len(url) == 61:
      appropriate_days_eis_number = int(url[-10])
    elif len(url) == 62:
      appropriate_days_eis_number = int(url[-11:-9])
  if url.endswith('ADSL'):
    if len(url) == 62:
      appropriate_days_eis_number = int(url[-11])
    elif len(url) == 63:
      appropriate_days_eis_number = int(url[-12:-10])
  if url.endswith('%D0%90%D0%A2%D0%A8'):
    if len(url) == 76:
      appropriate_days_eis_number = int(url[-25])
    elif len(url) == 77:
      appropriate_days_eis_number = int(url[-26:-24])
  
  try:
    return [days_eis[appropriate_days_eis_number-1], eis_res, stages, dates]
  except:
    return ["Some_days_interval", eis_res, stages, dates]


def create_urls_list(url_example):
  '''Method for creating URL list'''
  pon_urls, adsl_urls, atsh_urls = [], [], []
  for i in range(10):
    pon_urls.append(url_example[:51] + str(i+1) + url_example[52:])
    adsl_urls.append(url_example[:51] + str(i+1) + url_example[52:-3]+'ADSL')
    atsh_urls.append(url_example[:51] + str(i+1) + url_example[52:-3]+'%D0%90%D0%A2%D0%A8')
  
  urls = [pon_urls, adsl_urls, atsh_urls]
  # print(urls)
  return urls

def extract_all_data(urls, list_of_day_groups):
  '''Method for extracting all data from all pages'''
  pon_data, adsl_data, atsh_data = [], [], []
  for technology in range(len(urls)):
    for day_group in range(len(urls[technology])):
      if technology == 0:
        data = extract_data_from_page(urls[technology][day_group], list_of_day_groups)
        pon_data.append(data[0])
        pon_data.append("None")
        pon_data.append("None")
        pon_data.append("None")
        for i in range(len(data[1])):
          print("data_i {0}".format(data[1]))
          pon_data.append(data[1][i])
          pon_data.append(data[2][i])
          pon_data.append(data[3][i])
          pon_data.append('') # for cell comment
        pon_data.append("None")
        pon_data.append("None")
        pon_data.append("None")
        pon_data.append("None")
      if technology == 1:
        data = extract_data_from_page(urls[technology][day_group], list_of_day_groups)
        adsl_data.append(data[0])
        adsl_data.append("None")
        adsl_data.append("None")
        adsl_data.append("None")        
        for i in range(len(data[1])):
          adsl_data.append(data[1][i])
          adsl_data.append(data[2][i])
          adsl_data.append(data[3][i])
          adsl_data.append('') # for cell comment
        adsl_data.append("None")
        adsl_data.append("None")
        adsl_data.append("None")
        adsl_data.append("None")
      if technology == 2:
        data = extract_data_from_page(urls[technology][day_group], list_of_day_groups)
        atsh_data.append(data[0])
        atsh_data.append("None")
        atsh_data.append("None")
        atsh_data.append("None")
        for i in range(len(data[1])):
          atsh_data.append(data[1][i])
          atsh_data.append(data[2][i])
          atsh_data.append(data[3][i])
          atsh_data.append('') # for cell comment
        atsh_data.append("None")
        atsh_data.append("None")
        atsh_data.append("None")
        atsh_data.append("None")
        
  all_data = [pon_data, adsl_data, atsh_data]
  return all_data


def save_xlsx_file(parsed_data_lines, filename, stats_data):
  '''Method for saving parsed_data_lines structurwe to Excel file with filename.'''
  
  if filename.endswith('.xls'):
    filename = filename[:-4]
  elif filename.endswith('.xlsx'):
    filename = filename[:-5]
  new_filename = filename + '_result.xlsx'
  try:
    os.remove(new_filename)
  except:
    pass
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
   
   
  pon_data, adsl_data, atsh_data = [], [], []
  for sheet in range(len(parsed_data_lines)):  
    for i in range(0,len(parsed_data_lines[sheet]), 4):
      if parsed_data_lines[sheet][i+1] == 'Выезд' or parsed_data_lines[sheet][i+1] == 'Выезд IP ОПС':
        pass
      else:
        if sheet == 0:
          pon_data.append(parsed_data_lines[sheet][i])
          pon_data.append(parsed_data_lines[sheet][i+1])
          pon_data.append(parsed_data_lines[sheet][i+2])
          pon_data.append(parsed_data_lines[sheet][i+3])
        if sheet == 1:
          atsh_data.append(parsed_data_lines[sheet][i])
          atsh_data.append(parsed_data_lines[sheet][i+1])
          atsh_data.append(parsed_data_lines[sheet][i+2])
          atsh_data.append(parsed_data_lines[sheet][i+3])
        if sheet == 2:
          adsl_data.append(parsed_data_lines[sheet][i])
          adsl_data.append(parsed_data_lines[sheet][i+1])
          adsl_data.append(parsed_data_lines[sheet][i+2])
          adsl_data.append(parsed_data_lines[sheet][i+3])
  data = [pon_data, adsl_data, atsh_data]
  
  # print(parsed_data_lines[1])
  # input("sss")
  for sheet in range(-len(data)+1,1):
    print(len(data[sheet]))
    if sheet == 0:
      sh = ss.Worksheets.Add()
      sh.Name = "PON"
    elif abs(sheet) == 1:
      sh = ss.Worksheets.Add()
      sh.Name = "ADSL"
    elif abs(sheet) == 2:
      sh = ss.Worksheets.Add()
      sh.Name = "АТШ"
    
    for i in range(0,len(data[sheet]), 4):
      print(data[sheet][i], data[sheet][i+1], data[sheet][i+2], data[sheet][i+3])
      sh.Cells(i//4+1,1).Value = str(data[sheet][i])
      sh.Cells(i//4+1,2).Value = str(data[sheet][i+1])
      sh.Cells(i//4+1,3).Value = str(data[sheet][i+2])
      sh.Cells(i//4+1,4).Value = str(data[sheet][i+3])
      
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
  


def main():
  url_example = 'http://dts/ora/graph/tabs/tt_list.php?category=0&t=1&tech=PON'
  list_of_day_groups = ['> 30 дней', '25 - 30 дней', '20 - 25 дней', '15 - 20 дней', '10 - 15 дней', '7 - 10 дней', '5 - 7 дней', '3 - 5 дней', '2 - 3 дня', '< 48 часов']
  
  urls = create_urls_list(url_example)
  
  all_data = extract_all_data(urls, list_of_day_groups)
  
  stats_data=[]
  filename = os.path.abspath(os.path.join('.', '_excel_files')) + '\\pp.xlsx'
  
  # print(all_data[0])
  # input("dsds")
  save_xlsx_file(all_data, filename, stats_data)

  
  # url = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=1&tech=PON"
  # url2 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=2&tech=PON"
  # url3 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=3&tech=PON"
  # url4 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=8&tech=PON"
  # url5 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=10&tech=PON"
  # url6 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=10&tech=%D0%90%D0%A2%D0%A8"
  # url7 = "http://dts/ora/graph/tabs/tt_list.php?category=0&t=10&tech=ADSL"
  # path_to_result_file = os.path.abspath(os.path.join('.', '_excel_files')) + '\\result.xlsx'
  # print('\n')
  # print(path_to_result_file)
  # extract_data_from_page(url)
  # extract_data_from_page(url2)
  # extract_data_from_page(url3)
  # extract_data_from_page(url4)
  # extract_data_from_page(url5)
  # extract_data_from_page(url6)
  # extract_data_from_page(url7)
  # create_urls_list(url_example)
  
  pass
  
if __name__ == '__main__':
    main()
    

#NEED TO FIX Dates: now dated format is: 16-07-08 21:06. Must be: 08-07-2016 21:06