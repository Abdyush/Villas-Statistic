import os
import sys
sys.path.append(os.getcwd())

import openpyxl
from src.background import perfect_define_range
from data.recovery import separate_butler, determine_rate, define_number_villa, perfect_count_days

# Путь к файлу эксель с распределением вилл 
FILE_PATH = '/home/user/Рабочий стол/StudyProject/villas/distribution/Villas occupancy 2023.xlsx' 

# Путь к файлу эксель с распределением вилл 
FILE_24_PATH = '/home/user/Рабочий стол/StudyProject/villas/distribution/Villas occupancy 2024.xlsx' 

# Открываем файл
file = openpyxl.open(FILE_PATH)
file2 = openpyxl.open(FILE_24_PATH)

# определяем список со всеми доступными страницами файла
sheets = file.worksheets
sheets2 = file2.worksheets

list_sheet = [sheets, sheets2]

other_comments = []

#Перебирая листы файла за каждый месяц, добавляем каждый лист в функцию получения статистики, таким образом, получаем статистику за год
for sheets in list_sheet:
    for sheet in sheets:
       categories = perfect_define_range(sheet, sheets.index(sheet), list_sheet.index(sheets))
       for category in categories:
          for row in category[0]:
             for column in range(len(row)):
                if row[column].value != None:
                   try:
                      cmt = row[column].comment.text
                      but = separate_butler(cmt, other_comments)
                      rate = determine_rate(row[column])
                      number_villa = define_number_villa(index_sheet=(sheets.index(sheet)), category=(category[0]), 
                                                      number_row=(category[0].index(row)), index_sheets=list_sheet.index(sheets), 
                                                      sheet=sheet)
                      days = perfect_count_days(sheets=sheets, index_sheet=sheets.index(sheet), category=category[0],
                                          number_row=category[0].index(row), number_column=column, index_sheets=list_sheet.index(sheets), 
                                          number_villa=number_villa)
                      print(days)
                      #print(row[column].value)
                      #print(but)
                      #print(rate)
                      #print(number_villa)
                      #print(days)
                      #print()
                      
                      
                   except:
                      continue




# Закрываем файл
file.close()