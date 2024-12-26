import os
import sys
sys.path.append(os.getcwd())

import openpyxl
from data.collection import new_month_statistic, butlers
from data.processing import new_new_count_totals, new_prepare_dict
from data.data_output import new_fill_first_table, new_fill_second_table 


# Путь к файлу эксель с распределением вилл 
FILE_PATH = '/home/user/Рабочий стол/StudyProject/villas/distribution/Villas occupancy 2023.xlsx'
# Путь к файлу эксель с распределением вилл 
FILE_24_PATH = '/home/user/Рабочий стол/StudyProject/villas/distribution/Villas occupancy 2024.xlsx'  
# Путь к файлу со статистикой
STATISTIC_FILE_PATH = '/home/user/Рабочий стол/StudyProject/villas/Statistic/Butlers_statistic.xlsx'



#----------------------------------------РАБОТА С ФАЙЛОМ OCCUPANCY--------------------------------------------

# Открываем файл
file = openpyxl.open(FILE_PATH)
file2 = openpyxl.open(FILE_24_PATH)

# определяем список со всеми доступными страницами файла
sheets = file.worksheets
sheets2 = file2.worksheets

# Определяем список файлов с occupancy
list_sheet = [sheets, sheets2]


# Перебирая листы файла за каждый месяц, добавляем каждый лист в функцию получения статистики, таким образом, получаем статистику за год
for sheets in list_sheet:
    for sheet in sheets:
        new_month_statistic(sheets=sheets, sheet=sheet, sheet_index=sheets.index(sheet), sheets_index=list_sheet.index(sheets))

# Подводим словарь к финальному виду
new_prepare_dict(butlers)

# Закрываем файл
file.close()



#-------------------------------------------ПОДСЧЕТ ИТОГОВ------------------------------------------------------

generaly_dictionary = new_new_count_totals(butlers)



#----------------------------------------РАБОТА С ФАЙЛОМ BUTLERS STATISTIC--------------------------------------

# Открываем файл cо статистикой
file_stat = openpyxl.load_workbook(STATISTIC_FILE_PATH)

# Удаляем все листы из файла
all_sheet = file_stat.sheetnames
for sh in all_sheet:
    file_stat.remove(file_stat[sh])

# Создаем листы "Все виллы", "Статистика"
sheet = file_stat.create_sheet("Все виллы")
sheet1 = file_stat.create_sheet("Статистика")

# Заполняем первую таблицу всеми данными по виллам
new_fill_first_table(sheet, dict(sorted(butlers.items())))

# Заполняем вторую таблицу статистикой
new_fill_second_table(sheet1, generaly_dictionary)



#----------------------------------------------ЗАВЕРШЕНИЕ РАБОТЫ------------------------------------------------

# Выводим в терминал "Готово!", сохраняем и закрываем файл
print('Готово!')
file_stat.save(STATISTIC_FILE_PATH)
file_stat.close()


              


























#file_stat.save(STATISTIC_FILE_PATH)