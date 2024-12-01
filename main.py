import os
import sys
sys.path.append(os.getcwd())

import openpyxl
from data.collection import month_statistic, butlers
from data.processing import count_totals
from data.data_output import fill_first_table, fill_second_table


# Путь к файлу эксель с распределением вилл 
FILE_PATH = '/home/user/Рабочий стол/StudyProject/villas/distribution/Villas occupancy 2023.xlsx' 
# Путь к файлу со статистикой
STATISTIC_FILE_PATH = '/home/user/Рабочий стол/StudyProject/villas/Statistic/Butlers_statistic.xlsx'



#----------------------------------------РАБОТА С ФАЙЛОМ OCCUPANCY--------------------------------------------

# Открываем файл
file = openpyxl.open(FILE_PATH)

# определяем список со всеми доступными страницами файла
sheets = file.worksheets

#Перебирая листы файла за каждый месяц, добавляем каждый лист в функцию получения статистики, таким образом, получаем статистику за год
for sheet in sheets:
    month_statistic(sheets, sheet, sheets.index(sheet))

# Закрываем файл
file.close()



#-------------------------------------------ПОДСЧЕТ ИТОГОВ------------------------------------------------------

generaly_dictionary = count_totals(butlers)



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
fill_first_table(sheet, butlers)
# Заполняем вторую таблицу статистикой
fill_second_table(sheet1, generaly_dictionary)

# Выводим в терминал "Готово!", сохраняем и закрываем файл
print('Готово!')
file_stat.save(STATISTIC_FILE_PATH)
file_stat.close()


              


























#file_stat.save(STATISTIC_FILE_PATH)