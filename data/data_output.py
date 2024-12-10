from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side

from src.background import prepare_first_sheet, prepare_second_table

#------------------------------------ ВЫВОД ДАННЫХ В ECXEL ---------------------------------------------------------


# Функция заполнения первой таблицы словарем
def fill_first_table(sheet, butlers):
    # Подготавливаем таблицу
    prepare_first_sheet(sheet)
    
    # Заполняем словарь
    flag = True
    row = 2
    col = 3
    thin = Side(border_style="thin", color="000000")
    for butler, villa in butlers.items():
        cell = sheet.cell(row=row, column=1)
        cell.value = butler
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(name='Arial', size=12, bold=True)
        cell.border = Border(right=thin, bottom=thin)
    
        for guest, info in villa.items():
            cell = sheet.cell(row=row, column=2)
            cell.value = guest
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name='Arial', size=12, bold=True)
            cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
        
            for iterance in info:
                if info.index(iterance) > 0:
                    row += 1
                cell = sheet.cell(row=row, column=2)
                cell.value = guest
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name='Arial', size=12, bold=True)
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                col = 3
                for param in iterance:
                    if type(param) == list:
                        cell = sheet.cell(row=row, column=col)
                        cell.value = param[0][0].strftime('%d.%m.%y')
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name='Arial', size=12)
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                        col += 1
                    
                        cell = sheet.cell(row=row, column=col)
                        cell.value = param[0][1].strftime('%d.%m.%y')
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name='Arial', size=12)
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                        col += 1

                        cell = sheet.cell(row=row, column=col)
                        cell.value = param[1][1]
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name='Arial', size=12)
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                        col += 1

                    elif iterance.index(param) == 2:
                        cell = sheet.cell(row=row, column=col)
                        cell.value = param
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name='Arial', size=12)
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

                        if param == 'Открытый рынок':
                            for j in range(2, 10):
                                sheet.cell(row, j).fill = PatternFill('solid', fgColor='FFED7D31')
                        elif param == 'Сбер':
                            for j in range(2, 10):
                                sheet.cell(row, j).fill = PatternFill('solid', fgColor='FF00B050')
                        elif param == 'Комплиментарный тариф':
                            for j in range(2, 10):
                                sheet.cell(row, j).fill = PatternFill('solid', fgColor='FFFFFF00')
                        elif param == 'Апгрейд':
                            for j in range(2, 10):
                                sheet.cell(row, j).fill = PatternFill('solid', fgColor='FFD0CECE')
                        else:
                            for j in range(2, 10):
                                sheet.cell(row, j).fill = PatternFill('solid', fgColor='00C0C0C0')       
                    
                        col += 1
           
                    else:
                        cell = sheet.cell(row=row, column=col)
                        cell.value = param
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.font = Font(name='Arial', size=12)
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                        col += 1
                             
            row += 1

        row += 1   
         

# Функция заполнения второй таблицы словарем
def fill_second_table(sheet1, general_dictionary):
    # Настраиваем разметку таблицы
    prepare_second_table(sheet1=sheet1)

    # Заполняем таблицу статистикой
    row = 4
    col = 1
    thin = Side(border_style="thin", color="000000")
    for butler, stat in general_dictionary.items():
        cell = sheet1.cell(row=row, column=col)
        cell.value = butler
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.font = Font(name='Arial', size=12, bold=True)
        cell.border = Border(top=thin, right=thin, bottom=thin)
        col += 1
        for lst in stat:
            for dct in lst:
                for key, value in dct.items():
                    cell = sheet1.cell(row=row, column=col)
                    if type(value) == list:
                        value = ', '.join(value)
                    if value == 'Свободен':
                        cell.fill = PatternFill('solid', fgColor='00FF00')
                    elif value == 'С виллой':
                        cell.fill = PatternFill('solid', fgColor='FF0000')
                    cell.value = value
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name='Arial', size=12)
                    if col == 7 or col == 11 or col == 13 or col == 14 or col == 16 or col == 20 or col == 2:
                        cell.border = Border(top=thin, right=thin, bottom=thin)
                    else:
                        cell.border = Border(top=thin, bottom=thin)
                    col += 1
        row += 1
        col = 1




