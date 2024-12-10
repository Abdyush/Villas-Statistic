from openpyxl.cell import MergedCell
from openpyxl.styles import Font, PatternFill
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side


# ------------------------ ПОДГОТОВИТЕЛЬНЫЕ ФУНКЦИИ -----------------------------------------

# Функция, определяющая диапазоны категорий вилл для отдельной страницы
def define_range(sheet_):
   for column in range(1, sheet_.max_column):
      if sheet_[1][column].fill.start_color.index == '00000000' and type(sheet_[1][column]) != MergedCell:
         column_letters = str(sheet_[1][column - 1].coordinate)[:-1]
         veg = sheet_['C5': column_letters + '16']
         vps = sheet_['C18': column_letters + '19']
         vig = sheet_['C21': column_letters + '23']
         fwv = sheet_['C33': column_letters + '48']
         pwv = sheet_['C50': column_letters + '53']
         categories = [(veg, 'VEG'), (vps, 'VPS'), (vig, 'VIG'), (fwv, 'FWV'), (pwv, 'PWV')]
         return categories
      
# Совершенная функция, определяющая диапазоны категорий вилл для отдельной страницы подходящая для обоих файлов
def perfect_define_range(sheet, index_sheet, index_sheets):
   # Определяем ряды начала и окончания диапазонов
   ranges = {'veg': ['4001', '4012'],
             'vps': ['5001', '5002'],
             'vig': ['5003', '5005'],
             'fwv': ['4013', '4028'],
             'pwv': ['5006', '5009']}
   if index_sheet == 11 and index_sheets == 0:
       col = 0
   else:
       col = 1
   for row in range(1, 70):
       for key, value in ranges.items():
           if str(sheet[row][col].value) in value:
               i = value.index(str(sheet[row][col].value))
               value[i] = row

   # Определяем колонки окончания диапазонов
   if index_sheet == 0 and index_sheets == 1:
      end_column_letter = 'BL'
   else:
      for column in range(1, sheet.max_column):
         if sheet[1][column].fill.start_color.index == '00000000' and type(sheet[1][column]) != MergedCell:
            end_column_letter = str(sheet[1][column - 1].coordinate)[:-1]
            break
         
   if index_sheet == 11 and index_sheets == 0:
      start_column_letter = 'B'
   else:
      start_column_letter = 'C'
         
   veg = sheet[start_column_letter + str(ranges['veg'][0]): end_column_letter + str(ranges['veg'][1])]
   vps = sheet[start_column_letter + str(ranges['vps'][0]): end_column_letter + str(ranges['vps'][1])]
   vig = sheet[start_column_letter + str(ranges['vig'][0]): end_column_letter + str(ranges['vig'][1])]
   fwv = sheet[start_column_letter + str(ranges['fwv'][0]): end_column_letter + str(ranges['fwv'][1])]
   pwv = sheet[start_column_letter + str(ranges['pwv'][0]): end_column_letter + str(ranges['pwv'][1])]

   categories = [(veg, 'VEG'), (vps, 'VPS'), (vig, 'VIG'), (fwv, 'FWV'), (pwv, 'PWV')]
   return categories

        

# Функция настройки разметки первого листа excel
def prepare_first_sheet(sheet):
    # Настройка длины ячеек
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 45
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 17
    sheet.column_dimensions['E'].width = 30
    sheet.column_dimensions['F'].width = 30
    sheet.column_dimensions['G'].width = 25
    sheet.column_dimensions['H'].width = 22
    sheet.column_dimensions['I'].width = 20

    # Настройка высоты ячеек 
    for i in range(1, 1500):
        sheet.row_dimensions[i].height = 20

    # Cоздание заголовков
    sheet['A1'] = 'БАТЛЕР'
    sheet['B1'] = 'ГОСТЬ'
    sheet['C1'] = 'КАТЕГОРИЯ'
    sheet['D1'] = 'ВИЛЛА'
    sheet['E1'] = 'ТАРИФ'
    sheet['F1'] = 'ЗАЕЗД'
    sheet['G1'] = 'ВЫЕЗД'
    sheet['H1'] = 'КОЛИЧЕСТВО СУТОК'
    sheet['I1'] = 'CМЕННОСТЬ'

    # Форматирование и выравнивание заголовков
    for column in range(9):
        sheet[1][column].font = Font(name='Arial', bold=True)
        sheet[1][column].alignment = Alignment(horizontal="center", vertical="center")


# Функция настройки разметки второго листа excel
def prepare_second_table(sheet1):
    # Объединение ячеек
    sheet1.merge_cells('A1:T1')
    sheet1.merge_cells('A2:A3')
    sheet1.merge_cells('B2:G2')
    sheet1.merge_cells('H2:K2')
    sheet1.merge_cells('L2:M2')
    sheet1.merge_cells('N2:T2')

    # Вписываем названия столбцов
    sheet1['A1'] = 'Статистика'
    sheet1['A2'] = 'Батлер'

    sheet1['B2'] = 'Количество вилл'
    sheet1['B3'] = 'Всего:'
    sheet1['C3'] = 'veg'
    sheet1['D3'] = 'fwv'
    sheet1['E3'] = 'pwv'
    sheet1['F3'] = 'vps'
    sheet1['G3'] = 'vig'
    sheet1['H2'] = 'Тариф'

    sheet1['L2'] = 'Сменность'
    sheet1['L3'] = 'Один'
    sheet1['M3'] = '2/2'

    sheet1['N2'] = 'Занятость'
    sheet1['N3'] = 'Статус:'
    sheet1['O3'] = 'Дата выезда:'
    sheet1['P3'] = 'Гость:'
    sheet1['Q3'] = 'Планируемые заезды:'
    sheet1['R3'] = 'Ближайший заезд:'
    sheet1['S3'] = 'Дней до заезда:'
    sheet1['T3'] = 'Гость:'

    sheet1['H3'].fill = PatternFill('solid', fgColor='FFED7D31')
    sheet1['I3'].fill = PatternFill('solid', fgColor='FF00B050')
    sheet1['J3'].fill = PatternFill('solid', fgColor='FFD0CECE') 
    sheet1['K3'].fill = PatternFill('solid', fgColor='FFFFFF00') 

    # Настраиваем ширину ячеек
    sheet1.column_dimensions['A'].width = 20
    sheet1.column_dimensions['B'].width = 8
    for col in range(2, 11):
        letter = (sheet1[1][col].coordinate)[0]
        sheet1.column_dimensions[letter].width = 5
    sheet1.column_dimensions['L'].width = 8
    sheet1.column_dimensions['M'].width = 8
    sheet1.column_dimensions['N'].width = 15
    sheet1.column_dimensions['O'].width = 17
    sheet1.column_dimensions['P'].width = 22
    sheet1.column_dimensions['Q'].width = 38
    sheet1.column_dimensions['R'].width = 23
    sheet1.column_dimensions['S'].width = 20 
    sheet1.column_dimensions['T'].width = 20  

    # Настройка высоты ячеек 
    for i in range(1, 310):
        sheet1.row_dimensions[i].height = 20

    # Настраиваем размер, шрифт, расположение текста шапки
    for row in range(1, 4):
        for column in range(20):
            sheet1[row][column].font = Font(name='Arial', size=12, bold=True)
            sheet1[row][column].alignment = Alignment(horizontal="center", vertical="center")

    # Настраиваем обрамление ячеек
    def set_border(ws, cell_range):
        thin = Side(border_style="thin", color="000000")
        for row in ws[cell_range]:
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    set_border(sheet1, 'A1:T3') 



