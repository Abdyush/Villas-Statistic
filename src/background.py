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
         break


# Функция настройки разметки первого листа excel
def prepare_first_sheet(sheet):
    # Настройка длины ячеек
    sheet.column_dimensions['A'].width = 20
    sheet.column_dimensions['B'].width = 45
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 30
    sheet.column_dimensions['E'].width = 30
    sheet.column_dimensions['F'].width = 30
    sheet.column_dimensions['G'].width = 25
    sheet.column_dimensions['H'].width = 20

    # Настройка высоты ячеек 
    for i in range(1, 310):
        sheet.row_dimensions[i].height = 20

    # Cоздание заголовков
    sheet['A1'] = 'БАТЛЕР'
    sheet['B1'] = 'ГОСТЬ'
    sheet['C1'] = 'КАТЕГОРИЯ'
    sheet['D1'] = 'ТАРИФ'
    sheet['E1'] = 'ЗАЕЗД'
    sheet['F1'] = 'ВЫЕЗД'
    sheet['G1'] = 'КОЛИЧЕСТВО СУТОК'
    sheet['H1'] = 'CМЕННОСТЬ'

    # Форматирование и выравнивание заголовков
    for column in range(8):
        sheet[1][column].font = Font(name='Arial', bold=True)
        sheet[1][column].alignment = Alignment(horizontal="center", vertical="center")


# Функция настройки разметки второго листа excel
def prepare_second_table(sheet1):
    # Объединение ячеек
    sheet1.merge_cells('A1:S1')
    sheet1.merge_cells('A2:A3')
    sheet1.merge_cells('B2:G2')
    sheet1.merge_cells('H2:J2')
    sheet1.merge_cells('K2:L2')
    sheet1.merge_cells('M2:S2')

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

    sheet1['K2'] = 'Сменность'
    sheet1['K3'] = 'Один'
    sheet1['L3'] = '2/2'

    sheet1['M2'] = 'Занятость'
    sheet1['M3'] = 'Статус:'
    sheet1['N3'] = 'Дата выезда:'
    sheet1['O3'] = 'Гость:'
    sheet1['P3'] = 'Планируемые заезды:'
    sheet1['Q3'] = 'Ближайший заезд:'
    sheet1['R3'] = 'Дней до заезда:'
    sheet1['S3'] = 'Гость:'

    sheet1['H3'].fill = PatternFill('solid', fgColor='FFED7D31')
    sheet1['I3'].fill = PatternFill('solid', fgColor='FF00B050')
    sheet1['J3'].fill = PatternFill('solid', fgColor='FFFFFF00')

    # Настраиваем ширину ячеек
    sheet1.column_dimensions['A'].width = 20
    sheet1.column_dimensions['B'].width = 8
    for col in range(2, 10):
        letter = (sheet1[1][col].coordinate)[0]
        sheet1.column_dimensions[letter].width = 5
    sheet1.column_dimensions['K'].width = 8
    sheet1.column_dimensions['L'].width = 8
    sheet1.column_dimensions['M'].width = 15
    sheet1.column_dimensions['N'].width = 17
    sheet1.column_dimensions['O'].width = 25
    sheet1.column_dimensions['P'].width = 38
    sheet1.column_dimensions['Q'].width = 23
    sheet1.column_dimensions['R'].width = 20 
    sheet1.column_dimensions['S'].width = 20 

    # Настройка высоты ячеек 
    for i in range(1, 310):
        sheet1.row_dimensions[i].height = 20

    # Настраиваем размер, шрифт, расположение текста шапки
    for row in range(1, 4):
        for column in range(19):
            sheet1[row][column].font = Font(name='Arial', size=12, bold=True)
            sheet1[row][column].alignment = Alignment(horizontal="center", vertical="center")

    # Настраиваем обрамление ячеек
    def set_border(ws, cell_range):
        thin = Side(border_style="thin", color="000000")
        for row in ws[cell_range]:
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

    set_border(sheet1, 'A1:S3') 