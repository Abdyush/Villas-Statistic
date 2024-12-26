import re
from openpyxl.cell import MergedCell
from datetime import datetime, timedelta


# ------------------ФУНКЦИИ ИЗВЛЕЧЕНИЯ ПАРАМЕТРОВ ИЗ ОККУПАНСИ------------------------------------

# Извлечение фамилии батлера из комментария ячейки
def separate_butler(cmt, other_comments):
    all_butlers = ['Булгаков', 'Волгузов', 'Волков', 'Гетало', 'Гончар', 'Дембицкий', 'Диденко', 'Ларионов', 'Онищук', 'Орлов', 
    'Ляшов', 'Сергеев', 'Тараев', 'Абдюшев', 'Люфт', 'Макухин', 'Мартынов', 'Марченко', 'Нечипуренко', 'Старенький', 'Стибельский', 'Тарабанов', 
    'Федоренко', 'Черноштан', 'Шаповалов', 'Царенков', 'Потапов', 'Доронин', 'Мантуров', 'Саетов', 
    'Солодаев', 'Календжян', 'Буркеев', 'Карманов', 'Салединов', 'Тлашадзе', 'Соловьев', 'Дейнекин', 
    'Лесников', 'Гончаров', 'Попов', 'Пак', 'Курилко', 'Росихин']
    
    present_butlers = ['Булгаков', 'Волгузов', 'Волков', 'Гетало', 'Гончар', 'Дембицкий', 'Диденко', 'Онищук', 'Орлов', 
    'Ляшов', 'Сергеев', 'Тараев', 'Абдюшев', 'Люфт', 'Макухин', 'Мартынов', 'Марченко', 'Нечипуренко', 'Стибельский', 'Тарабанов', 
    'Федоренко', 'Черноштан', 'Шаповалов']

    but = []
    butler = []
    # Убираем лишнее из комментария
    if 'автор: мартынов' in cmt.lower():
        cmt = cmt[:cmt.lower().index('автор: мартынов')]
    elif 'автор мартынов' in cmt.lower():
        cmt = cmt[:cmt.lower().index('автор мартынов')]

    for name in all_butlers:
        if (name == 'Федоренко' and name.lower() in cmt.lower() and 'Н' in cmt) or ((name == 'Мартынов' and name.lower() in cmt.lower()) and 'Л' in cmt or 'Алексей' in cmt):
            continue
        else: 
            if name.lower() in cmt.lower():
                if name == 'Солодаев':
                    name = 'Ляшов'
                but.append(name)
    if but == []:
        other_comments.append(cmt)
    else:
      shift = len(but)
      if shift == 1:
         shift = 'Один'
      elif shift > 1:
         shift = '2/2'
     
      
      for name in present_butlers:
         if name in but:
            butler.append(name)
    if len(butler) > 0:
        return shift, butler


# Совершенная версия определения периода проживания виллы
def perfect_count_days(sheets, index_sheet, category, number_row, number_column, index_sheets, number_villa):

    # Определяем год
    if index_sheets == 0 and index_sheet <= 11:
        year = 2023
    elif index_sheets == 1 and index_sheet <= 11 or (index_sheets == 0 and index_sheet == 12):
        year = 2024
    elif index_sheets == 2 or (index_sheets == 1 and index_sheet == 12):
        year = 2025

    # Определяем месяц
    if index_sheet == 12:
        month = 1
    else:
        month = index_sheet + 1

    # Определяем день
    flag = False
    start_month = None
    end_month = None
    let = ''.join(list(filter(lambda x: x.isalpha(), str(category[number_row][number_column].coordinate))))
    day_arrival = sheets[index_sheet][let + '2'].value
    h = 0
    if day_arrival == None:
        let = ''.join(list(filter(lambda x: x.isalpha(), str(category[number_row][number_column - 1].coordinate))))
        day_arrival = sheets[index_sheet][let + '2'].value
        h = 12
        flag = True

    # Формируем дату заезда
    arrival = datetime(hour=h, day=day_arrival, month=month, year=year)
    if (arrival + timedelta(days=1)).month - arrival.month == 1 or (arrival + timedelta(days=1)).month - arrival.month == -11:
        check_out = arrival
        end_month = number_villa
        return [(arrival, check_out), ('end_month:', number_villa)]
    else:
        if arrival.day == 1 and flag == False:
            start_month = number_villa

    # Определяем дату выезда
    j = number_column + 1
    check_out = arrival + timedelta(hours=12)
    while type(category[number_row][j]) == MergedCell:
        check_out += timedelta(hours=12)
        if ((check_out + timedelta(hours=12)).month - check_out.month == 1 or (check_out + timedelta(hours=12)).month - check_out.month == -11) and type(category[number_row][j + 1]) == MergedCell:
            end_month = number_villa
            break             
        else:
            j += 1

    if start_month == None and end_month == None:
        return [(arrival, check_out), ('Количество суток:', (check_out - arrival).days)]
    elif start_month != None and end_month == None and index_sheet == 0 and index_sheets == 0:
        return [(arrival, check_out), ('Количество суток:', (check_out - arrival).days)]
    elif start_month != None and end_month == None:
        return [(arrival, check_out), ('start_month:', start_month)]
    elif end_month != None and start_month == None:
        return [(arrival, check_out), ('end_month:', end_month)]
    else:
        return [(arrival, check_out), ('start_month:', start_month), ('end_month:', end_month)]


# Функция, определяющая тариф
def determine_rate(cell):
   open_market = ['FFED7D31', 'FFFF9933', 'FFFF6600']
   complimentary = ['FFFFFF00']
   sber = ['FF00B050', 'FF70AD47']
   upgrade = ['FFD0CECE', 'FFD9D9D9', 'FFE7E6E6', 'FFBFBFBF']

   cell_fill = cell.fill.start_color.index #Получаем цвет ячейки
   if cell_fill in open_market:
      return 'Открытый рынок'
   elif cell_fill in complimentary:
      return 'Комплиментарный тариф'
   elif cell_fill in sber:
      return 'Сбер'
   elif cell_fill in upgrade:
       return 'Апгрейд'
   else:
      return cell_fill


# Функция определяющая номер виллы
def define_number_villa(index_sheet, category, number_row, index_sheets, sheet):
    # Определяем номер столбца
    if index_sheet == 11 and index_sheets == 0:
        col = 0
    else:
        col = 1
    # Определяем номер ряда
    row = int(str(category[number_row][0].coordinate)[1:])
    # Извлекаем номер виллы
    number_villa = sheet[row][col].value

    return number_villa
    

# Функция обработки фамилий гостя
def process_name_guest(name):
    # Преверяем есть ли цифры в строке, и убираем если есть
    m = re.search(r'\d', name)
    if m != None and (name[m.start() - 1] == ' ' or name[m.start() - 1] == '-'):
        name = name[:m.start() - 1]

    # Форматируем строку с заглавной буквы
    if name.isalpha():
        name = name.title()

    return name


