from data.recovery import separate_butler, determine_rate, perfect_count_days, define_number_villa, process_name_guest
from src.background import perfect_define_range




# ------------------------------ ОСНОВНЫЕ ФУНКЦИИ ---------------------------------------------

# Создаем словарь с будущей статистикой батлеров
butlers = {}

# Функция, формирующая всю статистику в словарь
def generate_statistics(butlers, butler, guest, categ, rate, days, shift, number_villa):
    if butler in butlers:
       if guest in butlers[butler]:
           if days[1][0] == 'start_month:':
               flag = False
               for villa in butlers[butler][guest]:
                   part_month = villa[3][1][0]
                   number = villa[3][1][1]
                   month_villa = villa[3][0][0].month
                   if part_month == 'end_month:' and number == days[1][1] and (days[0][0].month - month_villa == 1 or days[0][0].month - month_villa == -11) and villa[0] == categ:
                       days = [(villa[3][0][0], days[0][1]), ('Количество суток:', (days[0][1] - villa[3][0][0]).days)]
                       villa[3] = days
                       flag = True
               if flag == False:
                   butlers[butler][guest].append([categ, number_villa, rate, days, shift]) 

           elif days[1][0] == 'end_month:' or (len(days) == 3 and days[2][0] == 'end_month:'):
               flag = False
               for villa in butlers[butler][guest]:
                   part_month = villa[3][1][0]
                   number = villa[3][1][1]
                   month_villa = villa[3][0][0].month
                   if part_month == 'start_month:' and number == days[1][1] and month_villa - days[0][0].month == 1 and villa[0] == categ:
                       days = [(days[0][0], villa[3][0][1]), ('Количество суток:', (villa[3][0][1] - days[0][0]).days)]
                       villa[3] = days
                       flag = True
               if flag == False:
                   butlers[butler][guest].append([categ, number_villa, rate, days, shift])       
            
           else:
               butlers[butler][guest].append([categ, number_villa, rate, days, shift])       
       else:
           butlers[butler].setdefault(guest, [[categ, number_villa, rate, days, shift]])
    else:
       butlers.setdefault(butler, {guest: [[categ, number_villa, rate, days, shift]]})

   
other_comments = []


# Функция получения статистики месяца
def month_statistic(sheets, sheet, sheet_index, sheets_index):
    categories = perfect_define_range(sheet, sheet_index, sheets_index)
    for category in categories:
        for row in category[0]:
            for column in range(len(row)):
                if row[column].value != None:
                    try:
                        cmt = row[column].comment.text
                        if separate_butler(cmt, other_comments) == None:
                           continue
                        else:
                           shift, but = separate_butler(cmt, other_comments)
                           guest = process_name_guest(name=row[column].value)                         
                           categ = category[1]
                           rate = determine_rate(row[column])                         
                           number_villa = define_number_villa(index_sheet=sheet_index, category=category[0], 
                                                      number_row=(category[0].index(row)), index_sheets=sheets_index, 
                                                      sheet=sheet)
                           days = perfect_count_days(sheets=sheets, index_sheet=sheet_index, category=category[0],
                                          number_row=category[0].index(row), number_column=column, index_sheets=sheets_index, 
                                          number_villa=number_villa)
            
                           for butler in but:        
                               generate_statistics(butlers=butlers, butler=butler, guest=guest, categ=categ, 
                                                     rate=rate, days=days, shift=shift, number_villa=number_villa)
            
                    except:
                        continue




