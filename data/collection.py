from data.recovery import separate_butler, determine_rate, perfect_count_days
from src.background import define_range
from data.processing import prepare_dict


# ------------------------------ ОСНОВНЫЕ ФУНКЦИИ ---------------------------------------------

# Создаем словарь с будущей статистикой батлеров
butlers = {}

# Функция, формирующая всю статистику в словарь
def generate_statistics(butlers, butler, guest, categ, rate, days, shift):
    if butler in butlers:
       if guest in butlers[butler]:
           if days[1][0] == 'start_month:':
               flag = False
               for villa in butlers[butler][guest]:
                   part_month = villa[2][1][0]
                   row_villa = villa[2][1][1]
                   month_villa = villa[2][0][0].month
                   if part_month == 'end_month:' and row_villa == days[1][1] and days[0][0].month - month_villa == 1 and villa[0] == categ:
                       days = [(villa[2][0][0], days[0][1]), ('Количество суток:', (days[0][1] - villa[2][0][0]).days)]
                       villa[2] = days
                       flag = True
               if flag == False:
                   butlers[butler][guest].append([categ, rate, days, shift]) 

           elif days[1][0] == 'end_month:':
               flag = False
               for villa in butlers[butler][guest]:
                   part_month = villa[2][1][0]
                   row_villa = villa[2][1][1]
                   month_villa = villa[2][0][0].month
                   if part_month == 'start_month:' and row_villa == days[1][1] and month_villa - days[0][0].month == 1 and villa[0] == categ:
                       days = [(days[0][0], villa[2][0][1]), ('Количество суток:', (villa[2][0][1] - days[0][0]).days)]
                       villa[2] = days
                       flag = True
               if flag == False:
                   butlers[butler][guest].append([categ, rate, days, shift])
           # здесь должен быть вариант, когда есть и конец и начало месяца    
            
           else:
               butlers[butler][guest].append([categ, rate, days, shift])       
       else:
           butlers[butler].setdefault(guest, [[categ, rate, days, shift]])
    else:
       butlers.setdefault(butler, {guest: [[categ, rate, days, shift]]})

    prepare_dict(butlers)
   
other_comments = []


# Функция получения статистики месяца
def month_statistic(sheets, sheet, sheet_index):
    categories = define_range(sheet)
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
                           guest = row[column].value                          
                           categ = category[1]
                           rate = determine_rate(row[column])
                           days = perfect_count_days(sheets, sheet_index, category[0], category[0].index(row), column)
                        
                           for butler in but:
                               generate_statistics(butlers, butler, guest, categ, rate, days, shift)
                               
                    except:
                        continue




