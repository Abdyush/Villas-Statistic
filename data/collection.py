from datetime import datetime
import pickle
from data.recovery import separate_butler, determine_rate, perfect_count_days, define_number_villa, process_name_guest, determine_category
from src.background import perfect_perfect_define_range


# ------------------------------ ОСНОВНЫЕ ФУНКЦИИ ---------------------------------------------

# Создаем словарь с будущей статистикой батлеров

butlers = {}

def fill_dict(butlers):
    # open a pickle file
    filename1 = 'all_butlers.pk'
    filename2 = 'selected_butlers.pk'
    
    
    # load your data back to memory when you need it
    with open(filename2, 'rb') as fi:
        sel_but = pickle.load(fi)

    if sel_but != []:
        present_butlers = sel_but[0]
        present_butlers.extend(sel_but[1])
    else:
        with open(filename1, 'rb') as fi:
            all_but = pickle.load(fi)
        present_butlers = all_but[0]
        present_butlers.extend(all_but[1])

    for name in present_butlers:
        butlers[name] = {}


# Эксперементальная функция с подсчетом дней
def count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift):
    if days[1][0] == 'start_month:':
        flag = False
        # Создаем цикл из двух итераций на случай если для времени 'present' в словаре не найдется конец месяца
        for _ in range(2):
            for villa in butlers[butler][time][guest]:
                part_month = villa[3][1][0]
                number = villa[3][1][1]
                month_villa = villa[3][0][0].month
                # Проверяем не является ли переданная функцию вилла копией уже имеющейся вилле в словаре
                if number_villa == villa[1] and rate == villa[2] and shift == villa[4] and days[0][1] == villa[3][0][1]:
                    flag = True
                    break
                if part_month == 'end_month:' and number == days[1][1] and (days[0][0].month - month_villa == 1 or days[0][0].month - month_villa == -11) and villa[0] == categ:
                    days = [(villa[3][0][0], days[0][1]), ('Количество суток:', (days[0][1] - villa[3][0][0]).days)]
                    villa[3] = days
                    flag = True
                    break 
            if flag == False:
                if time == 'future':
                    time = 'present'
                else:
                    butlers[butler][time][guest].append([categ, number_villa, rate, days, shift])
                    break 

    elif days[1][0] == 'end_month:' or (len(days) == 3 and days[2][0] == 'end_month:'):
        flag = False
        for villa in butlers[butler][time][guest]:
            part_month = villa[3][1][0]
            number = villa[3][1][1]
            month_villa = villa[3][0][0].month
            # Проверяем не является ли переданная функцию вилла копией уже имеющейся вилле в словаре
            if number_villa == villa[1] and rate == villa[2] and shift == villa[4] and days[0][1] == villa[3][0][1]:
                flag = True
                break
            if part_month == 'start_month:' and number == days[1][1] and month_villa - days[0][0].month == 1 and villa[0] == categ:
                days = [(days[0][0], villa[3][0][1]), ('Количество суток:', (villa[3][0][1] - days[0][0]).days)]
                villa[3] = days
                flag = True
                break
        if flag == False:
            butlers[butler][time][guest].append([categ, number_villa, rate, days, shift])       
            
    else:
        butlers[butler][time][guest].append([categ, number_villa, rate, days, shift])


# Функция, формирующая всю статистику в словарь, делящая виллы на прошедкшие и будующие
def new_generate_statistics(butlers, time, butler, guest, categ, rate, days, shift, number_villa):
    if butler in butlers:
        if time not in butlers[butler]:
            if time == 'future' and days[1][0] == 'start_month:':
                time = 'present'
                count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
            else:
                butlers[butler].setdefault(time, {guest: [[categ, number_villa, rate, days, shift]]})
        else:
            if guest in butlers[butler][time]:
                count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
            else:
                if time == 'future' and days[1][0] == 'start_month:':
                    time = 'present'
                    count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
                else:
                    butlers[butler][time].setdefault(guest, [[categ, number_villa, rate, days, shift]])

    else:
        butlers.setdefault(butler, {time: {guest: [[categ, number_villa, rate, days, shift]]}})
    
other_comments = []


# Функция получения статистики месяца с учетом времени заезда виллы
def new_month_statistic(sheets, sheet, sheet_index, sheets_index):
    reference_date = datetime.today()
    categories = perfect_perfect_define_range(sheet, sheet_index, sheets_index)
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
                               
                               if days[0][0] > reference_date:
                                   time = 'future'                            
                               else:
                                   time = 'present'
                                   
                               new_generate_statistics(butlers=butlers, time=time, butler=butler, guest=guest, categ=categ, 
                                                     rate=rate, days=days, shift=shift, number_villa=number_villa)                                      
            
                    except:
                        continue



# Функция получения статистики месяца с учетом времени заезда виллы (совершенная)
def perfect_new_month_statistic(sheets, sheet, sheet_index, sheets_index):
    reference_date = datetime.today()
    diap = sheet.calculate_dimension()
    diapazone = sheet[diap]


    for row in diapazone:
        for column in range(len(row)):
            if row[column].value != None:
                try:
                    cmt = row[column].comment.text
                    if separate_butler(cmt, other_comments) == None:
                        continue
                    else:
                        shift, but = separate_butler(cmt, other_comments)
                        guest = process_name_guest(name=row[column].value)                       
                                        
                        rate = determine_rate(row[column])                         
                        number_villa = define_number_villa(index_sheet=sheet_index, category=diapazone, 
                                                    number_row=(diapazone.index(row)), index_sheets=sheets_index, 
                                                    sheet=sheet)
                        categ = determine_category(number_villa) 
                        days = perfect_count_days(sheets=sheets, index_sheet=sheet_index, category=diapazone,
                                        number_row=diapazone.index(row), number_column=column, index_sheets=sheets_index, 
                                        number_villa=number_villa)
                        

                        for butler in but:
                            
                            if days[0][0] > reference_date:
                                time = 'future'                            
                            else:
                                time = 'present'
                                
                            perfect_new_generate_statistics(butlers=butlers, time=time, butler=butler, guest=guest, categ=categ, 
                                                    rate=rate, days=days, shift=shift, number_villa=number_villa)                                      
        
                except:
                    continue


# Функция, формирующая всю статистику в словарь, делящая виллы на прошедкшие и будующие (новая)
def perfect_new_generate_statistics(butlers, time, butler, guest, categ, rate, days, shift, number_villa):
    
    if time not in butlers[butler]:
        if time == 'future' and days[1][0] == 'start_month:':
            time = 'present'
            count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
        else:
            butlers[butler].setdefault(time, {guest: [[categ, number_villa, rate, days, shift]]})
    else:
        if guest in butlers[butler][time]:
            count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
        else:
            if time == 'future' and days[1][0] == 'start_month:':
                time = 'present'
                count_final_days(butlers, butler, time, guest, days, categ, number_villa, rate, shift)
            else:
                butlers[butler][time].setdefault(guest, [[categ, number_villa, rate, days, shift]])

    
    
other_comments = []