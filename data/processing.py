from datetime import datetime
import locale

# Устанавливаем Русскую локализацию для datetime
locale.setlocale(locale.LC_ALL, 'ru_RU.UTF-8')


#------------------------------- ФУНКЦИИ РАБОТЫ СО СЛОВАРЕМ ---------------------------------------------------------------

# Поиск в словаре необработанных периодов проживания вилл, и приведение их к единообразному результату
def prepare_dict(butlers):
    for butler, villa in butlers.items():
        for guest, info in villa.items():
            for iterance in info:
                if iterance[3][1][0] != 'Количество суток:':
                    iterance[3][1] = ('Количество суток:', (iterance[3][0][1] - iterance[3][0][0]).days)


# Поиск в словаре необработанных периодов проживания вилл, и приведение их к единообразному результату
def new_prepare_dict(butlers):
    for butler, time in butlers.items():
        for guest, villa in time.items():
            for g, info in villa.items():
                for iterance in info:
                    if iterance[3][1][0] != 'Количество суток:':
                        iterance[3][1] = ('Количество суток:', (iterance[3][0][1] - iterance[3][0][0]).days)



# Функция распределения батлеров по сменам и подсчета итогов
def count_totals(butlers):
    first_shift = ['Булгаков', 'Волгузов', 'Волков', 'Гетало', 'Гончар', 'Дембицкий', 'Диденко', 'Ларионов', 'Онищук', 'Орлов', 
    'Ляшов', 'Сергеев', 'Тараев']

    second_shift = ['Абдюшев', 'Люфт', 'Макухин', 'Мартынов', 'Марченко', 'Нечипуренко', 'Старенький', 'Стибельский', 'Тарабанов', 
    'Федоренко', 'Черноштан', 'Шаповалов']

    stat_first_shift = {}
    stat_secon_shift = {}
    general_dictionary = {}

    reference_date = datetime.today()
    
    #reference_date = datetime(2023, 12, 5)

    for butler, villa in butlers.items():
        general_dictionary.setdefault(butler, [])

        veg, vps, vig, fwv, pwv = 0, 0, 0, 0, 0
        open_market, sber, complimentary, upgrade = 0, 0, 0, 0
        alone, twice = 0, 0
        total = 0
        planned_villas = 0
        status = 'Свободен'
        days_before_villa = 0
        dates_arrival = []
        last_check_out = []
        check_out_date = 0
        time_resource = []

        for guest, info in villa.items():
            for iterance in info:
                veg += iterance.count('VEG')
                vps += iterance.count('VPS')
                vig += iterance.count('VIG')
                fwv += iterance.count('FWV')
                pwv += iterance.count('PWV')

                open_market += iterance.count('Открытый рынок')
                sber += iterance.count('Сбер')
                complimentary += iterance.count('Комплиментарный тариф')
                upgrade += iterance.count('Апгрейд')

                alone += iterance.count('Один')
                twice += iterance.count('2/2')

                arrival = iterance[3][0][0]
                check_out = iterance[3][0][1]

                if check_out < reference_date:
                    last_check_out.append((check_out, guest))

                elif arrival < reference_date < check_out:
                    check_out_date = check_out
                    status = 'С виллой'
                    busy = guest

                elif arrival > reference_date:
                    dates_arrival.append((arrival, guest))
                    planned_villas += 1
                    
                total += 1 

        last_villas = sorted(last_check_out, reverse=True)
        next_arrival = sorted(dates_arrival)
        
        if status == 'Свободен' and planned_villas > 0:
            days_before_villa = (next_arrival[0][0] - reference_date).days
            days_after_last_villa = (reference_date - last_villas[0][0]).days
            time_resource = [{'Статус': status, 'Дата последнего выезда:': last_villas[0][0].strftime('%d.%m.%y'), 'Гость:': last_villas[0][1]}, 
                             {'Планируемые заезды': list(map(lambda x: x[1], next_arrival))}, 
                             {'Дата ближайшего заезда': next_arrival[0][0].strftime('%d.%m.%y'), 'Дней до заезда:': days_before_villa, 'Гость:': next_arrival[0][1]}]
            
        elif status == 'С виллой' and planned_villas > 0:
            days_before_villa = (next_arrival[0][0] - check_out_date).days
            time_resource = [{'Статус': status, 'Выезд:': check_out_date.strftime('%d.%m.%y'), 'Гость:': busy}, 
                             {'Планируемые заезды': list(map(lambda x: x[1], next_arrival))}, 
                             {'Дата ближайшего заезда': next_arrival[0][0].strftime('%d.%m.%y'), 'Дней до заезда:': days_before_villa, 'Гость:': next_arrival[0][1]}]
            
        elif status == 'Свободен' and planned_villas == 0:
            days_after_last_villa = (reference_date - last_villas[0][0]).days
            time_resource =[{'Статус': status, 'Дата последнего выезда:': last_villas[0][0].strftime('%d.%m.%y'), 'Гость:': last_villas[0][1]}, 
                            {'Планируемые заезды': '-'},
                            {'Дата ближайшего заезда': '-', 'Дней до заезда:': '-', 'Гость:': '-'}]
            
        elif status == 'С виллой' and planned_villas == 0:
            time_resource = [{'Статус': status, 'Выезд:': check_out_date.strftime('%d.%m.%y'), 'Гость:': busy}, 
                             {'Планируемые заезды': '-'},
                             {'Дата ближайшего заезда': '-', 'Дней до заезда:': '-', 'Гость:': '-'}]
            

        general_dictionary[butler].append([{'Всего:': total - planned_villas}])
        general_dictionary[butler].append([{'VEG': veg}, {'FWV': fwv}, {'PWV': pwv}, {'VPS': vps}, {'VIG': vig}])
        general_dictionary[butler].append([{'Открытый рынок': open_market}, {'Cбер': sber}, {'Апгрейд': upgrade}, 
                                           {'Комплиментарный тариф': complimentary}])
        general_dictionary[butler].append([{'Один': alone}, {'2/2': twice}])
        general_dictionary[butler].append(time_resource)


    
             
    return general_dictionary



# Функция подсчета итогов адапитрованная под новый словарь
def new_count_totals(butlers):
    first_shift = ['Булгаков', 'Волгузов', 'Волков', 'Гетало', 'Гончар', 'Дембицкий', 'Диденко', 'Ларионов', 'Онищук', 'Орлов', 
    'Ляшов', 'Сергеев', 'Тараев']

    second_shift = ['Абдюшев', 'Люфт', 'Макухин', 'Мартынов', 'Марченко', 'Нечипуренко', 'Старенький', 'Стибельский', 'Тарабанов', 
    'Федоренко', 'Черноштан', 'Шаповалов']

    
    general_dictionary = {'first_shift': {},
                          'second_shift': {}}

    reference_date = datetime.today()
    
    #reference_date = datetime(2023, 12, 5)
    for butler, villas in butlers.items():
        if butler in first_shift:
            shift = 'first_shift'
        else:
            shift = 'second_shift'
        
        general_dictionary[shift].setdefault(butler, [])

        veg, vps, vig, fwv, pwv = 0, 0, 0, 0, 0
        open_market, sber, complimentary, upgrade = 0, 0, 0, 0
        alone, twice = 0, 0
        total = 0
        planned_villas = 0
        status = 'Свободен'
        days_before_villa = 0
        dates_arrival = []
        last_check_out = []
        check_out_date = 0
        time_resource = []

        for time, guests in villas.items():
            for guest, info in guests.items():
                for iterance in info:
                    if time == 'present':
                        veg += iterance.count('VEG')
                        vps += iterance.count('VPS')
                        vig += iterance.count('VIG')
                        fwv += iterance.count('FWV')
                        pwv += iterance.count('PWV')

                        open_market += iterance.count('Открытый рынок')
                        sber += iterance.count('Сбер')
                        complimentary += iterance.count('Комплиментарный тариф')
                        upgrade += iterance.count('Апгрейд')

                        alone += iterance.count('Один')
                        twice += iterance.count('2/2')

                        arrival = iterance[3][0][0]
                        check_out = iterance[3][0][1]

                        if check_out < reference_date:
                            last_check_out.append((check_out, guest))

                        elif arrival < reference_date < check_out:
                            check_out_date = check_out
                            status = 'С виллой'
                            busy = guest

                        total += 1

                    elif time == 'future':
                        arrival = iterance[3][0][0]
                        dates_arrival.append((arrival, guest))
                        planned_villas += 1        

            last_villas = sorted(last_check_out, reverse=True)
            next_arrival = sorted(dates_arrival)
        
        if status == 'Свободен' and planned_villas > 0:
            days_before_villa = (next_arrival[0][0] - reference_date).days
            days_after_last_villa = (reference_date - last_villas[0][0]).days
            time_resource = [{'Статус': status, 'Дата последнего выезда:': last_villas[0][0].strftime('%d.%m.%y'), 'Гость:': last_villas[0][1]}, 
                            {'Планируемые заезды': list(map(lambda x: x[1], next_arrival))}, 
                            {'Дата ближайшего заезда': next_arrival[0][0].strftime('%d.%m.%y'), 'Дней до заезда:': days_before_villa, 'Гость:': next_arrival[0][1]}]
            
        elif status == 'С виллой' and planned_villas > 0:
            days_before_villa = (next_arrival[0][0] - check_out_date).days
            time_resource = [{'Статус': status, 'Выезд:': check_out_date.strftime('%d.%m.%y'), 'Гость:': busy}, 
                            {'Планируемые заезды': list(map(lambda x: x[1], next_arrival))}, 
                            {'Дата ближайшего заезда': next_arrival[0][0].strftime('%d.%m.%y'), 'Дней до заезда:': days_before_villa, 'Гость:': next_arrival[0][1]}]
            
        elif status == 'Свободен' and planned_villas == 0:
            days_after_last_villa = (reference_date - last_villas[0][0]).days
            time_resource =[{'Статус': status, 'Дата последнего выезда:': last_villas[0][0].strftime('%d.%m.%y'), 'Гость:': last_villas[0][1]}, 
                            {'Планируемые заезды': '-'},
                            {'Дата ближайшего заезда': '-', 'Дней до заезда:': '-', 'Гость:': '-'}]
            
        elif status == 'С виллой' and planned_villas == 0:
            time_resource = [{'Статус': status, 'Выезд:': check_out_date.strftime('%d.%m.%y'), 'Гость:': busy}, 
                            {'Планируемые заезды': '-'},
                            {'Дата ближайшего заезда': '-', 'Дней до заезда:': '-', 'Гость:': '-'}]
            

        general_dictionary[shift][butler].append([{'Всего:': total - planned_villas}])
        general_dictionary[shift][butler].append([{'VEG': veg}, {'FWV': fwv}, {'PWV': pwv}, {'VPS': vps}, {'VIG': vig}])
        general_dictionary[shift][butler].append([{'Открытый рынок': open_market}, {'Cбер': sber}, {'Апгрейд': upgrade}, 
                                        {'Комплиментарный тариф': complimentary}])
        general_dictionary[shift][butler].append([{'Один': alone}, {'2/2': twice}])
        general_dictionary[shift][butler].append(time_resource)


    general_dictionary['first_shift'] = dict(sorted(general_dictionary['first_shift'].items()))
    general_dictionary['second_shift'] = dict(sorted(general_dictionary['second_shift'].items()))
             
    return general_dictionary