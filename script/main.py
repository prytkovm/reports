# -*- coding: utf-8 -*-
import openpyxl as excel
from datetime import date, timedelta
import os
import random
import shutil


def put_base(file, exercises, flag, date_time, day_num):
    curr = -1
    file[flag] = 'Разминка'
    if flag == 'B10':
        flag = flag.replace('B10', 'C8')
    elif flag == 'H10':
        flag = flag.replace('H10', 'I8')
    else:
        exit(-1)
    file[flag] = date_time.strftime("%d.%m.20%y")
    code = ord(flag[0]) + ord('F') - ord('C')
    flag = flag.replace(flag, chr(code) + '10')
    file[flag] = '15 минут'
    code = ord(flag[0]) + 1
    flag = flag.replace(flag, chr(code) + '10')
    file[flag] = 1
    code = ord(flag[0]) - (ord('G') - ord('B'))
    flag = flag.replace(flag, chr(code) + '25')
    file[flag] = 'Растяжка'
    code = ord(flag[0]) + 4
    flag = flag.replace(flag, chr(code) + '25')
    file[flag] = '5 минут'
    code = ord(flag[0]) + 1
    flag = flag.replace(flag, chr(code) + '25')
    file[flag] = 1
    flag = flag.replace(flag, chr(code) + '7')
    file[flag] = day_num
    code = ord(flag[0]) - 1
    flag = flag.replace(flag, chr(code))
    iterator_header = flag[0]
    for i in range(11, 22):
        iterator_s = 0
        file[iterator_header + str(i)] = '1 минута'
        code = ord(flag[0])
        flag = flag.replace(flag, chr(code + 1) + str(i))
        file[flag] = 5
        while curr != 0:
            iterator_s += 1
            curr = exercises['C' + str(iterator_s)].value
            if iterator_s >= 137:
                iterator_s = 1
                for j in range(1, 138):
                    exercises['C' + str(j)] = 0
                break
        code = ord(flag[0]) - (ord('G') - ord('B'))
        flag = flag.replace(flag, chr(code) + str(i))
        exercises['C' + str(iterator_s)] = 1
        file[flag] = exercises['B' + str(iterator_s)].value
        code = ord(flag[0])
        flag = flag.replace(flag, chr(code + (ord('G') - ord('B')) - 1))
        curr = -1


def put_with_ball(file, flag, exercises):
    current = -1
    first = flag[0]
    for i in range(22, 25):
        iterator = 0
        while current != 0:
            iterator += 1
            current = exercises['G' + str(iterator)].value
            if iterator >= 69:
                iterator = 1
                for j in range(1, 69):
                    exercises['G' + str(j)] = 0
                break
        if first == 'B':
            file['B' + str(i)] = exercises['F' + str(iterator)].value
            file['F' + str(i)] = exercises['D' + str(iterator)].value
            file['G' + str(i)] = exercises['E' + str(iterator)].value
        else:
            file['H' + str(i)] = exercises['F' + str(iterator)].value
            file['L' + str(i)] = exercises['D' + str(iterator)].value
            file['M' + str(i)] = exercises['E' + str(iterator)].value
        exercises['G' + str(iterator)] = 1
        current = -1


def put_analise(file, flag):
    parameter = random.randint(1, 2)
    pulse = 65 + random.randint(3, 15)
    max_pulse = 150 + random.randint(4, 15)
    if parameter == 1:
        file[flag] = 'Самочувствие в норме, немного болят мышцы, утомился,' \
                 ' максимальный пульс - {}, пульс в покое - {}'.format(max_pulse, pulse)
    else:
        file[flag] = 'Самочувствие в норме, утомился, немного болят мышцы, ' \
                     ' пульс в покое - {}, максимальный пульс - {}'.format(pulse, max_pulse)


if __name__ == '__main__':
    path = os.getcwd()
    try:
        os.mkdir(path + '\\Отчеты')
    except FileExistsError:
        shutil.rmtree(path + '\\Отчеты')
        os.mkdir(path + '\\Отчеты')
    else:
        print("Директория \"Отчеты\" успешно создана!")
    counter = 1
    stud_num = int(input("Номер студ. билета: (Без незначащего нуля) "))
    student = input("ФИО: ")
    teacher = input("ФИО преподавателя: ")
    section = input("Название секции: ")
    stud_group = int(input("Группа: "))
    lambda_days = int(input("Промежуток между занятиями на неделе: "))  # Число крайнего на неделе минус число первого
    between = int(input("Промежуток между крайним занятием на неделе и следующим(на след. неделе):"))
    start = input("Стартовое число занятий в семестре(дд.мм.гг): ")
    start_day = int(start[0:2])
    start_month = int(start[3:5])
    start_year = int(start[6:10])
    start_date = date(start_year, start_month, start_day)
    days_count = int(input("Всего занятий в семестре (можно спросить у Арсения Чернова или поискать в беседе, он их за вас посчитал): "))
    database = 'exercises.xlsx'
    wb2 = excel.load_workbook(database)
    ex = wb2.active
    template = excel.load_workbook('template.xlsx')
    workbook = template.active
    for days in range(days_count // 2 + 1):
        os.chdir(path + '\\Отчеты')
        workbook['H4'] = section
        workbook['K4'] = teacher
        workbook['C4'] = student
        workbook['B4'] = stud_num
        workbook['G4'] = stud_group
        if start_date == date(2021, 4, 2):
            put_base(workbook, ex, 'B10', start_date, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%d-%m-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=2 * between)
        elif start_date == date(2021, 5, 9):
            put_base(workbook, ex, 'B10', start_date, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%d-%m-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=between)
        else:
            put_base(workbook, ex, 'B10', start_date, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            start_date += timedelta(days=lambda_days)
            put_base(workbook, ex, 'H10', start_date, counter)
            counter += 1
            put_with_ball(workbook, 'H22', ex)
            put_analise(workbook, 'H29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%d-%m-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=between)
        os.chdir(path)
