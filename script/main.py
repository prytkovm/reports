# -*- coding: utf-8 -*-
import openpyxl as excel
from datetime import date
from datetime import timedelta
import os
import random
import shutil


def put_base(file, exercises, flag, date_time, name, teacher_name, us_section, student_num, group, day_num):
    current = -1
    file['H4'] = us_section
    file['K4'] = teacher_name
    file['C4'] = name
    file['B4'] = student_num
    file['G4'] = group
    file[flag] = 'Разминка'
    if flag == 'B10':
        file['C8'] = date_time.strftime("%d.%m.20%y")
        file['F10'] = '15 минут'
        file['G10'] = 1
        file['B25'] = 'Растяжка'
        file['F25'] = '5 минут'
        file['G25'] = 1
        file['G7'] = day_num
        for i in range(11, 22):
            iterator = 0
            file['F' + str(i)] = '1 минута'
            file['G' + str(i)] = 5
            while current != 0:
                iterator += 1
                current = exercises['C' + str(iterator)].value
                if iterator >= 137:
                    iterator = 1
                    for j in range(1, 138):
                        exercises['C' + str(j)] = 0
                    break
            file['B' + str(i)] = exercises['B' + str(iterator)].value
            exercises['C' + str(iterator)] = 1
            current = -1
    else:
        file['I8'] = date_time.strftime("%d.%m.20%y")
        file['L10'] = '15 минут'
        file['M10'] = 1
        file['H25'] = 'Растяжка'
        file['L25'] = '5 минут'
        file['M25'] = 1
        file['M7'] = day_num
        for i in range(11, 22):
            iterator = 0
            file['L' + str(i)] = '1 минута'
            file['M' + str(i)] = 5
            while current != 0:
                iterator += 1
                if iterator >= 137:
                    iterator = 1
                    for j in range(1, 138):
                        exercises['C' + str(j+1)] = 0
                    break
                current = exercises['C' + str(iterator)].value
            file['H' + str(i)] = exercises['B' + str(iterator)].value
            exercises['C' + str(iterator)] = 1
            current = -1


def put_with_ball(file, flag, exercises):
    current = -1
    if flag == 'B22':
        for i in range(22, 25):
            iterator = 0
            while current != 0:
                iterator += 1
                current = exercises['G' + str(iterator)].value
                if iterator >= 41:
                    iterator = 1
                    for j in range(1, 42):
                        exercises['G' + str(j)] = 0
                    break
            file['B' + str(i)] = exercises['F' + str(iterator)].value
            file['F' + str(i)] = exercises['D' + str(iterator)].value
            file['G' + str(i)] = exercises['E' + str(iterator)].value
            exercises['G' + str(iterator)] = 1
            current = -1
    else:
        for i in range(22, 25):
            iterator = 0
            while current != 0:
                iterator += 1
                current = exercises['G' + str(iterator)].value
                if iterator >= 41:
                    iterator = 1
                    for j in range(1, 42):
                        exercises['G' + str(j)] = 0
                    break
            file['H' + str(i)] = exercises['F' + str(iterator)].value
            file['L' + str(i)] = exercises['D' + str(iterator)].value
            file['M' + str(i)] = exercises['E' + str(iterator)].value
            exercises['G' + str(iterator)] = 1
            current = -1


def put_analise(file, flag):
    pulse = 65 + random.randint(3, 15)
    max_pulse = 150 + random.randint(4, 15)
    file[flag] = 'Самочувствие в норме, немного болят мышцы, утомился,' \
                 ' максимальный пульс - {}, пульс в покое - {}'.format(max_pulse, pulse)


if __name__ == '__main__':
    path = os.getcwd()
    try:
        os.mkdir(path + '\\Отчеты')
    except FileExistsError:
        shutil.rmtree(path + '\\Отчеты')
        os.mkdir(path + '\\Отчеты')
    else:
        print("Директория \"Отчеты\" успешно создана!")
    stud_num = int(input("Номер студ. билета (Без незначащего нуля, если он есть): "))
    student = input("ФИО: ")
    teacher = input("ФИО преподавателя: ")
    section = input("Название секции: ")
    stud_group = int(input("Группа: "))
    lambda_days = int(input("Промежуток между занятиями на неделе: "))
    between = int(input("Промежуток между крайним занятием на неделе и следующим(на след. неделе):"))
    start_day = int(input("Стартовое число занятий в семестре: "))
    counter = 1
    start_date = date(2021, 2, start_day)
    database = 'exercises.xlsx'
    wb2 = excel.load_workbook(database)
    ex = wb2.active
    template = excel.load_workbook('template.xlsx')
    workbook = template.active
    for days in range(17):  # 17
        os.chdir(path + '\\Отчеты')
        if start_date == date(2021, 4, 2):
            put_base(workbook, ex, 'B10', start_date, student, teacher, section, stud_num, stud_group, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%m-%d-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=2 * between)
        elif start_date == date(2021, 5, 9):
            put_base(workbook, ex, 'B10', start_date, student, teacher, section, stud_num, stud_group, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%m-%d-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=between)
        else:
            put_base(workbook, ex, 'B10', start_date, student, teacher, section, stud_num, stud_group, counter)
            counter += 1
            put_with_ball(workbook, 'B22', ex)
            put_analise(workbook, 'B29')
            start_date += timedelta(days=lambda_days)
            put_base(workbook, ex, 'H10', start_date, student, teacher, section, stud_num, stud_group, counter)
            counter += 1
            put_with_ball(workbook, 'H22', ex)
            put_analise(workbook, 'H29')
            finished_file = 'Отчет.ТК.ФВиС.' + str(stud_num) + ' ' + start_date.strftime("%m-%d-20%y") + '.xlsx'
            template.save(filename=finished_file)
            start_date += timedelta(days=between)
        os.chdir(path)
