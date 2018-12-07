# -*- coding: utf-8 -*-
"""
asvt_fingerprint.excel

Created by prudkovskiy on 29.11.18 23:51
"""
from datetime import datetime, timedelta
import pandas as pd
import xlrd
import xlsxwriter
import calendar
from time import sleep
import re

__author__ = 'prudkovskiy'

week = {
    0: 'Пн',
    1: 'Вт',
    2: 'Ср',
    3: 'Чт',
    4: 'Пт',
    5: 'Сб',
    6: 'Вс'
}
alph = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
file_name = 'asvt.xlsx'


def make_start_excel(filename):
    """
    Создание шаблона таблицы
    :param filename:
    :return:
    """
    global week, alph

    sheet_name = datetime.now().strftime("%Y.%m")
    year_month = sheet_name.split('.')
    number_of_days = calendar.monthrange(int(year_month[0]), int(year_month[1]))[1]
    last_column_name = 'A' + alph[(number_of_days + 5) % 27]

    # Создаем новый excel файл и новый рабочий лист, где имя - это год и месяц
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)  # worksheet name 2018.11

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 20, bold)
    worksheet.set_column('C:E', 12)
    worksheet.set_column('F:{}'.format(last_column_name), 6)

    worksheet.write(0, 0, 'Номер')
    worksheet.write(0, 1, 'Фамилия')
    worksheet.write(0, 2, 'Должность')
    worksheet.write(0, 3, 'На работе')
    worksheet.write(0, 4, 'Время входа')

    for i in range(number_of_days):
        day_of_month = i + 1
        day_of_week = week[datetime.strptime('{}.{}'.format(sheet_name, day_of_month), '%Y.%m.%d').weekday()]
        worksheet.write(0, i + 5, '{}|{}'.format(day_of_week, day_of_month))

    worksheet.write(0, 5 + number_of_days, 'Отработано за месяц')
    worksheet.write(0, 6 + number_of_days, 'Осталось работать')
    worksheet.write(0, 7 + number_of_days, 'Перерасчет на каждый день')
    # worksheet.write(1, 0, 1)
    # worksheet.write(1, 1, 'Рафиков А.Г.')
    # worksheet.write(1, 2, 'Ведущий инженер')
    # worksheet.write(1, 3, 'Да')
    # worksheet.conditional_format('A1:B3'.format(last_column_name), {'type': '3_color_scale'})

    writer.save()


def create_new_sheet(filename):
    """
    Создает новый лист в excel для отчета в новый месяц
    :param filename:
    :return:
    """
    new_sheet_name = datetime.now().strftime("%Y.%m")
    year_month = new_sheet_name.split('.')
    number_of_days = calendar.monthrange(int(year_month[0]), int(year_month[1]))[1]

    # open the file for reading
    wbRD = xlrd.open_workbook(filename)
    sheets = wbRD.sheets()

    # open the same file for writing (just don't write yet)
    wb = xlsxwriter.Workbook(filename)

    # run through the sheets and store sheets in workbook
    # this still doesn't write to the file yet
    for sheet in sheets:  # write data from old file
        same_sheet = wb.add_worksheet(sheet.name)
        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                same_sheet.write(row, col, sheet.cell(row, col).value)

    new_sheet = wb.add_worksheet(new_sheet_name)

    prev_sheet = sheets.pop()
    for row in range(prev_sheet.nrows):
        for col in range(5):
            new_sheet.write(row, col, prev_sheet.cell(row, col).value)

    for i in range(number_of_days):
        day_of_month = i + 1
        day_of_week = week[datetime.strptime('{}.{}'.format(new_sheet_name, day_of_month), '%Y.%m.%d').weekday()]
        new_sheet.write(0, i + 5, '{}|{}'.format(day_of_week, day_of_month))

    new_sheet.write(0, 5 + number_of_days, 'Sum')

    wb.close()  # THIS writes


def make_writer(writer, sheet_name):
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
    # Widen the first column to make the text clearer.
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:B', 25, bold)
    worksheet.set_column('C:E', 12)

    year_month = sheet_name.split('.')
    number_of_days = calendar.monthrange(int(year_month[0]), int(year_month[1]))[1]
    last_column_name = 'A' + alph[(number_of_days + 5) % 27]
    worksheet.set_column('F:{}'.format(last_column_name), 18)
    last_column_name = 'A' + alph[(number_of_days + 5) % 27 + 1]
    worksheet.set_column('{}:{}'.format(last_column_name, last_column_name), 25)
    return


def create_new_employee(filename, name):
    """

    :param name: имя сотрудника
    :return:
    """
    global week, alph

    data = pd.read_excel(filename, sheet_name=None)
    sheet_name = datetime.now().strftime("%Y.%m")
    try:
        df = data[sheet_name]
    except KeyError:
        create_new_sheet(filename)
        create_new_employee(filename, name)
        return

    # s = df[df['Фамилия, инициалы'] == name]['Фамилия, инициалы']
    # if not s.get('Фамилия, инициалы'):
    new_id = len(df)+1
    val = [new_id, name, None, 'Да', datetime.now().strftime("%H:%M:%S")]

    for _ in df.columns[5:-3]:
        # val.append('0ч. 0мин. 0сек.')
        val.append('')

    val.append(0)  # sum

    df_add = pd.DataFrame([val], columns=df.columns)
    df = df.append(df_add)

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    for sheet, df in list(data.items())[:-1]:
        make_writer(writer, sheet)
        df.to_excel(writer, sheet_name=sheet, index=False)

    df.to_excel(writer, sheet_name=sheet_name, index=False)
    make_writer(writer, sheet_name)

    writer.save()
    return new_id


def hours_minutes_seconds(seconds):
    return seconds // 3600, (seconds // 60) % 60, seconds % 60


def get_sum(row):
    sheet_name = datetime.now().strftime("%Y.%m")
    year_month = sheet_name.split('.')
    number_of_days = calendar.monthrange(int(year_month[0]), int(year_month[1]))[1]
    # empty result timedelta
    td = timedelta()
    # сколько 8-часовых дней требовалось отработать в текущем месяце изначально
    days_cur = 0
    # сколько из них осталось отработать, включая текущий
    days_left = 0
    for i in range(number_of_days):
        day_of_month = i + 1
        day_of_week = week[datetime.strptime('{}.{}'.format(sheet_name, day_of_month), '%Y.%m.%d').weekday()]
        if day_of_week not in ('Сб', 'Вс'):
            days_cur += 1
            if day_of_month >= datetime.today().day:
                days_left += 1
        today = '{}|{}'.format(day_of_week, day_of_month)
        try:
            h, m, s = list(map(int, re.findall(r'\d+', row[today])))
        except Exception:
            h, m, s = 0, 0, 0
        td += timedelta(hours=h, minutes=m, seconds=s)
        
    # отработано
    d = td.days
    h, m, s = hours_minutes_seconds(td.seconds)
    current = '{}дн. {}ч. {}мин. {}сек.'.format(d, h, m, s)
    # осталось отработать
    time_left = timedelta(hours=days_cur*8) - td
    d = time_left.days
    h, m, s = hours_minutes_seconds(time_left.seconds)
    required = '{}дн. {}ч. {}мин. {}сек.'.format(d, h, m, s)
    # растолкать по оставшимся дням
    time_left_per_day = time_left / days_left
    h, m, s = hours_minutes_seconds(time_left_per_day.seconds)
    left = '{}ч. {}мин. {}сек.'.format(h, m, s)

    return pd.Series([current, required, left])


def enter_employee(filename, id):
    """

    :param filename:
    :param id:
    :return:
    """
    data = pd.read_excel(filename, sheet_name=None)
    sheet_name = datetime.now().strftime("%Y.%m")

    try:
        df = data[sheet_name]
    except KeyError:
        create_new_sheet(filename)
        enter_employee(filename, id)
        return

    if df[df['Номер'] == id].empty:
        return

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    for sheet, _df in list(data.items())[:-1]:
        make_writer(writer, sheet)
        _df.to_excel(writer, sheet_name=sheet, index=False)

    ser = df[df['Номер'] == id].iloc[0]

    if ser['На работе'] == 'Да':
        df.loc[id - 1, 'На работе'] = 'Нет'
        # ser['На работе'] = 'Нет'
        prev_time = datetime.strptime(ser['Время входа'], "%H:%M:%S")
        # ser['Время входа'] = None
        df.loc[id - 1, 'Время входа'] = None
        curr_time = datetime.now().replace(microsecond=0)
        td = curr_time - prev_time

        today = '{}|{}'.format(week[datetime.now().weekday()], datetime.now().day)
        try:
            h_old, m_old, s_old = list(map(int, re.findall(r'\d+', ser[today])))
        except Exception:
            h_old, m_old, s_old = 0, 0, 0

        td_old = timedelta(hours=h_old, minutes=m_old, seconds=s_old)
        td_res = td + td_old
        h_res, m_res, s_res = hours_minutes_seconds(td_res.seconds)
        df.loc[id - 1, today] = '{}ч. {}мин. {}сек.'.format(h_res, m_res, s_res)
        cols = ['Отработано за месяц', 'Осталось работать', 'Перерасчет на каждый день']
        df[cols] = df.apply(get_sum, axis=1)

    elif ser['На работе'] == 'Нет':
        df.loc[id - 1, 'На работе'] = 'Да'
        df.loc[id - 1, 'Время входа'] = datetime.now().strftime("%H:%M:%S")

    df.to_excel(writer, sheet_name=sheet_name, index=False)
    make_writer(writer, sheet_name)
    writer.close()
    return


if __name__ == "__main__":
    print('Run excel ({0})'.format(datetime.now()))
    make_start_excel(file_name)
