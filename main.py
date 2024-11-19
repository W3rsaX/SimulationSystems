from tkinter import *
import random
import datetime as dt
import pandas as pd
import openpyxl
from openpyxl.chart import Reference, LineChart
from openpyxl.styles import Alignment, Font

class day_class:  # Хранение данных

    def __init__(self, name):
        self.name = name
        self.status_truck = []
        self.status_crane = []
        self.status_crane2 = []
        self.status_mechanics = []
        self.time_status = []

truck_time = 4  # Мат. ожидание рабочего времени самосвала в часах
crane_time = 6  # Мат. ожидание рабочего времени крана в часах
crane2_time = 6  # Мат. ожидание рабочего времени крана 2 в часах
truck_repair1 = 1.5  # Мат. ожидание продолжительности ремонта в часах самосвала
crane_repair1 = 2.5  # Мат. ожидание продолжительности ремонта в часах крана
crane2_repair1 = 2.5  # Мат. ожидание продолжительности ремонта в часах крана 2
truck_repair2 = 0.5  # Мат. ожидание продолжительности ремонта в часах самосвала
crane_repair2 = 1.5  # Мат. ожидание продолжительности ремонта в часах крана
crane2_repair2 = 1.5  # Мат. ожидание продолжительности ремонта в часах крана 2
truck_loss = 500  # Убытки от простоя самосвала в рублях в час
crane_loss = 300  # Убытки от простоя крана в рублях в час
crane2_loss = 300  # Убытки от простоя крана 2 в рублях в час
truck_gain = 500  # Доход от работы самосвала в рублях в час
crane_gain = 300  # Доход от работы крана в рублях в час
crane2_gain = 300  # Доход от работы крана 2 в рублях в час
salary_6 = 100  # Зарплата слесаря 6-го разряда в рублях в час
salary_3 = 60  # Зарплата слесаря 3-го разряда в рублях в час
overhead = 50  # Накладные расхoоды на бригаду в рублях в час
k = 0  # Счетчик дней
count_days = 999  # Кол-во дней моделирования
num_day = 1  # Выбор дня для создания EXCEL
t = num_day - 1  # Пременная для задания дня файла EXCEL
days = [day_class(f"День_{i + 1}") for i in range(count_days + 1)]  # Создание листа объектов класса day_class
truck_work_time = 0  # Подсчет времени работы самосвала мин.
crane_work_time = 0  # Подсчет времени работы крана мин.
crane2_work_time = 0  # Подсчет времени работы крана 2 мин.
truck_afk_time = 0  # Подсчет времени простоя самосвала мин.
crane_afk_time = 0  # Подсчет времени простоя крана мин.
crane2_afk_time = 0  # Подсчет времени простоя крана 2 мин.
mechanics_work_time = 0  # Подсчет времени работы слесарей
Flag = 0  # Вариант решения 1 - 1 Слесарь, 2 - Слесаря


def export():
    global num_day, t, spinbox1
    num_day = int(spinbox.get())
    t = num_day - 1
    excel()


def sles1():
    global k, truck_work_time, crane_work_time, crane2_work_time, truck_afk_time, crane_afk_time, crane2_afk_time, mechanics_work_time, Flag

    Flag = 1
    k = 0  # Счетчик дней
    truck_work_time = 0  # Подсчет времени работы самосвала мин.
    crane_work_time = 0  # Подсчет времени работы крана мин.
    crane2_work_time = 0  # Подсчет времени работы крана 2 мин.
    truck_afk_time = 0  # Подсчет времени простоя самосвала мин.
    crane_afk_time = 0  # Подсчет времени простоя крана мин.
    crane2_afk_time = 0  # Подсчет времени простоя крана мин.
    mechanics_work_time = 0

    start_time = dt.datetime(2023, 1, 1, 8, 0, 0)  # Первый день
    end_time = start_time + dt.timedelta(days=count_days)  # Последний день

    simulate(start_time, end_time)
    # Расчеты показателей
    truck_profit = (truck_work_time / 60) * truck_gain
    truck_ubytki = (truck_afk_time / 60) * truck_loss
    truck_net_profit = truck_profit - truck_ubytki
    crane_profit = (crane_work_time / 60) * crane_gain
    crane_ubytki = (crane_afk_time / 60) * crane_loss
    crane_net_profit = crane_profit - crane_ubytki
    crane2_profit = (crane2_work_time / 60) * crane2_gain
    crane2_ubytki = (crane2_afk_time / 60) * crane2_loss
    crane2_net_profit = crane2_profit - crane2_ubytki
    mechanics_zp = (mechanics_work_time / 60) * (salary_6 + overhead)

    Label11 = Label(text=round(truck_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=0)
    Label12 = Label(text=round(truck_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=1)
    Label13 = Label(text=round(truck_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                             row=2)
    Label14 = Label(text=round(crane_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                         row=0)
    Label15 = Label(text=round(crane_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                         row=1)
    Label16 = Label(text=round(crane_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                             row=2)
    Label17 = Label(text=round(crane2_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                          row=0)
    Label18 = Label(text=round(crane2_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                          row=1)
    Label19 = Label(text=round(crane2_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                              row=2)
    Label20 = Label(text=round(mechanics_zp, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=3)
    Label21 = Label(text=round(truck_net_profit + crane_net_profit + crane2_net_profit - mechanics_zp, 2),
                    font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4, row=3)
    Button(root, text='Экспорт в EXCEL', font=('arial', 11, 'bold'), bg='#c9c945', state=["normal"],
           command=export).grid(row=7, column=3, columnspan=2)


def sles2():
    global k, truck_work_time, crane_work_time, crane2_work_time, truck_afk_time, crane_afk_time, crane2_afk_time, mechanics_work_time, Flag
    Flag = 2
    k = 0  # Счетчик дней
    truck_work_time = 0  # Подсчет времени работы самосвала мин.
    crane_work_time = 0  # Подсчет времени работы крана мин.
    crane2_work_time = 0
    truck_afk_time = 0  # Подсчет времени простоя самосвала мин.
    crane_afk_time = 0  # Подсчет времени простоя крана мин.
    crane2_afk_time = 0
    mechanics_work_time = 0

    start_time = dt.datetime(2023, 1, 1, 8, 0, 0)  # Первый день
    end_time = start_time + dt.timedelta(days=count_days)  # Последний день

    simulate(start_time, end_time)
    # Расчеты показателей
    truck_profit = round(truck_work_time / 60, 2) * truck_gain
    truck_ubytki = round(truck_afk_time / 60, 2) * truck_loss
    truck_net_profit = truck_profit - truck_ubytki
    crane_profit = round(crane_work_time / 60, 2) * crane_gain
    crane_ubytki = round(crane_afk_time / 60, 2) * crane_loss
    crane_net_profit = crane_profit - crane_ubytki
    crane2_profit = round(crane2_work_time / 60, 2) * crane2_gain
    crane2_ubytki = round(crane2_afk_time / 60, 2) * crane2_loss
    crane2_net_profit = crane2_profit - crane2_ubytki
    mechanics_zp = round(mechanics_work_time / 60, 2) * (salary_6 + salary_3 + overhead)

    Label11 = Label(text=round(truck_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=0)
    Label12 = Label(text=round(truck_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=1)
    Label13 = Label(text=round(truck_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                             row=2)
    Label14 = Label(text=round(crane_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                         row=0)
    Label15 = Label(text=round(crane_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                         row=1)
    Label16 = Label(text=round(crane_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4,
                                                                                                             row=2)
    Label17 = Label(text=round(crane2_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                          row=0)
    Label18 = Label(text=round(crane2_ubytki, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                          row=1)
    Label19 = Label(text=round(crane2_net_profit, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=6,
                                                                                                              row=2)
    Label20 = Label(text=round(mechanics_zp, 2), font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=1,
                                                                                                         row=3)
    Label21 = Label(text=round(truck_net_profit + crane_net_profit + crane2_net_profit - mechanics_zp, 2),
                    font=('arial', 11, 'bold'), width=8, bg='#e0e094').grid(column=4, row=3)
    Button(root, text='Экспорт в EXCEL', font=('arial', 11, 'bold'), bg='#c9c945', state=["normal"],
           command=export).grid(row=7, column=3, columnspan=2)


def day(start):  # Расчет одного дня
    global k, days
    global truck_work_time, truck_afk_time, mechanics_work_time, crane_work_time, crane_afk_time, crane2_work_time, crane2_afk_time
    seconds = 0
    end = start + dt.timedelta(days=1)
    truck_status = 1  # 1 - работает 2 - ремонт 3 - работы нет 4 - профилактика
    crane_status = 1  # 1 - работает 2 - ремонт 3 - работы нет 4 - профилактика
    crane2_status = 1
    mechanics_status = 2  # 1 - работает 2 - отдыхает
    truck_next_time = 0
    crane_next_time = 0
    crane2_next_time = 0
    shift = 1  # Смена 1 - работ 2 - не работа
    days[k].time_status.clear()
    days[k].status_truck.clear()
    days[k].status_crane.clear()
    days[k].status_crane2.clear()
    days[k].status_mechanics.clear()

    while (start != end):
        if (shift == 1):  # Рабочая смена
            if (truck_status == 1):  # Если работает самосвал
                truck_work_time = truck_work_time + 1

            if (crane_status == 1):  # Если работает кран
                crane_work_time = crane_work_time + 1

            if (crane2_status == 1):  # Если работает кран 2
                crane2_work_time = crane2_work_time + 1

            if (truck_status == 2) or (truck_status == 3):  # Если простаивает самосвал
                truck_afk_time = truck_afk_time + 1

            if (crane_status == 2) or (crane_status == 3):  # Если простаивает кран
                crane_afk_time = crane_afk_time + 1

            if (crane2_status == 2) or (crane2_status == 3):  # Если простаивает кран 2
                crane2_afk_time = crane2_afk_time + 1

            if (mechanics_status == 1):  # Если работает бригада
                mechanics_work_time = mechanics_work_time + 1

            if (truck_next_time == 0):  # Проверка на следующее событие самосвал
                if (truck_status == 1):
                    truck_next_time = start + dt.timedelta(hours=random.expovariate(1 / truck_time))  # Время работы

                elif (truck_status == 2) and (Flag == 1):
                    truck_next_time = start + dt.timedelta(hours=random.expovariate(1 / truck_repair1))  # Время ремонта

                elif (truck_status == 2) and (Flag == 2):
                    truck_next_time = start + dt.timedelta(hours=random.expovariate(1 / truck_repair2))  # Время ремонта

            if (crane_next_time == 0):  # Проверка на следующее событие кран
                if (crane_status == 1):
                    crane_next_time = start + dt.timedelta(hours=random.expovariate(1 / crane_time))  # Время работы

                elif (crane_status == 2) and (Flag == 1):
                    crane_next_time = start + dt.timedelta(hours=random.expovariate(1 / crane_repair1))  # Время ремонта

                elif (crane_status == 2) and (Flag == 2):
                    crane_next_time = start + dt.timedelta(hours=random.expovariate(1 / crane_repair2))  # Время ремонта

            if (crane2_next_time == 0):  # Проверка на следующее событие кран 2
                if (crane2_status == 1):
                    crane2_next_time = start + dt.timedelta(hours=random.expovariate(1 / crane2_time))  # Время работы

                elif ((crane2_status == 2) or (crane2_status == 3)) and (Flag == 1):
                    crane2_next_time = start + dt.timedelta(
                        hours=random.expovariate(1 / crane2_repair1))  # Время ремонта

                elif ((crane2_status == 2) or (crane2_status == 3)) and (Flag == 2):
                    crane2_next_time = start + dt.timedelta(
                        hours=random.expovariate(1 / crane2_repair2))  # Время ремонта

            if (truck_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) >= truck_next_time):  # Время подошло
                    if ((truck_status == 1) or (truck_status == 3)) and (
                            mechanics_status == 2):  # Работал или простаивает и бригада отдыхает
                        truck_status = 2  # Ремонт
                        truck_next_time = 0
                        mechanics_status = 1

                    elif ((truck_status == 1) or (truck_status == 3)) and (
                            mechanics_status == 1):  # Работал или простаивает и бригада отдыхает
                        truck_status = 2  # Ремонт
                        truck_next_time = 0
                        mechanics_status = 1
                        crane_status = 3

                    elif (truck_status == 2):
                        truck_status = 1  # Работа
                        truck_next_time = 0
                        mechanics_status = 2

                    elif (truck_status == 1) and (mechanics_status == 1):
                        truck_status = 3  # Простой

            if (crane_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) <= crane_next_time):
                    if crane_status == 3 and mechanics_status == 2:  # Работал или простаивает и бригада отдыхает
                        crane_status = 2  # Ремонт
                        crane_next_time = 0
                        mechanics_status = 1

                elif ((start + dt.timedelta(minutes=1)) >= crane_next_time):  # Время подошло
                    if ((crane_status == 1) or (crane_status == 3)) and (
                            mechanics_status == 2):  # Работал или простаивает и бригада отдыхает
                        crane_status = 2  # Ремонт
                        crane_next_time = 0
                        mechanics_status = 1

                    elif (crane_status == 2):
                        crane_status = 1  # Работа
                        crane_next_time = 0
                        mechanics_status = 2

                    elif (crane_status == 1) and (mechanics_status == 1):
                        crane_status = 3  # Простой

            if (crane2_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) <= crane2_next_time):
                    if crane2_status == 3 and mechanics_status == 2:  # Работал или простаивает и бригада отдыхает
                        crane2_status = 2  # Ремонт
                        crane2_next_time = 0
                        mechanics_status = 1

                elif ((start + dt.timedelta(minutes=1)) >= crane2_next_time):  # Время подошло
                    if ((crane2_status == 1) or (crane2_status == 3)) and (
                            mechanics_status == 2):  # Работал или простаивает и бригада отдыхает
                        crane2_status = 2  # Ремонт
                        crane2_next_time = 0
                        mechanics_status = 1

                    elif (crane2_status == 2):
                        crane2_status = 1  # Работа
                        crane2_next_time = 0
                        mechanics_status = 2

                    elif (crane2_status == 1) and (mechanics_status == 1):
                        crane2_status = 3  # Простой

        elif (shift == 2):
            if (mechanics_status == 1):  # Если работает бригада
                mechanics_work_time = mechanics_work_time + 1

            if (truck_status == 1):  # В работе
                truck_status = 4  # Ремонт
                truck_next_time = 0

            if (crane_status == 1):  # В работе
                crane_status = 4  # Ремонт
                crane_next_time = 0

            if (crane2_status == 1):  # В работе
                crane2_status = 4  # Ремонт
                crane2_next_time = 0

            if (truck_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) >= truck_next_time):  # Время подошло
                    if (truck_status == 1):  # В работе
                        truck_status = 4  # Профилактика
                        truck_next_time = 0

                    elif (truck_status == 2):  # В ремонте
                        truck_status = 4  # Профилактика
                        truck_next_time = 0
                        mechanics_status = 2

                    elif (truck_status == 3) and (mechanics_status == 2):
                        truck_status = 2  # Ремонт
                        truck_next_time = 0
                        mechanics_status = 1

            if (crane_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) >= crane_next_time):  # Время подошло
                    if (crane_status == 1):  # В работе
                        crane_status = 4  # Ремонт
                        crane_next_time = 0

                    elif (crane_status == 2):  # В ремонте
                        crane_status = 4  # Ремонт
                        crane_next_time = 0
                        mechanics_status = 2

                    elif (crane_status == 3) and (mechanics_status == 2):
                        crane_status = 2  # Ремонт
                        crane_next_time = 0
                        mechanics_status = 1

            if (crane2_next_time != 0):
                if ((start + dt.timedelta(minutes=1)) >= crane2_next_time):  # Время подошло
                    if (crane2_status == 1):  # В работе
                        crane2_status = 4  # Ремонт
                        crane2_next_time = 0

                    elif (crane2_status == 2):  # В ремонте
                        crane2_status = 4  # Ремонт
                        crane2_next_time = 0
                        mechanics_status = 2

                    elif (crane2_status == 3) and (mechanics_status == 2):
                        crane2_status = 2  # Ремонт
                        crane2_next_time = 0
                        mechanics_status = 1

        if (start.hour == 23) and (start.minute == 59):
            shift = 2
        days[k].time_status.insert(seconds, start.time())
        days[k].status_truck.insert(seconds, truck_status)
        days[k].status_crane.insert(seconds, crane_status)
        days[k].status_crane2.insert(seconds, crane2_status)
        days[k].status_mechanics.insert(seconds, mechanics_status)
        start = start + dt.timedelta(minutes=1)
        seconds = seconds + 1


def simulate(start_date, end_date):  # Моделирование всех дней
    while (start_date <= end_date):
        global k, day_id, days
        day(start_date)
        start_date = start_date + dt.timedelta(days=1)
        k = k + 1


def excel():  # Создание файла EXCEL
    exel = pd.DataFrame(
        {
            "Время": days[t].time_status,
            "Статус самосвала": days[t].status_truck,
            "Статус крана": days[t].status_crane,
            "Статус крана 2": days[t].status_crane2,
            "Статус бригады": days[t].status_mechanics
        }
    )
    if Flag == 1:
        exel.to_excel(days[t].name + "_1слесарь.xlsx", sheet_name="day", index=False)
        book = openpyxl.load_workbook(days[t].name + "_1слесарь.xlsx")
    elif Flag == 2:
        exel.to_excel(days[t].name + "_2слесаря.xlsx", sheet_name="day", index=False)
        book = openpyxl.load_workbook(days[t].name + "_2слесарь.xlsx")
    sheet = book.active
    sheet.column_dimensions['A'].width = 8
    sheet.column_dimensions['B'].width = 17
    sheet.column_dimensions['C'].width = 12
    sheet.column_dimensions['D'].width = 15

    sheet.column_dimensions['A'].width = 8
    sheet.column_dimensions['B'].width = 17
    sheet.column_dimensions['C'].width = 12
    sheet.column_dimensions['D'].width = 15
    sheet.column_dimensions['F'].width = 10
    sheet.column_dimensions['G'].width = 47

    sheet['G1'] = "Статусы"
    sheet['G1'].alignment = Alignment(horizontal='center')
    sheet['G1'].font = Font(bold=True)
    sheet['F2'] = "Машины"
    sheet['F2'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['F2'].font = Font(bold=True)
    if Flag == 1:
        sheet['F3'] = "Слесарь"
    elif Flag == 2:
        sheet['F3'] = "Слесари"
    sheet['F3'].alignment = Alignment(horizontal='center', vertical='center')
    sheet['F3'].font = Font(bold=True)

    sheet['G2'] = "1-Работает, 2-Ремонт, 3-Простой, 4-Профилактика"
    sheet['G2'].alignment = Alignment(horizontal='left')
    sheet['G3'] = "1-Работает, 2-Отдыхает"
    sheet['G3'].alignment = Alignment(horizontal='left')

    chart1 = LineChart()  # График самосвала
    chart1.width = 1000
    chart1.height = 8
    chart1.anchor = "I1"
    chart1.title = "График состояний самосвала"
    chart1.y_axis.title = "Состояние самосвала"
    chart1.y_axis.majorUnit = 1
    chart1.x_axis.title = "Время"
    data1 = Reference(sheet, min_col=2, max_col=2, min_row=2, max_row=1441)
    chart1.add_data(data1)

    chart1.series[0].graphicalProperties.line.solidFill = "FFC0CB"

    dates = Reference(sheet, min_col=1, min_row=2, max_row=1441)
    chart1.set_categories(dates)

    chart2 = LineChart()  # График крана
    chart2.width = 1000
    chart2.height = 8
    chart2.anchor = "I16"
    chart2.title = "График состояний крана"
    chart2.y_axis.title = "Состояние крана"
    chart2.y_axis.majorUnit = 1
    chart2.x_axis.title = "Время"
    data2 = Reference(sheet, min_col=3, max_col=3, min_row=2, max_row=1441)
    chart2.add_data(data2)
    dates = Reference(sheet, min_col=1, min_row=2, max_row=1441)
    chart2.set_categories(dates)

    chart3 = LineChart()  # График крана 2
    chart3.width = 1000
    chart3.height = 8
    chart3.anchor = "I31"
    chart3.title = "График состояний крана 2"
    chart3.y_axis.title = "Состояние крана 2"
    chart3.y_axis.majorUnit = 1
    chart3.x_axis.title = "Время"
    data3 = Reference(sheet, min_col=3, max_col=3, min_row=2, max_row=1441)
    chart3.add_data(data3)
    dates = Reference(sheet, min_col=1, min_row=2, max_row=1441)
    chart3.set_categories(dates)

    chart4 = LineChart()  # График слесарей
    chart4.width = 1000
    chart4.height = 8
    chart4.anchor = "I45"
    chart4.title = "График состояний слесаря"
    chart4.y_axis.title = "Состояние слесаря"
    chart4.y_axis.majorUnit = 1
    chart4.x_axis.title = "Время"
    data4 = Reference(sheet, min_col=4, max_col=4, min_row=2, max_row=1441)
    chart4.add_data(data4)
    dates = Reference(sheet, min_col=1, min_row=2, max_row=1441)
    chart4.set_categories(dates)

    sheet.add_chart(chart1)
    sheet.add_chart(chart2)
    sheet.add_chart(chart3)
    sheet.add_chart(chart4)
    if Flag == 1:
        book.save(days[t].name + "_1слесарь.xlsx")
    elif Flag == 2:
        book.save(days[t].name + "_2слесаря.xlsx")


root = Tk()

root.geometry('1100x220')
root.resizable(width=False, height=False)
root.configure(bg='#e0e094')
root.title('Системотехника строительства')

var = IntVar()
var.set(0)
btn_var = BooleanVar()

Label_space0 = Label(text=" ", width=8, font=('arial', 11, 'bold'), bg='#e0e094').grid(column=2, row=0)
Label_space00 = Label(text=" ", width=10, font=('arial', 11, 'bold'), bg='#e0e094').grid(column=1, row=0)
Label_space000 = Label(text=" ", width=10, font=('arial', 11, 'bold'), bg='#e0e094').grid(column=4, row=0)

Label1 = Label(text="Самосвал прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=0, row=0)
Label2 = Label(text="Самосвал убытки, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=0, row=1)
Label3 = Label(text="Самосвал чистая прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=0, row=2)
Label4 = Label(text="Кран прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(column=3,
                                                                                                             row=0)
Label5 = Label(text="Кран убытки, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(column=3,
                                                                                                            row=1)
Label6 = Label(text="Кран чистая прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=3, row=2)
Label7 = Label(text="Кран 2 прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(column=5,
                                                                                                               row=0)
Label8 = Label(text="Кран 2 убытки, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(column=5,
                                                                                                              row=1)
Label9 = Label(text="Кран 2 чистая прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=5, row=2)
Label10 = Label(text="Слесари затраты, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=0, row=3)
Label11 = Label(text="Общая чистая прибыль, р:", width=24, font=('arial', 11, 'bold'), anchor='w', bg='#e0e094').grid(
    column=3, row=3)

Label_space1 = Label(text="  ", width=8, bg='#e0e094').grid(column=2, row=4)

Button(root, text='1 слесарь', font=('arial', 11, 'bold'), bg='#c9c945', command=sles1).grid(row=5, column=0,
                                                                                             columnspan=3)
Button(root, text='2 слесаря', font=('arial', 11, 'bold'), bg='#c9c945', command=sles2).grid(row=5, column=2,
                                                                                             columnspan=3)

Label_space2 = Label(text="    ", width=10, font=('arial', 11, 'bold'), bg='#e0e094').grid(column=0, row=6)

Label28 = Label(text="Выберите день для вывода", width=30, font=('arial', 11, 'bold'), bg='#e0e094').grid(column=0,
                                                                                                          row=7,
                                                                                                          columnspan=2)

spinbox = Spinbox(from_=1.0, to=count_days + 1, width=4, font=('arial', 11, 'bold'), wrap="true")
spinbox.grid(column=2, row=7)

Button(root, text='Экспорт в EXCEL', font=('arial', 11, 'bold'), bg='#c9c945', state=["disable"]).grid(row=7, column=3,
                                                                                                       columnspan=2)

root.mainloop()
