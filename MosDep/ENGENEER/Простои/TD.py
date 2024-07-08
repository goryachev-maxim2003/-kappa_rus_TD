# pyinstaller --onefile --windowed TD.py //Для того, чтобы сделать exe файл без консоли

import tkinter as tk
import xlwings as xw
import pandas as pd
import numpy as np

def middle_of_day(d):
    return pd.to_datetime(f'{d.year}-{d.month}-{d.day} 12:00:00')

def end_of_day(d):
    d += pd.Timedelta(days=1)
    return pd.to_datetime(f'{d.year}-{d.month}-{d.day} 00:00:00')

def end_of_shift(d):
    return end_of_day(d) if (d >= middle_of_day(d)) else middle_of_day(d)
def get_plan_time(machine, shift, day, month, year):
    date = pd.to_datetime(f'{year}-{month}-{day}')
    try:
        return plan_time_dict[machine][
            (plan_time_dict[machine]["Date"] == date) & (plan_time_dict[machine]["Shift"] == shift)][
            "working time-stops for maintenance"].iloc[0]
    except IndexError:
        return np.nan
def cell_to_tuple(s):
    numbers_list = ['0','1','2','3','4','5','6','7','8','9']
    letters = ''
    number = ''
    for i in s:
        if (i in numbers_list):
            number += i
        else:
            letters += i
    return (letters, int(number))

def convert_shift_to_number(s):
    if (s == for_TD.loc["первая смена", "Значение"]):
        return 1
    elif (s == for_TD.loc["вторая смена", "Значение"]):
        return 2
    else:
        raise ValueError("Изменено одно из значений смены (19-07 или 07-19)")


#Глобальные переменные
for_TD = None
path_to_KPI_file = None
path_to_plan_time_file = None
for_TD_machine_names = None
machine = None
machine_in_rus = None
machine_in_rus_keys = None
TableDowntime = None
KPI_book = None
KPI_sheet = None

day_grouped_Downtime = None
month_grouped_Downtime = None
plan_time_dict = None
grouped_Downtime = None
target_KPI = None


was_main_calc = False
was_open_all = False

def open_all():
    global for_TD
    global path_to_KPI_file
    global path_to_plan_time_file
    global for_TD_machine_names
    global machine
    global machine_in_rus
    global machine_in_rus_keys
    global TableDowntime
    global KPI_book
    global KPI_sheet
    was_open_all = True
    for_TD = pd.read_excel('Файл для TD exe.xlsx', 'Параметры', index_col=0, dtype = str)
    path_to_KPI_file = for_TD.loc["Путь к файлу KPI", "Значение"]
    path_to_plan_time_file = for_TD.loc["Путь к файлу Daily", "Значение"]
    for_TD_machine_names = pd.read_excel('Файл для TD exe.xlsx', 'Названия станков', dtype = str)

    #Словари для соответствия названий станков
    machine = dict((str(for_TD_machine_names["Названия станков на листе downtime from CPMS"].values[i]), str(for_TD_machine_names["Названия станков в файле Daily"].values[i])) for i in range(len(for_TD_machine_names)))
    machine_in_rus = dict((str(for_TD_machine_names["Названия станков в файле Daily"].values[i]), str(for_TD_machine_names["Названия станков на листе KPI"].values[i])) for i in range(len(for_TD_machine_names)))
    machine_in_rus_keys = list(machine_in_rus.keys())

    # Чтение файла KPI
    TableDowntime = pd.read_excel(path_to_KPI_file, sheet_name=for_TD.loc["Имя листа downtime from CPMS", "Значение"],
                    usecols=list(map(lambda x: x.strip(), for_TD.loc["Название колонок в файле KPI", "Значение"].split(','))))
    TableDowntime.columns = ['Станок', 'Код простоя', 'Месяц', 'Число', 'Начало простоя', 'Конец простоя','простой в мин.', 'Проблема', 'Узел']
    
    # Открытие файла для записи
    KPI_book = xw.Book(path_to_KPI_file)
    try:
        KPI_sheet = KPI_book.sheets[for_TD.loc["Имя листа KPI", "Значение"]]
    except:
        list_name = for_TD.loc["Имя листа KPI", "Значение"]
        raise FileNotFoundError(f"Не найден лист {list_name}") 

def main_calc(): #Основное вычисление. Вычисляет по сменам
    global was_main_calc
    global TableDowntime
    global plan_time_dict
    global grouped_Downtime
    global target_KPI
    was_main_calc = True
    # Меняем формат месяца из Apr в 04 и т.д. Цифру месяца берём из столбца месяц
    TableDowntime['Начало простоя'] = pd.to_datetime(TableDowntime['Начало простоя'])
    TableDowntime['Конец простоя'] = pd.to_datetime(TableDowntime['Конец простоя'])  # pandas понимает формат Apr, Mar и т.д.

    # Разделяем простои длинной в несколько дней на несколько записей или если простой затрагивает 2 дня
    # Пример 16.06 - период с 7-00 16.06.24 до 7-00 17.06.24
    # Временно смещаем время на 7 часов, чтобы корректно вычислять число и месяц в формате с 7 до 7
    TableDowntime['Начало простоя'] = pd.to_datetime(TableDowntime['Начало простоя'], format='%d %m %Y %H:%M:%S') - pd.Timedelta(hours=7)
    TableDowntime['Конец простоя'] = pd.to_datetime(TableDowntime['Конец простоя'], format='%d %m %Y %H:%M:%S') - pd.Timedelta(hours=7)

    # Делим запись с длинным простоем на записи с более короткими простоями по сменам (в днях не учитываем график с 7 до 7, т.к. даты временно смещены на 7 часов)
    for i in range(len(TableDowntime)):  # цикл не пробегает по созданным в цикле записям
        cur_i = i
        # Пока затрагивается новая смена отрезаем запись и проверяем новую запись в таблице
        while (end_of_shift(TableDowntime.loc[cur_i, 'Начало простоя']) != end_of_shift(
                TableDowntime.loc[cur_i, 'Конец простоя'])):
            new_end = end_of_shift(TableDowntime.loc[cur_i, 'Начало простоя'])
            new_series = TableDowntime.loc[cur_i].copy()
            new_series.loc['Начало простоя'] = new_end
            TableDowntime.loc[cur_i, 'Конец простоя'] = new_end
            cur_i = len(TableDowntime)  # задаём индекс новой записи (только что созданной)
            TableDowntime.loc[len(TableDowntime)] = new_series
    # Для новых записей также определяем число и месяц
    TableDowntime["Число"] = TableDowntime['Начало простоя'].dt.day
    TableDowntime["Месяц"] = TableDowntime['Начало простоя'].dt.month
    TableDowntime["Год"] = TableDowntime['Начало простоя'].dt.year
    TableDowntime['Смена'] = 1 + (TableDowntime['Начало простоя'] >= TableDowntime['Начало простоя'].apply(
        middle_of_day))  # чтобы можно было сортировать по смене

    # Смещаем обратно на 7 часов
    TableDowntime['Начало простоя'] = pd.to_datetime(TableDowntime['Начало простоя']) + pd.Timedelta(hours=7)
    TableDowntime['Конец простоя'] = pd.to_datetime(TableDowntime['Конец простоя']) + pd.Timedelta(hours=7)
    TableDowntime["простой в мин."] = ((TableDowntime['Конец простоя'] - TableDowntime[
        'Начало простоя']).dt.total_seconds() / 60).round().astype(int)
    TableDowntime["Станок"] = TableDowntime["Станок"].apply(lambda x: machine[x])
    #Берём время работы for maintenance
    plan_time_dict = pd.read_excel(path_to_plan_time_file,
        skiprows=[1,2,3,4],
        usecols = list(map(lambda x: x.strip(), for_TD.loc["Название колонок в Daily", "Значение"].split(','))),
        sheet_name=list(for_TD_machine_names['Названия станков в файле Daily']))
    for sheet_name in plan_time_dict.keys():
        # Переименовываем колонки
        plan_time_dict[sheet_name].columns = ['Date', 'Shift', 'working time-stops for maintenance']
        plan_time_dict[sheet_name] = plan_time_dict[sheet_name].dropna()
        #Изменяем 07-19, 19-07 на 1 и 2 смену
        plan_time_dict[sheet_name]["Shift"] = plan_time_dict[sheet_name]["Shift"].apply(convert_shift_to_number)
    # группировка по столбцам месяц, число и станок
    grouped_Downtime = TableDowntime.groupby(['Станок', 'Смена', 'Месяц', 'Число', 'Год']).agg({'простой в мин.': 'sum'}).reset_index()
    # сопоставляем время работы for maintenance с записями о застоях
    grouped_Downtime["Время работы for maintenance"] = grouped_Downtime.apply(lambda row: get_plan_time(row['Станок'], row['Смена'], row['Число'], row['Месяц'], row['Год']), axis=1)
    grouped_Downtime["TD"] = grouped_Downtime["простой в мин."] / grouped_Downtime["Время работы for maintenance"] / 60 * 100
    grouped_Downtime["Дата"] = grouped_Downtime.apply(lambda row: pd.to_datetime(f'{row["Год"]}-{row["Месяц"]}-{row["Число"]}'), axis=1)
    # Возвращаем смену в исходный формат
    grouped_Downtime["Смена"] = grouped_Downtime["Смена"].apply(lambda n: for_TD.loc["первая смена", "Значение"] if (n == 1) else for_TD.loc["вторая смена", "Значение"])
    grouped_Downtime = grouped_Downtime.sort_values(by='Дата', ascending=False)

    #Считываем целевой KPI
    target_KPI = pd.DataFrame(KPI_sheet.range(for_TD.loc["Диапазон ячеек с целевым KPI", "Значение"]).value, columns=['Целевой KPI'], index=KPI_sheet.range(for_TD.loc["Диапазон ячеек с целевым KPI", "Значение"]).offset(0, -1).value)

def is_in_target(row):
    return row['TD'] <= target_KPI.loc[machine_in_rus[row['Станок']]]

#cell_tuple - ячейка начала таблицы, shift сдвиг до столбца TD
def colorize(df, cell_tuple, shift):
    green = (226, 239, 218)
    red = (252, 228, 214)
    is_in_target_ser = df.apply(is_in_target, axis = 1)
    i = 0
    for c in KPI_sheet.range(f'{cell_tuple[0]+str(cell_tuple[1])}:{cell_tuple[0]+str(cell_tuple[1]+len(is_in_target_ser)-1)}').offset(1, shift):
        if (is_in_target_ser.iloc[i]["Целевой KPI"]):
            c.color = green
        else:
            c.color = red
        i+=1
def clear_range(ran):
    ran.value = None
    ran.color = None
def clear():
    clear_range(KPI_sheet.range(for_TD.loc["Ячейка для таблицы по сменам", "Значение"]).expand('table'))
    clear_range(KPI_sheet.range(for_TD.loc["Ячейка для таблицы по дням", "Значение"]).expand('table'))
    clear_range(KPI_sheet.range(for_TD.loc["Ячейка для таблицы по месяцам", "Значение"]).expand('table'))
    clear_range(KPI_sheet.range(for_TD.loc["Диапазон ячеек с целевым KPI", "Значение"]).offset(0,1))
def grouped_by_shift():
    #Записываем таблицу в файл
    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по сменам", "Значение"]).offset(0,-1).value = grouped_Downtime[
        ["Станок", "Смена", "Дата", "TD"]]
    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по сменам", "Значение"]).offset(1,-1).expand('down').value = None  # чистим индексы
def Today():
    # Выводим сегодняшнюю статистику
    today = pd.to_datetime(KPI_sheet.range(for_TD.loc["Ячейка для просмотра вчерашней даты", "Значение"]).value) - pd.Timedelta("1 day")
    day_grouped_Downtime = grouped_Downtime.groupby(['Станок', 'Дата']).agg(
        {"Время работы for maintenance": 'sum', "простой в мин.": "sum"}).reset_index()
    day_grouped_Downtime["TD"] = day_grouped_Downtime["простой в мин."] / (
                day_grouped_Downtime["Время работы for maintenance"] * 60) * 100
    today_TD = day_grouped_Downtime.loc[day_grouped_Downtime["Дата"] == today, ['Станок', 'TD']]
    today_TD.set_index('Станок', inplace=True)  # inplace - вместо возвращения нового датафрейма измениться текущий
    for machine_name in (
            set(machine_in_rus_keys) - set(today_TD.index)):  # Для машин по которым данных не было заполняем нулями
        today_TD.loc[machine_name] = 0
    #имена станков в том порядке в котором они написаны на листе ['Corrugator', 616, 924, ...]
    try:
        names_sorted_by_target_KPI = list(target_KPI.index.map(lambda x: for_TD_machine_names["Названия станков в файле Daily"].loc[for_TD_machine_names["Названия станков на листе KPI"] == x].values[0]))
    except:
        cell_name = for_TD.loc["Диапазон ячеек с целевым KPI", "Значение"]
        raise ValueError(f"Изменены названия станков в ячейках слева от {cell_name}") 
    # сортировка прихотливым индексированием
    sorted_today_TD = np.array(today_TD.loc[names_sorted_by_target_KPI]["TD"])  

    KPI_sheet.range(for_TD.loc["Диапазон ячеек с целевым KPI", "Значение"]).offset(0, 1).value = sorted_today_TD.reshape(len(sorted_today_TD), 1)
def grouped_by_day():
    global day_grouped_Downtime
    day_grouped_Downtime = grouped_Downtime.groupby(['Станок', 'Дата']).agg(
        {"Время работы for maintenance": 'sum', "простой в мин.": "sum"}).reset_index()
    day_grouped_Downtime["TD"] = day_grouped_Downtime["простой в мин."] / (
                day_grouped_Downtime["Время работы for maintenance"] * 60) * 100
    day_grouped_Downtime = day_grouped_Downtime.sort_values(by='Дата', ascending=False)
    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по дням", "Значение"]).offset(0,-1).value = day_grouped_Downtime[["Станок", "Дата", "TD"]]
    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по дням", "Значение"]).offset(1,-1).expand('down').value = None  # чистим индексы
def grouped_by_month():
    global month_grouped_Downtime
    # По Месяцам
    month_grouped_Downtime = grouped_Downtime.groupby(['Станок', 'Год', 'Месяц']).agg(
        {"Время работы for maintenance": 'sum', "простой в мин.": "sum"}).reset_index()
    month_grouped_Downtime["TD"] = month_grouped_Downtime["простой в мин."] / (
                month_grouped_Downtime["Время работы for maintenance"] * 60) * 100

    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по месяцам", "Значение"]).offset(0,-1).value = month_grouped_Downtime[["Станок", "Год", "Месяц", "TD"]]
    KPI_sheet.range(for_TD.loc["Ячейка для таблицы по месяцам", "Значение"]).offset(1,-1).expand('down').value = None  # чистим индексы

def TD_all():
    Today()
    grouped_by_shift()
    grouped_by_day()
    grouped_by_month()

def colorize_grouped_by_shift():
    global grouped_Downtime
    if (grouped_Downtime is not None):
        colorize(grouped_Downtime, cell_to_tuple(for_TD.loc["Ячейка для таблицы по сменам", "Значение"]), 3) #E133
def colorize_grouped_by_day():
    global day_grouped_Downtime
    if (day_grouped_Downtime is not None):
        colorize(day_grouped_Downtime, cell_to_tuple(for_TD.loc["Ячейка для таблицы по дням", "Значение"]), 2) #J133
def colorize_grouped_by_month():
    global month_grouped_Downtime
    if (month_grouped_Downtime is not None):
        colorize(month_grouped_Downtime, cell_to_tuple(for_TD.loc["Ячейка для таблицы по месяцам", "Значение"]), 3) #P133
def colorize_all():
    colorize_grouped_by_shift()
    colorize_grouped_by_day()
    colorize_grouped_by_month()
def do_nothing():
    pass
def upadte():
    global was_main_calc
    global was_open_all
    was_main_calc = False
    was_open_all = False
    execute(do_nothing)
def execute(f): #Добавляет проверки перед выполнением функции
    errors.delete('1.0', 'end')
    try:
        if (not was_open_all):
            open_all()
        if (not was_main_calc):
            main_calc()
        f()
    except Exception as e:
        errors.insert(1.0, str(e)
            +'\n1. Проверьте корректность названия в файле: "Файл для TD exe.xlsx"\n\
2. Исправьте найденную ошибку и сохраните файл\n\
3. Нажмите кнопку обновить данные или перезагрузите приложение')


root = tk.Tk()
root.geometry("700x500")
bt_TD_today = tk.Button(root, text="Расчёт TD на вчерашний день", width=30, height=1, command=lambda : execute(Today))
bt_TD_today.place(x=50, y=25)
bt_TD_by_shift = tk.Button(root, text="TD по сменам", width=30, height=1, command=lambda : execute(grouped_by_shift))
bt_TD_by_shift.place(x=50, y=55)
colorize_bt_TD_by_shift = tk.Button(root, text="Раскрасить TD по сменам", width=30, height=1, command=lambda : execute(colorize_grouped_by_shift))
colorize_bt_TD_by_shift.place(x=300, y=55)
bt_TD_by_day = tk.Button(root, text="TD по дням", width=30, height=1, command=lambda : execute(grouped_by_day))
bt_TD_by_day.place(x=50, y=85)
colorize_bt_TD_by_day = tk.Button(root, text="Раскрасить TD по дням", width=30, height=1, command=lambda : execute(colorize_grouped_by_day))
colorize_bt_TD_by_day.place(x=300, y=85)
bt_TD_by_month = tk.Button(root, text="TD по месяцам", width=30, height=1, command=lambda : execute(grouped_by_month))
bt_TD_by_month.place(x=50, y=115)
bt_TD_by_month = tk.Button(root, text="Раскрасить TD по месяцам", width=30, height=1, command=lambda : execute(colorize_grouped_by_month))
bt_TD_by_month.place(x=300, y=115)
bt_TD_by_shift_clear = tk.Button(root, text="Отчистить всё", width=20, height=1, command=lambda : execute(clear))
bt_TD_by_shift_clear.place(x=50, y=145)
bt_TD_by_shift_clear = tk.Button(root, text="Раскрасить всё", width=20, height=1, command=lambda : execute(colorize_all))
bt_TD_by_shift_clear.place(x=300, y=145)
bt_TD_all = tk.Button(root, text="Рассчитать всё", width=20, height=2, command=lambda : execute(TD_all))
bt_TD_all.place(x=50, y=175)
update = tk.Button(root, text="Обновить данные", width=20, height=2, command=lambda : upadte())
update.place(x=300, y=175)
errors = tk.Text(root, width=70, height=10, foreground="red")
errors.place(x=50, y=235)

root.mainloop()