from datetime import datetime as dt
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import (column_index_from_string,
                                 coordinate_from_string)

# заполнить вручную
file_name = 'Приложение 1__(2).xlsx' #название файла
sheet_name = 'Волгоградский 32 кор' #название листа в файле
name_start_column = 'A' #даты начала работ
name_end_column = 'B' #даты конца работ
name_costs_column ='E' #стоимости работ
name_table_start_column = 'L' #столбец, где начинается заполнение графика
ind_start_row = 12 #строка с 1-й работой
ind_end_row = 48 #строка с последней работой
ind_sum_row = 51 #строка, где суммируются плановые затраты по месяцам

#далее - работает программа
starts = []
ends = []
work_in_month = []
costs = []
start_month = []
end_month = []
costs_per_month = []
PV_sum_in_month = []
PV_sum_of_sums = []

ind_start_column = ind_end_column = ind_costs_column = ind_table_start_column = 0
start_date_project = end_date_project = 0


#функции
def calc_columns_index():
    global ind_start_column, ind_end_column, ind_costs_column, ind_table_start_column
    xy = coordinate_from_string(name_start_column+str(ind_start_row))
    ind_start_column = column_index_from_string(xy[0])
    xy = coordinate_from_string(name_end_column+str(ind_start_row))
    ind_end_column = column_index_from_string(xy[0])
    xy = coordinate_from_string(name_costs_column+str(ind_start_row))
    ind_costs_column = column_index_from_string(xy[0])
    xy = coordinate_from_string(name_table_start_column+str(ind_start_row))
    ind_table_start_column = column_index_from_string(xy[0])

def time_delta_in_month(d1, d2):
    m1 = d1.month
    y1 = d1.year
    m2 = d2.month
    y2 = d2.year
    if y1 == y2:
        return m2-m1+1
    elif y2 > y1:
        return (y2-y1-1)*12 + (12-m1+1)+m2
    else:
        return -111

def fill_month(m, y):
    global start_month, end_month
    if m==1 or m==3 or m==5 or m==7 or m==8 or m==10 or m==12:
        start_month.append(dt(y, m, 1).date())
        end_month.append(dt(y, m, 31).date())
    elif m==2:
        start_month.append(dt(y, m, 1).date())
        if y%4 == 0: #високосный
            end_month.append(dt(y, m, 29).date())
        else:
            end_month.append(dt(y, m, 28).date())
    elif m==4 or m==6 or m==9 or m==11:
        start_month.append(dt(y, m, 1).date())
        end_month.append(dt(y, m, 30).date())

def fill_month2(st_date, count_month):
    global start_month, end_month
    month = st_date.month
    year = st_date.year
    for i in range(0, count_month):
        fill_month(month, year)
        if month == 12:
            month = 1
            year += 1
        else:
            month += 1

def find_start_end_project(st, end):
    global start_date_project, end_date_project
    st_ = st.copy()
    st_.sort()
    end_ = end.copy()
    end_.sort()
    start_date_project = st_[0]
    end_date_project = end_[-1]

#открыть файл с нужным листом
book_graf=openpyxl.open(file_name) #, data_only=True)
sheet_graf=book_graf[sheet_name]
#вычислить индексы столбцов, откуда считывать данные (чтобы не использовать буквы в названиях столбцов)
calc_columns_index()
#считать данные о
# - начале работ (каждой работы)
# - конце работ
# - сколько месяцев длиться работа (с расчетом в функции)
# - стоимости работы
# - стоимости работы в месяц, но с учетом, что 1-й и посл-й месяцы идут с расходом "/2", остальные месяцы с полной суммой
for i in range(ind_start_row, ind_end_row+1):
    ind_row = i-ind_start_row
    starts.append(dt.date(sheet_graf.cell(row=i, column=ind_start_column).value))
    ends.append(dt.date(sheet_graf.cell(row=i, column=ind_end_column).value))
    work_in_month.append(time_delta_in_month(starts[ind_row], ends[ind_row])) #append((ends[j]-starts[j]).days//27)
    costs.append(sheet_graf.cell(row=i, column=ind_costs_column).value)
    costs_per_month.append(costs[ind_row]/(work_in_month[ind_row]))#-1)) при -1 можно сделать затраты на 1й и последний месяц в два раза меньше, чем со 2го по предпоследний
#находим самую раннюю и подзнюю даты, это начало и конец проекта; а также продолжительность в месяцах всего проекта
find_start_end_project(starts, ends)
count_month_in_project = time_delta_in_month(start_date_project, end_date_project)
# заполняем даты для каждой ячейки в укрупненном графике
#for i in range(0, count_month_in_project):
fill_month2(start_date_project, count_month_in_project)
#заполнить укрупненный график с суммами по месяцам
#begin_to_fill = False
#begin_j = -1
for i in range(ind_start_row, ind_end_row+1):
    ind_row = i-ind_start_row
    for j in range(ind_table_start_column, ind_table_start_column+count_month_in_project):
        ind_col = j-ind_table_start_column
        if ((starts[ind_row]>=start_month[ind_col] and starts[ind_row]<=end_month[ind_col]) or 
            (ends[ind_row]>=start_month[ind_col] and ends[ind_row]<=end_month[ind_col]) or
            (starts[ind_row]<=start_month[ind_col] and ends[ind_row]>=end_month[ind_col])):
            #if begin_to_fill == False:
            #    begin_to_fill = True
            #    begin_j = j
            #if j == begin_j or j == begin_j+work_in_month[ind_row]: #-1:
            sheet_graf.cell(row=i, column=j).value = costs_per_month[ind_row] #/2.0
            #else:
            #    sheet_graf.cell(row=i, column=j).value = costs_per_month[ind_row]
            sheet_graf.cell(row=i, column=j).fill = openpyxl.styles.PatternFill(start_color="0033CCFF", end_color="0033CCFF", fill_type = "solid")
        else:
            sheet_graf.cell(row=i, column=j).value = ""
            sheet_graf.cell(row=i, column=j).fill = openpyxl.styles.PatternFill(start_color="00FFFFFF", end_color="00FFFFFF", fill_type = "solid")
    #begin_to_fill = False
    #begin_j = -1
#заполняем PV по месяцам и накопительным итогом
for j in range(ind_table_start_column, ind_table_start_column+count_month_in_project):
    ind_col = j-ind_table_start_column
    PV_sum_in_month.append(0)
    for i in range(ind_start_row, ind_end_row+1):
        if sheet_graf.cell(row=i, column=j).data_type == 'n':
            PV_sum_in_month[ind_col] += sheet_graf.cell(row=i, column=j).value
    #заполняем по месяцам
    sheet_graf.cell(row=ind_sum_row, column=j).value = PV_sum_in_month[ind_col]
    #заполняем накопительно
    if ind_col == 0:
        PV_sum_of_sums.append(PV_sum_in_month[ind_col])
        sheet_graf.cell(row=ind_sum_row+1, column=j).value = PV_sum_of_sums[ind_col]
    else:
        PV_sum_of_sums.append(0)
        PV_sum_of_sums[ind_col] = PV_sum_of_sums[ind_col-1] + PV_sum_in_month[ind_col]
        sheet_graf.cell(row=ind_sum_row+1, column=j).value = PV_sum_of_sums[ind_col]

book_graf.save(file_name)