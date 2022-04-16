import pandas as pd

# Пишем переменную для считывания файла
date_data = '28.03.2022'
filepath = '~/Desktop/Script check/РО '+ date_data + '.xlsm'
# Считываем файл
df_ro = pd.read_excel(filepath, sheet_name='TDSheet', skiprows=5)

# Фильтруем не нужные штуки
df_ro = df_ro[df_ro['ID Должности'] != 99999999]
# Обрезаем нужные столбцы
df_ro = df_ro.iloc[:, 0:22]

# Применяем фильтры
df_zero_seg = df_ro[df_ro['Сегмент'].isnull()]
df_no_ff = df_ro[df_ro['Function for Forecasting'].isnull()]
df_no_col = df_ro[df_ro['Collars'].isnull()]
df_no_gr = df_ro[df_ro['Грейд Korn Ferry'].isnull()]
df_dec_zero = df_ro[(df_ro['Признак занятости'] == 'НН') | (df_ro['Признак занятости'] == 'СОВ/ДЕК') & (df_ro['Степень занятости'] == 0)]
df_rab_zero = df_ro[(df_ro['Признак занятости'] == 'РАБ') & (df_ro['Степень занятости']) == 0]

# Переменная для записи экселя
writer = pd.ExcelWriter('~/Desktop/Script check/Report.xlsx', engine='xlsxwriter')

# Записываем листы экселя
df_zero_seg.to_excel(writer, sheet_name='Пуст_сег')
df_no_ff.to_excel(writer, sheet_name='Нет_функций')
df_no_col.to_excel(writer, sheet_name='Нет_coll')
df_no_gr.to_excel(writer, sheet_name='Нет_грейд')
df_dec_zero.to_excel(writer, sheet_name='Декреты')
df_rab_zero.to_excel(writer, sheet_name='РАБ и 0')
df_ro.to_excel(writer, sheet_name='DF')

# Записываем файл
writer.save()


