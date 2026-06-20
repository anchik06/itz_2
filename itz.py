import pandas as pd

input_file = r'нагрузка- вар2-ТиСАПРМП-2023-2024.xlsx'
df_past = pd.read_excel(input_file, sheet_name='Нагрузка')

print(f"строк в файле прошлого года: {len(df_past)}")
print(f"столбцы: {df_past.columns.tolist()}")

#разделяем данные на осень и весну по столбцу "семестр"
df_autumn = df_past[df_past['семестр'] == 'осень'].copy()
df_spring = df_past[df_past['семестр'] == 'весна'].copy()

print(f"\nстрок в осеннем семестре: {len(df_autumn)}")
print(f"строк в весеннем семестре: {len(df_spring)}")

#очищаем столбцы с ФИО и связанные поля, преподаватели будут расставлены автоматически позже (по усл)
columns_clear = ['ФИО', 'Табельный №', 'Должность', '65+', 'Желаемая аудитория']
for col in columns_clear:
    if col in df_autumn.columns:
        df_autumn[col] = None
    if col in df_spring.columns:
        df_spring[col] = None

#объединяем осень и весну в один сводный файл (структура строк сохраняется)
df_current = pd.concat([df_autumn, df_spring], ignore_index=True)

print(f"итого строк в сводном файле текущего года: {len(df_current)}")

#сохраняем сводный файл текущего года
output_current = 'сводный_файл_текущего_года.xlsx'
df_current.to_excel(output_current, sheet_name='Нагрузка', index=False)
print(f"сводный файл: {output_current}")

#3 задача

#ключи для сопоставления (только 3 признака)
keys = ['Группы', 'Дисциплина', 'Вид нагрузки']

#выполняем сопоставление
#суффиксы _current и _past помогают различать столбцы из разных файлов
df_matched = pd.merge(
    df_current,
    df_past,
    on=keys,
    how='left',
    suffixes=('_current', '_past')
)

#проверяем результат сопоставления
matched_count = df_matched['ФИО_past'].notna().sum()
total_count = len(df_matched)

print(f"\nрезультаты сопоставления:")
print(f"всего строк в текущем году: {total_count}")
print(f"найдено совпадений в прошлом году: {matched_count}")
print(f"не найдено совпадений: {total_count - matched_count}")
print(f"процент покрытия: {matched_count / total_count * 100:.1f}%")

#несколько примеров сопоставления
print("\nпримеры сопоставления (первые 5 строк):")
sample = df_matched[df_matched['ФИО_past'].notna()].head(5)
print(sample[['Группы', 'Дисциплина', 'Вид нагрузки', 'ФИО_past']].to_string(index=False))

#сохраняем результат сопоставления в отдельный файл
output_matched = 'сопоставление_с_прошлым_годом.xlsx'
df_matched.to_excel(output_matched, sheet_name='Сопоставление', index=False)

print(f"результат сопоставления : {output_matched}")

#задача 4. Первичное копирование обычной нагрузки из прошлого года

#ключевые слова для исключения
exclude_keywords = [
    'диплом', 'руководство', 'практика',
    'представление научного доклада', 'ига',
    'итоговая гос', 'государственный экзамен'
]
pattern = '|'.join(exclude_keywords) #берет каждую строку в файле и проверяет, встречается ли какое-то слово из списка

#берет столбец из таблицы, заполняет пустые строки и переводит в нижний регистр, результат сохраняется в новую переменную
vid_nagruzki_lower = df_matched['Вид нагрузки'].fillna('').str.lower()
disciplina_lower = df_matched['Дисциплина'].fillna('').str.lower()

#проверяет каждую строку, содержит ли она что-то из шаблона, сохраняет True если найдено и False - если нет
mask_exclude = vid_nagruzki_lower.str.contains(pattern) | disciplina_lower.str.contains(pattern)
#меняем True на False и наоборот, т к нам нужны обычные строчки, а не те, которые мы исключаем
df_matched['is_standard'] = ~mask_exclude

#считает обычные(подходящие нам) строки
standard_count = df_matched['is_standard'].sum()
#считаем строки, которые мы исключаем
excluded_count = (~df_matched['is_standard']).sum()

#выводим, сколько каких строк
print(f"обычных строк нагрузки (для первичного копирования): {standard_count}")
print(f"строк с дипломами/руководством/практиками (отложены для добалансировки): {excluded_count}")

#задача 5. копирование строки, если в прошлом году был один преподаватель

#проверяет каждую строку на отсутствие пропуска, создает копию таблицы, в которой только те строки, гд есть ФИО
df_past_valid = df_past[df_past['ФИО'].notna()].copy()

grouped = df_past_valid.groupby(['Группы', 'Дисциплина', 'Вид нагрузки'])['ФИО'] #группирует все строки таблицы по признакам
past_teachers = grouped.apply(lambda x: list(x.unique())) #убирает дубликаты и превращает результат в список
past_teachers = past_teachers.reset_index() #превращает результат обратно в список
past_teachers.rename(columns={'ФИО': 'past_teachers_list'}, inplace=True) #меняет название столбца

past_teachers['num_teachers'] = past_teachers['past_teachers_list'].apply(len)# считаем количество преподавателей в каждой строке
#оставляем только те строки, где один преподавателб
mask_single = past_teachers['num_teachers'] == 1
past_teachers_single = past_teachers[mask_single].copy()
past_teachers_single['ФИО_single'] = past_teachers_single['past_teachers_list'].apply(lambda x: x[0]) #извлекаем имя преподавателя

#создаем кортежи из трех элементов, далее обьединяем фио и кортеж, получаем словарь из этих пар
single_teacher_map = dict(zip(
    zip(past_teachers_single['Группы'], past_teachers_single['Дисциплина'], past_teachers_single['Вид нагрузки']),
    past_teachers_single['ФИО_single']
))

# определяем, как называется столбец с фио
fio_col_current = 'ФИО_current' if 'ФИО_current' in df_matched.columns else 'ФИО'

#функция для присвоения ФИО
def assign_single_teacher(row):
    if row['is_standard']:#проверяем, подходит ли нам строка (по заданию 4)
        key = (row['Группы'], row['Дисциплина'], row['Вид нагрузки']) #создаем ключ - кортеж из трех значений
        if key in single_teacher_map:
            return single_teacher_map[key] # проверяем, что ключ есть в словаре и возвращаем фио преподавателя
    return row[fio_col_current]

#применяем функцию к каждой строке
df_matched[fio_col_current] = df_matched.apply(assign_single_teacher, axis=1)

#считаем статистику выполненной работы
filled_teachers = df_matched[df_matched['is_standard']][fio_col_current].notna().sum() #считает заполненные обычные строки
total_standard = df_matched['is_standard'].sum() #считает все обычные строки

print(f"всего обычных строк: {total_standard}")
print(f"строк, где в прошлом году был 1 преподаватель и ФИО успешно скопировано: {filled_teachers}")

# сохраняем промежуточный результат
output_task4_5 = 'сводный_файл_задачи_4_5.xlsx'
df_matched.to_excel(output_task4_5, sheet_name='Задачи_4_5', index=False)
print(f"промежуточный результат сохранен в: {output_task4_5}")