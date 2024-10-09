import pandas as pd

# Загружаем данные из Excel файла
plan = pd.read_excel(r"Task.xlsx", sheet_name="Plan")
fact = pd.read_excel(r"Task.xlsx", sheet_name="Fact")

# Преобразуем дату в формат даты
plan['дата'] = pd.to_datetime(plan['дата'], dayfirst=True)
fact['дата'] = pd.to_datetime(fact['дата'], dayfirst=True)

# Создаем столбцы с неделями, месяцами и годами
plan['неделя'] = plan['дата'].dt.isocalendar().week
plan['месяц'] = plan['дата'].dt.month
plan['год'] = plan['дата'].dt.year

fact['неделя'] = fact['дата'].dt.isocalendar().week
fact['месяц'] = fact['дата'].dt.month
fact['год'] = fact['дата'].dt.year

# Группируем данные по неделям и объединяем их
plan_by_week = plan.groupby(['год', 'неделя', 'юнит', 'локация_производства'])['количество_план'].sum().reset_index().rename(columns={'количество_план': 'количество_план'})
fact_by_week = fact.groupby(['год', 'неделя', 'юнит', 'локация_производства'])['количество'].sum().reset_index().rename(columns={'количество': 'количество_факт'})

weekly_data = pd.merge(plan_by_week, fact_by_week, on=['год', 'неделя', 'юнит', 'локация_производства'], how='outer')

# Группируем данные по месяцам и объединяем их
plan_by_month = plan.groupby(['год', 'месяц', 'юнит', 'локация_производства'])['количество_план'].sum().reset_index().rename(columns={'количество_план': 'количество_план'})
fact_by_month = fact.groupby(['год', 'месяц', 'юнит', 'локация_производства'])['количество'].sum().reset_index().rename(columns={'количество': 'количество_факт'})

monthly_data = pd.merge(plan_by_month, fact_by_month, on=['год', 'месяц', 'юнит', 'локация_производства'], how='outer')

# Группируем данные по годам и объединяем их
plan_by_year = plan.groupby(['год', 'юнит', 'локация_производства'])['количество_план'].sum().reset_index().rename(columns={'количество_план': 'количество_план'})
fact_by_year = fact.groupby(['год', 'юнит', 'локация_производства'])['количество'].sum().reset_index().rename(columns={'количество': 'количество_факт'})

yearly_data = pd.merge(plan_by_year, fact_by_year, on=['год', 'юнит', 'локация_производства'], how='outer')

# Заменяем NaN на 0 в каждой таблице
weekly_data[['количество_план', 'количество_факт']] = weekly_data[['количество_план', 'количество_факт']].fillna(0)
monthly_data[['количество_план', 'количество_факт']] = monthly_data[['количество_план', 'количество_факт']].fillna(0)
yearly_data[['количество_план', 'количество_факт']] = yearly_data[['количество_план', 'количество_факт']].fillna(0)


# Сохраняем данные в Excel файлы
with pd.ExcelWriter('Aggregated_Data.xlsx') as writer:
    weekly_data.to_excel(writer, sheet_name='Weekly_Data', index=False)
    monthly_data.to_excel(writer, sheet_name='Monthly_Data', index=False)
    yearly_data.to_excel(writer, sheet_name='Yearly_Data', index=False)


print("Data saved to Excel")
