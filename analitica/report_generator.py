import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from jinja2 import Template
import os
import re
import base64
import numpy as np
import unicodedata

# helper slugify
def slugify(text: str):
    text = text.replace(' ', '_').replace(',', '').replace('₽', 'RUB')
    return ''.join(c for c in unicodedata.normalize('NFKD', text) if c.isalnum() or c=='_')

# === Файлы кабинетов ===
# Автоматически берём все .xlsx в текущем каталоге
files = sorted([f for f in os.listdir('.') if f.lower().endswith('.xlsx')])

all_data = []
account_names = []
for file in files:
    # Извлекаем логин из названия файла
    fname = os.path.basename(file)
    match = re.search(r'id-([\w\-]+)', fname)
    if match:
        acc_name = match.group(1)
    else:
        acc_name = os.path.splitext(fname)[0]
    account_names.append(acc_name)
    df = pd.read_excel(file)
    header_row = df[df.iloc[:,0] == 'Месяц'].index[0]
    data = pd.read_excel(file, skiprows=header_row+1)
    data = data.dropna(subset=['Месяц'])
    data = data[data['Месяц'] != 'Итого']
    # Проверяем наличие всех нужных столбцов, если нет — добавляем с NaN
    for col in ['Конверсии', 'CR, %', 'CPA, ₽', 'Расход, ₽', 'Показы', 'Клики', 'CPC, ₽']:
        if col not in data.columns:
            data[col] = float('nan')
        data[col] = pd.to_numeric(data[col].astype(str).str.replace(',', '.'), errors='coerce')
    # Пересчитываем CR, %
    data['CR, %'] = data.apply(lambda row: (row['Конверсии'] / row['Клики'] * 100) if row['Клики'] and row['Клики'] > 0 else 0, axis=1)
    data['cabinet'] = acc_name
    all_data.append(data)

# Объединяем данные
full = pd.concat(all_data, ignore_index=True)

# Пересчитываем CR, % для summary
summary = full.groupby('Месяц').agg({
    'Конверсии': 'sum',
    'CR, %': 'mean',  # временно, ниже пересчитаем
    'CPA, ₽': 'mean',
    'Расход, ₽': 'sum',
    'Показы': 'sum',
    'Клики': 'sum',
    'CPC, ₽': 'mean'
}).reset_index()
# Пересчет CR, % для summary
summary['CR, %'] = summary.apply(lambda row: (row['Конверсии'] / row['Клики'] * 100) if row['Клики'] and row['Клики'] > 0 else 0, axis=1)

# --- вспомогательная функция сортировки месяцев (русские названия) ---
months_ru = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']

def month_key(month_str: str):
    # Принимает строку вида "Август, 2024"
    name, year = [p.strip() for p in month_str.split(',')]
    return int(year), months_ru.index(name)

# Сортировка по времени (используем month_key)
summary = summary.sort_values('Месяц', key=lambda col: col.map(month_key))

# Графики сохраняются в папку
os.makedirs('report_imgs', exist_ok=True)

# --- Графики по каждому кабинету ---
for metric, ylabel in [('Конверсии','Конверсии'), ('Расход, ₽','Расход, ₽'), ('Клики','Клики'), ('Показы','Показы')]:
    plt.figure(figsize=(10,5))
    sns.barplot(data=full, x='Месяц', y=metric, hue='cabinet')
    plt.title(f'{metric} по месяцам (сравнение аккаунтов)')
    plt.ylabel(ylabel)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(f'report_imgs/compare_{metric}.png')
    plt.close()

# --- Суммарные графики ---
for metric, ylabel in [('Конверсии','Конверсии'), ('Расход, ₽','Расход, ₽'), ('Клики','Клики'), ('Показы','Показы')]:
    plt.figure(figsize=(10,5))
    plt.bar(summary['Месяц'], summary[metric])
    plt.title(f'Суммарные {metric} по месяцам (все аккаунты)')
    plt.ylabel(ylabel)
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.savefig(f'report_imgs/total_{metric}.png')
    plt.close()

def format_val(val, col_name):
    """Форматирование чисел в таблицах.
    - Для CR, % выводим с двумя десятичными знаками.
    - Для остальных метрик – целые с пробелами в качестве разделителя тысяч.
    """
    try:
        if pd.isna(val):
            return ''
        if col_name == 'CR, %':
            return f"{val:.2f}"
        return '{:,.0f}'.format(val).replace(',', ' ')
    except Exception:
        return val

# --- Универсальные таблицы по каждому аккаунту ---
tables_cab = []
for i, df in enumerate(all_data):
    cols_order = ['Месяц','Конверсии','CR, %','CPA, ₽','Расход, ₽','Показы','Клики','CPC, ₽']
    df_sorted = df.sort_values('Месяц', key=lambda col: col.map(month_key), ascending=False)
    table_fmt = df_sorted[cols_order].copy()
    for col in cols_order:
        if col != 'Месяц':
            table_fmt[col] = table_fmt[col].apply(lambda v, c=col: format_val(v, c))
    # --- Итоговая строка ---
    total_values = {
        'Месяц': 'Итого',
        'Конверсии': df['Конверсии'].sum(),
        'CR, %': df['CR, %'].mean(),
        'CPA, ₽': df['CPA, ₽'].mean(),
        'Расход, ₽': df['Расход, ₽'].sum(),
        'Показы': df['Показы'].sum(),
        'Клики': df['Клики'].sum(),
        'CPC, ₽': df['CPC, ₽'].mean()
    }
    total_fmt = {k: (format_val(v,k) if k!='Месяц' else v) for k,v in total_values.items()}

    # --- Формируем HTML ---
    html = '<table><thead><tr>' + ''.join(f'<th>{c}</th>' for c in cols_order) + '</tr></thead><tbody>'
    for _, row in table_fmt.iterrows():
        html += '<tr>' + ''.join(f'<td>{row[c]}</td>' for c in cols_order) + '</tr>'
    # Итоговая строка жирным
    html += '<tr>' + ''.join(f'<td style="font-weight:bold;">{total_fmt[c]}</td>' for c in cols_order) + '</tr>'
    html += '</tbody></table>'
    tables_cab.append(html)

# --- Сводная таблица по аккаунтам ---
def agg_row(df, name):
    return pd.Series({
        'Аккаунт': name,
        'Конверсии': df['Конверсии'].sum(),
        'CR, %': df['CR, %'].mean(),
        'CPA, ₽': df['CPA, ₽'].mean(),
        'Расход, ₽': df['Расход, ₽'].sum(),
        'Показы': df['Показы'].sum(),
        'Клики': df['Клики'].sum(),
        'CPC, ₽': df['CPC, ₽'].mean()
    })
rows = [agg_row(df, name) for df, name in zip(all_data, account_names)]
sum_row = pd.Series({
    'Аккаунт': 'Суммарно',
    'Конверсии': sum(r['Конверсии'] for r in rows),
    'CR, %': sum(r['CR, %'] for r in rows)/len(rows),
    'CPA, ₽': sum(r['CPA, ₽'] for r in rows)/len(rows),
    'Расход, ₽': sum(r['Расход, ₽'] for r in rows),
    'Показы': sum(r['Показы'] for r in rows),
    'Клики': sum(r['Клики'] for r in rows),
    'CPC, ₽': sum(r['CPC, ₽'] for r in rows)/len(rows)
})
sum_table_df = pd.DataFrame(rows + [sum_row])
sum_table_html = '<table><thead><tr>' + ''.join(f'<th>{col}</th>' for col in sum_table_df.columns) + '</tr></thead><tbody>'
for _, row in sum_table_df.iterrows():
    sum_table_html += '<tr>'
    for col in sum_table_df.columns:
        val = row[col]
        if col != 'Аккаунт':
            val = format_val(val, col)
        sum_table_html += f'<td>{val}</td>'
    sum_table_html += '</tr>'
sum_table_html += '</tbody></table>'

# --- Таблицы ---
table_total = summary[['Месяц','Конверсии','CR, %','CPA, ₽','Расход, ₽','Показы','Клики','CPC, ₽']].copy()
# сортировка summary desc
table_total = table_total.sort_values('Месяц', key=lambda col: col.map(month_key), ascending=False)
for col in table_total.columns:
    if col != 'Месяц':
        table_total[col] = table_total[col].apply(lambda v, c=col: format_val(v, c))
# добавляем итоговую строку
tot_total = {
    'Месяц': 'Итого',
    'Конверсии': format_val(summary['Конверсии'].sum(), 'Конверсии'),
    'CR, %': format_val(summary['CR, %'].mean(), 'CR, %'),
    'CPA, ₽': format_val(summary['CPA, ₽'].mean(), 'CPA, ₽'),
    'Расход, ₽': format_val(summary['Расход, ₽'].sum(), 'Расход, ₽'),
    'Показы': format_val(summary['Показы'].sum(), 'Показы'),
    'Клики': format_val(summary['Клики'].sum(), 'Клики'),
    'CPC, ₽': format_val(summary['CPC, ₽'].mean(), 'CPC, ₽')
}
table_total = pd.concat([table_total, pd.DataFrame([tot_total])], ignore_index=True)
table_total = table_total.to_html(index=False, escape=False)

def img_to_base64(path):
    with open(path, 'rb') as f:
        return 'data:image/png;base64,' + base64.b64encode(f.read()).decode('utf-8')

# ... после генерации графиков ...
img_paths = {
    'compare_Конверсии': 'report_imgs/compare_Конверсии.png',
    'compare_Расход, ₽': 'report_imgs/compare_Расход, ₽.png',
    'compare_Клики': 'report_imgs/compare_Клики.png',
    'compare_Показы': 'report_imgs/compare_Показы.png',
    'total_Конверсии': 'report_imgs/total_Конверсии.png',
    'total_Расход, ₽': 'report_imgs/total_Расход, ₽.png',
    'total_Клики': 'report_imgs/total_Клики.png',
    'total_Показы': 'report_imgs/total_Показы.png',
}
img_b64 = {k: img_to_base64(v) for k, v in img_paths.items()}

# --- Список месяцев в хронологическом порядке (desc: последний месяц сверху) ---
months_sorted_desc = sorted({m for df in all_data for m in df['Месяц']}, key=month_key, reverse=True)

# --- Сравнительная таблица по месяцам для всех кабинетов ---
def make_monthly_comparison_table(all_data, account_names):
    # Собираем все уникальные месяцы
    months = sorted(set(m for df in all_data for m in df['Месяц']))
    metrics = ['Конверсии', 'CR, %', 'CPA, ₽', 'Расход, ₽', 'Показы', 'Клики', 'CPC, ₽']
    # Формируем заголовки
    columns = ['Месяц']
    for name in account_names:
        for metric in metrics:
            columns.append(f'{metric} ({name})')
    rows = []
    for month in months:
        row = [month]
        for df in all_data:
            df_month = df[df['Месяц'] == month]
            for metric in metrics:
                if not df_month.empty:
                    val = df_month.iloc[0][metric]
                else:
                    val = ''
                if metric != 'Месяц' and metric != 'cabinet':
                    val = format_val(val, metric)
                row.append(val)
        rows.append(row)
    # Формируем HTML-таблицу
    html = '<table><thead><tr>' + ''.join(f'<th>{col}</th>' for col in columns) + '</tr></thead><tbody>'
    for row in rows:
        html += '<tr>' + ''.join(f'<td>{cell}</td>' for cell in row) + '</tr>'
    html += '</tbody></table>'
    return html

monthly_comparison_table = make_monthly_comparison_table(all_data, account_names)

# --- Функция таблицы по одному метрику ---
def make_metric_table(all_data, account_names, metric, metric_name, months):
    columns = ['Месяц'] + [f'{metric_name} ({name})' for name in account_names]
    rows = []
    for month in months:
        row = [month]
        for df in all_data:
            df_m = df[df['Месяц'] == month]
            if not df_m.empty and not pd.isna(df_m.iloc[0][metric]):
                val_num = df_m.iloc[0][metric]
            else:
                val_num = None
            row.append(val_num)
        rows.append(row)
    # --- HTML ---
    html = f'<h3>{metric_name}</h3>'
    html += '<table><thead><tr>' + ''.join(f'<th>{c}</th>' for c in columns) + '</tr></thead><tbody>'
    for row in rows:
        month = row[0]
        values = row[1:]
        # Определяем максимум
        numeric_vals = [v for v in values if v is not None]
        max_val = max(numeric_vals) if numeric_vals else None
        html += '<tr><td>' + month + '</td>'
        for val in values:
            if val is None:
                disp = ''
            else:
                disp = format_val(val, metric_name)
            style = 'font-weight:bold;' if max_val is not None and val == max_val else ''
            html += f'<td style="{style}">{disp}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    # --- Линейный график ---
    slug = slugify(metric_name)
    chart_key = f'line_{slug}'
    chart_file = f'report_imgs/{chart_key}.png'
    try:
        pivot = full.pivot_table(index='Месяц', columns='cabinet', values=metric, aggfunc='sum').reindex(months)
        pivot.plot(figsize=(8,5))
        plt.title(f'{metric_name} по месяцам')
        plt.ylabel(metric_name)
        plt.xticks(rotation=45, ha='right')
        plt.tight_layout()
        plt.savefig(chart_file)
        plt.close()
        # Добавим тег img после таблицы
        img_b64_local = img_to_base64(chart_file)
        html += f'<br/><img src="{img_b64_local}"><br/>'
    except Exception as e:
        print('Не смог построить линейный график', metric_name, e)
    return html

metric_tables_html = ''
# динамические картинки (линейные графики)
img_paths_dynamic = []
metric_mappings = [
    ('Расход, ₽', 'Расход, ₽'),
    ('Клики', 'Клики'),
    ('Конверсии', 'Конверсии'),
    ('CPA, ₽', 'CPA, ₽'),
    ('CPC, ₽', 'CPC, ₽')
]
for m, name in metric_mappings:
    metric_tables_html += make_metric_table(all_data, account_names, m, name, months_sorted_desc)

# img_paths_dynamic больше не нужен, т.к. base64 уже встроен в HTML

# --- График зависимость Расход vs Конверсии (scatter + линейная регрессия) ---
scatter_path = 'report_imgs/spend_vs_conversions.png'
try:
    # Сортируем по хронологии месяцев (asc)
    summary_sorted = summary.sort_values('Месяц', key=lambda col: col.map(month_key))
    months_axis = range(len(summary_sorted))
    fig, ax1 = plt.subplots(figsize=(10,6))
    color1 = 'tab:blue'
    ax1.set_xlabel('Месяцы (хронология)')
    ax1.set_ylabel('Расход, ₽', color=color1)
    ax1.plot(months_axis, summary_sorted['Расход, ₽'], marker='o', linestyle='-', color=color1, label='Расход')
    ax1.tick_params(axis='y', labelcolor=color1)

    ax2 = ax1.twinx()
    color2 = 'tab:orange'
    ax2.set_ylabel('Конверсии', color=color2)
    ax2.plot(months_axis, summary_sorted['Конверсии'], marker='o', linestyle='-', color=color2, label='Конверсии')
    ax2.tick_params(axis='y', labelcolor=color2)

    plt.title('Динамика Расходов и Конверсий по месяцам')
    ax1.set_xticks(months_axis)
    ax1.set_xticklabels(summary_sorted['Месяц'], rotation=45, ha='right')
    fig.tight_layout()
    plt.savefig(scatter_path)
    plt.close()
except Exception as e:
    print('Не удалось построить scatter график:', e)

img_paths = {
    'compare_Конверсии': 'report_imgs/compare_Конверсии.png',
    'compare_Расход, ₽': 'report_imgs/compare_Расход, ₽.png',
    'compare_Клики': 'report_imgs/compare_Клики.png',
    'compare_Показы': 'report_imgs/compare_Показы.png',
    'total_Конверсии': 'report_imgs/total_Конверсии.png',
    'total_Расход, ₽': 'report_imgs/total_Расход, ₽.png',
    'total_Клики': 'report_imgs/total_Клики.png',
    'total_Показы': 'report_imgs/total_Показы.png',
    'scatter_Спенд_Конверсии': scatter_path,
}
img_b64 = {k: img_to_base64(v) for k, v in img_paths.items()}

html_template = '''
<html>
<head>
    <meta charset="utf-8">
    <title>Сравнительный аналитический отчет</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1, h2 { color: #2c3e50; }
        table { border-collapse: collapse; width: 80%; margin-bottom: 30px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: right; }
        th { background: #f4f4f4; }
        img { max-width: 700px; margin-bottom: 30px; }
    </style>
</head>
<body>
    <h1>Сравнительный аналитический отчет по аккаунтам</h1>
    <h2>Период: 29.06.2024 - 28.06.2025</h2>
    {% for acc in account_names %}
    <h2>{{ acc }}</h2>
    {{ tables_cab[loop.index0]|safe }}
    {% endfor %}
    <h2>Сравнительные графики</h2>
    <h3>Конверсии</h3>
    <img src="{{ img_b64['compare_Конверсии'] }}">
    <h3>Расходы</h3>
    <img src="{{ img_b64['compare_Расход, ₽'] }}">
    <h3>Клики</h3>
    <img src="{{ img_b64['compare_Клики'] }}">
    <h3>Показы</h3>
    <img src="{{ img_b64['compare_Показы'] }}">
    <h2>Суммарные показатели по всем аккаунтам</h2>
    {{ table_total|safe }}
    <h2>Суммарные графики</h2>
    <h3>Конверсии</h3>
    <img src="{{ img_b64['total_Конверсии'] }}">
    <h3>Расходы</h3>
    <img src="{{ img_b64['total_Расход, ₽'] }}">
    <h3>Клики</h3>
    <img src="{{ img_b64['total_Клики'] }}">
    <h3>Показы</h3>
    <img src="{{ img_b64['total_Показы'] }}">
    <h2>Сравнительные таблицы по месяцам</h2>
    {{ metric_tables_html|safe }}
    <h2>Зависимость "Расходы" vs "Конверсии"</h2>
    <img src="{{ img_b64['scatter_Спенд_Конверсии'] }}">
</body>
</html>
'''

report = Template(html_template).render(
    account_names=account_names,
    tables_cab=tables_cab,
    table_total=table_total,
    sum_table=sum_table_html,
    metric_tables_html=metric_tables_html,
    img_b64=img_b64
)

with open('analitica_report.html', 'w', encoding='utf-8') as f:
    f.write(report)

print('Сравнительный отчет успешно создан: analitica_report.html') 