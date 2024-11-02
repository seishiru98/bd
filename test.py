
from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT

import pandas as pd
from openpyxl import load_workbook

import datetime
#-----------------------------------------------------------------------------------------------------------------------

doc = Document()

section = doc.sections[0]

section.left_margin = Cm(2)  # Левое поле
section.right_margin = Cm(1)  # Правое поле
section.top_margin = Cm(2)  # Верхнее поле
section.bottom_margin = Cm(2)  # Нижнее поле
#-----------------------------------------------------------------------------------------------------------------------
# Функции

def set_font(run, font_name, font_size, italic=False, bold=False):
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.italic = italic
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

    # Это нужно для корректного отображения шрифта на всех платформах
    r = run._element
    rPr = r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), font_name)
    rPr.append(rFonts)


def set_paragraph_format(paragraph, left_indent=0, right_indent=0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0):
    paragraph_format = paragraph.paragraph_format
    paragraph_format.left_indent = Cm(left_indent)
    paragraph_format.right_indent = Cm(right_indent)
    paragraph_format.first_line_indent = Cm(first_line_indent)
    paragraph_format.line_spacing = Pt(line_spacing)
    paragraph_format.space_after = Cm(space_after)
    paragraph_format.space_before = Cm(space_before)

def add_header(doc, header_text):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(header_text)
    run.font.size = Pt(14)
    run.font.name = 'Times New Roman'
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

def add_table(doc, df, merged_ranges, max_rows_per_page=25):
    total_width = Cm(18.5)  # Ширина страницы

    # Определение максимальной длины текста для каждой колонки
    column_widths = [0] * len(df.columns)
    for _, row in df.iterrows():
        for i, value in enumerate(row):
            if pd.notna(value) and str(value).strip() != "":
                column_widths[i] = max(column_widths[i], len(str(value)))
    total_text_length = sum(column_widths)

    def add_page_table(df_page, merged_ranges):
        table = doc.add_table(rows=len(df_page) + 1, cols=len(df.columns))
        table.style = 'Table Grid'

        # Заголовки таблицы
        hdr_cells = table.rows[0].cells
        for i, column_name in enumerate(df.columns):
            hdr_cells[i].width = Cm(total_width.cm * (column_widths[i] / total_text_length))
            hdr_cells[i].text = column_name
            set_paragraph_format(hdr_cells[i].paragraphs[0], line_spacing=18)

        # Заполнение данных
        for idx, row in df_page.iterrows():
            if idx + 1 < len(table.rows):
                row_cells = table.rows[idx + 1].cells
                for i, value in enumerate(row):
                    row_cells[i].width = Cm(total_width.cm * (column_widths[i] / total_text_length))
                    cell_paragraph = row_cells[i].paragraphs[0]
                    cell_paragraph.text = str(value) if pd.notna(value) else ""
                    set_paragraph_format(cell_paragraph, line_spacing=18)

        # Объединение ячеек для текущей страницы
        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            if min_row - 1 < len(df_page) + 1 and max_row - 1 < len(df_page) + 1:
                start_cell = table.cell(min_row - 1, min_col - 1)
                end_cell = table.cell(max_row - 1, max_col - 1)
                start_cell.merge(end_cell)

    # Разбиение таблицы на страницы
    for start_row in range(0, len(df), max_rows_per_page):
        df_page = df.iloc[start_row:start_row + max_rows_per_page]
        add_page_table(df_page, merged_ranges)
        if start_row + max_rows_per_page < len(df):
            doc.add_page_break()


def read_excel_with_merged_cells(filename, sheet_name):
    # Загружаем Excel-файл
    wb = load_workbook(filename, data_only=True)

    # Проверка наличия листа
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Лист '{sheet_name}' не найден в файле '{filename}'.")

    ws = wb[sheet_name]

    # Извлечение данных
    data = []
    for row in ws.iter_rows(values_only=True):
        data_row = []
        for cell in row:
            data_row.append(cell if cell is not None else "")
        data.append(data_row)

    # Получение объединенных диапазонов
    merged_ranges = ws.merged_cells.ranges

    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        merged_value = ws.cell(row=min_row, column=min_col).value
        for row in range(min_row - 1, max_row):
            for col in range(min_col - 1, max_col):
                if row == min_row - 1 and col == min_col - 1:
                    data[row][col] = merged_value
                else:
                    data[row][col] = ''

    # Преобразование в DataFrame pandas
    df = pd.DataFrame(data[1:], columns=data[0])
    return df, merged_ranges


class Counter:
    def __init__(self, start_value, step):
        self.value = start_value
        self.step = step

    def increment(self):
        current_value = self.value
        self.value += self.step
        return current_value

class HeadingCounter(Counter):
    def __init__(self, start_value, paragraph_counter, table_counter, fig_counter):
        super().__init__(start_value, 1)
        self.paragraph_counter = paragraph_counter
        self.table_counter = table_counter
        self.fig_counter = fig_counter

    def increment(self):
        current_value = super().increment()
        self.paragraph_counter.reset(current_value + 0.1)
        self.table_counter.reset(current_value + 0.1)
        self.fig_counter.reset(current_value + 0.1)
        return current_value

class ParagraphCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

class TableCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

class FigCounter(Counter):
    def reset(self, new_start_value):
        self.value = new_start_value

# Инициализация счетчиков
n_heading = 1
n_paragraph_start = n_heading + 0.1
n_table_start = n_heading + 0.1
n_fig_start = n_heading + 0.1

par_counter = ParagraphCounter(n_paragraph_start, 0.1)
table_counter = TableCounter(n_table_start, 0.1)
fig_counter = FigCounter(n_fig_start, 0.1)
head_counter = HeadingCounter(n_heading, par_counter, table_counter, fig_counter)

def read_excel_data(filename, sheet_name):
    try:
        return pd.read_excel(filename, sheet_name=sheet_name)
    except FileNotFoundError:
        print(f"ОШИБКА: Файл {filename} не найден.")
    except ValueError:
        print(f"ОШИБКА: Лист {sheet_name} не найден в файле {filename}.")
    except Exception as e:
        print(f"ОШИБКА: Произошла ошибка при чтении файла {filename}: {e}")


#-----------------------------------------------------------------------------------------------------------------------
# Переменные
database = pd.read_excel('database.xlsx', sheet_name='1')

work_time = database.iloc[0, 3]
print(work_time)

mass_frac_tiols = database.iloc[15, 4]
print(mass_frac_tiols)
mass_frac_sulphur = database.iloc[16, 4]
print(mass_frac_sulphur)

ppm_tiols = mass_frac_tiols * 10000
print(ppm_tiols)
ppm_sulphur = mass_frac_sulphur * 10000
print(ppm_sulphur)

min_flow_rate = database.iloc[18, 3]
print(min_flow_rate)
flow_rate = database.iloc[19, 3]
print(flow_rate)

MPS_calc_p, MPS_calc_t = database.iloc[55, 1], database.iloc[55, 2]
print(MPS_calc_p, MPS_calc_t)
MPS_work_p, MPS_work_t = database.iloc[56, 1], database.iloc[56, 2]
print(MPS_work_p, MPS_work_t)

LPS_calc_p, LPS_calc_t = database.iloc[60, 1], database.iloc[60, 2]
print(LPS_calc_p, LPS_calc_t)
LPS_work_p, LPS_work_t = database.iloc[61, 1], database.iloc[61, 2]
print(LPS_work_p, LPS_work_t)

water_direct_p, water_direct_t = database.iloc[65, 1], database.iloc[65, 2]
print(water_direct_p, water_direct_t)
water_reversed_p, water_reversed_t = database.iloc[75, 1], database.iloc[75, 2]
print(water_reversed_p, water_reversed_t)

LPG_Nitrogen_calc_p, LPG_Nitrogen_calc_t = database.iloc[105, 1], database.iloc[105, 2]
print(LPG_Nitrogen_calc_p, LPG_Nitrogen_calc_t)
LPG_Nitrogen_work_p, LPG_Nitrogen_work_t = database.iloc[106, 1], database.iloc[106, 2]
print(LPG_Nitrogen_work_p, LPG_Nitrogen_work_t)

HPG_Nitrogen_calc_p, HPG_Nitrogen_calc_t = database.iloc[112, 1], database.iloc[112, 2]
print(HPG_Nitrogen_calc_p, HPG_Nitrogen_calc_t)
HPG_Nitrogen_work_p, HPG_Nitrogen_work_t = database.iloc[113, 1], database.iloc[113, 2]
print(HPG_Nitrogen_work_p, HPG_Nitrogen_work_t)

air_calc_p, air_calc_t_min, air_calc_t_max = database.iloc[119, 1], database.iloc[119, 2], database.iloc[119, 3]
print(air_calc_p, air_calc_t_min, air_calc_t_max)
air_work_p, air_work_t = database.iloc[120, 1], database.iloc[120, 2]
print(air_work_p, air_work_t)


#-----------------------------------------------------------------------------------------------------------------------
ch_10 = head_counter.increment()

heading = doc.add_heading(f'{ch_10:.0f} МАТЕРИАЛЬНЫЙ БАЛАНС ПРОЦЕССА', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

text = [f'',
        f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

text = [f'Исходные данные для расчета материального баланса',
        f'',
        f'Материальный баланс установки демеркаптанизации керосиновой фракции («Demerus-Jet») составлен в соответствии со следующим расчетом:',
        f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для альбомной ориентации
if new_section.page_width < new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

table10_1 = table_counter.increment()
table10_2 = table_counter.increment()
table10_3 = table_counter.increment()

df10_1, merged_ranges = read_excel_with_merged_cells('database.xlsx', '10.1')
add_header(doc, f'Таблица {table10_1:.1f} – Материальный баланс установки демеркаптанизации керосиновой фракции')
add_table(doc, df10_1, merged_ranges)

#-----------------------------------------------------------------------------------------------------------------------

# Сохраняем документ
doc.save('test.docx')
