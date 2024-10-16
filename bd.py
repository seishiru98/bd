from docx import Document
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION, WD_ORIENT

import pandas as pd

import datetime

# Создаем новый документ
doc = Document()

section = doc.sections[0]

# Установка полей страницы
section.left_margin = Cm(2)  # Левое поле
section.right_margin = Cm(1)  # Правое поле
section.top_margin = Cm(2)  # Верхнее поле
section.bottom_margin = Cm(2)  # Нижнее поле

# Функция для изменения шрифта и размера шрифта
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


# Функция для установки отступов и междустрочного интервала
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
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

def add_table(doc, df, start_row, end_row):
    table = doc.add_table(rows=1, cols=len(df.columns))
    table.style = 'Table Grid'

    hdr_cells = table.rows[0].cells
    for i, column_name in enumerate(df.columns):
        cell_paragraph = hdr_cells[i].paragraphs[0]
        cell_paragraph.text = column_name
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

    for index in range(start_row, end_row):
        row = df.iloc[index]
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            cell_paragraph = row_cells[i].paragraphs[0]
            cell_paragraph.text = str(value)
            cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
            for run in cell_paragraph.runs:
                set_font(run, 'Times New Roman', 12)
            set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                                 line_spacing=18, space_after=0, space_before=0)

        for col_idx in range(len(df.columns)):
            merge_start = 2
            for row_idx in range(2, len(table.rows) + 1):
                cell = table.cell(row_idx - 1, col_idx)
                if cell.text == "":
                    continue
                if merge_start < row_idx - 1:
                    table.cell(merge_start - 1, col_idx).merge(table.cell(row_idx - 2, col_idx))
                merge_start = row_idx
            if merge_start < len(table.rows):
                table.cell(merge_start - 1, col_idx).merge(table.cell(len(table.rows) - 1, col_idx))

def insert_page_break(doc):
    doc.add_page_break()

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

text = ['ООО «НТЦ «Ахмадуллины»',
        '']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 18)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

months_in_russian = {
    'January': 'Январь',
    'February': 'Февраль',
    'March': 'Март',
    'April': 'Апрель',
    'May': 'Май',
    'June': 'Июнь',
    'July': 'Июль',
    'August': 'Август',
    'September': 'Сентябрь',
    'October': 'Октябрь',
    'November': 'Ноябрь',
    'December': 'Декабрь'
}

current_date = datetime.datetime.now()

day = current_date.day
month_english = current_date.strftime("%B")
month = months_in_russian[month_english]
year = current_date.year

text = ['УТВЕРЖДАЮ',
        'Генеральный директор',
        '__________Р.М. Ахмадуллин',
       f'«{day}» {month} {year} года',
        '']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

text = ['',
        'Базовый проект',
        'установки демеркаптанизации керосиновой фракции АО ««ННК-Хабаровский Нефтеперерабатывающий завод» по технологии «Demerus Jet»',
        '',
        '']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14, bold=False)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

text = ['']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14, bold=True)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

text = ['В настоящем документе содержится конфиденциальная информация относительно технологии «Demerus-Jet», включая эксплуатационные условия и технологические возможности, которые не могут быть раскрыты неуполномоченным лицам. Представленные материалы являются собственностью Лицензиара. Получая настоящую информацию, вы соглашаетесь не использовать ее ни для каких других целей, кроме тех, которые согласованы с Лицензиаром в письменной форме, не воспроизводить этот документ полностью или частично и не раскрывать его содержимое третьим лицам без письменного разрешения Лицензиара.',
        '',
        '']

# Добавление абзацев и установка их форматирования
for line in text:
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    run = paragraph.add_run(line)
    set_font(run, 'Times New Roman', 12, italic=True)  # Установить курсив
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=18,
                         space_after=0, space_before=0)

table = doc.add_table(rows=6, cols=3)
table.style = 'Table Grid'

header_text = ['№ п/п', 'Ревизия', 'Дата выдачи']

hdr_cells = table.rows[0].cells
for i, text in enumerate(header_text):
    cell_paragraph = hdr_cells[i].paragraphs[0]
    cell_paragraph.text = text
    cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in cell_paragraph.runs:
        set_font(run, 'Times New Roman', 14, bold=True)
    set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                         line_spacing=22, space_after=0, space_before=0)

text = ['',
        '',
        '',
        '',
        f'Казань – {year}']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_1 = head_counter.increment()

heading = doc.add_heading(f'{ch_1:.0f} ВВЕДЕНИЕ', level=1)
heading.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
for run in heading.runs:
    set_font(run, 'Times New Roman', 14)
    set_paragraph_format(heading, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

min_flow_rate = 26.8
flow_rate = 45.3
work_time = 8760

text = ['',
        '',
        'Настоящий Базовый проект на проектирование установки очистки керосиновой фракции от меркаптанов и кислых примесей выполнен в соответствии с договором № 61 от 26 марта 2020г. для АО ««ННК-Хабаровский Нефтеперерабатывающий завод».',
        'Керосиновые фракции установок ЭЛОУ-АТ и ЭЛОУ-АВТ АО «ННК-Хабаровский Нефтеперерабатывающий завод» отличаются повышенным содержанием меркаптановой серы (от 300,0 до 400,0 ppm), не соответствующим требованиям ГОСТ 10227-86 «Топлива для реактивных двигателей» на авиационное топливо марки ТС-1 (не более 30,0 ppm по меркаптановой сере, сероводород - отсутствие).',
        'Меркаптановая сера в керосиновой фракции представлена высокомолекулярными соединениями, трудно извлекаемыми водно-щелочными растворами. Поэтому процесс их демеркаптанизации сводится к дезодорации содержащихся в них коррозионно-активных меркаптанов путем их окисления в инертные дисульфиды, остающиеся в очищаемом топливе. Такой подход оправдан приемлемым содержанием общей серы в этих фракциях.',
        f'Блок щелочной демеркаптанизации керосиновой фракции предназначен для окисления меркаптановых соединений и рассчитан на переработку по номинальной производительности {flow_rate} т/ч.',
        'В состав блока щелочной очистки «Demerus-Jet» входят:',
        '– узел окислительной демеркаптанизации керосиновой фракции;',
        '– узел адсорбционной очистки керосиновой фракции;',
        '– узел регенерации и концентрирования промотора;',
        '– реагентное хозяйство.',
        f'Режим работы блока демеркаптанизации керосиновой фракции «Demerus-Jet» - непрерывный, {work_time} часов в год. Расчетный период непрерывной эксплуатации установки между остановками на капитальный ремонт – 24 месяца. Срок службы оборудования не менее 20 лет. При расчете и подборе оборудования, согласно заданию на проектирование, был принят диапазон устойчивой производительности от {min_flow_rate} до {flow_rate} т/ч.',
        '']

for line in text:
    paragraph = doc.add_paragraph(line)
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=1.25, line_spacing=22,
                         space_after=0, space_before=0)

#-----------------------------------------------------------------------------------------------------------------------
# Добавление нового раздела
new_section = doc.add_section(WD_SECTION.NEW_PAGE)

# Установка ориентации
new_section.orientation = WD_ORIENT.PORTRAIT

# Убедимся, что размеры страницы корректны для книжной ориентации
if new_section.page_width > new_section.page_height:
    new_section.page_width, new_section.page_height = new_section.page_height, new_section.page_width

ch_8 = head_counter.increment()

heading = doc.add_heading(f'{ch_8:.0f} УСЛОВИЯ ПРОВЕДЕНИЯ ПРОЦЕССА', level=1)
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

df8_1 = read_excel_data('database.xlsx', '8.1')
df8_1 = df8_1.fillna('')

table8_1 = table_counter.increment()
header_text = f'Таблица {table8_1} – Условия проведения процесса демеркаптанизации'

rows_per_page_first = 100  # Количество строк для первой таблицы
rows_per_page_next = 100  # Количество строк для следующих таблиц

total_rows = len(df8_1)
start_row = 0

# Первая таблица
end_row = min(start_row + rows_per_page_first, total_rows)
add_header(doc, header_text)
add_table(doc, df8_1, start_row, end_row)
start_row = end_row

# Последующие таблицы
while start_row < total_rows:
    end_row = min(start_row + rows_per_page_next, total_rows)
    insert_page_break(doc)
    add_header(doc, header_text)
    add_table(doc, df8_1, start_row, end_row)
    start_row = end_row

text = [f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

# Сохраняем документ
doc.save('Мат баланс.docx')
