
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
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY


def add_table(doc, df, merged_ranges):
    # Общая ширина таблицы в сантиметрах
    total_width = Cm(18.5)  # Примерная ширина текста на странице A4 с полями

    # Определяем таблицу и стиль
    table = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    table.style = 'Table Grid'

    # Игнорируем первую строку, чтобы определить максимальную длину текста в каждой колонке
    column_widths = [0] * len(df.columns)

    # Добавляем строки с данными и определяем максимальную длину текста в каждом столбце
    for index, row in df.iterrows():
        for i, value in enumerate(row):
            if pd.notna(value) and str(value).strip() != "":
                column_widths[i] = max(column_widths[i], len(str(value)))

    # Вычисляем общую длину текста для пропорциональной настройки ширины столбцов
    total_text_length = sum(column_widths)

    # Добавление заголовков таблицы
    hdr_cells = table.rows[0].cells
    for i, column_name in enumerate(df.columns):
        hdr_cells[i].width = Cm(total_width.cm * (column_widths[i] / total_text_length))  # Задаем ширину на основе данных
        cell_paragraph = hdr_cells[i].paragraphs[0]
        cell_paragraph.text = column_name
        cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
        for run in cell_paragraph.runs:
            set_font(run, 'Times New Roman', 12)  # Шрифт для заголовков
        set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                             line_spacing=18, space_after=0, space_before=0)

    # Добавление строк таблицы
    for index, row in df.iterrows():
        row_cells = table.rows[index + 1].cells
        for i, value in enumerate(row):
            row_cells[i].width = Cm(total_width.cm * (column_widths[i] / total_text_length))  # Задаем ширину для ячеек на основе данных
            if pd.notna(value) and str(value).strip() != "":
                cell_paragraph = row_cells[i].paragraphs[0]
                cell_paragraph.text = str(value)
                cell_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
                for run in cell_paragraph.runs:
                    set_font(run, 'Times New Roman', 12)
                set_paragraph_format(cell_paragraph, left_indent=0.0, right_indent=0.0, first_line_indent=0.0,
                                     line_spacing=18, space_after=0, space_before=0)

    # Объединение ячеек в Word на основе объединённых диапазонов из Excel
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if row == min_row and col == min_col:
                    start_cell = table.cell(min_row - 1, min_col - 1)
                    end_cell = table.cell(max_row - 1, max_col - 1)
                    start_cell.merge(end_cell)


def insert_page_break(doc):
    doc.add_page_break()


def read_excel_with_merged_cells(filename, sheet_name):
    # Загружаем Excel-файл с помощью openpyxl
    wb = load_workbook(filename, data_only=True)
    ws = wb[sheet_name]

    # Сохраняем данные в список строк (list of lists)
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append(list(row))

    # Получаем объединенные ячейки
    merged_ranges = ws.merged_cells.ranges

    # Проходим по каждому объединенному диапазону и заполняем его данные
    for merged_range in merged_ranges:
        # Получаем диапазон объединённых ячеек
        min_col, min_row, max_col, max_row = merged_range.bounds

        # Получаем значение из первой ячейки диапазона
        merged_value = ws.cell(row=min_row, column=min_col).value

        # Присваиваем это значение только первой ячейке, остальные оставляем пустыми
        for row in range(min_row - 1, max_row):
            for col in range(min_col - 1, max_col):
                if row == min_row - 1 and col == min_col - 1:
                    data[row][col] = merged_value
                else:
                    data[row][col] = ''  # Оставляем остальные ячейки пустыми

    # Преобразуем данные в DataFrame pandas
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
# main
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

ch_5 = head_counter.increment()

heading = doc.add_heading(f'{ch_5:.0f} ХАРАКТЕРИСТИКА ИСХОДНОГО СЫРЬЯ, ПРОДУКТОВ, ОСНОВНЫХ И ВСПОМОГАТЕЛЬНЫХ МАТЕРИАЛОВ', level=1)
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

text = [f'Исходным сырьем блока демеркаптанизации керосиновой фракции "Demerus Jet" является прямогонный дистиллят (керосиновая фракция) в количестве от {min_flow_rate} до {flow_rate} т/ч и содержанием меркаптановой серы до {mass_frac_tiols}% мас. ({ppm_tiols} ppm).',
        f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

table5_1 = table_counter.increment()
table5_2 = table_counter.increment()
table5_3 = table_counter.increment()
table5_4 = table_counter.increment()
table5_5 = table_counter.increment()

df5_1, merged_ranges = read_excel_with_merged_cells('database.xlsx', '5.1')
add_header(doc, f'Таблица {table5_1:.1f} Физико-химические показатели качества сырья, поступающего на блок "Demerus Jet"')
add_table(doc, df5_1, merged_ranges)

text = [f'',
        f'',
        f'',
        f'',
        f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

df5_2, merged_ranges = read_excel_with_merged_cells('database.xlsx', '5.2')
add_header(doc, f'Таблица {table5_2:.1f} Характеристика керосиновой фракции - сырья блока «Demerus Jet»')
add_table(doc, df5_2, merged_ranges)

text = [f'',
        f'Целевым продуктом блока "Demerus Jet" является керосиновая фракция с массовой долей меркаптановой серы не более 30 ppm, сероводород – отсутствие. Концентрация общей серы остается без изменений в диапазоне 0,114÷0,116 % мас.',
        f'',
        f'К основным материалам относятся (характеристики представлены в таблице {table5_5:.1f}):',
        f'- гетерогенный катализатор КСМ-Х, изготавливаемый в соответствии с ТУ 2175-001-40655797-2014',
        f'- глина отбеливающая (бентонитовая); ',
        f'- γ – оксид алюминия по ТУ 6-09-426-75;',
        f'- шары фарфоровые номинальный диаметр шара 3 мм, изготовляются в соответствии с ТУ 4328-030-07608911-2015. Материал - фарфор по ГОСТ 20419-83;',
        f'- воздух сжатый (КИП, технологический);',
        f'- пар среднего давления;',
        f'- пар низкого давления;',
        f'- оборотная вода прямая;',
        f'- оборотная вода обратная;',
        f'- инертный газ низкого давления (Азот);',
        f'- инертный газ высокого давления (Азот);',
        f'- деминерализованная вода для приготовления водных растворов NaOH и КОН;',
        f'- промотор КСП(ж), соответствует ТУ 0258-015-00151638-ОП-99. В качестве промотора КСП (тв.) используется калия гидрат окиси твердый – КОН. Промотор КСП(ж) образуется в ходе эксплуатации установки из продуктов взаимодействия кислых примесей керосина с гидроксидом калия и кислородом воздуха на поверхности гетерогенного катализатора КСМ-Х. Необходимость в закупки КСП(ж) отсутствует. ',
        f'Промотор КСП(ж) представляет собой темно-коричневую жидкость с плотностью не менее 1,3 кг/дм3. При гравиметрическом отстаивании он расслаивается на два слоя: светлый тяжелый (КСП(ж)) и темный легкий (калиевые соли нафтеновых кислот). Хранится при температуре не ниже 5оС. ',
        f'В состав промотора КСП(ж) входит спектр органических кислых примесей, извлеченных щелочью из керосиновой фракции и окисленных на гетерогенном катализаторе КСМ-Х воздухом до алкилтиосульфонатов, солей сульфокислот и др. кислородсодержащих продуктов.',
        f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

df5_3, merged_ranges = read_excel_with_merged_cells('database.xlsx', '5.3')
add_header(doc, f'Таблица {table5_3:.1f} Характеристика керосиновой фракции - сырья блока «Demerus Jet»')
add_table(doc, df5_3, merged_ranges)

text = [f'']

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)

df5_4, merged_ranges = read_excel_with_merged_cells('database.xlsx', '5.4')
add_header(doc, f'Таблица {table5_4:.1f} – Физико-химические характеристики КСП (ж)')
add_table(doc, df5_4, merged_ranges)

text = [f''
       ]

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

df5_5, merged_ranges = read_excel_with_merged_cells('database.xlsx', '5.5')
add_header(doc, f'Таблица {table5_4:.1f} – Характеристика основных и вспомогательных материалов')
add_table(doc, df5_5, merged_ranges)

text = [f''
       ]

for line in text:
    paragraph_after_break = doc.add_paragraph(line)
    paragraph_after_break.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY
    for run in paragraph_after_break.runs:
        set_font(run, 'Times New Roman', 14)
    set_paragraph_format(paragraph_after_break, left_indent=0.0, right_indent=0.0, first_line_indent=1.25,
                         line_spacing=22, space_after=0, space_before=0)
#-----------------------------------------------------------------------------------------------------------------------

# Сохраняем документ
doc.save('БП.docx')
