import sys
import os
import pandas
from openpyxl import load_workbook


class Document():
    # проверяем существует ли указаный файл
    def __new__(cls, *args, **kwargs):
        instance = super().__new__(cls)
        this_directory = os.getcwd()
        path_document = os.path.join(this_directory, kwargs['read_document'])

        if os.path.exists(path_document):
            return instance
        else:
            print('\n', path_document)
            print("\n ⛔️ No file to directory\n")
            sys.exit()
    
    def __init__(self, read_document: str, sheet_name: str):
        self.read_document: str = read_document
        self.data_frame: object = pandas.read_excel(self.read_document)
        self.work_book: object = load_workbook(self.read_document)
        self.work_sheet: str = self.work_book[sheet_name]
        self.data_column = None
        self.is_title = False
        self.search_text = None
        self.add_text_cell = None
        self.dell_text_cell = None

    
    @staticmethod
    def print_text(text: str, color_name: str = None, pause: bool = None) -> None:
        if color_name:
            color = color_name
            if color_name == 'orange':
                color = f"\033[33m{text}\033[0m"
            if color_name == 'blue':
                color = f"\033[34m{text}\033[0m"
            else:
                color = text
        else:
            color = text

        if pause:
            print(input(color))
        else:
            print(color)
    
    # возвращаем список нумерованых строк с данными в указаной ячейке
    def return_list_data_row(self, column_name: str) -> None:
        all_rows_to_column = []
        if isinstance(self.data_frame, pandas.core.frame.DataFrame):
            self.print_text(text="\nColumn:", color_name="orange")
            self.print_text(text=column_name)
            print()
            number_row = 2
            for row in self.data_frame[column_name]:
                if isinstance(row, str):
                    row_to_line = []
                    row_to_line.append(number_row)
                    row_to_line.append(row)
                    all_rows_to_column.append(row_to_line)
                number_row += 1
            self.data_column = all_rows_to_column
            # return self.data_column
    
    # проверяем строку на первую заглавную букву
    def is_title_line(self, line_text: str) -> None:
        try:
            first_leter = line_text[0]
            if first_leter.istitle() == False and first_leter[0].isdigit() == False:
                self.print_text(text="\n✅ Litle letter to start string -> ", color_name="orange", pause=True)

        except IndexError:
            self.print_text(text="\n✅ Обнаружен символ в [начале или конце] строки -> ", color_name="orange", pause=True)

    # ищим фрагмент текста в ячейке
    def serch_to_line(self, line_text: str, pause: bool = False) -> None:
        search_word_lower = [word.lower() for word in self.search_text]
        for search in search_word_lower:
            if (line_text.lower()).find(search) != -1:
                print()
                self.print_text(text=f"✅ Найдено совпадение: -> {search}", color_name="orange", pause=pause)
                return True

    # удалить фрагмент текста в ячейке
    def delete_data_to_column(self, letters_colunm, number_row):
        cell_col_row = f'{letters_colunm}{number_row}'
        old_data = str(self.work_sheet[cell_col_row].value)
        new_data = old_data.replace(self.search_text[0], '').replace(';;', ';').lstrip()
        
        try:
            if new_data[0] == ';':
                new_data = new_data.replace(';', '')
        except IndexError:
            pass
        
        try:
            if new_data[-1] == ';':
                new_data = new_data.replace(';', '')
        except IndexError:
            pass

        self.work_sheet[cell_col_row] = new_data.lstrip()
    
    # добавление фрагмента текста в ячейку
    def add_data_to_colums(self, letters_colunm, number_row):
        cell_col_row = f'{letters_colunm}{number_row}'
        current_text_cell = str(self.work_sheet[cell_col_row].value)
        if current_text_cell != "None":
            self.work_sheet[cell_col_row] = current_text_cell + ";" + str(self.search_text[0])
        else:
            self.work_sheet[cell_col_row] = str(self.search_text[0])

    # проверяем есть ли данные в столбце
    def check_data_to_column(self):
        if len(self.data_column) < 1:
            print("⛔️ Not data to this column\n")

    # работа с одним атриботов в тексте ячейке
    def parse_atributeto_row(self, row_text):
        atribute_to_row = row_text.split(';')
        for atribute in atribute_to_row:
            print(atribute)
            if self.is_title:
                self.is_title_line(atribute)
            if self.search_text:
                if self.add_text_cell and self.dell_text_cell:
                    self.serch_to_line(atribute)
                else:
                    self.serch_to_line(atribute, pause=True)

    # работа со всем текстом в ячейке
    def parse_text_row(self, row_number, row_text):
        print(row_text)
        if self.is_title:
            self.is_title_line(row_text)
        if self.search_text:
            if self.add_text_cell and self.dell_text_cell:
                if self.serch_to_line(row_text):
                    self.add_data_to_colums(self.add_text_cell, row_number)
                    self.delete_data_to_column(self.dell_text_cell, row_number)
            else:
                self.serch_to_line(row_text, pause=True)

    # парсим данные со всех строк ячейки
    def parse_data_column(self, how_to_parse: str):
        self.check_data_to_column()
        for row in self.data_column:
            row_number = row[0]
            row_text = row[1]
            print(row_number)
            self.print_text(text="-"*15)

            if how_to_parse == 'line':
                self.parse_atributeto_row(row_text)
            if how_to_parse == 'text':
                self.parse_text_row(row_number, row_text)
            
            print()
    
    # save new file
    def save_new_file(self, new_name_file):
        if self.add_text_cell and self.dell_text_cell:
            try:
                output_file = new_name_file + ".xlsx"
                self.work_book.save(output_file)
                print('💾 Document [Save]\n')
            except:
                print('⛔️ Error [Save]\n')
        else:
            print("✅ No data to save\n")



# указываем название файла и название листа
document = Document(read_document='new_name_file.xlsx', sheet_name='Worksheet')
print('✅', document.read_document)

# указываем название колонки для проверки
document.return_list_data_row(column_name='Особенности-25704')

# Для проверки на заглавную букву в строке  ## (не обязательно)
# document.is_title = True

# указываем текст или слова по отдельности для поиска  ## (не обязательно)
document.search_text = ['Стерео-динамики (2 шт)']

# удаляем текст поиска с ячейки  ## (не обязательно)
# document.dell_text_cell = 'AV'

# добавляем текст поиска в ячейку  ## (не обязательно)
# document.add_text_cell = 'AU'

# парсим данные со всех строк ячейки
document.parse_data_column(how_to_parse='text')

# указываем название нового файла для сохранения
document.save_new_file('new_name_file_2')
