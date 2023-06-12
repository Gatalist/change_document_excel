import sys
import os
import pandas
from openpyxl import load_workbook


class OpenDocument:
    # проверяем существует ли указаный файл
    def __new__(cls, *args, **kwargs):
        instance = super().__new__(cls)
        this_directory = os.getcwd()
        path_document = os.path.join(this_directory, kwargs['read_document'])

        if os.path.exists(path_document):
            return instance
        else:
            print('\n', path_document)
            print("⛔️ No file to directory\n")
            sys.exit()

    def __init__(self, read_document: str, sheet_name: str):
        self.read_document: str = read_document
        self.data_frame: object = pandas.read_excel(self.read_document)
        self.work_book: object = load_workbook(self.read_document)
        self.work_sheet: str = self.work_book[sheet_name]
    
    # вывод результата
    def __repr__(self) -> str:
        return f'\n✅ Open file [{self.read_document}]\n'
    

class SaveDocument:
    def __init__(self, document: object):
        self.document = document

    # save new file
    def save(self, new_name):
        try:
            output_file = new_name + ".xlsx"
            self.document.work_book.save(output_file)
            print('💾 Document [Save]\n')
        except:
            print('⛔️ Error [Save]\n')


class ReadDocument:
    # проверяем есть ли столбец
    def __new__(cls, *args, **kwargs):
        instance = super().__new__(cls)
        read_document = kwargs['document']
        column_name = kwargs['column_name']
        columns = read_document.data_frame.columns.values.tolist()

        if column_name in columns:
            print(f"✅ Read column [{column_name}]\n")
            return instance
        else:
            print(f"⛔️ Not column [{column_name}]\n")
            sys.exit()

    def __init__(self, document: object, column_name: str):
        self.read_document = document
        self.column_name: str = column_name

    # возвращаем список нумерованых строк с данными в указаной ячейке
    def list_data_row(self) -> list:
        data_rows = []
        start_number_row = 2
        for row in self.read_document.data_frame[self.column_name]:
            # if isinstance(row, str):
            row_to_line = []
            row_to_line.append(start_number_row)
            row_to_line.append(row)
            data_rows.append(row_to_line)
            start_number_row += 1
    
        return data_rows
    
    # Выводим данные ячейки
    # Аргумент "read_line" разобъет текст на строки по символу ";"
    def read_list_data(self, list_data: str, read_line: bool):
        if len(list_data) > 0:
            for row in list_data:
                row_number = row[0]
                row_text = row[1]
                
                if type(row_text) == str:
                    print(f'-----[ Line  {row_number} ]-----')

                    if read_line:
                        call_data = row_text.split(';')
                        for line in call_data:
                            print(line)
                    else:
                        print(row_text)

                    print('\n\n')
        else:
            print("⛔️ Not data this column")


class SearchText:
    def __init__(self, list_data: list, pause: bool = None):
        self.list_data = list_data
        self.pause = pause

    # проверяем строку на первую заглавную букву
    def checking_is_title(self) -> None:
        errors = 0
        for row in self.list_data:
            row_number = row[0]
            row_text = row[1]
            call_data = row_text.split(';')
            
            for line in call_data:
                try:
                    first_leter = line[0]
                    if first_leter.istitle() == False and first_leter[0].isdigit() == False:
                        print(f'-----[ Line  {row_number} ]-----')
                        print(line)
                        if first_leter == ' ':
                            print("\n❌ [ Удалите символ ' ' в начале строки ]")
                        if self.pause:
                            input("Next -> ")
                        errors += 1
                        print("\n")

                except IndexError:
                    print(f'-----[ Line  {row_number} ]-----')
                    print(line)
                    print("\n❌ Обнаружен символ в [начале или конце] строки -> ")
                    print("\n")
                    if self.pause:
                        input("Next -> ")
                    errors += 1
                    print("\n")

        print("Ошибок ->", errors, '\n')

    # ищим фрагмент текста в ячейке
    def serch_text(self, text: str) -> None:
        result = 0
        for row in self.list_data:
            row_number = row[0]
            row_text = row[1]
            call_data = row_text.split(';')
            
            for line in call_data:
                search_word_lower = [word.lower() for word in text]
                for search in search_word_lower:
                    if (line.lower()).find(search) != -1:
                        print(f'\n-----[ Line  {row_number} ]-----')
                        print(line)
                        print(f"\n✅ Найдено совпадение ->: {search}")
                        if self.pause:
                            input("Next -> ")
                        result += 1


class ChangeDocument:
    def __init__(self, document: object, list_data: list, pause: bool = None):
        self.document = document
        self.list_data = list_data
        self.pause = pause

    # получаем ячейку в столбце
    def get_cell_in_column(self, cell_letter, cell_number):
        cell = f'{cell_letter}{cell_number}'
        return self.document.work_sheet[cell]

    # сохраняем новые данные в ячейку
    def save_new_data_in_cell(self, cell_letter, cell_number, new_data):
        cell = f'{cell_letter}{cell_number}'
        self.document.work_sheet[cell] = new_data
        return self.document

    # заменяем сымволы в строке
    def replace_symbol(self, text: str):
        new_text = text.replace(';;', ';').replace('; ;', ';').replace('  ', ' ').replace('   ', ' ').strip()
        if len(new_text) > 0:
            if new_text[0] == ';':
                new_text = new_text[1:]
        if len(new_text) > 0:
            if new_text[-1] == ';':
                new_text = new_text[:-1]
        return new_text

    # обьеденяем ячейки в одну
    def join_columns_text(self, save_column: str, join_columns: list, join_separator: str, end_text):
        for row in self.list_data:
            row_number = row[0]
            new_list_join = [col_row + str(row_number) for col_row in join_columns]
            print(new_list_join)
        
            new_data = []
            for column in new_list_join:
                old_data = str(self.document.work_sheet[column].value).replace(end_text, '').strip()
                new_data.append(old_data)
            
            new_text = join_separator.join(new_data)
            if new_text:
                new_text = new_text + ' ' + end_text

            self.document.work_sheet[save_column + str(row_number)] = new_text
            print(new_text)

    # добавление фрагмента текста в каждую ячейку
    def add_data_to_colums(self, cell_past: str, text: str):
        for row in self.list_data:
            number_string = row[0]
            cell_move_obj = self.get_cell_in_column(cell_past, number_string)

            if cell_move_obj.value is not None:
                new_text = str(cell_move_obj.value) + text
                self.save_new_data_in_cell(cell_past, number_string, new_text)
            else:
                self.save_new_data_in_cell(cell_past, number_string, text)

        print(f"✅ Text add in [{cell_past}]\n")

    # удалить фрагмент текста у всех ячейках
    def delete_data_to_column(self, cell_move: str, text: str):
        for row in self.list_data:
            number_string = row[0]
            cell_move_obj = self.get_cell_in_column(cell_move, number_string)
            
            cell_move_obj_data = cell_move_obj.value
            if cell_move_obj_data is not None:
                new_data = cell_move_obj_data.replace(text, '')
                cell_move_txt_new = self.replace_symbol(new_data)

                self.save_new_data_in_cell(cell_move, number_string, cell_move_txt_new)

        print(f"✅ Text dell in [{cell_move}]\n")
    
    # удаяем текст поиска с ячейки и добавляем в другую ячейку 
    def serch_move_past(self, search: str, cell_move: str, cell_past: str):
        for row in self.list_data:
            number_string = row[0]
            cell_move_obj = self.get_cell_in_column(cell_move, number_string)

            search_text_lower = search.lower()

            if cell_move_obj.value is not None and cell_move_obj.value.lower().find(search_text_lower) != -1:
                print(f'\n\n-----[ Line  {number_string} ]-----')
                print(cell_move_obj.value)
                print("\n--- [ new text cell_move] ---\n")
                # удалить фрагмент текста в ячейке
                cell_move_txt = cell_move_obj.value.replace(search, '')
                cell_move_txt_new = self.replace_symbol(cell_move_txt)
                self.save_new_data_in_cell(cell_move, number_string, cell_move_txt_new)
                print(cell_move_txt_new)

                # # добавление фрагмента текста в ячейку
                cell_past_obj = self.get_cell_in_column(cell_past, number_string)
                if cell_move_obj.value is not None:
                    curent_text = cell_past_obj.value
                    if str(curent_text) == "None":
                        self.save_new_data_in_cell(cell_past, number_string, search)
                    else:
                        self.save_new_data_in_cell(cell_past, number_string, f"{cell_past_obj.value};{search}")
                else:
                    self.save_new_data_in_cell(cell_past, number_string, search)

    # добавление фрагмента текста в не пустую ячейку в начало
    def add_data_start(self, cell_past: str, text: str):
        for row in self.list_data:
            number_string = row[0]
            cell_past_obj = self.get_cell_in_column(cell_past, number_string)

            if cell_past_obj.value is not None:
                new_text = text + str(cell_past_obj.value)
                self.save_new_data_in_cell(cell_past, number_string, new_text)

        print(f"✅ Text add to start in [{cell_past}]\n")

    # удаяем весь текст с одной ячееки и добавляем в другую ячейку
    def move_to_other_cell(self, cell_move, cell_past):
        for row in self.list_data:
            number_string = row[0]

            cell_move_obj = self.get_cell_in_column(cell_move, number_string)
            # вырезаем данные с ячейки если она не пустая
            if cell_move_obj.value is not None:

                cell_past_obj = self.get_cell_in_column(cell_past, number_string)
                # вставляем данные в другую ячейку
                if cell_past_obj.value is not None:
                    self.save_new_data_in_cell(cell_past, number_string, f'{cell_past_obj.value};{cell_move_obj.value}')
                elif cell_past_obj.value is None:
                    self.save_new_data_in_cell(cell_past, number_string, cell_move_obj.value)

            # очищаем ячейку откуда копируем текст
            self.save_new_data_in_cell(cell_move, number_string, '')
        
        print(f"✅ Text dell cell [{cell_move}] and add cell [{cell_past}]\n")
