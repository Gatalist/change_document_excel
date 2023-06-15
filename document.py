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
            print("\n\n⛔️ No file to directory\n")
            sys.exit()

    def __init__(self, read_document: str, sheet_name: str):
        self.read_document: str = read_document
        self.data_frame: object = pandas.read_excel(self.read_document) # получаем data_frame документа в pandas
        self.work_book: object = load_workbook(self.read_document)  # открываем документ в openpyxl
        self.work_sheet: str = self.work_book[sheet_name] # получаем рабочий лист в документе
        self.len_strings = len(self.data_frame.index) + 1 # получаем список строк в документе
        self.start_number_string = 2 # с какой строки начинать
        self.list_len_string = [string for string in range(2, self.len_strings + 1)] # список строк

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
            print('\n\n💾 Save [ Document ]\n')
        except:
            print('\n\n⛔️ Failed to save [ Document ]\n')


class ReadDocument:
    # проверяем есть ли столбец
    def __new__(cls, *args, **kwargs):
        instance = super().__new__(cls)
        read_document = kwargs['document']
        column_name = kwargs['column_name']
        all_columns = read_document.data_frame.columns.values.tolist() # получаем список колонок
       
        if column_name in all_columns:
            print(f"\n✅ Read column [{column_name}]\n")
            return instance
        else:
            print(f"\n\n⛔️ No column with that name [{column_name}]\n")
            sys.exit()

    def __init__(self, document: object, column_name: str):
        self.read_document = document
        self.column_name: str = column_name

    # возвращаем список нумерованых строк с данными в указаной ячейке
    def generate_list_data_row(self) -> list:
        data_rows = []
        start_number_string = 2
        for row in self.read_document.data_frame[self.column_name]:
            row_to_line = []
            if type(row) == str:
                row_to_line.append(start_number_string)
                row_to_line.append(row)
                data_rows.append(row_to_line)
            start_number_string += 1
        return data_rows

    # Выводим данные ячейки
    # Аргумент "read_line" разобъет текст на строки по символу ";"
    def read_list_data_row(self, list_data: str, read_line: bool):
        number = 0
        if len(list_data) > 0:
            for number_string, text in list_data:              
                print(f'-----[ Строка: {number_string} ]-----')
                if read_line:
                    call_data = text.split(';')
                    for line in call_data:
                        print(line + '\n\n')
                else:
                    print(text + '\n\n')
                number += 1
        else:
            print("⛔️ Колонка пустая")

        print("Прочитано: ", number, "строк")


class SearchText:
    def __init__(self, list_data: list, pause: bool = None):
        self.list_data = list_data
        self.pause = pause

    # проверяем строку на первую заглавную букву
    def checking_is_title(self) -> None:
        errors = 0
        for number_string, text in self.list_data:            
            for line in text.split(';'):
                try:
                    first_leter = line[0]
                    if first_leter.istitle() == False and first_leter[0].isdigit() == False:
                        print(f'-----[ Строка {number_string} ]-----')
                        print(line + "\n")
                        if first_leter == ' ':
                            print("\n❌ [ Удалите символ ' ' в начале строки ]\n")
                        if self.pause:
                            input("Next -> ")
                        errors += 1

                except IndexError:
                    print(f'-----[ Строка {number_string} ]-----')
                    print(line + "\n")
                    print("\n❌ Обнаружен символ в [начале или конце] строки -> \n")
                    if self.pause:
                        input("Next -> ")
                    errors += 1

        print("Ошибок ->", errors, '\n')

    # ищим фрагмент текста в ячейке
    def serch_text(self, text: str) -> None:
        result = 0
        for number_string, text in self.list_data:            
            for line in text.split(';'):
                search_word_lower = [word.lower() for word in text]
                for search in search_word_lower:
                    if (line.lower()).find(search) != -1:
                        print(f'\n-----[ Line  {number_string} ]-----')
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
        self.symbols = [';;', '; ;', '  ', '   ']

    # делаем первую букву каждой строки заглавной
    def upper_first_letter_in_text(self, text: str) -> str:
        split_text = text.split(';')
        new_list = []
        for line in split_text:
            capitalized = line[0:1].upper() + line[1:]
            new_list.append(capitalized)
        return ';'.join(new_list)
    
    # заменяем сымволы в строке
    def replace_symbol(self, text: str) -> str:
        for symbol in self.symbols:
            text.replace(symbol, ';')
        text.strip()

        if len(text) > 0:
            if text[0] == ';':
                text = text[1:]
        if len(text) > 0:
            if text[-1] == ';':
                text = text[:-1]
        return text.strip()
    
    # получаем обьект ячейки по координатам: например 'AC4'
    def get_cell_obj(self, cell_letter: str, cell_number: str) -> object:
        cell = f'{cell_letter}{cell_number}'
        return self.document.work_sheet[cell]

    # сохраняем результат в ячейку
    def save_result_in_cell(self, cell_object: object, text: str) -> object:
        self.document.work_sheet[cell_object.coordinate] = text
        return self.document
    
    # добавление фрагмента текста в ячейку
    def add_text_to_cell(self, cell_object: object, text: str) -> object:
        if text:
            cell_object_value = cell_object.value
            if cell_object_value is not None:
                cell_object_value = f'{cell_object_value};{text}'
            else:
                cell_object_value = text

            clear_txt = self.replace_symbol(cell_object_value)
            upper_first_letter = self.upper_first_letter_in_text(clear_txt)
            self.save_result_in_cell(cell_object, upper_first_letter)
        else:
            self.save_result_in_cell(cell_object, None)
        return self.document

    # добавление фрагмента текста в начало ячейки если она не пустая
    def add_text_to_cell_stert(self, cell_past: str, text: str):
        for number_string in self.document.list_len_string:
            cell_past_obj = self.get_cell_obj(cell_past, number_string)
            if cell_past_obj.value is not None:
                new_text = text + cell_past_obj.value
                self.save_result_in_cell(cell_past_obj, new_text)

        print(f"✅ Text add to start in [{cell_past}]\n")

    # обьеденяем ячейки в одну
    def join_columns_text(self, save_column: str, join_columns: list, join_separator: str, end_text: str) -> None:
        for number_string in self.document.list_len_string:
            new_list_join = [col_row + str(number_string) for col_row in join_columns]
            new_data = []

            for column in new_list_join:
                old_data = self.get_cell_obj(column).value.replace(end_text, '').strip()
                new_data.append(old_data)
            
            new_text = join_separator.join(new_data)
            if new_text:
                new_text = new_text + ' ' + end_text
            
            cell_save = self.get_cell_obj(save_column, number_string)
            self.add_text_to_cell(cell_save, new_text)

    # добавление фрагмента текста в каждую ячейку
    def add_text_to_column(self, cell_past: str, text: str) -> None:
        for number_string in self.document.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_past, number_string)
            self.add_text_to_cell(cell_move_obj, text)

        print(f"✅ Text added to each cell [{cell_past}]\n")

    # удалить фрагмент текста со всех ячейк в столбце
    def remove_text_from_cell(self, cell_remove: str, text: str) -> None:
        for number_string in self.document.list_len_string:
            cell_remove_obj = self.get_cell_obj(cell_remove, number_string)
            if cell_remove_obj.value is not None:
                new_data = cell_remove_obj.value.replace(text, '')
                self.add_text_to_cell(cell_remove_obj, None)
                self.add_text_to_cell(cell_remove_obj, new_data)

        print(f"✅ Text dell in [{cell_remove}]\n")
    
    # удаяем текст поиска с ячейки и добавляем в другую ячейку
    def serch_move_text_to_another_cell(self, cell_move: str, cell_past: str, method_remove: str, search: list) -> None:

        for number_string in self.document.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_move, number_string)
            search_text_lower = [word.lower() for word in search]
            
            if method_remove == 'str' and cell_move_obj.value is not None:
                current_text = cell_move_obj.value
                cell_past_list_new_text = []

                # перебераем список совпадений
                for search_word in search_text_lower:
                    # перебераем список строк
                    for line in cell_move_obj.value.split(';'):
                        if line.lower().find(search_word) != -1:
                            if line not in cell_past_list_new_text:
                                cell_past_list_new_text.append(line)
                            # save text in current cell
                            current_text = cell_move_obj.value.replace(line, '')
                            self.add_text_to_cell(cell_move_obj, None)
                            self.add_text_to_cell(cell_move_obj, current_text)

                # добавление фрагмента текста в ячейку
                cell_past_txt_add = ";".join(cell_past_list_new_text)
                cell_past_obj = self.get_cell_obj(cell_past, number_string)
                self.add_text_to_cell(cell_past_obj, cell_past_txt_add)
                print(f'\n\n-----[ Line  {number_string} ]-----')
                print(cell_move_obj.value)     
            
            if method_remove == 'txt':
                if cell_move_obj.value is not None and cell_move_obj.value.lower().find(search_text_lower[0]) != -1:
                    print(f'\n\n-----[ Line  {number_string} ]-----')
                    print(cell_move_obj.value)

                    # удалить фрагмент текста в ячейке
                    cell_move_txt = cell_move_obj.value.replace(search, '')
                    self.add_text_to_cell(cell_move_obj, cell_move_txt)

                    # добавление фрагмента текста в ячейку
                    cell_past_obj = self.get_cell_obj(cell_past, number_string)
                    self.add_text_to_cell(cell_past_obj, search_text_lower[0])

    # вырезаем весь текст с одной ячееки и добавляем в другую ячейку
    def move_text_to_another_cell(self, cell_move: str, cell_past: str) -> None:
        for number_string in self.document.list_len_string:
            cell_move_obj = self.get_cell_obj(cell_move, number_string)
            # вырезаем данные с ячейки если она не пустая
            if cell_move_obj.value is not None:
                cell_past_obj = self.get_cell_obj(cell_past, number_string)
                # вставляем данные в другую ячейку
                self.add_text_to_cell(cell_past_obj, text=cell_move_obj.value)

            # очищаем ячейку откуда копируем текст
            self.add_text_to_cell(cell_move_obj, text=None)
        
        print(f"✅ Text dell cell [{cell_move}] and add cell [{cell_past}]\n")
