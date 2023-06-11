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
                print(f'-----[ Line  {row_number} ]-----')

                if read_line:
                    call_data = row_text.split(';')
                    for line in call_data:
                        print(line)
                else:
                    print(row_text)

                print()
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
                            print("\n❌ [ Удалите символ ' ' начале строки ]")
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

    # удаляем пробелы в начале, в конце строки и в между атрибутами
    def delete_space(self, call: str):
        for row in self.list_data:
            row_number = row[0]
            cell_col_row = f'{call}{row_number}'

            old_data = str(self.document.work_sheet[cell_col_row].value)
            new_data = old_data.replace("; ", ";").strip()
            self.document.work_sheet[cell_col_row] = new_data
            print(f"✅ Space delete [{call}]\n")

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
    def add_data_to_colums(self, letters_colunm: str, text: str):
        for row in self.list_data:
            row_number = row[0]
            row_text = row[1]
            cell_col_row = f'{letters_colunm}{row_number}'

            current_text_cell = self.document.work_sheet[cell_col_row]
            # print(row_number, type(row_text))

            if type(row_text) == str:
                self.document.work_sheet[cell_col_row] = str(current_text_cell.value) + ";" + text
            else:
                self.document.work_sheet[cell_col_row] = text

        print(f"✅ Text add in [{letters_colunm}]\n")

    # удалить фрагмент текста у всех ячейках
    def delete_data_to_column(self, letters_colunm: str, text: str):
        for row in self.list_data:
            row_number = row[0]
            row_text = row[1]
            cell_col_row = f'{letters_colunm}{row_number}'

            if type(row_text) == str:
                current_text_cell = str(self.document.work_sheet[cell_col_row].value)
                new_data = current_text_cell.replace(text, '').replace(';;', ';').lstrip()
            try:
                if new_data[0] == ';':
                    # new_data[0] = new_data[0].replace(';', '')
                    new_data[0] = ''
            except IndexError:
                pass

            try:
                if new_data[-1] == ';':
                    # new_data = new_data.replace(';', '')
                    new_data[-1] = ''
            except IndexError:
                pass

            self.document.work_sheet[cell_col_row] = new_data.lstrip()

        print(f"✅ Text dell in [{letters_colunm}]\n")

    # удаяем текст поиска с ячеек и добавляем в другую ячейку
    def serch_add_dell(self, search: str, add_colunm, dell_colunm):
        for row in self.list_data:
            row_number = row[0]
            row_text = row[1]

            # ищим фрагмент текста в ячейке
            search_text_lower = search.lower()
            if type(row_text) == str and row_text.lower().find(search_text_lower) != -1:
                print(f'\n-----[ Line  {row_number} ]-----')
                print(row_text)

                # удалить фрагмент текста в ячейке
                dell_cell = f'{dell_colunm}{row_number}'
                dell_cell_data = str(self.document.work_sheet[dell_cell].value)
                dell_cell_data_new = dell_cell_data.replace(search, '').replace(';;', ';').lstrip()
                self.document.work_sheet[dell_cell] = dell_cell_data_new
            
                # # добавление фрагмента текста в ячейку    
                add_cell = f'{add_colunm}{row_number}'       
                add_cell_data = str(self.document.work_sheet[add_cell].value)
                if type(row_text) == str:
                    self.document.work_sheet[add_cell] = add_cell_data + ";" + search
                else:
                    self.document.work_sheet[add_cell] = search
