from document import OpenDocument, SaveDocument, ReadDocument, SearchText, ChangeDocument


# указываем название файла и название листа
document = OpenDocument(read_document='smart-chasy.xlsx', sheet_name='Worksheet')
print(document)
# print(document.list_len_string)
# указываем название колонки
read_document = ReadDocument(document=document, column_name='Интерфейсы и подключение-25840')

# получаем список строк с данными
list_data = read_document.list_data_row()

search = SearchText(list_data=list_data, pause=True)
change = ChangeDocument(document=document, list_data=list_data, pause=True)
save = SaveDocument(document=document)


# Выводим данные ячейки
read_document.read_list_data(list_data=list_data, read_line=True)
# read_document.get_len_strings()

# Для проверки на заглавную букву в строке
# search.checking_is_title()


# указываем текст или слова по отдельности для поиска
# search.serch_text(text=["NFC"])


# добавляем текст в начало ячейки
# change.add_data_start(cell_past='AE', text="В режиме ожидания: ")


# добавляем текст во все ячейки в столбце
# change.add_data_to_colums(cell_past='AG', text=";-2-test")


# удаляем текст со всех ячейк в столбце
# change.delete_data_to_column(cell_move='AI', text="При интенсивном использовании: До 24 ч")


# удаяем весь текст с одной ячееки и добавляем в другую ячейку
# change.move_to_other_cell(cell_move='AC', cell_past='AB')


# удаляем текст что ищем с одной ячейки и добавляем в другую
search_text = ['камеры', 'камера']

# change.serch_move_past(cell_move="AC", cell_past="AG", method_remove='str', search=search_text)
#

# обьяденяем данные столбцов в один столбец
# change.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')


# указываем название нового файла для сохранения
save.save(new_name="new")
