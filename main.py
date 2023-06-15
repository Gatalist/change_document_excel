from document import OpenDocument, SaveDocument, ReadDocument, SearchText, ChangeDocument


# указываем название файла и название листа
document = OpenDocument(read_document='smart-chasy.xlsx', sheet_name='Worksheet')
print(document)
# print(document.list_len_string)

# указываем название колонки
read_document = ReadDocument(document=document, column_name='Особенности-25849')
# получаем список строк с данными
list_data = read_document.generate_list_data_row()

search = SearchText(list_data=list_data, pause=True)
change = ChangeDocument(document=document, list_data=list_data, pause=True)
save = SaveDocument(document=document)

# Выводим данные ячейки
read_document.read_list_data_row(list_data=list_data, read_line=True)

# Для проверки на заглавную букву в строке
# search.checking_is_title()

# указываем текст или слова по отдельности для поиска
# search.serch_text(text=["NFC"])

# добавление фрагмента текста в начало ячейки если она не пустая
# change.add_text_to_cell_stert(cell_past='AD', text="мАч")

# добавление фрагмента текста в каждую ячейку в конец
# change.add_text_to_column(cell_past='AD', text="мАч")

# удалить фрагмент текста со всех ячейк в столбце
# change.remove_text_from_cell(cell_remove='AD', text="мАч")

# удаяем весь текст с одной ячееки и добавляем в другую ячейку
# change.move_text_to_another_cell(cell_move='AE', cell_past='AD')

# удаяем текст поиска с ячейки и добавляем в другую ячейку
search_text = ['мАч']
# change.serch_move_text_to_another_cell(cell_move="AD", cell_past="AG", method_remove='str', search=search_text)

# обьяденяем данные столбцов в один столбец
# change.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')

# указываем название нового файла для сохранения
save.save(new_name="new")
