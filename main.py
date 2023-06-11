from document import OpenDocument, SaveDocument, ReadDocument, SearchText, ChangeDocument


# указываем название файла и название листа
document = OpenDocument(read_document='smart-chasy-2.xlsx', sheet_name='Worksheet')
print(document)

# указываем название колонки
read_document = ReadDocument(document=document, column_name='Защита от влаги и пыли-11111')

# получаем список строк с данными
list_data = read_document.list_data_row()

search = SearchText(list_data=list_data, pause=True)
change = ChangeDocument(document=document, list_data=list_data, pause=True)
save = SaveDocument(document=document)


# Выводим данные ячейки
# read_document.read_list_data(list_data=list_data, read_line=False)

# Для проверки на заглавную букву в строке
# search.checking_is_title()

# указываем текст или слова по отдельности для поиска
# search.serch_text(text=["NFC"])

# change.delete_space("AH")

# обьяденяем данные столбцов в один столбец
# change.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')

# удаляем текст что ищем с одной ячейки и добавляем в другую
# change.serch_add_dell(search="Да", add_colunm="S", dell_colunm="V")

# удаляем текст со всех ячейк
# change.delete_data_to_column(letters_colunm='U', text="rfvcx")

# добавляем текст во все ячейки
# change.add_data_to_colums(letters_colunm='U', text="rfvcx")


# указываем название нового файла для сохранения

# save.save(new_name="new")
