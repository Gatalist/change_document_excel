from document import OpenDocument, SaveDocument, ReadDocument, SearchText, ChangeDocument


# указываем название файла и название листа
document = OpenDocument(read_document='smart-chasy.xlsx', sheet_name='Worksheet')
print(document)

# указываем название колонки
read_document = ReadDocument(document=document, column_name='Особенности-14311')

# получаем список строк с данными
list_data = read_document.list_data_row()

search = SearchText(list_data=list_data, pause=True)
change = ChangeDocument(document=document, list_data=list_data, pause=True)
save = SaveDocument(document=document)


# Выводим данные ячейки
# read_document.read_list_data(list_data=list_data, read_line=True)


# добавляем текст в начало ячейки
# change.add_data_start(cell_past='AE', text="В режиме ожидания: ")


# удаяем весь текст с одной ячееки и добавляем в другую ячейку
# change.move_to_other_cell(cell_move='AG', cell_past='AI')


# Для проверки на заглавную букву в строке
# search.checking_is_title()


# указываем текст или слова по отдельности для поиска
# search.serch_text(text=["NFC"])


# обьяденяем данные столбцов в один столбец
# change.join_columns_text(save_column='AH', join_columns=['AE', 'AF', 'AG'], join_separator=' x ', end_text='см')


# удаляем текст что ищем с одной ячейки и добавляем в другую
# change.serch_move_past(search="Функции: время, звонки, будильник, GPS трекер, Anti-Lost, шагомер, сигнал SOS, мониторинг передвижения, Geo – зоны", cell_move="AC", cell_past="AE")


# удаляем текст со всех ячейк в столбце
# change.delete_data_to_column(cell_move='AI', text="При интенсивном использовании: До 24 ч")


# добавляем текст во все ячейки в столбце
change.add_data_to_colums(cell_past='AE', text="test")


# указываем название нового файла для сохранения
save.save(new_name="new")
